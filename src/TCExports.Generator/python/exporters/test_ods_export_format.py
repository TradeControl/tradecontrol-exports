# pip install odfdo
import sys, argparse, base64, io, zipfile
import pyodbc
from lxml import etree as ET
from odfdo import Document
from odfdo.table import Table, Row, Cell, Column

OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
STYLE_NS  = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
TABLE_NS  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TEXT_NS   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
FO_NS     = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"

def _col_letter(index_1based: int) -> str:
    dividend = index_1based
    name = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        name = chr(65 + modulo) + name
        dividend = (dividend - modulo) // 26
    return name

def build_format_template_sheet(doc: Document, cur, sheet_name: str = "FormatTemplates") -> Table:
    """
    Per-cell formatting (styles injected post-save).
    Supports:
      - Num0/Num1/Num2  -> numbers (float)
      - Pct0/Pct1/Pct2  -> percentages
      - Cash0/Cash2     -> accounting numbers; negatives red + parentheses
    Columns for Cash expand to pairs: CashX+ and CashX-.
    Row 2:
      - CashX+: constant 1234.567
      - CashX-: formula = -(<CashX+ col>2)
    Row 3:
      - CashX+: formula = <CashX+ col>2 (echo)
      - CashX-: formula = -(<CashX+ col>3)
    """
    table = Table(sheet_name)
    templates = cur.fetchall()
    if not templates:
        r = Row(); r.append(Cell(text="(no templates)")); table.append(r)
        return table

    # Expand Cash into +/-
    headers = []
    for tpl in templates:
        code = (tpl[0] if not isinstance(tpl, dict) else tpl.get("TemplateCode","")).strip()
        ucode = code.upper()
        if ucode.startswith("CASH"):
            headers.append((f"{code}+", f"{ucode}_POS_CELL", ucode))
            headers.append((f"{code}-", f"{ucode}_NEG_CELL", ucode))
        else:
            headers.append((code, f"{ucode}_CELL", ucode))

    for _ in headers:
        table.append(Column())

    # Row 1
    hdr = Row()
    for title, _, _ in headers:
        hdr.append(Cell(text=title))
    table.append(hdr)

    # Row 2: write true numeric values, choose style by sign
    row2 = Row()
    for title, style_name, ucode in headers:
        c = Cell()
        if ucode.startswith("PCT"):
            c.set_attribute("office:value-type", "percentage")
            c.set_attribute("office:value", "0.23456")
        elif ucode.startswith("CASH"):
            val = "1234.567" if title.endswith("+") else "-1234.567"
            c.set_attribute("office:value-type", "float")
            c.set_attribute("office:value", val)
        else:
            c.set_attribute("office:value-type", "float")
            c.set_attribute("office:value", "1.2345")
        c.set_attribute("table:style-name", style_name)
        row2.append(c)
    table.append(row2)

    # Row 3: echo row 2 with correct sign (no ABS), style already applied per cell
    row3 = Row()
    for idx, (title, style_name, ucode) in enumerate(headers, start=1):
        letter = _col_letter(idx)
        f = Cell()
        if ucode.startswith("PCT"):
            f.set_attribute("office:value-type", "percentage")
        else:
            f.set_attribute("office:value-type", "float")
        f.set_attribute("office:value", "0")
        f.set_attribute("table:formula", f"of:={letter}2")
        f.set_attribute("table:style-name", style_name)
        row3.append(f)
    table.append(row3)

    return table

def _inject_styles_into_content(content_xml: bytes) -> bytes:
    """
    Inject styles into content.xml/office:automatic-styles for actually used styles:
      - numbers (Num*): number-style with dp, grouping
      - percentages (Pct*): percentage-style with dp and trailing %
      - cash (Cash*): explicit POS/NEG data styles and cell styles (no style:map)
        NEG is red with parentheses and no leading minus (via display-factor).
    """
    parser = ET.XMLParser(remove_blank_text=False)
    root = ET.fromstring(content_xml, parser=parser)

    auto = root.find(ET.QName(OFFICE_NS, "automatic-styles"))
    if auto is None:
        auto = ET.Element(ET.QName(OFFICE_NS, "automatic-styles"))
        root.insert(0, auto) if len(root) else root.append(auto)

    body = root.find(ET.QName(OFFICE_NS, "body"))
    if body is not None:
        ss = body.find(ET.QName(OFFICE_NS, "spreadsheet"))
        if ss is not None:
            for tbl in list(ss.findall(ET.QName(TABLE_NS, "table"))):
                if tbl.get(ET.QName(TABLE_NS, "name")) == "Feuille1":
                    ss.remove(tbl)

    used_styles: dict[str, tuple[str, int]] = {}
    for cell in root.findall(f".//{ET.QName(TABLE_NS, 'table-cell')}"):
        sname = cell.get(ET.QName(TABLE_NS, "style-name"))
        if not sname:
            continue
        uname = sname.strip().upper()
        if uname.startswith("NUM"):
            try: dp = int(uname[3])
            except Exception: dp = 0
            used_styles[uname] = ("number", dp)
        elif uname.startswith("PCT"):
            try: dp = int(uname[3])
            except Exception: dp = 0
            used_styles[uname] = ("percentage", dp)
        elif uname.startswith("CASH"):
            try: dp = int(uname[4])
            except Exception: dp = 0
            if uname.endswith("_POS_CELL"):
                used_styles[uname] = ("cash_pos", dp)
            elif uname.endswith("_NEG_CELL"):
                used_styles[uname] = ("cash_neg", dp)
        if cell.find(ET.QName(TEXT_NS, "p")) is None:
            ET.SubElement(cell, ET.QName(TEXT_NS, "p"))

    def find_style(tag_ns: str, tag_local: str, name_attr_ns: str, name: str):
        for el in auto.findall(ET.QName(tag_ns, tag_local)):
            if el.get(ET.QName(name_attr_ns, "name")) == name:
                return el
        return None

    def ensure_number_ds(name: str, dp: int):
        ds = find_style(NUMBER_NS, "number-style", STYLE_NS, name)
        if ds is None:
            ds = ET.SubElement(auto, ET.QName(NUMBER_NS, "number-style"))
            ds.set(ET.QName(STYLE_NS, "name"), name)
            pos = ET.SubElement(ds, ET.QName(NUMBER_NS, "number"))
            pos.set(ET.QName(NUMBER_NS, "decimal-places"), str(dp))
            pos.set(ET.QName(NUMBER_NS, "min-decimal-places"), str(dp))
            pos.set(ET.QName(NUMBER_NS, "min-integer-digits"), "1")
            pos.set(ET.QName(NUMBER_NS, "grouping"), "true")
        return ds

    def ensure_percent_ds(name: str, dp: int):
        ds = find_style(NUMBER_NS, "percentage-style", STYLE_NS, name)
        if ds is None:
            ds = ET.SubElement(auto, ET.QName(NUMBER_NS, "percentage-style"))
            ds.set(ET.QName(STYLE_NS, "name"), name)
            num = ET.SubElement(ds, ET.QName(NUMBER_NS, "number"))
            num.set(ET.QName(NUMBER_NS, "decimal-places"), str(dp))
            num.set(ET.QName(NUMBER_NS, "min-decimal-places"), str(dp))
            num.set(ET.QName(NUMBER_NS, "min-integer-digits"), "1")
            ET.SubElement(ds, ET.QName(NUMBER_NS, "text")).text = "%"
        return ds

    def ensure_cell_style(name: str, data_style_name: str, neg_red: bool = False):
        cs = find_style(STYLE_NS, "style", STYLE_NS, name)
        if cs is None:
            cs = ET.SubElement(auto, ET.QName(STYLE_NS, "style"))
            cs.set(ET.QName(STYLE_NS, "name"), name)
            cs.set(ET.QName(STYLE_NS, "family"), "table-cell")
            cs.set(ET.QName(STYLE_NS, "parent-style-name"), "Default")
            ET.SubElement(cs, ET.QName(STYLE_NS, "table-cell-properties"))
        cs.set(ET.QName(STYLE_NS, "data-style-name"), data_style_name)
        if neg_red:
            tp = cs.find(ET.QName(STYLE_NS, "text-properties"))
            if tp is None:
                tp = ET.SubElement(cs, ET.QName(STYLE_NS, "text-properties"))
            tp.set(ET.QName(FO_NS, "color"), "#FF0000")
        return cs

    def ensure_cash_neg_ds(name: str, dp: int):
        ds = find_style(NUMBER_NS, "number-style", STYLE_NS, name)
        if ds is None:
            ds = ET.SubElement(auto, ET.QName(NUMBER_NS, "number-style"))
            ds.set(ET.QName(STYLE_NS, "name"), name)
            ET.SubElement(ds, ET.QName(NUMBER_NS, "text")).text = "("
            neg = ET.SubElement(ds, ET.QName(NUMBER_NS, "number"))
            neg.set(ET.QName(NUMBER_NS, "decimal-places"), str(dp))
            neg.set(ET.QName(NUMBER_NS, "min-decimal-places"), str(dp))
            neg.set(ET.QName(NUMBER_NS, "min-integer-digits"), "1")
            neg.set(ET.QName(NUMBER_NS, "grouping"), "true")
            ET.SubElement(ds, ET.QName(NUMBER_NS, "text")).text = ")"
        return ds

    for cell_style_name, (kind, dp) in used_styles.items():
        if kind == "number":
            ds_name = cell_style_name.replace("_CELL", "_DS")
            ensure_number_ds(ds_name, dp)
            ensure_cell_style(cell_style_name, ds_name)
        elif kind == "percentage":
            ds_name = cell_style_name.replace("_CELL", "_DS")
            ensure_percent_ds(ds_name, dp)
            ensure_cell_style(cell_style_name, ds_name)
        elif kind == "cash_pos":
            ds_name = cell_style_name.replace("_POS_CELL", "_POS_DS")
            ensure_number_ds(ds_name, dp)
            ensure_cell_style(cell_style_name, ds_name, neg_red=False)
        elif kind == "cash_neg":
            ds_name = cell_style_name.replace("_NEG_CELL", "_NEG_DS")
            ensure_cash_neg_ds(ds_name, dp)
            ensure_cell_style(cell_style_name, ds_name, neg_red=True)

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")

def _finalize_ods_add_template_styles(ods_bytes: bytes) -> bytes:
    in_mem = io.BytesIO(ods_bytes)
    with zipfile.ZipFile(in_mem, "r") as zin:
        content_xml = zin.read("content.xml")
        new_content = _inject_styles_into_content(content_xml)
        out_mem = io.BytesIO()
        with zipfile.ZipFile(out_mem, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "content.xml":
                    zout.writestr("content.xml", new_content)
                else:
                    zout.writestr(item, zin.read(item.filename))
        return out_mem.getvalue()

def _apply_default_language_to_styles(ods_bytes: bytes, lang: str = "en", country: str = "GB") -> bytes:
    """
    Set default language in styles.xml (per odfdo recipe) so document inherits en-GB locale.
    """
    in_mem = io.BytesIO(ods_bytes)
    with zipfile.ZipFile(in_mem, "r") as zin:
        try:
            styles_xml = zin.read("styles.xml")
            s_root = ET.fromstring(styles_xml)
        except KeyError:
            s_root = ET.Element(ET.QName(OFFICE_NS, "document-styles"))
            ET.SubElement(s_root, ET.QName(OFFICE_NS, "styles"))
        office_styles = s_root.find(ET.QName(OFFICE_NS, "styles"))
        if office_styles is None:
            office_styles = ET.SubElement(s_root, ET.QName(OFFICE_NS, "styles"))

        def ensure_default_style(family: str):
            for ds in office_styles.findall(ET.QName(STYLE_NS, "default-style")):
                if ds.get(ET.QName(STYLE_NS, "family")) == family:
                    return ds
            ds = ET.SubElement(office_styles, ET.QName(STYLE_NS, "default-style"))
            ds.set(ET.QName(STYLE_NS, "family"), family)
            return ds

        def ensure_text_props(parent):
            tp = parent.find(ET.QName(STYLE_NS, "text-properties"))
            if tp is None:
                tp = ET.SubElement(parent, ET.QName(STYLE_NS, "text-properties"))
            tp.set(ET.QName(FO_NS, "language"), lang)
            tp.set(ET.QName(FO_NS, "country"), country)
            tp.set(ET.QName(STYLE_NS, "language-asian"), lang)
            tp.set(ET.QName(STYLE_NS, "country-asian"), country)
            tp.set(ET.QName(STYLE_NS, "language-complex"), lang)
            tp.set(ET.QName(STYLE_NS, "country-complex"), country)
            return tp

        for family in ("paragraph", "text", "table-cell"):
            ds = ensure_default_style(family)
            ensure_text_props(ds)

        new_styles = ET.tostring(s_root, xml_declaration=True, encoding="UTF-8")

        out_mem = io.BytesIO()
        with zipfile.ZipFile(out_mem, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "styles.xml":
                    zout.writestr("styles.xml", new_styles)
                else:
                    zout.writestr(item.filename, zin.read(item.filename))
            if "styles.xml" not in {i.filename for i in zin.infolist()}:
                zout.writestr("styles.xml", new_styles)
        return out_mem.getvalue()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--conn", required=True)
    parser.add_argument("--query", required=True)
    parser.add_argument("--filename", required=True)
    args = parser.parse_args()

    conn = pyodbc.connect(args.conn)
    cur = conn.cursor()
    cur.execute(args.query)

    doc = Document("spreadsheet")
    template_sheet = build_format_template_sheet(doc, cur, sheet_name="FormatTemplates")
    doc.body.append(template_sheet)

    buf = io.BytesIO()
    try:
        doc.save(buf)
        content = buf.getvalue()
    except TypeError:
        content = doc.save()
    finally:
        conn.close()

    content = _finalize_ods_add_template_styles(content)
    content = _apply_default_language_to_styles(content, lang="en", country="GB")
    encoded = base64.b64encode(content).decode("ascii")
    print(f"{args.filename}|{encoded}")