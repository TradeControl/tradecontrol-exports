import sys, argparse, base64, io, zipfile
import pyodbc
from lxml import etree as ET

from odfdo import Document, Style
from odfdo.table import Table, Row, Cell, Column
from odfdo.element import Element
from style_factory import apply_styles_bytes

def add_stylesheet(doc):
    # Column header: 8pt bold
    bold_style = Style(family='table-cell', name='ColumnHeader')
    bold_style.set_properties(area='text', **{
        'fo:font-size': '8pt',
        'fo:font-weight': 'bold',
        'style:font-weight-complex': 'bold'
    })
    bold_style.set_properties(area='table-cell', **{
        'style:vertical-align': 'middle',
        'style:wrap-option': 'no-wrap',
        'style:shrink-to-fit': 'false'
    })
    doc.insert_style(bold_style, automatic=True)

    # ce1: default borders (thin left, thick right) for unformatted cells
    ce1 = Style(family='table-cell', name='ce1')
    ce1.set_properties(area='table-cell', **{
        'fo:border-left': '0.5pt solid #000000',
        'fo:border-right': '1.5pt solid #000000'
    })
    doc.insert_style(ce1, automatic=True)

    return doc

def run_test(doc: Document) -> Document:
    """
    Column B primary test + extend across columns:
    - B4: header 'Column B Test'
    - B5: 1234.5678 as CASH0_CELL
    - B6: -8765.4321 as CASH0_CELL
    - B7: 4321.00 as CASH0_CELL (positive total)
    - B8: -4321.00 as CASH0_CELL (negative total)
    - B9: =B8 as CASH0_CELL (formula should render negative with red/parentheses)
    - B10: =B7 as CASH0_CELL (formula should render positive)
    Column borders:
    - Column B default cell style = ce1 (thin-left, thick-right)
    """
    tbl = Table(name="Sheet1")

    # Create A..H columns; set B default borders
    for col_idx in range(1, 9):  # A..H
        col = Column()
        if col_idx == 2:  # B
            col.set_attribute('table:default-cell-style-name', 'ce1')
        tbl.append(col)

    # Rows 1..3: blanks across A..H
    for _ in range(3):
        r = Row()
        for _c in range(8):
            r.append(Element.from_tag("table:table-cell"))
        tbl.append(r)

    # Row 4: header in B4
    r4 = Row()
    r4.append(Element.from_tag("table:table-cell"))  # A4 blank
    hdr = Cell(text="Column B Test")
    hdr.set_attribute("table:style-name", "ColumnHeader")
    r4.append(hdr)  # B4
    for _ in range(8 - 2):  # C..H blanks
        r4.append(Element.from_tag("table:table-cell"))
    tbl.append(r4)

    # Row 5: B5 = 1234.5678 (neutral style)
    r5 = Row()
    r5.append(Element.from_tag("table:table-cell"))  # A5
    b5 = Cell()
    b5.set_attribute("office:value-type", "float")
    b5.set_attribute("office:value", str(float(1234.5678)))
    b5.set_attribute("table:style-name", "CASH0_CELL")
    r5.append(b5)
    for _ in range(8 - 2):
        r5.append(Element.from_tag("table:table-cell"))
    tbl.append(r5)

    # Row 6: B6 = -8765.4321 (neutral style; style:map handles negative)
    r6 = Row()
    r6.append(Element.from_tag("table:table-cell"))  # A6
    b6 = Cell()
    b6.set_attribute("office:value-type", "float")
    b6.set_attribute("office:value", str(float(-8765.4321)))
    b6.set_attribute("table:style-name", "CASH0_CELL")
    r6.append(b6)
    for _ in range(8 - 2):
        r6.append(Element.from_tag("table:table-cell"))
    tbl.append(r6)

    # Row 7: B7 = 4321.00 (neutral style)
    r7 = Row()
    r7.append(Element.from_tag("table:table-cell"))  # A7
    b7 = Cell()
    b7.set_attribute("office:value-type", "float")
    b7.set_attribute("office:value", str(float(4321.00)))
    b7.set_attribute("table:style-name", "CASH0_CELL")
    r7.append(b7)
    for _ in range(8 - 2):
        r7.append(Element.from_tag("table:table-cell"))
    tbl.append(r7)

    # Row 8: B8 = -4321.00 (neutral style)
    r8 = Row()
    r8.append(Element.from_tag("table:table-cell"))  # A8
    b8 = Cell()
    b8.set_attribute("office:value-type", "float")
    b8.set_attribute("office:value", str(float(-4321.00)))
    b8.set_attribute("table:style-name", "CASH0_CELL")
    r8.append(b8)
    for _ in range(8 - 2):
        r8.append(Element.from_tag("table:table-cell"))
    tbl.append(r8)

    # Row 9: B9 = B8 (formula; cached value fixed in post-process)
    r9 = Row()
    r9.append(Element.from_tag("table:table-cell"))  # A9
    b9 = Cell()
    b9.set_attribute("table:formula", "of:=B8")
    b9.set_attribute("office:value-type", "float")
    b9.set_attribute("office:value", "0")
    b9.set_attribute("table:style-name", "CASH0_CELL")
    r9.append(b9)
    for _ in range(8 - 2):
        r9.append(Element.from_tag("table:table-cell"))
    tbl.append(r9)

    # Row 10: B10 = B7 (formula; cached value fixed in post-process)
    r10 = Row()
    r10.append(Element.from_tag("table:table-cell"))  # A10
    b10 = Cell()
    b10.set_attribute("table:formula", "of:=B7")
    b10.set_attribute("office:value-type", "float")
    b10.set_attribute("office:value", "0")
    b10.set_attribute("table:style-name", "CASH0_CELL")
    r10.append(b10)
    for _ in range(8 - 2):
        r10.append(Element.from_tag("table:table-cell"))
    tbl.append(r10)

    doc.body.append(tbl)
    return doc

def _post_process_styles_add_borders(content: bytes) -> bytes:
    # Unzip content.xml and cache entries
    src_zip = zipfile.ZipFile(io.BytesIO(content), 'r')
    namelist = src_zip.namelist()
    if 'content.xml' not in namelist:
        src_zip.close()
        return content
    content_xml = src_zip.read('content.xml')

    # Parse XML
    root = ET.fromstring(content_xml)
    ns = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    }

    auto_styles = root.find('office:automatic-styles', ns)
    if auto_styles is None:
        src_zip.close()
        return content

    # Column B borders base: ce1 (thin-left, thick-right)
    ce1_style = auto_styles.find("style:style[@style:name='ce1'][@style:family='table-cell']", ns)
    ce1_props = ce1_style.find('style:table-cell-properties', ns) if ce1_style is not None else None
    if ce1_props is None:
        src_zip.close()
        return content

    # Helpers
    def ensure_bordered_clone(src_style_name: str) -> str:
        """Clone a cell style and add ce1 L/R borders; returns new style name."""
        if not src_style_name:
            return src_style_name
        new_name = f"{src_style_name}_BORDERED"
        existing = auto_styles.find(f"style:style[@style:name='{new_name}'][@style:family='table-cell']", ns)
        if existing is not None:
            return new_name
        src = auto_styles.find(f"style:style[@style:name='{src_style_name}'][@style:family='table-cell']", ns)
        if src is None:
            return src_style_name
        new = ET.fromstring(ET.tostring(src))
        new.set(f"{{{ns['style']}}}name", new_name)
        props = new.find('style:table-cell-properties', ns)
        if props is None:
            props = ET.Element(f"{{{ns['style']}}}table-cell-properties")
            new.append(props)
        for attr in ('border-left', 'border-right'):
            val = ce1_props.get(f"{{{ns['fo']}}}{attr}")
            if val:
                props.set(f"{{{ns['fo']}}}{attr}", val)
        auto_styles.append(new)
        return new_name

    # Locate table and rows
    table = root.find('.//table:table', ns)
    if table is not None:
        rows = table.findall('table:table-row', ns)

        # A. Column B accumulation: convert B4..B10 cells' styles to bordered variants
        if len(rows) >= 4:
            b4 = rows[3].find("./table:table-cell[2]", ns)
            if b4 is not None:
                sname = b4.get(f"{{{ns['table']}}}style-name") or ""
                bordered = ensure_bordered_clone(sname)
                if bordered and bordered != sname:
                    b4.set(f"{{{ns['table']}}}style-name", bordered)

        for ridx in (4, 5, 6, 7, 8, 9):
            if len(rows) > ridx:
                cell = rows[ridx].find("./table:table-cell[2]", ns)
                if cell is not None:
                    sname = cell.get(f"{{{ns['table']}}}style-name") or ""
                    bordered = ensure_bordered_clone(sname)
                    if bordered and bordered != sname:
                        cell.set(f"{{{ns['table']}}}style-name", bordered)

        # B. Fix cached values for simple direct-reference formulas so style:map applies immediately
        def col_letters_to_index(letters: str) -> int:
            idx = 0
            for ch in letters:
                idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
            return idx

        def find_cell_by_index(row_elem: ET._Element, col_index_1based: int):
            count = 0
            for child in row_elem:
                tag = child.tag
                if not (tag.endswith('table-cell') or tag.endswith('covered-table-cell')):
                    continue
                repeat = int(child.get(f"{{{ns['table']}}}number-columns-repeated", "1"))
                next_count = count + repeat
                if col_index_1based <= next_count:
                    return child
                count = next_count
            return None

        import re
        direct_ref_patterns = [
            re.compile(r"^of:=\.?\$?([A-Za-z]+)\$?(\d+)$"),
            re.compile(r"^of:=\[\.\$?([A-Za-z]+)\$?(\d+)\]$"),
        ]

        for row in rows:
            cells = row.findall('table:table-cell', ns)
            for cell in cells:
                formula = cell.get(f"{{{ns['table']}}}formula")
                if not formula:
                    continue
                m = None
                for pat in direct_ref_patterns:
                    m = pat.match(formula)
                    if m:
                        break
                if not m:
                    continue
                col_letters, row_num = m.group(1), int(m.group(2))
                ref_col = col_letters_to_index(col_letters)
                if 1 <= row_num <= len(rows):
                    ref_row = rows[row_num - 1]
                    ref_cell = find_cell_by_index(ref_row, ref_col)
                    if ref_cell is not None:
                        vtype = ref_cell.get(f"{{{ns['office']}}}value-type")
                        if vtype:
                            cell.set(f"{{{ns['office']}}}value-type", vtype)
                        val = ref_cell.get(f"{{{ns['office']}}}value")
                        if val is not None:
                            cell.set(f"{{{ns['office']}}}value", val)
                        cur = ref_cell.get(f"{{{ns['office']}}}currency")
                        if cur:
                            cell.set(f"{{{ns['office']}}}currency", cur)

        # C. Enforce applied POS/NEG bordered style based on cached numeric value for column B cells
        # This guarantees parentheses/red for negatives, even on formula cells.
        def enforce_pos_neg_bordered(cell_elem: ET._Element):
            sname = cell_elem.get(f"{{{ns['table']}}}style-name") or ""
            if not sname:
                return
            val_str = cell_elem.get(f"{{{ns['office']}}}value")
            vtype = cell_elem.get(f"{{{ns['office']}}}value-type")
            if vtype != "float" or val_str is None:
                return
            try:
                num = float(val_str)
            except ValueError:
                return
            # Ensure we have bordered variants of POS/NEG
            pos_bordered = ensure_bordered_clone("CASH0_POS_CELL")
            neg_bordered = ensure_bordered_clone("CASH0_NEG_CELL")
            # Switch the applied style to explicit POS/NEG bordered to force format
            cell_elem.set(f"{{{ns['table']}}}style-name", neg_bordered if num < 0 else pos_bordered)

        # Apply to B5..B10 (column index 2)
        for ridx in (4, 5, 6, 7, 8, 9):  # rows are 0-based in code; B5..B10
            if len(rows) > ridx:
                b_cell = rows[ridx].find("./table:table-cell[2]", ns)
                if b_cell is not None:
                    enforce_pos_neg_bordered(b_cell)

        # D. Repeater row to last, covering A..H (8 columns)
        rep = ET.Element(f"{{{ns['table']}}}table-row")
        rep.set(f"{{{ns['table']}}}number-rows-repeated", "1048568")
        rep_cell = ET.Element(f"{{{ns['table']}}}table-cell")
        rep_cell.set(f"{{{ns['table']}}}number-columns-repeated", "8")
        rep.append(rep_cell)
        table.append(rep)

    # Serialize back
    new_content_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    # Cache entries
    entries = {}
    for name in namelist:
        if name == 'content.xml':
            continue
        data = src_zip.read(name)
        entries[name] = data
    src_zip.close()

    # Write new zip (mimetype first stored, then content and others deflated)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as zf:
        if 'mimetype' in entries:
            zi = zipfile.ZipInfo('mimetype')
            zi.compress_type = zipfile.ZIP_STORED
            zf.writestr(zi, entries['mimetype'])
            del entries['mimetype']
        zf.writestr('content.xml', new_content_xml)
        for name, data in entries.items():
            zf.writestr(name, data)

    return out.getvalue()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--conn", required=True)
    parser.add_argument("--query", required=False)
    parser.add_argument("--filename", required=True)
    args = parser.parse_args()
    conn = pyodbc.connect(args.conn)

    doc = Document("spreadsheet")
    doc = add_stylesheet(doc)
    doc = run_test(doc)

    # Save to bytes
    raw = io.BytesIO()
    doc.save(raw)
    content = raw.getvalue()

    # Materialize semantic styles (NUM/PCT/CASH)
    try:
        content = apply_styles_bytes(content, locale=("en", "GB"))
    except Exception:
        pass

    # Post-process: add borders to column default, clone CASH styles with borders, fix B5/B6 style names, append repeated row
    content = _post_process_styles_add_borders(content)

    conn.close()
    encoded = base64.b64encode(content).decode("ascii")
    print(f"{args.filename}|{encoded}")

if __name__ == "__main__":
    main()