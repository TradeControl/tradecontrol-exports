# pip install odfdo lxml
import argparse
import base64
import io
import sys
import zipfile
from pathlib import Path

from odfdo import Document
from odfdo.table import Table, Row, Cell, Column
from style_factory import apply_styles_bytes
from lxml import etree as ET

# Copy of add_number_cell: identical behavior to the exporter
def add_number_cell(row: Row, value: float = None, style: str | None = None, formula: str | None = None, display_text: str | None = None):
    def resolve_cash_style(base: str, is_negative: bool) -> str:
        u = (base or "").strip().upper()
        if not u.startswith("CASH") or not u.endswith("_CELL"):
            return base or "CASH0_CELL"
        if u.endswith("_POS_CELL") or u.endswith("_NEG_CELL"):
            return u
        return u.replace("_CELL", "_NEG_CELL" if is_negative else "_POS_CELL")

    cell = Cell()
    is_negative = False
    if formula:
        f = formula.strip().upper()
        if f.startswith("-") or "*-1" in f or "=-" in f:
            is_negative = True
        cell.set_attribute("table:formula", f"of:={formula}")
        cell.set_attribute("office:value-type", "float")
        cell.set_attribute("office:value", "0")
    else:
        num = float(value or 0.0)
        is_negative = num < 0
        cell.set_attribute("office:value-type", "float")
        cell.set_attribute("office:value", str(num))
    stamped_style = resolve_cash_style(style or "CASH0_CELL", is_negative)
    cell.set_attribute("table:style-name", stamped_style)
    row.append(cell)

def _build_add_number_cell_check_doc() -> bytes:
    doc = Document("spreadsheet")
    tbl = Table(name="AddNumberCellCheck")
    for _ in range(6):
        tbl.append(Column())
    r1 = Row()
    for _ in range(4):
        r1.append(Cell())
    add_number_cell(r1, value=1234.5678, style="CASH0_CELL")   # E1
    add_number_cell(r1, formula="E1*-1", style="CASH0_CELL")    # F1
    tbl.append(r1)
    doc.body.append(tbl)
    buf = io.BytesIO()
    try:
        doc.save(buf)
        return buf.getvalue()
    except TypeError:
        return doc.save()

def _ns():
    return {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    }

def _extract_content_xml(ods_bytes: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(ods_bytes), "r") as z:
        return z.read("content.xml")

def test_cash0_cell_injection_creates_base_and_maps():
    original = _build_add_number_cell_check_doc()
    patched = apply_styles_bytes(original, locale=("en", "GB"), strip_defaults=True)

    # Parse content.xml from the patched ODS
    content_xml = _extract_content_xml(patched)
    root = ET.fromstring(content_xml)

    ns = _ns()
    auto = root.find("office:automatic-styles", ns)
    assert auto is not None, "automatic-styles must exist"

    # Verify number data styles exist
    cash_pos_ds = auto.find("number:number-style[@style:name='CASH0_POS_DS']", ns)
    assert cash_pos_ds is not None, "CASH0_POS_DS must be generated"
    num_node = cash_pos_ds.find("number:number", ns)
    assert num_node is not None, "CASH0_POS_DS must contain a <number:number>"
    assert num_node.get("{%s}grouping" % ns["number"]) == "true", "Grouping should be enabled"
    assert num_node.get("{%s}decimal-places" % ns["number"]) == "0", "Decimals must match CASH0"

    cash_neg_ds = auto.find("number:number-style[@style:name='CASH0_NEG_DS']", ns)
    assert cash_neg_ds is not None, "CASH0_NEG_DS must be generated"
    neg_num = cash_neg_ds.find("number:number", ns)
    assert neg_num is not None, "CASH0_NEG_DS must contain a <number:number>"
    assert neg_num.get("{%s}display-factor" % ns["number"]) == "-1", "Negative display factor required"
    neg_texts = [e.text or "" for e in cash_neg_ds if e.tag == "{%s}text" % ns["number"]]
    assert "(" in neg_texts and ")" in neg_texts, "Negative style should wrap with parentheses"

    # Verify POS/NEG cell styles exist and reference data styles
    pos_cell = auto.find("style:style[@style:name='CASH0_POS_CELL']", ns)
    assert pos_cell is not None, "CASH0_POS_CELL must be created"
    assert pos_cell.get("{%s}data-style-name" % ns["style"]) == "CASH0_POS_DS"

    neg_cell = auto.find("style:style[@style:name='CASH0_NEG_CELL']", ns)
    assert neg_cell is not None, "CASH0_NEG_CELL must be created"
    assert neg_cell.get("{%s}data-style-name" % ns["style"]) == "CASH0_NEG_DS"

    # Verify base CASH0_CELL exists with a default data-style and conditional maps
    base_cell = auto.find("style:style[@style:name='CASH0_CELL']", ns)
    assert base_cell is not None, "CASH0_CELL base style must be created"
    assert base_cell.get("{%s}data-style-name" % ns["style"]) == "CASH0_POS_DS", "Base should default to POS data style"

    maps = base_cell.findall("style:map", ns)
    assert any(m.get("{%s}condition" % ns["style"]) == "value() < 0" and m.get("{%s}apply-style-name" % ns["style"]) == "CASH0_NEG_CELL" for m in maps), "NEG map required"
    assert any(m.get("{%s}condition" % ns["style"]) == "value() >= 0" and m.get("{%s}apply-style-name" % ns["style"]) == "CASH0_POS_CELL" for m in maps), "POS map required"

def test_cash0_cell_applied_to_numeric_cells():
    original = _build_add_number_cell_check_doc()
    patched = apply_styles_bytes(original, locale=("en", "GB"), strip_defaults=True)

    content_xml = _extract_content_xml(patched)
    root = ET.fromstring(content_xml)
    ns = _ns()

    table = root.find(".//table:table[@table:name='AddNumberCellCheck']", ns)
    assert table is not None, "AddNumberCellCheck table must exist"
    row = table.find("table:table-row", ns)
    assert row is not None, "First row must exist"
    cells = row.findall("table:table-cell", ns)
    assert len(cells) >= 6, "Row should have at least 6 cells (A..F)"

    e1 = cells[4]
    f1 = cells[5]
    # E1 is a positive numeric value (1234.5678), should be POS style
    assert e1.get("{%s}value-type" % ns["office"]) == "float"
    assert e1.get("{%s}style-name" % ns["table"]) == "CASH0_POS_CELL"

    # F1 is a negative formula (E1*-1), should be NEG style
    assert f1.get("{%s}value-type" % ns["office"]) == "float"
    assert f1.get("{%s}style-name" % ns["table"]) == "CASH0_NEG_CELL"

def _parse_locale_tuple(locale_str: str) -> tuple[str, str]:
    s = (locale_str or "").strip()
    if not s:
        return ("en", "GB")
    alias = {
        "france": "fr-FR",
        "germany": "de-DE",
        "spain": "es-ES",
        "united kingdom": "en-GB",
        "uk": "en-GB"
    }
    s = alias.get(s.lower(), s)
    s = s.replace("_", "-")
    parts = s.split("-", 1)
    if len(parts) == 1:
        lang = parts[0].lower()
        defaults = {
            "en": "GB",
            "fr": "FR",
            "de": "DE",
            "es": "ES",
        }
        country = defaults.get(lang, lang.upper())
        return (lang, country)
    lang = parts[0].lower()
    country = parts[1].upper()
    return (lang, country)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--filename", default="AddNumberCellCheck.ods", help="logical filename to print alongside base64")
    parser.add_argument("--locale", default="en-GB", help="locale like en-GB, fr-FR")
    parser.add_argument("--keep-defaults", action="store_true", help="Do not strip default sheet artifacts like 'Feuille1'")
    # Backward-compat: accept and ignore legacy args
    parser.add_argument("--conn", help="ignored (legacy)", default=None)
    parser.add_argument("--query", help="ignored (legacy)", default=None)
    args = parser.parse_args()

    # Run the two focused tests (no DB required)
    test_cash0_cell_injection_creates_base_and_maps()
    test_cash0_cell_applied_to_numeric_cells()

    # Produce a minimal ODS to visually inspect, then apply styles
    content = _build_add_number_cell_check_doc()
    lang, country = _parse_locale_tuple(args.locale)
    content = apply_styles_bytes(
        content,
        locale=(lang, country),
        strip_defaults=not args.keep_defaults
    )

    encoded = base64.b64encode(content).decode("ascii")
    print(f"{args.filename}|{encoded}")