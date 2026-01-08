# pip install odfdo
import argparse
import base64
import io
import sys
from pathlib import Path

import pyodbc
from odfdo import Document
from odfdo.table import Table, Row, Cell, Column
from style_factory import apply_styles_bytes

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
    Per-cell formatting using semantic style names (injected post-save by Style Factory).
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

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--conn", required=True, help="pyodbc connection string")
    parser.add_argument("--query", required=True, help="SQL returning TemplateCode in first column")
    parser.add_argument("--filename", required=True, help="logical filename to print alongside base64")
    parser.add_argument("--lang", default="en")
    parser.add_argument("--country", default="GB")
    parser.add_argument("--keep-defaults", action="store_true", help="Do not strip default sheet artifacts like 'Feuille1'")
    args = parser.parse_args()

    conn = pyodbc.connect(args.conn)
    try:
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

    content = apply_styles_bytes(
        content,
        locale=(args.lang, args.country),
        strip_defaults=not args.keep_defaults
    )

    encoded = base64.b64encode(content).decode("ascii")
    print(f"{args.filename}|{encoded}")