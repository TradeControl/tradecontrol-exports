# pip install pyodbc odfpy
import json, sys, io, base64
from datetime import datetime
import os
from pathlib import Path

# Bootstrap: add parent 'python' folder to sys.path so 'data' and 'i18n' imports resolve
_here = Path(__file__).resolve().parent        # .../python/exporters
_python_root = _here.parent                    # .../python
if str(_python_root) not in sys.path:
    sys.path.insert(0, str(_python_root))

from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableCellProperties, TextProperties
from odf import config
from data.factory import create_repo
from i18n.resources import ResourceManager

def build_styles(doc):
    header = Style(name="HeaderCell", family="table-cell")
    header.addElement(TextProperties(fontweight="bold"))
    header.addElement(TableCellProperties(backgroundcolor="#D9D9D9", borderbottom="0.75pt solid #808080"))
    total = Style(name="TotalCell", family="table-cell")
    total.addElement(TextProperties(fontweight="bold"))
    total.addElement(TableCellProperties(backgroundcolor="#EEEEEE", borderbottom="0.75pt solid #808080"))
    data = Style(name="DataCell", family="table-cell")
    data.addElement(TableCellProperties())
    doc.automaticstyles.addElement(header); doc.automaticstyles.addElement(total); doc.automaticstyles.addElement(data)
    return header, total, data

def freeze_first_rows(doc, table_name, freeze_row_index=2):
    view_settings = config.ConfigItemSet(name="ooo:view-settings")
    views = config.ConfigItemMapIndexed(name="Views")
    view = config.ConfigItemMapEntry()
    tables = config.ConfigItemMapNamed(name="Tables")
    table_entry = config.ConfigItemMapEntry(name=table_name)
    def add_item(name, typ, text):
        itm = config.ConfigItem(name=name, type=typ); itm.addText(text); table_entry.addElement(itm)
    add_item("HorizontalSplitMode", "short", "2")
    add_item("HorizontalSplitPosition", "int", str(freeze_row_index))
    add_item("VerticalSplitMode", "short", "0")
    add_item("VerticalSplitPosition", "int", "0")
    add_item("ActiveSplitRange", "short", "2")
    tables.addElement(table_entry); view.addElement(tables); views.addElement(view); view_settings.addElement(views)
    doc.settings.addElement(view_settings)

def generate_ods(payload: dict) -> tuple[str, bytes]:
    params = payload.get("Params") or payload.get("params") or {}
    conn = payload.get("SqlConnection") or payload.get("sqlConnection") or payload.get("connectionString")
    locale = params.get("locale") or "en-GB"
    res = ResourceManager(locale)
    repo = create_repo(conn, params)

    include_active = params.get("includeActivePeriods") == "true"
    include_orderbook = params.get("includeOrderBook") == "true"
    include_tax_accruals = params.get("includeTaxAccruals") == "true"
    include_vat_details = params.get("includeVatDetails") == "true"
    include_bank_balances = params.get("includeBankBalances") == "true"
    include_balance_sheet = params.get("includeBalanceSheet") == "true"

    active = repo.get_active_period() or {}
    years = repo.get_active_years()
    months = repo.get_months()
    company_name = repo.get_company_name()

    doc = OpenDocumentSpreadsheet()
    header_style, total_style, data_style = build_styles(doc)

    table_name = "Cash Flow"
    table = Table(name=table_name)

    # Header
    tr = TableRow(); cell = TableCell(stylename=header_style)
    cell.addElement(P(text=res.t("TextStatementTitle").format(active.get("MonthName",""), active.get("Description",""))))
    tr.addElement(cell); table.addElement(tr)
    tr = TableRow(); cell = TableCell(stylename=header_style); cell.addElement(P(text=company_name or "")); tr.addElement(cell); table.addElement(tr)
    tr = TableRow()
    cellA = TableCell(stylename=header_style); cellA.addElement(P(text=res.t("TextDate")))
    cellB = TableCell(stylename=header_style); cellB.addElement(P(text=datetime.now().strftime("%d %b %H:%M:%S")))
    tr.addElement(cellA); tr.addElement(cellB); table.addElement(tr)
    tr = TableRow()
    cellA = TableCell(stylename=header_style); cellA.addElement(P(text=res.t("TextCode")))
    cellB = TableCell(stylename=header_style); cellB.addElement(P(text=res.t("TextName")))
    tr.addElement(cellA); tr.addElement(cellB); table.addElement(tr)

    # Period grid header row
    period_header = TableRow()
    period_header.addElement(TableCell()); period_header.addElement(TableCell())
    for _y in years:
        for m in months:
            c = TableCell(stylename=header_style); c.addElement(P(text=str(m.get("MonthName","")))); period_header.addElement(c)
        ctot = TableCell(stylename=header_style); ctot.addElement(P(text=res.t("TextTotals"))); period_header.addElement(ctot)
    table.addElement(period_header)

    # Placeholders for remaining blocks (to be ported next)
    table.addElement(TableRow())
    if include_bank_balances:
        tr = TableRow(); c = TableCell(stylename=header_style); c.addElement(P(text=res.t("TextClosingBalances"))); tr.addElement(c); table.addElement(tr)
    if include_vat_details:
        tr = TableRow(); c = TableCell(stylename=header_style); c.addElement(P(text=res.t("TextVatDueTotals"))); tr.addElement(c); table.addElement(tr)
    if include_balance_sheet:
        tr = TableRow(); c = TableCell(stylename=header_style); c.addElement(P(text=res.t("TextBalanceSheet"))); tr.addElement(c); table.addElement(tr)

    doc.spreadsheet.addElement(table)
    freeze_first_rows(doc, table_name, freeze_row_index=2)

    buf = io.BytesIO(); doc.save(buf)
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return f"Cash_Flow_{ts}.ods", buf.getvalue()

if __name__ == "__main__":
    # Accept payload as path argument (preferred) or stdin fallback
    if len(sys.argv) >= 2:
        with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
    else:
        payload = json.loads(sys.stdin.read())
    filename, content = generate_ods(payload)
    print(f"{filename}|{base64.b64encode(content).decode('ascii')}")