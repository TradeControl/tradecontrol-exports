# pip install pyodbc odfpy
import sys, argparse, base64, io
import pyodbc
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell, TableHeaderRows
from odf.text import P
from odf.style import Style, TableCellProperties, TextProperties
from odf import config

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--conn", required=True)
    parser.add_argument("--query", required=True)
    parser.add_argument("--filename", required=True)
    args = parser.parse_args()

    conn = pyodbc.connect(args.conn)
    cur = conn.cursor()
    cur.execute(args.query)

    doc = OpenDocumentSpreadsheet()

    # Header style: bold + light gray background + bottom border
    header_style = Style(name="HeaderCell", family="table-cell")
    header_style.addElement(TextProperties(fontweight="bold"))
    header_style.addElement(TableCellProperties(
        backgroundcolor="#D9D9D9",
        borderbottom="0.75pt solid #808080"
    ))

    data_style = Style(name="DataCell", family="table-cell")
    data_style.addElement(TableCellProperties())

    doc.automaticstyles.addElement(header_style)
    doc.automaticstyles.addElement(data_style)

    table_name = "Export"
    table = Table(name=table_name)

    # Header row in a header group (repeats on print)
    header_row = TableRow()
    for col in cur.description:
        cell = TableCell(stylename=header_style)
        cell.addElement(P(text=str(col[0])))
        header_row.addElement(cell)
    header_group = TableHeaderRows()
    header_group.addElement(header_row)
    table.addElement(header_group)

    # Data rows
    for row in cur.fetchall():
        tr = TableRow()
        for val in row:
            cell = TableCell(stylename=data_style)
            cell.addElement(P(text="" if val is None else str(val)))
            tr.addElement(cell)
        table.addElement(tr)

    doc.spreadsheet.addElement(table)

    # Freeze the first row in Calc via view settings (settings.xml)
    # Note: These are LibreOffice-specific config keys.
    view_settings = config.ConfigItemSet(name="ooo:view-settings")
    views = config.ConfigItemMapIndexed(name="Views")
    view = config.ConfigItemMapEntry()
    tables = config.ConfigItemMapNamed(name="Tables")

    table_entry = config.ConfigItemMapEntry(name=table_name)

    def add_item(name, typ, text):
        itm = config.ConfigItem(name=name, type=typ)
        itm.addText(text)
        table_entry.addElement(itm)

    add_item("HorizontalSplitMode", "short", "2")     # 2 = freeze
    add_item("HorizontalSplitPosition", "int", "2")   # freeze at row 2 -> keeps row 1 visible
    add_item("VerticalSplitMode", "short", "0")
    add_item("VerticalSplitPosition", "int", "0")
    add_item("ActiveSplitRange", "short", "2")

    tables.addElement(table_entry)
    view.addElement(tables)
    views.addElement(view)
    view_settings.addElement(views)
    doc.settings.addElement(view_settings)

    buf = io.BytesIO()
    doc.save(buf)
    conn.close()

    encoded = base64.b64encode(buf.getvalue()).decode("ascii")
    print(f"{args.filename}|{encoded}")

if __name__ == "__main__":
    main()