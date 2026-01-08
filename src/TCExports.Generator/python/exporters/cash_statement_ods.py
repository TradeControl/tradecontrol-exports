# pip install pyodbc odfpy
import json, sys, io, base64
from datetime import datetime
from pathlib import Path
from typing import Optional, Union

from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableCellProperties, TextProperties
from odf.number import NumberStyle, Number, PercentageStyle
from odf.namespaces import STYLENS, TABLENS, OFFICENS
from odf import config
from data.enums import CashType
from data.factory import create_repo
from i18n.resources import ResourceManager

def build_styles(doc):
    # Number style: zero decimal places
    int_style = NumberStyle(name="Int0")
    int_style.addElement(Number(decimalplaces=0, minintegerdigits=1))
    doc.automaticstyles.addElement(int_style)

    # Percentage style: 2 decimal places
    pct_style_num = PercentageStyle(name="Pct2")
    pct_style_num.addElement(Number(decimalplaces=2, minintegerdigits=1))
    doc.automaticstyles.addElement(pct_style_num)

    # Header style (text only)
    header = Style(name="HeaderCell", family="table-cell")
    header.addElement(TextProperties(fontweight="bold"))
    header.addElement(TableCellProperties(backgroundcolor="#D9D9D9", borderbottom="0.75pt solid #808080"))

    # Total style (bind number style for zero-dp display)
    total = Style(name="TotalCell", family="table-cell", datastylename="Int0")
    total.addElement(TextProperties(fontweight="bold"))
    total.addElement(TableCellProperties(backgroundcolor="#EEEEEE", borderbottom="0.75pt solid #808080"))

    # Data style (bind number style for zero-dp display)
    data = Style(name="DataCell", family="table-cell", datastylename="Int0")
    data.addElement(TableCellProperties())

    doc.automaticstyles.addElement(header)
    doc.automaticstyles.addElement(total)
    doc.automaticstyles.addElement(data)
    return header, total, data


def _col_letter(col_index_1based: int) -> str:
    # Excel-like column letters (A=1)
    dividend = col_index_1based
    name = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        name = chr(65 + modulo) + name
        dividend = (dividend - modulo) // 26
    return name

def add_text_cell(row, text: str, style=None):
    cell = TableCell(stylename=style) if style else TableCell()
    # Only add text when not empty; allows visual overflow from A to empty B in UI
    if text is not None and str(text) != "":
        cell.addElement(P(text=str(text)))
    row.addElement(cell)

def add_number_cell(row, value: float = None, style=None, formula: str = None, display_text: str = None):
    """
    Write a numeric or formula cell that Calc treats as numeric.
    - numeric: valuetype="float" and value
    - formula: formula="of:=..." and valuetype="float"
    Do not add visible text for numeric cells (prevents string casting).
    """
    cell = TableCell(stylename=style) if style else TableCell()

    if formula:
        cell.setAttribute("formula", f"of:={formula}")
        cell.setAttribute("valuetype", "float")
        if display_text is not None and str(display_text) != "":
            cell.addElement(P(text=str(display_text)))
    else:
        num = float(value or 0.0)
        cell.setAttribute("valuetype", "float")
        cell.setAttribute("value", str(num))

    row.addElement(cell)

def freeze_first_rows(doc, table_name, freeze_row_index=4):
    """
    Freeze rows 1-4 and columns A-C (months row stays visible; column D scrolls).
    Uses split positions equal to the number of frozen rows/columns.
    """
    view_settings = config.ConfigItemSet(name="ooo:view-settings")
    views = config.ConfigItemMapIndexed(name="Views")
    view = config.ConfigItemMapEntry()
    tables = config.ConfigItemMapNamed(name="Tables")
    table_entry = config.ConfigItemMapEntry(name=table_name)

    def add_item(name, typ, text):
        itm = config.ConfigItem(name=name, type=typ)
        itm.addText(text)
        table_entry.addElement(itm)

    # Freeze columns: first 3 columns (A..C)
    add_item("HorizontalSplitMode", "short", "2")          # 2 = split/freeze
    add_item("HorizontalSplitPosition", "int", "3")        # freeze columns A..C

    # Freeze rows: top 4 rows
    add_item("VerticalSplitMode", "short", "2")            # 2 = split/freeze
    add_item("VerticalSplitPosition", "int", freeze_row_index)          # freeze rows 1..4

    # Keep initial view at top-left; don't force the cursor into another pane
    add_item("ActiveSplitRange", "short", "2")             # keep focus on left/bottom pane

    tables.addElement(table_entry)
    view.addElement(tables)
    views.addElement(view)
    view_settings.addElement(views)
    doc.settings.addElement(view_settings)

def cols_for_year_block(month_count: int, year_index: int) -> tuple[int, int]:
    """
    Column indices within a row for the given year block, zero-based, counting cells created in order.
    We reserve columns A,B,C at the start of each row, so the first month is D.
    start: index of first month cell for the year block
    totals: index of the totals cell for the year block
    """
    start = 3 + year_index * (month_count + 1)
    totals = start + month_count
    return start, totals

def render_summary_after_categories(table: Table,
                                    repo,
                                    res: ResourceManager,
                                    years,
                                    months,
                                    categories: list,
                                    header_style,
                                    total_style,
                                    totals_row_by_category: dict[str, int] | None = None):
    if not categories or len(categories) < 2:
        return

    # Header
    hdr = TableRow()
    add_text_cell(hdr, res.t("TextSummary"), header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())  # C
    table.addElement(hdr)

    firstCol = 4
    lastCol = firstCol + (len(years) * (len(months) + 1)) - 1

    # Remember the first summary row index for period total calculation
    start_row_index = len(table.getElementsByType(TableRow)) + 1

    # Summary rows: mirror Excel by referencing each category's totals row in the same column
    for cat in categories:
        code = (cat.get("CategoryCode") or "").strip()
        name = cat.get("Category", "") or ""

        r = TableRow()
        add_text_cell(r, code)   # A
        add_text_cell(r, name)   # B
        r.addElement(TableCell())  # C reserved

        target_row = (totals_row_by_category or {}).get(code, -1)

        for col in range(firstCol, lastCol + 1):
            col_letter = _col_letter(col)
            if target_row > 0:
                add_number_cell(r, style=total_style, formula=f"{col_letter}{target_row}")
            else:
                add_number_cell(r, value=0.0, style=total_style)

        table.addElement(r)

    # Period Total row: SUM of the summary rows per column
    pr = TableRow()
    add_text_cell(pr, res.t("TextPeriodTotal"), header_style)
    add_text_cell(pr, "", header_style)
    pr.addElement(TableCell())
    end_row_index = len(table.getElementsByType(TableRow)) + 1

    for col in range(firstCol, lastCol + 1):
        col_letter = _col_letter(col)
        add_number_cell(pr, style=header_style,
                        formula=f"SUM([.{col_letter}{start_row_index}:.{col_letter}{end_row_index - 1}])")

    table.addElement(pr)

def render_categories_and_summary(table: Table,
                                  repo,
                                  res: ResourceManager,
                                  years,
                                  months,
                                  cash_type: Union[CashType, int],
                                  include_active: bool,
                                  include_orderbook: bool,
                                  include_tax_accruals: bool,
                                  header_style,
                                  total_style,
                                  data_style,
                                  totals_row_by_category: Optional[dict[str, int]] = None):
    table.addElement(TableRow())
    categories = repo.get_categories(cash_type)

    row_counter = len(table.getElementsByType(TableRow))  # current count

    for cat in categories:
        # Category name row
        cat_row = TableRow()
        add_text_cell(cat_row, cat.get("Category",""), header_style)
        add_text_cell(cat_row, "", header_style)
        cat_row.addElement(TableCell())
        table.addElement(cat_row)
        row_counter += 1

        codes = repo.get_cash_codes(cat.get("CategoryCode",""))

        for code in codes:
            r = TableRow()
            add_text_cell(r, code.get("CashCode",""))
            add_text_cell(r, code.get("CashDescription",""))
            r.addElement(TableCell())
            cur_row_index = row_counter + 1

            for y_idx, y in enumerate(years):
                year_num = int(y.get("YearNumber"))
                vals = repo.get_cash_code_values(code.get("CashCode",""),
                                                 year_num,
                                                 include_active, include_orderbook, include_tax_accruals)
                mm = { int(v.get("MonthNumber")): float(v.get("InvoiceValue", 0) or 0) for v in vals }

                # Month cells
                for m in months:
                    v = mm.get(int(m.get("MonthNumber")), 0.0)
                    add_number_cell(r, value=v, style=data_style)

                # Year total formula: SUM of that year's months
                start_col = 4 + (y_idx * (len(months) + 1))
                end_col = start_col + len(months) - 1
                start_letter = _col_letter(start_col)
                end_letter = _col_letter(end_col)
                add_number_cell(r, style=total_style,
                                formula=f"SUM([.{start_letter}{cur_row_index}:.{end_letter}{cur_row_index}])")

            table.addElement(r)
            row_counter += 1

        # Category totals row: SUM down each period column, apply polarity like Excel
        tot = TableRow()
        add_text_cell(tot, res.t("TextTotals"), total_style)
        add_text_cell(tot, "", total_style)
        # Column C marker for summary lookups
        cat_code = (cat.get("CategoryCode","") or "").strip()
        add_number_cell(tot, style=total_style, formula=f"\"{cat_code}\"", display_text="")
        cur_row_index = row_counter + 1

        # Determine polarity factor: 0 => multiply by -1, 1 or others => as-is
        cash_polarity = cat.get("CashPolarityCode")
        factor = -1 if cash_polarity == 0 or cash_polarity == "0" else 1

        total_cols = len(years) * (len(months) + 1)
        for i in range(total_cols):
            col = 4 + i
            col_letter = _col_letter(col)
            first_code_row = cur_row_index - len(codes)
            last_code_row = cur_row_index - 1
            base_sum = f"SUM([.{col_letter}{first_code_row}:.{col_letter}{last_code_row}])"
            if factor == -1:
                add_number_cell(tot, style=total_style, formula=f"{base_sum}*-1")
            else:
                add_number_cell(tot, style=total_style, formula=base_sum)

        table.addElement(tot)
        if totals_row_by_category is not None and cat_code:
            totals_row_by_category[cat_code] = cur_row_index

        row_counter += 1        

def render_summary_totals_block(table: Table, repo, res: ResourceManager,
                                cash_type: Union[CashType, int],
                                header_style,
                                totals_row_by_category: Optional[dict[str, int]] = None):
    """
    Render totals block header and codes, aligned with A,B,C then months.
    Also places a code marker in column C and optionally records the row index per total code.
    """
    totals = repo.get_categories_by_type(cash_type, "Total") if hasattr(repo, "get_categories_by_type") else []
    if not totals or len(totals) < 2:
        return
    table.addElement(TableRow())

    hdr = TableRow()
    heading = f"{totals[0].get('CashType','')} {res.t('TextTotals')}".strip()
    add_text_cell(hdr, heading, header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())
    table.addElement(hdr)

    for t in totals:
        r = TableRow()
        code = (t.get("CategoryCode", "") or "").strip()
        desc = t.get("Category", "") or ""
        add_text_cell(r, code)                 # A
        add_text_cell(r, desc)                 # B
        # C: marker equals code for lookup
        add_number_cell(r, formula=f"\"{code}\"", display_text="", style=None)
        # record row index if requested
        if totals_row_by_category is not None and code:
            row_index = len(table.getElementsByType(TableRow)) + 1
            table.addElement(r)
            totals_row_by_category[code] = row_index
        else:
            table.addElement(r)

def render_totals_formula(table: Table, repo, res: ResourceManager, years=None, months=None,
                          header_style=None, total_style=None, data_style=None,
                          totals_row_by_category: Optional[dict[str, int]] = None):
    """
    Totals: A=CategoryCode, B=Category, C=CategoryCode marker; then per-period formulas.
    Month columns use data_style; annual totals use total_style.
    """
    table.addElement(TableRow())
    hdr = TableRow()
    add_text_cell(hdr, res.t("TextTotals"), header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())
    table.addElement(hdr)

    if not hasattr(repo, "get_category_totals") or not hasattr(repo, "get_category_total_codes"):
        return

    totals = repo.get_category_totals() or []
    if not totals:
        return

    first_col = 4
    month_count = len(months or [])
    year_count = len(years or [])
    last_col = first_col + (year_count * (month_count + 1)) - 1
    if last_col < first_col:
        return

    for t in totals:
        code = (t.get("CategoryCode") or "").strip()
        desc = (t.get("Category") or "").strip()
        if not code:
            continue

        r = TableRow()
        add_text_cell(r, code)         # A: CategoryCode
        add_text_cell(r, desc)         # B: Category
        add_number_cell(r, style=total_style, formula=f"\"{code}\"", display_text="")  # C marker

        # Register this total row so expressions like [Gross Profit]=001 resolve (e.g., D97)
        row_index = len(table.getElementsByType(TableRow)) + 1
        if totals_row_by_category is not None:
            totals_row_by_category[code] = row_index

        # Build per-column formulas; can reference both categories and other totals
        sum_codes_rows = repo.get_category_total_codes(code) or []
        src_codes = [row.get("SourceCategoryCode", row.get("CategoryCode", "")) for row in sum_codes_rows if row]

        for col in range(first_col, last_col + 1):
            col_letter = _col_letter(col)
            zero_based = col - first_col
            is_year_total_col = (month_count > 0) and ((zero_based + 1) % (month_count + 1) == 0)
            style_for_cell = total_style if is_year_total_col else data_style

            terms = []
            for sc in src_codes:
                sc = (sc or "").strip()
                if totals_row_by_category and sc in totals_row_by_category:
                    terms.append(f"{col_letter}{totals_row_by_category[sc]}")

            if terms:
                add_number_cell(r, style=style_for_cell, formula="+".join(terms))
            else:
                add_number_cell(r, value=0.0, style=style_for_cell)

        table.addElement(r)

# Cache for format-driven styles
_format_style_cache: dict[str, Style] = {}

def get_style_for_format(doc, fmt: str) -> Optional[Style]:
    """
    Create or retrieve an ODF style for the given Excel-like format string.
    Supported:
      - '%'                      -> PercentageStyle with 0 dp
      - '0', '0.0', '0.00', ...  -> NumberStyle with matching dp
      - '0%', '0.0%', '0.00%'    -> PercentageStyle with matching dp
    Returns a table-cell Style bound to the number style, or None if not supported.
    """
    key = (fmt or "").strip()
    if not key:
        return None
    if key in _format_style_cache:
        return _format_style_cache[key]

    from odf.number import NumberStyle, Number, PercentageStyle
    from odf.style import Style, TableCellProperties

    try:
        if key.endswith("%"):
            dp = 0
            base = key[:-1]
            if "." in base:
                dp = len(base.split(".")[1])
            num = PercentageStyle(name=f"FmtPct_{len(_format_style_cache)}")
            num.addElement(Number(decimalplaces=dp, minintegerdigits=1))
            doc.automaticstyles.addElement(num)
            cell = Style(name=f"CellPct_{len(_format_style_cache)}", family="table-cell", datastylename=num.getAttribute("name"))
            cell.addElement(TableCellProperties())
            doc.automaticstyles.addElement(cell)
            _format_style_cache[key] = cell
            return cell

        if key == "%":
            num = PercentageStyle(name=f"FmtPct_{len(_format_style_cache)}")
            num.addElement(Number(decimalplaces=0, minintegerdigits=1))
            doc.automaticstyles.addElement(num)
            cell = Style(name=f"CellPct_{len(_format_style_cache)}", family="table-cell", datastylename=num.getAttribute("name"))
            cell.addElement(TableCellProperties())
            doc.automaticstyles.addElement(cell)
            _format_style_cache[key] = cell
            return cell

        if key.startswith("0"):
            dp = 0
            if "." in key:
                dp = len(key.split(".")[1])
            num = NumberStyle(name=f"FmtNum_{len(_format_style_cache)}")
            num.addElement(Number(decimalplaces=dp, minintegerdigits=1))
            doc.automaticstyles.addElement(num)
            cell = Style(name=f"CellNum_{len(_format_style_cache)}", family="table-cell", datastylename=num.getAttribute("name"))
            cell.addElement(TableCellProperties())
            doc.automaticstyles.addElement(cell)
            _format_style_cache[key] = cell
            return cell

        return None
    except Exception:
        return None

def render_expressions(table: Table,
                       repo,
                       res: ResourceManager,
                       years=None,
                       months=None,
                       header_style=None,
                       total_style=None,
                       data_style=None,
                       totals_row_by_category: Optional[dict[str, int]] = None,
                       doc=None):
    """
    Expressions:
    - A: Category (name)
    - B: blank
    - C: CategoryCode marker (if resolvable)
    - D..: formulas referencing totals rows in the same column.
    Applies per-row cell style derived from the expression's Format string.
    """
    table.addElement(TableRow())
    hdr = TableRow()
    add_text_cell(hdr, res.t("TextAnalysis"), header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())
    table.addElement(hdr)

    if not hasattr(repo, "get_category_expressions"):
        return

    exprs = repo.get_category_expressions() or []
    if not exprs:
        return

    # Build description->code map for totals (e.g. "Gross Profit"->"001")
    totals_by_name = {}
    if hasattr(repo, "get_category_totals"):
        for t in repo.get_category_totals() or []:
            nm = (t.get("Category") or "").strip()
            cd = (t.get("CategoryCode") or "").strip()
            if nm and cd:
                totals_by_name[nm] = cd

    first_col = 4
    month_count = len(months or [])
    year_count = len(years or [])
    last_col = first_col + (year_count * (month_count + 1)) - 1
    if last_col < first_col:
        return

    import re
    def normalize_for_calc(formula_text: str) -> str:
        f = re.sub(r'\bif\s*\(', 'IF(', formula_text, flags=re.IGNORECASE)
        return f.replace(',', ';')

    for expr in exprs:
        r = TableRow()
        category_name = (expr.get("Category") or "").strip()
        template = (expr.get("Expression") or "").strip()
        fmt = (expr.get("Format") or "").strip()

        # A, B
        add_text_cell(r, category_name)
        add_text_cell(r, "")

        # Column C: expression CategoryCode marker
        expr_code = (expr.get("CategoryCode") or "").strip()
        if not expr_code and hasattr(repo, "get_category_code_from_name"):
            expr_code = (repo.get_category_code_from_name(category_name) or "").strip()
        if not expr_code:
            expr_code = totals_by_name.get(category_name, "")
        if expr_code:
            add_number_cell(r, style=total_style, formula=f"\"{expr_code}\"", display_text="")
        else:
            r.addElement(TableCell())

        # Extract [Name] tokens in order
        tokens: list[str] = []
        i = 0
        while i < len(template):
            lb = template.find("[", i)
            if lb < 0: break
            rb = template.find("]", lb + 1)
            if rb < 0: break
            token = template[lb + 1:rb].strip()
            if token and token not in tokens:
                tokens.append(token)
            i = rb + 1

        # Map names -> codes (cash categories via repo lookup, totals via totals_by_name)
        name_to_code: dict[str, str] = {}
        for name in tokens:
            code = ""
            if hasattr(repo, "get_category_code_from_name"):
                code = (repo.get_category_code_from_name(name) or "").strip()
            if not code:
                code = totals_by_name.get(name, "")
            if not code:
                code = name
            name_to_code[name] = code

        # Replace [Name] -> [Code] and normalize for Calc
        normalized = template
        for name in tokens:
            normalized = normalized.replace(f"[{name}]", f"[{name_to_code[name]}]")
        normalized = normalize_for_calc(normalized)

        errors = []

        # Per-expression cell style derived from format string
        # If fmt = "0%", "0.00%", etc., create Percentage style; "0", "0.00", etc., create Number style.
        expr_cell_style = get_style_for_format(doc, fmt) if doc is not None else None
        if fmt and expr_cell_style is None:
            # Unsupported format: we still render values with default numeric style
            errors.append(f"Unsupported format: {fmt}")
        cell_style = expr_cell_style or data_style

        for col in range(first_col, last_col + 1):
            col_letter = _col_letter(col)
            formula = normalized

            # Replace each [Code] with same-column A1 reference
            for code in {c for c in name_to_code.values()}:
                row_index = (totals_row_by_category or {}).get(code, -1)
                if row_index <= 0:
                    errors.append(f"Reference not found: [{code}]")
                    formula = formula.replace(f"[{code}]", "0")
                else:
                    formula = formula.replace(f"[{code}]", f"{col_letter}{row_index}")

            try:
                add_number_cell(r, style=cell_style, formula=formula)
            except Exception as ex:
                errors.append(f"Calc formula error: {str(ex)}")
                add_number_cell(r, value=0.0, style=cell_style)

        table.addElement(r)

        # Report status (optional)
        if hasattr(repo, "set_category_expression_status"):
            is_error = len(errors) > 0
            msg = "; ".join(sorted(set(errors))) if is_error else None
            try:
                repo.set_category_expression_status(expr_code or category_name, is_error, msg)
            except Exception:
                pass

def render_vat_recurrence_totals(table, repo, res, years, months, include_active_periods, include_tax_accruals, header_style, data_style, total_style):
    """
    VAT Due by recurrence periods. Align headers A,B,C; add spacer after section.
    """
    hdr = TableRow()
    vat_type = (repo.get_vat_recurrence_type() or "").upper()
    add_text_cell(hdr, f"{res.t('TextVatDueTitle')} {vat_type}".upper(), header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())
    table.addElement(hdr)

    labels = [
        res.t("TextVatHomeSales"),
        res.t("TextVatHomePurchases"),
        res.t("TextVatExportSales"),
        res.t("TextVatExportPurchases"),
        res.t("TextVatHomeSalesVat"),
        res.t("TextVatHomePurchasesVat"),
        res.t("TextVatExportSalesVat"),
        res.t("TextVatExportPurchasesVat"),
        res.t("TextVatAdjustment"),
        res.t("TextVatDue")
    ]

    recurrence = repo.get_vat_recurrence()
    by_year = {}
    for p in recurrence:
        y = int(p.get("YearNumber"))
        by_year.setdefault(y, []).append(p)

    accruals_by_year = {}
    if include_tax_accruals:
        for a in repo.get_vat_recurrence_accruals():
            y = int(a.get("YearNumber"))
            accruals_by_year.setdefault(y, []).append(a)

    for li, label in enumerate(labels):
        r = TableRow()
        add_text_cell(r, label.upper(), header_style if li in (0, 9) else data_style)
        add_text_cell(r, "", data_style)
        r.addElement(TableCell())

        for y in years:
            ynum = int(y.get("YearNumber"))
            periods = by_year.get(ynum, [])

            period_vals = []
            for p in periods:
                v = 0.0
                if include_active_periods or (p.get("StartOn") <= datetime.utcnow()):
                    if li == 0: v = float(p.get("HomeSales", 0) or 0)
                    elif li == 1: v = float(p.get("HomePurchases", 0) or 0)
                    elif li == 2: v = float(p.get("ExportSales", 0) or 0)
                    elif li == 3: v = float(p.get("ExportPurchases", 0) or 0)
                    elif li == 4: v = float(p.get("HomeSalesVat", 0) or 0)
                    elif li == 5: v = float(p.get("HomePurchasesVat", 0) or 0)
                    elif li == 6: v = float(p.get("ExportSalesVat", 0) or 0)
                    elif li == 7: v = float(p.get("ExportPurchasesVat", 0) or 0)
                    elif li == 8: v = float(p.get("VatAdjustment", 0) or 0)
                    elif li == 9: v = float(p.get("VatDue", 0) or 0)
                period_vals.append(v)

            if include_tax_accruals and li in (4, 5, 6, 7, 9):
                accs = accruals_by_year.get(ynum, [])
                for idx, a in enumerate(accs):
                    if idx >= len(period_vals):
                        break
                    add_val = 0.0
                    if li == 4 and a.get("HomeSalesVat") is not None: add_val = float(a.get("HomeSalesVat"))
                    elif li == 5 and a.get("HomePurchasesVat") is not None: add_val = float(a.get("HomePurchasesVat"))
                    elif li == 6 and a.get("ExportSalesVat") is not None: add_val = float(a.get("ExportSalesVat"))
                    elif li == 7 and a.get("ExportPurchasesVat") is not None: add_val = float(a.get("ExportPurchasesVat"))
                    elif li == 9 and a.get("VatDue") is not None: add_val = float(a.get("VatDue"))
                    period_vals[idx] += add_val

            for v in period_vals:
                add_number_cell(r, v, data_style)
            add_number_cell(r, sum(period_vals), total_style)

        table.addElement(r)

    # Spacer after VAT recurrence section
    table.addElement(TableRow())

def render_vat_period_totals(table, repo, res, years, months, include_active_periods, include_tax_accruals, header_style, data_style, total_style):
    """
    VAT monthly totals; align headers A,B,C; add spacer after section.
    """
    hdr = TableRow()
    add_text_cell(hdr, f"{res.t('TextVatDueTitle')} {res.t('TextTotals')}".upper(), header_style)
    add_text_cell(hdr, "", header_style)
    hdr.addElement(TableCell())
    table.addElement(hdr)

    labels = [
        res.t("TextVatHomeSales"),
        res.t("TextVatHomePurchases"),
        res.t("TextVatExportSales"),
        res.t("TextVatExportPurchases"),
        res.t("TextVatHomeSalesVat"),
        res.t("TextVatHomePurchasesVat"),
        res.t("TextVatExportSalesVat"),
        res.t("TextVatExportPurchasesVat"),
        res.t("TextVatDue")
    ]

    monthly = repo.get_vat_period_totals()
    by_year = {}
    for p in monthly:
        y = int(p.get("YearNumber"))
        by_year.setdefault(y, []).append(p)

    accruals_by_year = {}
    if include_tax_accruals:
        for a in repo.get_vat_period_accruals():
            y = int(a.get("YearNumber"))
            accruals_by_year.setdefault(y, []).append(a)

    for li, label in enumerate(labels):
        r = TableRow()
        add_text_cell(r, label.upper(), header_style if li == len(labels) - 1 else data_style)
        add_text_cell(r, "", data_style)
        r.addElement(TableCell())

        for y in years:
            ynum = int(y.get("YearNumber"))
            periods = by_year.get(ynum, [])
            month_vals = [0.0] * len(months)

            for p in periods:
                if include_active_periods or (p.get("StartOn") <= datetime.utcnow()):
                    mdt = p.get("StartOn")
                    mnum = mdt.month if hasattr(mdt, "month") else None
                    if mnum is None:
                        continue
                    idx = next((i for i, m in enumerate(months) if int(m.get("MonthNumber")) == mnum), None)
                    if idx is None:
                        continue
                    v = 0.0
                    if li == 0: v = float(p.get("HomeSales", 0) or 0)
                    elif li == 1: v = float(p.get("HomePurchases", 0) or 0)
                    elif li == 2: v = float(p.get("ExportSales", 0) or 0)
                    elif li == 3: v = float(p.get("ExportPurchases", 0) or 0)
                    elif li == 4: v = float(p.get("HomeSalesVat", 0) or 0)
                    elif li == 5: v = float(p.get("HomePurchasesVat", 0) or 0)
                    elif li == 6: v = float(p.get("ExportSalesVat", 0) or 0)
                    elif li == 7: v = float(p.get("ExportPurchasesVat", 0) or 0)
                    elif li == 8: v = float(p.get("VatDue", 0) or 0)
                    month_vals[idx] += v

            if include_tax_accruals and li in (4, 5, 6, 7, 8):
                accs = accruals_by_year.get(ynum, [])
                for idx, a in enumerate(accs):
                    if idx >= len(month_vals):
                        break
                    add_val = 0.0
                    if li == 4 and a.get("HomeSalesVat") is not None: add_val = float(a.get("HomeSalesVat"))
                    elif li == 5 and a.get("HomePurchasesVat") is not None: add_val = float(a.get("HomePurchasesVat"))
                    elif li == 6 and a.get("ExportSalesVat") is not None: add_val = float(a.get("ExportSalesVat"))
                    elif li == 7 and a.get("ExportPurchasesVat") is not None: add_val = float(a.get("ExportPurchasesVat"))
                    elif li == 8 and a.get("VatDue") is not None: add_val = float(a.get("VatDue"))
                    month_vals[idx] += add_val

            for v in month_vals:
                add_number_cell(r, v, data_style)
            add_number_cell(r, sum(month_vals), total_style)

        table.addElement(r)

    # Spacer after VAT period totals
    table.addElement(TableRow())

def render_bank_balances(table, repo, res, years, months, header_style, data_style, total_style):
    """
    Bank balances section header alignment (A,B,C), values unchanged.
    """
    table.addElement(TableRow())
    hr = TableRow()
    add_text_cell(hr, res.t("TextClosingBalances").upper(), header_style)
    add_text_cell(hr, "", header_style)
    hr.addElement(TableCell())
    table.addElement(hr)

    accounts = repo.get_bank_accounts()
    if not accounts:
        r = TableRow()
        add_text_cell(r, "(no bank accounts)", header_style)
        add_text_cell(r, "", header_style)
        r.addElement(TableCell())
        table.addElement(r)
        return

    cols_per_year = len(months) + 1
    total_cols = len(years) * cols_per_year
    company_totals = [0.0] * total_cols

    for acct in accounts:
        r = TableRow()
        add_text_cell(r, acct.get("AccountCode", ""), data_style)
        add_text_cell(r, acct.get("AccountName", ""), data_style)
        r.addElement(TableCell())

        balances = repo.get_bank_balances(acct.get("AccountCode", ""))
        bal_map = {(int(b["YearNumber"]), int(b["MonthNumber"])): float(b["Balance"] or 0) for b in balances}

        for y_idx, y in enumerate(years):
            year_num = int(y.get("YearNumber"))
            last_val = None
            for m_idx, m in enumerate(months):
                v = bal_map.get((year_num, int(m.get("MonthNumber"))), 0.0)
                add_number_cell(r, v, data_style)
                company_totals[y_idx * cols_per_year + m_idx] += v
                last_val = v
            carry = bal_map.get((year_num, 12), last_val if last_val is not None else 0.0)
            add_number_cell(r, carry, total_style)
            company_totals[y_idx * cols_per_year + len(months)] += carry

        table.addElement(r)

    tr = TableRow()
    add_text_cell(tr, res.t("TextCompanyBalance").upper(), header_style)
    add_text_cell(tr, "", header_style)
    tr.addElement(TableCell())
    for val in company_totals:
        add_number_cell(tr, val, total_style)
    table.addElement(tr)

def render_balance_sheet(table, repo, res, years, months, header_style, data_style, total_style):
    hr = TableRow(); add_text_cell(hr, res.t("TextBalanceSheet").upper(), header_style); table.addElement(hr)
    entries = repo.get_balance_sheet()
    if not entries: return

    order = []
    groups = {}
    for e in entries:
        key = f"{e['AssetCode']}\u001F{e['AssetName']}"
        if key not in groups:
            groups[key] = {}
            order.append((e['AssetCode'], e['AssetName'], key))
        groups[key][(int(e['YearNumber']), int(e['MonthNumber']))] = float(e['Balance'] or 0)

    for code, name, key in order:
        r = TableRow()
        add_text_cell(r, code, data_style)      # A
        add_text_cell(r, name, data_style)      # B
        r.addElement(TableCell())               # C reserved so months start at D

        # Write months and year total per year (numeric cells; zero-dp display via column styles)
        cur_row_index = len(table.getElementsByType(TableRow)) + 1

        for y_idx, y in enumerate(years):
            year_num = int(y.get("YearNumber"))
            year_start_col = 4 + (y_idx * (len(months) + 1))  # D is 4
            last_non_empty_col_letter = None

            # months
            for m_idx, m in enumerate(months):
                mm = int(m.get("MonthNumber"))
                v = groups[key].get((year_num, mm))
                add_number_cell(r, value=(v or 0.0), style=data_style)
                if v is not None:
                    last_non_empty_col_letter = _col_letter(year_start_col + m_idx)

            # year total: reference last non-empty month cell if any, else 0
            if last_non_empty_col_letter:
                add_number_cell(r, style=total_style,
                                formula=f"{last_non_empty_col_letter}{cur_row_index}",
                                display_text="")
            else:
                add_number_cell(r, value=0.0, style=total_style)

        table.addElement(r)

    cap = TableRow()
    add_text_cell(cap, res.t("TextCapital").upper(), header_style)
    add_text_cell(cap, "", header_style)
    cap.addElement(TableCell())  # C reserved so totals begin at D

    # Capital per column: SUM of asset rows in this section
    first_asset_row = len(table.getElementsByType(TableRow)) - len(order) + 1
    last_row_index = len(table.getElementsByType(TableRow)) + 1
    for col_offset in range(len(years) * (len(months) + 1)):
        col = 4 + col_offset
        letter = _col_letter(col)
        add_number_cell(cap, style=total_style,
                        formula=f"SUM([.{letter}{first_asset_row}:.{letter}{last_row_index - 1}])")
    table.addElement(cap)

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

    # Track totals rows for categories (CategoryCode -> row index)
    totals_row_by_category: dict[str, int] = {}

    # Zero-dp number style and a table-cell style bound to it
    from odf.number import NumberStyle, Number
    from odf.style import Style, TableCellProperties, TableColumnProperties
    int0 = NumberStyle(name="Int0")
    int0.addElement(Number(decimalplaces=0, minintegerdigits=1))
    doc.automaticstyles.addElement(int0)

    zero_dp_cell = Style(name="ZeroDpCell", family="table-cell", datastylename="Int0")
    zero_dp_cell.addElement(TableCellProperties())
    doc.automaticstyles.addElement(zero_dp_cell)

    table_name = "Cash Flow"
    table = Table(name=table_name)

    # Define columns: A (Code), B (Name), C hidden
    from odf.table import TableColumn
    table.addElement(TableColumn())                           # A
    table.addElement(TableColumn())                           # B
    table.addElement(TableColumn(visibility="collapse"))      # C hidden

    # Period columns (D -> ...): apply zero-dp display via column default cell style
    cols_per_year = (len(months) + 1) if years else 0
    period_cols = len(years) * cols_per_year
    for _ in range(period_cols):
        table.addElement(TableColumn(defaultcellstylename=zero_dp_cell))

    # Row 1: Title
    tr = TableRow()
    add_text_cell(tr, res.t("TextStatementTitle").format(active.get("MonthName",""), active.get("Description","")), header_style)
    add_text_cell(tr, "", header_style)
    tr.addElement(TableCell())
    table.addElement(tr)

    # Row 2: Company
    tr = TableRow()
    add_text_cell(tr, company_name or "", header_style)
    add_text_cell(tr, "", header_style)
    tr.addElement(TableCell())
    table.addElement(tr)

    # Row 3: Year labels
    year_row = TableRow()
    add_text_cell(year_row, res.t("TextDate"), header_style)
    add_text_cell(year_row, datetime.now().strftime("%d %b %H:%M:%S"), header_style)
    year_row.addElement(TableCell())  # C
    for y in years:
        desc = y.get("Description") or str(y.get("YearNumber"))
        status = y.get("CashStatus") or ""
        add_text_cell(year_row, f"{desc} ({status})" if status else desc, header_style)
        for _ in range(len(months) - 1):
            year_row.addElement(TableCell())
        add_text_cell(year_row, desc, header_style)
    table.addElement(year_row)

    # Row 4: Column headers
    hdr = TableRow()
    add_text_cell(hdr, res.t("TextCode"), header_style)
    add_text_cell(hdr, res.t("TextName"), header_style)
    hdr.addElement(TableCell())
    for _y in years:
        for m in months:
            add_text_cell(hdr, str(m.get("MonthName","")), header_style)
        add_text_cell(hdr, res.t("TextTotals"), header_style)
    table.addElement(hdr)

    # Sections (with summaries after categories, matching Excel Workbook behavior)
    # Trade categories + summary
    render_categories_and_summary(table, repo, res, years, months, CashType.Trade,
                                  include_active, include_orderbook, False,
                                  header_style, total_style, data_style,
                                  totals_row_by_category)
    render_summary_after_categories(table, repo, res, years, months,
                                    repo.get_categories(CashType.Trade),
                                    header_style, total_style,
                                    totals_row_by_category)

    # Money categories + summary
    render_categories_and_summary(table, repo, res, years, months, CashType.Money,
                                  False, False, False,
                                  header_style, total_style, data_style,
                                  totals_row_by_category)
    render_summary_after_categories(table, repo, res, years, months,
                                    repo.get_categories(CashType.Money),
                                    header_style, total_style,
                                    totals_row_by_category)

    # Totals block for Trade
    render_summary_totals_block(table, repo, res, CashType.Trade, header_style, totals_row_by_category)

    # Totals-of-totals formulas (requires totals rows tracked)
    render_totals_formula(table, repo, res, years, months,
                          header_style, total_style, data_style,
                          totals_row_by_category)

    # Tax categories + summary
    render_categories_and_summary(table, repo, res, years, months, CashType.Tax,
                                  include_active, False, include_tax_accruals,
                                  header_style, total_style, data_style,
                                  totals_row_by_category)
    render_summary_after_categories(table, repo, res, years, months,
                                    repo.get_categories(CashType.Tax),
                                    header_style, total_style,
                                    totals_row_by_category)

    # Totals block for Tax
    render_summary_totals_block(table, repo, res, CashType.Tax, header_style, totals_row_by_category)

    # Expressions
    render_expressions(table, repo, res, years, months,
                           header_style, total_style, data_style,
                           totals_row_by_category, doc=doc)

    # Spacer before optional sections
    # table.addElement(TableRow())

    if include_bank_balances:
        render_bank_balances(table, repo, res, years, months, header_style, data_style, total_style)
        table.addElement(TableRow())

    if include_vat_details:
        render_vat_recurrence_totals(table, repo, res, years, months, include_active, include_tax_accruals, header_style, data_style, total_style)
        render_vat_period_totals(table, repo, res, years, months, include_active, include_tax_accruals, header_style, data_style, total_style)

    if include_balance_sheet:
        render_balance_sheet(table, repo, res, years, months, header_style, data_style, total_style)

    doc.spreadsheet.addElement(table)
    freeze_first_rows(doc, table_name, freeze_row_index=4)

    buf = io.BytesIO()
    doc.save(buf)
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return f"Cash_Flow_{ts}.ods", buf.getvalue()

if __name__ == "__main__":
    if len(sys.argv) >= 2:
        with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
    else:
        payload = json.loads(sys.stdin.read())
    filename, content = generate_ods(payload)
    print(f"{filename}|{base64.b64encode(content).decode('ascii')}")