# pip install pyodbc odfdo
import json, sys, io, base64
from datetime import datetime
from pathlib import Path
from typing import Optional, Union

from odfdo import Document
from odfdo.table import Table, Row, Cell, Column

from data.enums import CashType
from data.factory import create_repo
from i18n.resources import ResourceManager
from style_factory import apply_styles_bytes

# ********** TEMP **************************************
def _append_add_number_cell_check(doc: Document) -> None:
    """
    Verify formatting via add_number_cell:
    - E1: 1234.5678
    - F1: -E1
    Both with base accounting style CASH2_CELL (maps to POS/NEG).
    """
    tbl = Table(name="AddNumberCellCheck")
    # Ensure columns up to F
    for _ in range(6):
        tbl.append(Column())

    r1 = Row()
    # A-D empty to reach E
    for _ in range(4):
        r1.append(Cell())

    # E1 via add_number_cell (value)
    add_number_cell(r1, value=1234.5678, style="CASH0_CELL")
    # F1 via add_number_cell (formula)
    add_number_cell(r1, formula="E1*-1", style="CASH0_CELL")

    tbl.append(r1)
    doc.body.append(tbl)
# ********** TEMP **************************************


def _col_letter(index_1based: int) -> str:
    dividend = index_1based
    name = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        name = chr(65 + modulo) + name
        dividend = (dividend - modulo) // 26
    return name

def add_text_cell(row: Row, text: str, style: Optional[str] = None):
    # Only add text when not empty; allows visual overflow in UI
    if text is not None and str(text) != "":
        row.append(Cell(text=str(text)))
    else:
        row.append(Cell())

def add_number_cell(row: Row, value: float = None, style: Optional[str] = None, formula: str = None, display_text: str = None):
    """
    Write a numeric or formula cell that Calc treats as numeric.
    - numeric: office:value-type="float" and office:value
    - formula: table:formula="of:=..." and office:value-type="float" with office:value="0"
    Do not add visible text for numeric cells (prevents string casting).
    Pragmatic override: when style is a neutral CASHx_CELL, stamp POS/NEG style directly.
    """
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
        # Heuristic: detect simple negative formulas
        f = formula.strip().upper()
        # negatives like "-A1", "A1*-1", "SUM(...)*-1", "(-A1)", etc.
        if f.startswith("-"):
            is_negative = True
        elif "*-1" in f or "=-" in f:
            is_negative = True
        cell.set_attribute("table:formula", f"of:={formula}")
        cell.set_attribute("office:value-type", "float")
        cell.set_attribute("office:value", "0")
    else:
        num = float(value or 0.0)
        is_negative = num < 0
        cell.set_attribute("office:value-type", "float")
        cell.set_attribute("office:value", str(num))

    # Apply resolved style
    stamped_style = resolve_cash_style(style or "CASH0_CELL", is_negative)
    cell.set_attribute("table:style-name", stamped_style)
    row.append(cell)

def _to_semantic_cell_style(template_code: Optional[str]) -> str:
    """
    Map a Template Code from the database (e.g., 'Cash0','Num2','Pct1')
    to a canonical semantic cell style name used by the Style Factory.
    Defaults to CASH0_CELL if missing or unrecognized.
    """
    if not template_code:
        return "CASH0_CELL"
    u = template_code.strip().upper()
    # accept already-suffixed names
    if u.endswith("_CELL") or u.endswith("_POS_CELL") or u.endswith("_NEG_CELL"):
        return u
    if u.startswith("CASH") or u.startswith("NUM") or u.startswith("PCT"):
        return f"{u}_CELL"
    return "CASH0_CELL"

def freeze_first_rows(doc: Document, table_name: str, freeze_row_index: int = 4):
    """
    No-op in this migration step.
    TODO: Reintroduce via settings.xml manipulation in a follow-up.
    """
    return

def cols_for_year_block(month_count: int, year_index: int) -> tuple[int, int]:
    start = 3 + year_index * (month_count + 1)
    totals = start + month_count
    return start, totals

class SheetBuilder:
    """
    Tracks rows appended to compute 1-based row indices for formulas.
    """
    def __init__(self, name: str):
        self.table = Table(name=name)
        self._row_index = 0

    def append_row(self, row: Row):
        self.table.append(row)
        self._row_index += 1

    def current_row_index(self) -> int:
        return self._row_index

def render_summary_after_categories(sb: SheetBuilder,
                                    repo,
                                    res: ResourceManager,
                                    years,
                                    months,
                                    categories: list,
                                    totals_row_by_category: dict[str, int] | None = None):
    if not categories or len(categories) < 2:
        return

    # Header
    hdr = Row()
    add_text_cell(hdr, res.t("TextSummary"))
    add_text_cell(hdr, "")
    hdr.append(Cell())  # C
    sb.append_row(hdr)

    firstCol = 4
    lastCol = firstCol + (len(years) * (len(months) + 1)) - 1

    # First summary row index (for period total)
    start_row_index = sb.current_row_index() + 1

    # Summary rows
    for cat in categories:
        code = (cat.get("CategoryCode") or "").strip()
        name = cat.get("Category", "") or ""

        r = Row()
        add_text_cell(r, code)   # A
        add_text_cell(r, name)   # B
        r.append(Cell())  # C reserved

        target_row = (totals_row_by_category or {}).get(code, -1)

        for col in range(firstCol, lastCol + 1):
            col_letter = _col_letter(col)
            if target_row > 0:
                add_number_cell(r, formula=f"{col_letter}{target_row}")
            else:
                add_number_cell(r, value=0.0)

        sb.append_row(r)

    # Period Total row: SUM of the summary rows per column
    pr = Row()
    add_text_cell(pr, res.t("TextPeriodTotal"))
    add_text_cell(pr, "")
    pr.append(Cell())
    end_row_index = sb.current_row_index()

    for col in range(firstCol, lastCol + 1):
        col_letter = _col_letter(col)
        add_number_cell(pr, formula=f"SUM([.{col_letter}{start_row_index}:.{col_letter}{end_row_index}])")

    sb.append_row(pr)

def render_categories_and_summary(sb: SheetBuilder,
                                  repo,
                                  res: ResourceManager,
                                  years,
                                  months,
                                  cash_type: Union[CashType, int],
                                  include_active: bool,
                                  include_orderbook: bool,
                                  include_tax_accruals: bool,
                                  totals_row_by_category: Optional[dict[str, int]] = None):
    sb.append_row(Row())
    categories = repo.get_categories(cash_type)

    for cat in categories:
        # Category name row
        cat_row = Row()
        add_text_cell(cat_row, cat.get("Category",""))
        add_text_cell(cat_row, "")
        cat_row.append(Cell())
        sb.append_row(cat_row)

        codes = repo.get_cash_codes(cat.get("CategoryCode",""))

        for code in codes:
            r = Row()
            add_text_cell(r, code.get("CashCode",""))
            add_text_cell(r, code.get("CashDescription",""))
            r.append(Cell())
            # Correct current row index for formulas (avoid drift)
            cur_row_index = sb.current_row_index() + 1

            for y_idx, y in enumerate(years):
                year_num = int(y.get("YearNumber"))
                vals = repo.get_cash_code_values(code.get("CashCode",""),
                                                 year_num,
                                                 include_active, include_orderbook, include_tax_accruals)
                mm = { int(v.get("MonthNumber")): float(v.get("InvoiceValue", 0) or 0) for v in vals }

                # Month cells
                for m in months:
                    v = mm.get(int(m.get("MonthNumber")), 0.0)
                    add_number_cell(r, value=v)

                # Year total formula: SUM of that year's months
                start_col = 4 + (y_idx * (len(months) + 1))
                end_col = start_col + len(months) - 1
                start_letter = _col_letter(start_col)
                end_letter = _col_letter(end_col)
                add_number_cell(r, formula=f"SUM([.{start_letter}{cur_row_index}:.{end_letter}{cur_row_index}])")

            sb.append_row(r)

        # Category totals row: SUM down each period column, apply polarity like Excel
        tot = Row()
        add_text_cell(tot, res.t("TextTotals"))
        add_text_cell(tot, "")
        # Column C marker (not relied upon programmatically)
        cat_code = (cat.get("CategoryCode","") or "").strip()
        add_number_cell(tot, formula=f"\"{cat_code}\"", display_text="")
        # Correct index for totals row
        cur_row_index = sb.current_row_index() + 1

        # Determine polarity factor: 0 => multiply by -1, 1 or others => as-is
        cash_polarity = cat.get("CashPolarityCode")
        factor = -1 if cash_polarity == 0 or cash_polarity == "0" else 1

        total_cols = len(years) * (len(months) + 1)
        for i in range(total_cols):
            col = 4 + i
            col_letter = _col_letter(col)
            # Sum the block of cash code rows just appended
            # We can compute the number of codes per category from 'codes'
            first_code_row = cur_row_index - len(codes)
            last_code_row = cur_row_index - 1
            base_sum = f"SUM([.{col_letter}{first_code_row}:.{col_letter}{last_code_row}])"
            if factor == -1:
                add_number_cell(tot, formula=f"{base_sum}*-1")
            else:
                add_number_cell(tot, formula=base_sum)

        sb.append_row(tot)
        sb.append_row(Row())
        if totals_row_by_category is not None and cat_code:
            totals_row_by_category[cat_code] = cur_row_index

def render_summary_totals_block(sb: SheetBuilder, repo, res: ResourceManager,
                                cash_type: Union[CashType, int],
                                totals_row_by_category: Optional[dict[str, int]] = None):
    totals = repo.get_categories_by_type(cash_type, "Total") if hasattr(repo, "get_categories_by_type") else []
    if not totals or len(totals) < 2:
        return
    sb.append_row(Row())

    hdr = Row()
    heading = f"{totals[0].get('CashType','')} {res.t('TextTotals')}".strip()
    add_text_cell(hdr, heading)
    add_text_cell(hdr, "")
    hdr.append(Cell())
    sb.append_row(hdr)

    for t in totals:
        r = Row()
        code = (t.get("CategoryCode", "") or "").strip()
        desc = t.get("Category", "") or ""
        add_text_cell(r, code)                 # A
        add_text_cell(r, desc)                 # B
        # C: marker equals code for lookup
        add_number_cell(r, formula=f"\"{code}\"", display_text="")
        # record row index if requested
        if totals_row_by_category is not None and code:
            row_index = sb.current_row_index() + 1
            sb.append_row(r)
            totals_row_by_category[code] = row_index
        else:
            sb.append_row(r)

def render_totals_formula(sb: SheetBuilder, repo, res: ResourceManager, years=None, months=None,
                          totals_row_by_category: Optional[dict[str, int]] = None):
    sb.append_row(Row())
    hdr = Row()
    add_text_cell(hdr, res.t("TextTotals"))
    add_text_cell(hdr, "")
    hdr.append(Cell())
    sb.append_row(hdr)

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

        r = Row()
        add_text_cell(r, code)         # A: CategoryCode
        add_text_cell(r, desc)         # B: Category
        add_number_cell(r, formula=f"\"{code}\"", display_text="")  # C marker

        # Register row for expressions block
        row_index = sb.current_row_index() + 1
        if totals_row_by_category is not None:
            totals_row_by_category[code] = row_index

        sum_codes_rows = repo.get_category_total_codes(code) or []
        src_codes = [row.get("SourceCategoryCode", row.get("CategoryCode", "")) for row in sum_codes_rows if row]

        for col in range(first_col, last_col + 1):
            col_letter = _col_letter(col)
            terms = []
            for sc in src_codes:
                sc = (sc or "").strip()
                if totals_row_by_category and sc in totals_row_by_category:
                    terms.append(f"{col_letter}{totals_row_by_category[sc]}")

            if terms:
                add_number_cell(r, formula="+".join(terms))
            else:
                add_number_cell(r, value=0.0)

        sb.append_row(r)

def render_expressions(sb: SheetBuilder,
                repo,
                res: ResourceManager,
                years=None,
                months=None,
                totals_row_by_category: Optional[dict[str, int]] = None):
    """
    Expressions:
    - A: Category (name)
    - B: blank
    - C: CategoryCode marker (if resolvable)
    - D..: formulas referencing totals rows in the same column.
    """
    hdr = Row()
    add_text_cell(hdr, res.t("TextAnalysis"))
    add_text_cell(hdr, "")
    hdr.append(Cell())
    sb.append_row(hdr)

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
        r = Row()
        category_name = (expr.get("Category") or "").strip()
        template = (expr.get("Expression") or "").strip()

        # Determine style override from Format when SyntaxTypeCode is Libre or Both
        syntax = (expr.get("SyntaxTypeCode") or "").strip().lower()
        fmt = expr.get("Format") or expr.get("TemplateCode")  # support alt field names
        override_style = _to_semantic_cell_style(fmt) if syntax in ("libre", "both") else "CASH0_CELL"

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
            add_number_cell(r, formula=f"\"{expr_code}\"", display_text="")
        else:
            r.append(Cell())

        # Extract [Name] tokens
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

        # Map names -> codes
        totals_by_name_local = totals_by_name
        name_to_code: dict[str, str] = {}
        for name in tokens:
            code = ""
            if hasattr(repo, "get_category_code_from_name"):
                code = (repo.get_category_code_from_name(name) or "").strip()
            if not code:
                code = totals_by_name_local.get(name, "")
            if not code:
                code = name
            name_to_code[name] = code

        # Replace [Name] -> [Code] and normalize for Calc
        normalized = template
        for name in tokens:
            normalized = normalized.replace(f"[{name}]", f"[{name_to_code[name]}]")
        normalized = normalize_for_calc(normalized)

        # Write per-column formulas with style override
        for col in range(first_col, last_col + 1):
            col_letter = _col_letter(col)
            formula = normalized
            for code in {c for c in name_to_code.values()}:
                row_index = (totals_row_by_category or {}).get(code, -1)
                if row_index <= 0:
                    formula = formula.replace(f"[{code}]", "0")
                else:
                    formula = formula.replace(f"[{code}]", f"{col_letter}{row_index}")
            add_number_cell(r, formula=formula, style=override_style)

        sb.append_row(r)

def render_vat_recurrence_totals(sb: SheetBuilder, repo, res, years, months, include_active_periods, include_tax_accruals):
    hdr = Row()
    vat_type = (repo.get_vat_recurrence_type() or "").upper()
    add_text_cell(hdr, f"{res.t('TextVatDueTitle')} {vat_type}".upper())
    add_text_cell(hdr, "")
    hdr.append(Cell())
    sb.append_row(hdr)

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
        r = Row()
        add_text_cell(r, label.upper())
        add_text_cell(r, "")
        r.append(Cell())

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
                add_number_cell(r, v)
            add_number_cell(r, sum(period_vals))

        sb.append_row(r)

    sb.append_row(Row())  # spacer

def render_vat_period_totals(sb: SheetBuilder, repo, res, years, months, include_active_periods, include_tax_accruals):
    hdr = Row()
    add_text_cell(hdr, f"{res.t('TextVatDueTitle')} {res.t('TextTotals')}".upper())
    add_text_cell(hdr, "")
    hdr.append(Cell())
    sb.append_row(hdr)

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
        r = Row()
        add_text_cell(r, label.upper() if li == len(labels) - 1 else label.upper())
        add_text_cell(r, "")
        r.append(Cell())

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
                add_number_cell(r, v)
            add_number_cell(r, sum(month_vals))

        sb.append_row(r)

    sb.append_row(Row())  # spacer

def render_bank_balances(sb: SheetBuilder, repo, res, years, months):
    sb.append_row(Row())
    hr = Row()
    add_text_cell(hr, res.t("TextClosingBalances").upper())
    add_text_cell(hr, "")
    hr.append(Cell())
    sb.append_row(hr)

    accounts = repo.get_bank_accounts()
    if not accounts:
        r = Row()
        add_text_cell(r, "(no bank accounts)")
        add_text_cell(r, "")
        r.append(Cell())
        sb.append_row(r)
        return

    cols_per_year = len(months) + 1
    total_cols = len(years) * cols_per_year
    company_totals = [0.0] * total_cols

    for acct in accounts:
        r = Row()
        add_text_cell(r, acct.get("AccountCode", ""))
        add_text_cell(r, acct.get("AccountName", ""))
        r.append(Cell())

        balances = repo.get_bank_balances(acct.get("AccountCode", ""))
        bal_map = {(int(b["YearNumber"]), int(b["MonthNumber"])): float(b["Balance"] or 0) for b in balances}

        for y_idx, y in enumerate(years):
            year_num = int(y.get("YearNumber"))
            last_val = None
            for m_idx, m in enumerate(months):
                v = bal_map.get((year_num, int(m.get("MonthNumber"))), 0.0)
                add_number_cell(r, v)
                company_totals[y_idx * cols_per_year + m_idx] += v
                last_val = v
            carry = bal_map.get((year_num, 12), last_val if last_val is not None else 0.0)
            add_number_cell(r, carry)
            company_totals[y_idx * cols_per_year + len(months)] += carry

        sb.append_row(r)

    tr = Row()
    add_text_cell(tr, res.t("TextCompanyBalance").upper())
    add_text_cell(tr, "")
    tr.append(Cell())
    for val in company_totals:
        add_number_cell(tr, val)
    sb.append_row(tr)

def render_balance_sheet(sb: SheetBuilder, repo, res, years, months):
    hr = Row()
    add_text_cell(hr, res.t("TextBalanceSheet").upper())
    sb.append_row(hr)
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
        r = Row()
        add_text_cell(r, code)      # A
        add_text_cell(r, name)      # B
        r.append(Cell())            # C reserved so months start at D

        cur_row_index = sb.current_row_index() + 1

        for y_idx, y in enumerate(years):
            year_num = int(y.get("YearNumber"))
            year_start_col = 4 + (y_idx * (len(months) + 1))  # D is 4
            last_non_empty_col_letter = None

            # months
            for m_idx, m in enumerate(months):
                mm = int(m.get("MonthNumber"))
                v = groups[key].get((year_num, mm))
                add_number_cell(r, value=(v or 0.0))
                if v is not None:
                    last_non_empty_col_letter = _col_letter(year_start_col + m_idx)

            # year total: reference last non-empty month cell if any, else 0
            if last_non_empty_col_letter:
                add_number_cell(r, formula=f"{last_non_empty_col_letter}{cur_row_index}", display_text="")
            else:
                add_number_cell(r, value=0.0)

        sb.append_row(r)

    cap = Row()
    add_text_cell(cap, res.t("TextCapital").upper())
    add_text_cell(cap, "")
    cap.append(Cell())  # C reserved so totals begin at D

    # Capital per column: SUM of asset rows in this section
    first_asset_row = sb.current_row_index() - len(order) + 1
    last_row_index = sb.current_row_index() + 1
    for col_offset in range(len(years) * (len(months) + 1)):
        col = 4 + col_offset
        letter = _col_letter(col)
        add_number_cell(cap, formula=f"SUM([.{letter}{first_asset_row}:.{letter}{last_row_index - 1}])")
    sb.append_row(cap)

def _parse_locale_tuple(locale_str: str) -> tuple[str, str]:
    s = (locale_str or "").strip()
    if not s:
        return ("en", "GB")

    # Normalize common aliases before splitting
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
        # Default country per language when none provided
        defaults = {
            "en": "GB",
            "fr": "FR",
            "de": "DE",
            "es": "ES",
        }
        country = defaults.get(lang, lang.upper())
        return (lang, country)

    # If a country was explicitly provided, respect it
    lang = parts[0].lower()
    country = parts[1].upper()
    return (lang, country)

def generate_ods(payload: dict) -> tuple[str, bytes]:
    params = payload.get("Params") or payload.get("params") or {}
    conn = payload.get("SqlConnection") or payload.get("sqlConnection") or payload.get("connectionString")
    locale = params.get("locale") or "en-GB"
    # Normalize locale once and reuse
    lang, country = _parse_locale_tuple(locale)
    normalized_locale = f"{lang}-{country}"

    res = ResourceManager(normalized_locale)
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

    table_name = "Cash Flow"
    sb = SheetBuilder(name=table_name)

    # Columns: A (Code), B (Name), C hidden (we just create the columns; visibility skip for now)
    sb.table.append(Column())  # A
    sb.table.append(Column())  # B
    sb.table.append(Column())  # C (was hidden in odfpy; can be handled later if needed)

    # Row 1: Title
    tr = Row()
    add_text_cell(tr, res.t("TextStatementTitle").format(active.get("MonthName",""), active.get("Description","")))
    add_text_cell(tr, "")
    tr.append(Cell())
    sb.append_row(tr)

    # Row 2: Company
    tr = Row()
    add_text_cell(tr, company_name or "")
    add_text_cell(tr, "")
    tr.append(Cell())
    sb.append_row(tr)

    # Row 3: Year labels
    year_row = Row()
    add_text_cell(year_row, res.t("TextDate"))
    add_text_cell(year_row, datetime.now().strftime("%d %b %H:%M:%S"))
    year_row.append(Cell())  # C
    for y in years:
        desc = y.get("Description") or str(y.get("YearNumber"))
        status = y.get("CashStatus") or ""
        add_text_cell(year_row, f"{desc} ({status})" if status else desc)
        for _ in range(len(months) - 1):
            year_row.append(Cell())
        add_text_cell(year_row, desc)
    sb.append_row(year_row)

    # Row 4: Column headers
    hdr = Row()
    add_text_cell(hdr, res.t("TextCode"))
    add_text_cell(hdr, res.t("TextName"))
    hdr.append(Cell())
    for _y in years:
        for m in months:
            add_text_cell(hdr, str(m.get("MonthName","")))
        add_text_cell(hdr, res.t("TextTotals"))
    sb.append_row(hdr)

    # Trade categories + summary
    totals_row_by_category: dict[str, int] = {}
    render_categories_and_summary(sb, repo, res, years, months, CashType.Trade,
                                  include_active, include_orderbook, False,
                                  totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months,
                                    repo.get_categories(CashType.Trade),
                                    totals_row_by_category)

    # Money categories + summary
    render_categories_and_summary(sb, repo, res, years, months, CashType.Money,
                                  False, False, False,
                                  totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months,
                                    repo.get_categories(CashType.Money),
                                    totals_row_by_category)

    # Totals block for Trade
    render_summary_totals_block(sb, repo, res, CashType.Trade, totals_row_by_category)

    # Totals-of-totals formulas
    render_totals_formula(sb, repo, res, years, months, totals_row_by_category)

    # Tax categories + summary
    render_categories_and_summary(sb, repo, res, years, months, CashType.Tax,
                                  include_active, False, include_tax_accruals,
                                  totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months,
                                    repo.get_categories(CashType.Tax),
                                    totals_row_by_category)

    # Totals block for Tax
    render_summary_totals_block(sb, repo, res, CashType.Tax, totals_row_by_category)

    # Expressions
    render_expressions(sb, repo, res, years, months, totals_row_by_category)

    if include_bank_balances:
        render_bank_balances(sb, repo, res, years, months)
        sb.append_row(Row())

    if include_vat_details:
        render_vat_recurrence_totals(sb, repo, res, years, months, include_active, include_tax_accruals)
        render_vat_period_totals(sb, repo, res, years, months, include_active, include_tax_accruals)

    if include_balance_sheet:
        render_balance_sheet(sb, repo, res, years, months)

    doc = Document("spreadsheet")
    doc.body.append(sb.table)
    
    _append_add_number_cell_check(doc) # verification via add_number_cell

    freeze_first_rows(doc, table_name, freeze_row_index=4)  # no-op for now

    buf = io.BytesIO()
    try:
        doc.save(buf)
        content = buf.getvalue()
    except TypeError:
        content = doc.save()

    # Apply Style Factory: set locale and strip default artifacts
    content = apply_styles_bytes(content, locale=(lang, country), strip_defaults=True)

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return f"Cash_Flow_{ts}.ods", content

if __name__ == "__main__":
    if len(sys.argv) >= 2:
        with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
    else:
        payload = json.loads(sys.stdin.read())
    filename, content = generate_ods(payload)
    print(f"{filename}|{base64.b64encode(content).decode('ascii')}")