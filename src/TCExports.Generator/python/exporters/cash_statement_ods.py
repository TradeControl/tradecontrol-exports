# pip install pyodbc odfdo
import json, sys, io, base64
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Union

from odfdo import Document, Settings, Style
from odfdo.table import Table, Row, Cell, Column
from odfdo.element import Element

from data.enums import CashType
from data.factory import create_repo
from i18n.resources import ResourceManager
from style_factory import apply_styles_bytes

def _col_letter(index_1based: int) -> str:
    dividend = index_1based
    name = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        name = chr(65 + modulo) + name
        dividend = (dividend - modulo) // 26
    return name

# def _empty_cell(style: Optional[str] = None) -> Element:
#     # Truly empty cell: no text:p child
#     e = Element.from_tag("table:table-cell")
#     if style:
#         e.set_attribute("table:style-name", style)
#     return e

def add_spanned_text_cell(row: Row, text: str, span: int = 1, style: Optional[str] = None) -> None:
    # Create a table cell with text and span it across span columns.
    tc = Element.from_tag("table:table-cell")
    if style:
        tc.set_attribute("table:style-name", style)
    if span and span > 1:
        tc.set_attribute("table:number-columns-spanned", str(span))
    p = Element.from_tag("text:p")
    p.text = str(text or "")
    tc.append(p)
    row.append(tc)
    # For each additional spanned column, add a covered cell (never contains text:p)
    for _ in range(max(0, span - 1)):
        row.append(Element.from_tag("table:covered-table-cell"))

def add_text_cell(row: Row, text: str, style: Optional[str] = None):
    # If there is actual text, create a normal cell
    if text is not None and str(text).strip() != "":
        cell = Cell(text=str(text))
        if style:
            cell.set_attribute("table:style-name", style)
        row.append(cell)
    else:
        # Method 5
        # append a raw empty table cell (no text:p)
        # row.append(_empty_cell(style))

        # Method 4
        # empty = Element.from_tag("table:table-cell")
        # if style:
        #     empty.set_attribute("table:style-name", style)
        # Do NOT add any text child
        # row.append(empty)

        # Method 3
        # cell = Cell()
        # row.append(cell)
        
        # last_cell = row.get_cells()[-1]
        # paragraphs = last_cell.get_elements('text:p')
        # for p in paragraphs:
        #     last_cell.delete_element(p)

        # Method 2
        #cell = Cell(text=None)
        #for p in cell.get_elements('text:p'):
        #    cell.delete_element(p)
        #row.append(cell)

        # Method 1
        null_cell = Element.from_tag('table:table-cell')
        row.append(null_cell)

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

def freeze_and_rename_active_sheet(doc: Document, new_name: str, freeze_row_index: int = 4):
    settings = getattr(doc, "settings", None)
    if settings is None:
        settings = Settings()
        doc.settings = settings

    def item(name: str, typ: str, val: str) -> Element:
        e = Element.from_tag("config:config-item")
        e.set_attribute("config:name", name)
        e.set_attribute("config:type", typ)
        e.text = str(val)
        return e

    # 1) Locate the main view-settings set (do not remove it)
    view_sets = settings.root.xpath(".//config:config-item-set[@config:name='ooo:view-settings']")
    if not view_sets:
        # If absent, create a minimal container
        view_set = Element.from_tag("config:config-item-set")
        view_set.set_attribute("config:name", "ooo:view-settings")
        settings.root.append(view_set)
    else:
        view_set = view_sets[0]

    # 2) Get or create the first view entry
    views_list = view_set.xpath(".//config:config-item-map-indexed[@config:name='Views']")
    if not views_list:
        views = Element.from_tag("config:config-item-map-indexed")
        views.set_attribute("config:name", "Views")
        view_set.append(views)
    else:
        views = views_list[0]

    entries = views.xpath("./config:config-item-map-entry")
    if entries:
        view_entry = entries[0]
    else:
        view_entry = Element.from_tag("config:config-item-map-entry")
        views.append(view_entry)

    # 3) Ensure ActiveTable points to new_name; remove old one if present
    for at in view_entry.xpath("./config:config-item[@config:name='ActiveTable']"):
        at.parent.delete(at)
    view_entry.append(item("ActiveTable", "string", new_name))

    # 4) Find/create Tables map, remove ghost entries, and upsert the Cash Flow freeze keys
    tables_map_list = view_entry.xpath("./config:config-item-map-named[@config:name='Tables']")
    if not tables_map_list:
        tables_map = Element.from_tag("config:config-item-map-named")
        tables_map.set_attribute("config:name", "Tables")
        view_entry.append(tables_map)
    else:
        tables_map = tables_map_list[0]

    # Remove ghost table entries like Feuille1/Sheet1
    for ghost in list(tables_map.xpath("./config:config-item-map-entry")):
        nm = ghost.get_attribute("config:name") or ""
        if nm in ("Feuille1", "Sheet1"):
            tables_map.delete(ghost)

    # Get or create the table entry for new_name
    target_entries = tables_map.xpath(f"./config:config-item-map-entry[@config:name='{new_name}']")
    if target_entries:
        table_entry = target_entries[0]
    else:
        table_entry = Element.from_tag("config:config-item-map-entry")
        table_entry.set_attribute("config:name", new_name)
        tables_map.append(table_entry)

    def upsert(ci_name: str, typ: str, val: str):
        for ci in table_entry.xpath(f"./config:config-item[@config:name='{ci_name}']"):
            ci.parent.delete(ci)
        table_entry.append(item(ci_name, typ, val))

    # D5 freeze and pane focus
    upsert("CursorPositionX", "int", "3")
    upsert("CursorPositionY", "int", str(freeze_row_index))
    upsert("HorizontalSplitMode", "short", "2")
    upsert("VerticalSplitMode", "short", "2")
    upsert("HorizontalSplitPosition", "int", "3")
    upsert("VerticalSplitPosition", "int", str(freeze_row_index))
    upsert("ActiveSplitRange", "short", "3")
    upsert("PositionLeft", "int", "0")
    upsert("PositionRight", "int", "3")
    upsert("PositionTop", "int", "0")
    upsert("PositionBottom", "int", str(freeze_row_index))

    # 5) Clean up ScriptConfiguration ghosts (optional but avoids Calc renaming)
    cfg_sets = settings.root.xpath(".//config:config-item-set[@config:name='ooo:configuration-settings']")
    if cfg_sets:
        cfg_set = cfg_sets[0]
        script_maps = cfg_set.xpath(".//config:config-item-map-named[@config:name='ScriptConfiguration']")
        if script_maps:
            script_map = script_maps[0]
            for ghost in list(script_map.xpath("./config:config-item-map-entry")):
                nm = ghost.get_attribute("config:name") or ""
                if nm in ("Feuille1", "Sheet1"):
                    script_map.delete(ghost)
            # Upsert CodeName for new_name if needed
            target = script_map.xpath(f"./config:config-item-map-entry[@config:name='{new_name}']")
            if not target:
                entry = Element.from_tag("config:config-item-map-entry")
                entry.set_attribute("config:name", new_name)
                entry.append(item("CodeName", "string", "Sheet1"))
                script_map.append(entry)

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
                # Use default numeric style so Style Factory formats are applied
                add_number_cell(r, formula=f"{col_letter}{target_row}", style="CASH0_CELL")
            else:
                add_number_cell(r, value=0.0, style="CASH0_CELL")

        sb.append_row(r)

    # Period Total row: SUM of the summary rows per column
    pr = Row()
    add_text_cell(pr, res.t("TextPeriodTotal"))
    add_text_cell(pr, "")
    pr.append(Cell())
    end_row_index = sb.current_row_index()

    for col in range(firstCol, lastCol + 1):
        col_letter = _col_letter(col)
        add_number_cell(pr, formula=f"SUM([.{col_letter}{start_row_index}:.{col_letter}{end_row_index}])", style="CASH0_CELL")

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

        # Always honor Format/TemplateCode (e.g., Pct0 -> PCT0_CELL) — no SyntaxTypeCode gating
        fmt = expr.get("Format") or expr.get("TemplateCode")
        override_style = _to_semantic_cell_style(fmt) if fmt else "NUM0_CELL"

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

        # Write per-column formulas with style override (Pct0 -> PCT0_CELL, Num2 -> NUM2_CELL, etc.)
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
                start_on = p.get("StartOn")
                if isinstance(start_on, datetime) and start_on.tzinfo is None:
                    start_on = start_on.replace(tzinfo=timezone.utc)
                if include_active_periods or (start_on <= datetime.now(timezone.utc)):
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
                start_on = p.get("StartOn")
                if isinstance(start_on, datetime) and start_on.tzinfo is None:
                    start_on = start_on.replace(tzinfo=timezone.utc)
                if include_active_periods or (start_on <= datetime.now(timezone.utc)):
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

def add_stylesheet(doc):
    bold_style = Style(family='table-cell', name='BoldHeaderStyle')
    bold_style.set_properties(area='text', **{
        'fo:font-weight': 'bold',
        'style:font-weight-complex': 'bold'
    })
    bold_style.set_properties(area='table-cell', **{
        'style:vertical-align': 'middle',
        'style:wrap-option': 'no-wrap',
        'style:shrink-to-fit': 'false'
    })
    doc.insert_style(bold_style, automatic=True)

    col_b_style = Style(family='table-column', name='ColBWidth')
    col_b_style.set_properties(area='table-column', **{
        'style:column-width': '8.5cm'
    })
    doc.insert_style(col_b_style, automatic=True)

    row4_hdr = Style(family='table-cell', name='Row4HeaderCell')
    row4_hdr.set_properties(area='text', **{
        'fo:font-size': '8pt',
        'fo:font-weight': 'bold',
        'style:font-weight-complex': 'bold'
    })
    row4_hdr.set_properties(area='table-cell', **{
        'style:vertical-align': 'middle',
        'style:wrap-option': 'no-wrap',
        'style:shrink-to-fit': 'false',
        'fo:border-top': '0.5pt solid #000000',
        'fo:border-bottom': '0.5pt solid #000000'
    })
    doc.insert_style(row4_hdr, automatic=True)

    totals_col_header = Style(family='table-cell', name='TotalsColHeaderCell')
    totals_col_header.set_properties(area='text', **{
        'fo:font-size': '8pt',
        'fo:font-weight': 'bold',
        'style:font-weight-complex': 'bold'
    })
    totals_col_header.set_properties(area='table-cell', **{
        'style:vertical-align': 'middle',
        'style:wrap-option': 'no-wrap',
        'style:shrink-to-fit': 'false',
        'fo:border-top': '0.5pt solid #000000',
        'fo:border-bottom': '0.5pt solid #000000',
        'fo:border-left': '0.5pt solid #000000',
        'fo:border-right': '1.5pt solid #000000'
    })
    doc.insert_style(totals_col_header, automatic=True)

    totals_col_default = Style(family='table-cell', name='TotalsColDefaultCell')
    totals_col_default.set_properties(area='table-cell', **{
        'fo:border-left': '0.5pt solid #000000',
        'fo:border-right': '1.5pt solid #000000'
    })
    doc.insert_style(totals_col_default, automatic=True)

    return doc

def initialise_ods(payload: dict) -> dict:
    params = payload.get("Params") or payload.get("params") or {}
    conn = payload.get("SqlConnection") or payload.get("sqlConnection") or payload.get("connectionString")
    locale = params.get("locale") or "en-GB"
    lang, country = _parse_locale_tuple(locale)

    res = ResourceManager(f"{lang}-{country}")
    repo = create_repo(conn, params)

    active = repo.get_active_period() or {}
    years = repo.get_active_years()
    months = repo.get_months()
    company_name = repo.get_company_name()

    include_active = params.get("includeActivePeriods") == "true"
    include_orderbook = params.get("includeOrderBook") == "true"
    include_tax_accruals = params.get("includeTaxAccruals") == "true"
    include_vat_details = params.get("includeVatDetails") == "true"
    include_bank_balances = params.get("includeBankBalances") == "true"
    include_balance_sheet = params.get("includeBalanceSheet") == "true"

    return {
        "params": params,
        "conn": conn,
        "lang": lang,
        "country": country,
        "res": res,
        "repo": repo,
        "active": active,
        "years": years,
        "months": months,
        "company_name": company_name,
        "table_name": "Cash Flow",
        "include_active": include_active,
        "include_orderbook": include_orderbook,
        "include_tax_accruals": include_tax_accruals,
        "include_vat_details": include_vat_details,
        "include_bank_balances": include_bank_balances,
        "include_balance_sheet": include_balance_sheet,
    }

def build_cashflow_table(sb: SheetBuilder, ctx: dict) -> SheetBuilder:
    res = ctx["res"]
    repo = ctx["repo"]
    years = ctx["years"]
    months = ctx["months"]
    active = ctx["active"]
    company_name = ctx["company_name"]

    include_active = ctx["include_active"]
    include_orderbook = ctx["include_orderbook"]
    include_tax_accruals = ctx["include_tax_accruals"]
    include_vat_details = ctx["include_vat_details"]
    include_bank_balances = ctx["include_bank_balances"]
    include_balance_sheet = ctx["include_balance_sheet"]

    # A,B,C
    colA = Column()
    colB = Column()
    colC = Column()
    colB.set_attribute("table:style-name", "ColBWidth")
    colC.set_attribute("table:visibility", "collapse")
    sb.table.append(colA)
    sb.table.append(colB)
    sb.table.append(colC)

    # Append period grid columns (months + totals per year)
    month_count = len(months)
    total_period_cols = len(years) * (month_count + 1)
    for i in range(total_period_cols):
        col = Column()
        # Totals column in each block gets default borders (applies to blank cells)
        if (i % (month_count + 1)) == month_count:
            col.set_attribute("table:default-cell-style-name", "TotalsColDefaultCell")
        sb.table.append(col)

    # Row 1
    r1 = Row()
    title_text = res.t("TextStatementTitle").format(active.get("MonthName", ""), active.get("Description", ""))
    add_spanned_text_cell(r1, title_text, span=3, style="BoldHeaderStyle")
    sb.append_row(r1)

    # Row 2
    r2 = Row()
    add_spanned_text_cell(r2, company_name or "", span=3, style="BoldHeaderStyle")
    sb.append_row(r2)

    # Row 3
    r3 = Row()
    add_text_cell(r3, res.t("TextDate"), style="BoldHeaderStyle")
    add_text_cell(r3, datetime.now().strftime("%d %b %H:%M:%S"), style="BoldHeaderStyle")
    r3.append(Cell(text=None))
    for y in years:
        desc = y.get("Description") or str(y.get("YearNumber"))
        status = y.get("CashStatus") or ""
        add_spanned_text_cell(r3, f"{desc} ({status})" if status else desc, span=2, style="BoldHeaderStyle")
        for _ in range(month_count - 2):
            r3.append(Cell(text=None))
        add_text_cell(r3, desc, style="BoldHeaderStyle")
    sb.append_row(r3)

    # Row 4 headers
    r4 = Row()
    add_text_cell(r4, res.t("TextCode"), style="Row4HeaderCell")
    add_text_cell(r4, res.t("TextName"), style="Row4HeaderCell")
    c_cell = Cell(text=None)
    c_cell.set_attribute("table:style-name", "Row4HeaderCell")
    r4.append(c_cell)
    # Month headers use Row4HeaderCell; Totals header uses TotalsColHeaderCell
    for _ in years:
        for m in months:
            add_text_cell(r4, str(m.get("MonthName", "")), style="Row4HeaderCell")
        add_text_cell(r4, res.t("TextTotals"), style="TotalsColHeaderCell")
    sb.append_row(r4)

    # Sections
    totals_row_by_category: dict[str, int] = {}
    render_categories_and_summary(sb, repo, res, years, months, CashType.Trade, include_active, include_orderbook, False, totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months, repo.get_categories(CashType.Trade), totals_row_by_category)

    render_categories_and_summary(sb, repo, res, years, months, CashType.Money, False, False, False, totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months, repo.get_categories(CashType.Money), totals_row_by_category)

    render_summary_totals_block(sb, repo, res, CashType.Trade, totals_row_by_category)
    render_totals_formula(sb, repo, res, years, months, totals_row_by_category)

    render_categories_and_summary(sb, repo, res, years, months, CashType.Tax, include_active, False, include_tax_accruals, totals_row_by_category)
    render_summary_after_categories(sb, repo, res, years, months, repo.get_categories(CashType.Tax), totals_row_by_category)
    render_summary_totals_block(sb, repo, res, CashType.Tax, totals_row_by_category)

    render_expressions(sb, repo, res, years, months, totals_row_by_category)

    if include_bank_balances:
        render_bank_balances(sb, repo, res, years, months)
        sb.append_row(Row())

    if include_vat_details:
        render_vat_recurrence_totals(sb, repo, res, years, months, include_active, include_tax_accruals)
        render_vat_period_totals(sb, repo, res, years, months, include_active, include_tax_accruals)

    if include_balance_sheet:
        render_balance_sheet(sb, repo, res, years, months)

    return sb

def save_cashflow(doc: Document, ctx: dict) -> tuple[str, bytes]:
    table_name = ctx["table_name"]
    lang = ctx["lang"]
    country = ctx["country"]

    # Keep working freeze logic
    freeze_and_rename_active_sheet(doc, table_name, freeze_row_index=4)

    buf = io.BytesIO()
    try:
        doc.save(buf)
        content = buf.getvalue()
    except TypeError:
        content = doc.save()

    # 1) Materialize semantic styles (NUM/PCT/CASH, maps, etc.)
    content = apply_styles_bytes(content, locale=(lang, country), strip_defaults=True)

    # 2) Column-first totals borders 
    months = ctx["months"]
    years = ctx["years"]
    content = _post_process_totals_borders(content, month_count=len(months), years_count=len(years))

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    return f"Cash_Flow_{ts}.ods", content

def _post_process_totals_borders(content: bytes, month_count: int, years_count: int) -> bytes:
    import io, zipfile, re
    from lxml import etree as ET

    src_zip = zipfile.ZipFile(io.BytesIO(content), 'r')
    namelist = src_zip.namelist()
    if 'content.xml' not in namelist:
        src_zip.close()
        return content
    content_xml = src_zip.read('content.xml')

    ns = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    }
    root = ET.fromstring(content_xml)
    auto_styles = root.find('office:automatic-styles', ns)
    if auto_styles is None:
        src_zip.close()
        return content

    def ensure_totals_default():
        existing = auto_styles.find("style:style[@style:name='TotalsColDefaultCell'][@style:family='table-cell']", ns)
        if existing is not None:
            return
        st = ET.Element(f"{{{ns['style']}}}style", {
            f"{{{ns['style']}}}name": "TotalsColDefaultCell",
            f"{{{ns['style']}}}family": "table-cell",
            f"{{{ns['style']}}}parent-style-name": "Default"
        })
        props = ET.Element(f"{{{ns['style']}}}table-cell-properties")
        props.set(f"{{{ns['fo']}}}border-left", "0.5pt solid #000000")
        props.set(f"{{{ns['fo']}}}border-right", "1.5pt solid #000000")
        st.append(props)
        auto_styles.append(st)

    def ensure_bordered_clone(src_style_name: str) -> str:
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
        props.set(f"{{{ns['fo']}}}border-left", "0.5pt solid #000000")
        props.set(f"{{{ns['fo']}}}border-right", "1.5pt solid #000000")
        auto_styles.append(new)
        return new_name

    def col_letters_to_index(letters: str) -> int:
        idx = 0
        for ch in letters.upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx

    def find_cell_by_index(row_elem: ET._Element, col_index_1based: int):
        count = 0
        for child in row_elem:
            if not hasattr(child, 'tag'):
                continue
            tag = child.tag
            if not (tag.endswith('table-cell') or tag.endswith('covered-table-cell')):
                continue
            repeat = int(child.get(f"{{{ns['table']}}}number-columns-repeated", "1"))
            next_count = count + repeat
            if col_index_1based <= next_count:
                return child
            count = next_count
        return None

    def is_totals_col(col_index_1based: int) -> bool:
        rel = col_index_1based - 4  # D=4 starts period grid
        if rel < 0:
            return False
        block = month_count + 1
        return (rel % block) == (block - 1)

    def cash_root(style_name: str):
        m = re.match(r'^(CASH\d+)_(?:POS_|NEG_)?CELL(?:_BORDERED)?$', (style_name or '').upper())
        return m.group(1) if m else None

    ensure_totals_default()

    table = root.find('.//table:table', ns)
    if table is None:
        src_zip.close()
        return content

    total_period_cols = years_count * (month_count + 1)
    total_visual_cols = 3 + total_period_cols

    cols = table.findall('table:table-column', ns)
    for i in range(total_period_cols):
        col_1based = 4 + i
        if col_1based - 1 >= len(cols):
            break
        if (i % (month_count + 1)) == month_count:
            cols[col_1based - 1].set(f"{{{ns['table']}}}default-cell-style-name", "TotalsColDefaultCell")

    rows = table.findall('table:table-row', ns)

    # 1) Bordered clones for explicitly styled totals-column cells
    for row in rows:
        cells = row.findall('table:table-cell', ns)
        for col_index_1based, cell in enumerate(cells, start=1):
            if cell.tag.endswith('covered-table-cell'):
                continue
            if not is_totals_col(col_index_1based):
                continue
            sname = cell.get(f"{{{ns['table']}}}style-name") or ""
            if sname:
                bordered = ensure_bordered_clone(sname)
                if bordered and bordered != sname:
                    cell.set(f"{{{ns['table']}}}style-name", bordered)

    # 2) Compute cached values: direct refs, SUM row-ranges, simple +/- refs, and SUM down a column
    direct_ref_patterns = [
        re.compile(r"^of:=\.?\$?([A-Za-z]+)\$?(\d+)$"),
        re.compile(r"^of:=\[\.\$?([A-Za-z]+)\$?(\d+)\]$"),
    ]
    sum_row_patterns = [
        re.compile(r"^of:=SUM\(\[\.\$?([A-Za-z]+)\$?(\d+):\.\$?([A-Za-z]+)\$?\2\]\)(\*\-1)?$"),
        re.compile(r"^of:=SUM\(\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?\2\)(\*\-1)?$"),
    ]
    # New: SUM down a single column between two row indices, optional *-1
    sum_col_patterns = [
        re.compile(r"^of:=SUM\(\[\.\$?([A-Za-z]+)\$?(\d+):\.\$?\1\$(\d+)\]\)(\*\-1)?$"),
        re.compile(r"^of:=SUM\(\$?([A-Za-z]+)\$?(\d+):\$?\1\$(\d+)\)(\*\-1)?$"),
    ]
    add_sub_pat = re.compile(r'([+\-]?)\s*(?:\[\.\$?([A-Za-z]+)\$?(\d+)\]|\$?([A-Za-z]+)\$?(\d+))')

    for row in rows:
        cells = row.findall('table:table-cell', ns)
        for c_idx, cell in enumerate(cells, start=1):
            if not is_totals_col(c_idx):
                continue
            formula = cell.get(f"{{{ns['table']}}}formula")
            if not formula:
                continue
            computed = None

            # Direct ref
            for pat in direct_ref_patterns:
                m = pat.match(formula)
                if m:
                    col_letters, row_num = m.group(1), int(m.group(2))
                    ref_col = col_letters_to_index(col_letters)
                    ref_row = rows[row_num - 1] if 1 <= row_num <= len(rows) else None
                    ref_cell = find_cell_by_index(ref_row, ref_col) if ref_row is not None else None
                    if ref_cell is not None:
                        v = ref_cell.get(f"{{{ns['office']}}}value")
                        if v is not None:
                            try:
                                computed = float(v)
                            except ValueError:
                                pass
                    break
            if computed is not None:
                cell.set(f"{{{ns['office']}}}value-type", "float")
                cell.set(f"{{{ns['office']}}}value", str(computed))
                continue

            # SUM of row-range (same row)
            for pat in sum_row_patterns:
                m = pat.match(formula)
                if not m:
                    continue
                start_col, row_num, end_col = m.group(1), int(m.group(2)), m.group(3)
                mult = -1.0 if (m.group(4) or "") == "*-1" else 1.0
                if 1 <= row_num <= len(rows):
                    start_ci = col_letters_to_index(start_col)
                    end_ci = col_letters_to_index(end_col)
                    if end_ci < start_ci:
                        start_ci, end_ci = end_ci, start_ci
                    ref_row = rows[row_num - 1]
                    total = 0.0
                    for ci in range(start_ci, end_ci + 1):
                        rc = find_cell_by_index(ref_row, ci)
                        if rc is None:
                            continue
                        v = rc.get(f"{{{ns['office']}}}value")
                        if v is None:
                            continue
                        try:
                            total += float(v)
                        except ValueError:
                            continue
                    computed = total * mult
                    break
            if computed is not None:
                cell.set(f"{{{ns['office']}}}value-type", "float")
                cell.set(f"{{{ns['office']}}}value", str(computed))
                continue

            # SUM down a column between two row indices (category totals)
            for pat in sum_col_patterns:
                m = pat.match(formula)
                if not m:
                    continue
                col_letters, start_row_num, end_row_num = m.group(1), int(m.group(2)), int(m.group(3))
                mult = -1.0 if (m.group(4) or "") == "*-1" else 1.0
                ref_ci = col_letters_to_index(col_letters)
                total = 0.0
                for rnum in range(min(start_row_num, end_row_num), max(start_row_num, end_row_num) + 1):
                    ref_row = rows[rnum - 1] if 1 <= rnum <= len(rows) else None
                    rc = find_cell_by_index(ref_row, ref_ci) if ref_row is not None else None
                    if rc is None:
                        continue
                    v = rc.get(f"{{{ns['office']}}}value")
                    if v is None:
                        continue
                    try:
                        total += float(v)
                    except ValueError:
                        continue
                computed = total * mult
                break
            if computed is not None:
                cell.set(f"{{{ns['office']}}}value-type", "float")
                cell.set(f"{{{ns['office']}}}value", str(computed))
                continue

            # +/- list of refs
            total = 0.0
            matched = False
            for sign, c1, r1, c2, r2 in add_sub_pat.findall(formula):
                matched = True
                col_letters = c1 or c2
                row_num = int(r1 or r2)
                ref_col = col_letters_to_index(col_letters)
                rc_row = rows[row_num - 1] if 1 <= row_num <= len(rows) else None
                rc = find_cell_by_index(rc_row, ref_col) if rc_row is not None else None
                if rc is None:
                    continue
                v = rc.get(f"{{{ns['office']}}}value")
                if v is None:
                    continue
                try:
                    valf = float(v)
                except ValueError:
                    continue
                total += (-valf if sign == '-' else valf)
            if matched:
                cell.set(f"{{{ns['office']}}}value-type", "float")
                cell.set(f"{{{ns['office']}}}value", str(total))

    # 3) Enforce CASH POS/NEG bordered style from cached value
    for row in rows:
        cells = row.findall('table:table-cell', ns)
        for col_index_1based, cell in enumerate(cells, start=1):
            if not is_totals_col(col_index_1based):
                continue
            val_str = cell.get(f"{{{ns['office']}}}value")
            vtype = cell.get(f"{{{ns['office']}}}value-type")
            if vtype != "float" or val_str is None:
                continue
            try:
                num = float(val_str)
            except ValueError:
                continue
            sname = (cell.get(f"{{{ns['table']}}}style-name") or "").upper()
            root_name = cash_root(sname)
            if not root_name:
                continue
            pos = ensure_bordered_clone(f"{root_name}_POS_CELL")
            neg = ensure_bordered_clone(f"{root_name}_NEG_CELL")
            cell.set(f"{{{ns['table']}}}style-name", neg if num < 0 else pos)

    # 4) Normalize spacer/short rows
    for row in rows:
        visual_cols = 0
        children = list(row)
        for child in children:
            if not hasattr(child, 'tag'):
                row.remove(child)
                continue
            tag = child.tag
            if not (tag.endswith('table-cell') or tag.endswith('covered-table-cell')):
                continue
            repeat = int(child.get(f"{{{ns['table']}}}number-columns-repeated", "1"))
            visual_cols += repeat
        if visual_cols < total_visual_cols:
            deficit = total_visual_cols - visual_cols
            rc = ET.Element(f"{{{ns['table']}}}table-cell")
            rc.set(f"{{{ns['table']}}}number-columns-repeated", str(deficit))
            row.append(rc)

    # 5) End-of-sheet repeater
    rep = ET.Element(f"{{{ns['table']}}}table-row")
    rep.set(f"{{{ns['table']}}}number-rows-repeated", "1048568")
    rep_cell = ET.Element(f"{{{ns['table']}}}table-cell")
    rep_cell.set(f"{{{ns['table']}}}number-columns-repeated", str(total_visual_cols))
    rep.append(rep_cell)
    table.append(rep)

    # Final defensive cleanup
    for elem in root.iter():
        for child in list(elem):
            if not hasattr(child, 'tag'):
                elem.remove(child)

    new_content_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    entries = {}
    for name in namelist:
        if name == 'content.xml':
            continue
        entries[name] = src_zip.read(name)
    src_zip.close()

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

def generate_ods(payload: dict) -> tuple[str, bytes]:

    ctx = initialise_ods(payload)

    doc = Document("spreadsheet")
    doc = add_stylesheet(doc)

    sb = SheetBuilder(name=ctx["table_name"])
    sb = build_cashflow_table(sb, ctx)

    doc.body.append(sb.table)
    return save_cashflow(doc, ctx)

if __name__ == "__main__":
    if len(sys.argv) >= 2:
        with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
    else:
        payload = json.loads(sys.stdin.read())
    filename, content = generate_ods(payload)
    print(f"{filename}|{base64.b64encode(content).decode('ascii')}")