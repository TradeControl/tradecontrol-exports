# pip install pyodbc
from typing import Any, Dict, Iterable, List, Optional
from data.enums import CashType
import pyodbc

def _query_all(conn_str: str, sql: str, params: Iterable[Any] = ()) -> List[Dict[str, Any]]:
    with pyodbc.connect(conn_str) as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]

def _query_one(conn_str: str, sql: str, params: Iterable[Any] = ()) -> Optional[Dict[str, Any]]:
    rows = _query_all(conn_str, sql, params)
    return rows[0] if rows else None

class SqlServerRepository:
    def __init__(self, conn_str: str):
        self.conn_str = conn_str

    def _cash_type_code(self, cash_type: Any) -> int:
        if isinstance(cash_type, int): return cash_type
        if isinstance(cash_type, str):
            s = cash_type.strip()
            if s.isdigit(): return int(s)
            m = {"trade": 0, "money": 1, "tax": 2}
            if s.lower() in m: return m[s.lower()]
        raise ValueError(f"Unknown cash_type '{cash_type}'. Use numeric code or Trade/Money/Tax.")

    # Core periods/company (App schema)
    def get_active_period(self) -> Optional[Dict[str, Any]]:
        return _query_one(self.conn_str, "SELECT YearNumber, MonthNumber, StartOn, MonthName, Description FROM App.vwActivePeriod")

    def get_active_years(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT YearNumber, Description, CashStatus FROM App.vwActiveYears ORDER BY YearNumber")

    def get_months(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT MonthNumber, MonthName, StartOn FROM App.vwMonths ORDER BY StartOn")

    def get_company_name(self) -> str:
        row = _query_one(self.conn_str, "SELECT TOP (1) SubjectName FROM App.vwHomeAccount")
        return row["SubjectName"] if row else ""

    # Categories/codes/values
    def get_categories(self, cash_type: CashType | int) -> list[dict]:
        code = int(cash_type)  # IntEnum -> int
        return _query_all(self.conn_str,
            "SELECT CategoryCode, Category, CashPolarityCode, DisplayOrder FROM Cash.fnFlowCategory(?) ORDER BY DisplayOrder, Category",
            (code,))

    def get_cash_codes(self, category_code: str) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT CashCode, CashDescription FROM Cash.fnFlowCategoryCashCodes(?) ORDER BY CashDescription", (category_code,))

    def get_cash_code_values(self, cash_code: str, year_number: int, include_active: bool, include_orderbook: bool, include_tax_accruals: bool) -> List[Dict[str, Any]]:
        # Stored procedure: Cash.proc_FlowCashCodeValues
        with pyodbc.connect(self.conn_str) as conn:
            cur = conn.cursor()
            cur.execute(
                "{CALL Cash.proc_FlowCashCodeValues(?, ?, ?, ?, ?)}",
                (cash_code, int(year_number), 1 if include_active else 0, 1 if include_orderbook else 0, 1 if include_tax_accruals else 0)
            )
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
        # Expected shape: StartOn, InvoiceValue, InvoiceTax, ForecastValue, ForecastTax
        # Project to MonthNumber + InvoiceValue for ODS categories (month alignment)
        result = []
        for r in rows:
            row = dict(zip(cols, r))
            # MonthNumber from StartOn
            start_on = row.get("StartOn")
            month_num = start_on.month if hasattr(start_on, "month") else None
            result.append({
                "MonthNumber": month_num,
                "InvoiceValue": row.get("InvoiceValue", 0)
            })
        return result

    # Totals and expressions
    def get_category_totals(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT CategoryCode, Category FROM Cash.vwCategoryTotals ORDER BY DisplayOrder, Category")

    def get_category_total_codes(self, category_code: str) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT CategoryCode AS SourceCategoryCode FROM Cash.fnFlowCategoryTotalCodes(?) ORDER BY CategoryCode", (category_code,))

    def get_category_expressions(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT DisplayOrder, CategoryCode, Category, Expression, Format FROM Cash.vwCategoryExpressions WHERE SyntaxTypeCode IN (0,1) ORDER BY DisplayOrder, Category")

    def get_category_code_from_name(self, name: str) -> Optional[str]:
        row = _query_one(self.conn_str, "SELECT CategoryCode FROM Cash.vwFlowCategories WHERE Category = ?", (name,))
        return row["CategoryCode"] if row else None

    def set_category_expression_status(self, category_code: str, is_error: bool, message: Optional[str] = None) -> None:
        # Minimal implementation: write to App.proc_EventLog (Error=0, Information=2)
        event_type = 0 if is_error else 2
        msg = f"Expression {category_code}: {message or 'OK'}"
        with pyodbc.connect(self.conn_str) as conn:
            cur = conn.cursor()
            # LogCode is OUTPUT in SQL; we can ignore it here
            cur.execute("{CALL App.proc_EventLog(?, ?, ?)}", (msg, event_type, None))
            conn.commit()

    # VAT
    def get_vat_recurrence_type(self) -> str:
        row = _query_one(self.conn_str, "SELECT TOP (1) UPPER(Recurrence) AS VatType FROM Cash.vwFlowTaxType WHERE TaxTypeCode = 1")
        return row["VatType"] if row else ""

    def get_vat_recurrence(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT YearNumber, StartOn, HomeSales, HomePurchases, ExportSales, ExportPurchases, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatAdjustment, VatDue FROM Cash.vwFlowVatRecurrence ORDER BY YearNumber, StartOn")

    def get_vat_recurrence_accruals(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT YearNumber, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue FROM Cash.vwFlowVatRecurrenceAccruals ORDER BY YearNumber")

    def get_vat_period_totals(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT YearNumber, StartOn, HomeSales, HomePurchases, ExportSales, ExportPurchases, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue FROM Cash.vwFlowVatPeriodTotals ORDER BY YearNumber, StartOn")

    def get_vat_period_accruals(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT YearNumber, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue FROM Cash.vwFlowVatPeriodAccruals ORDER BY YearNumber")

    # Bank
    def get_bank_accounts(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT AccountCode, AccountName FROM Cash.vwBankAccounts ORDER BY DisplayOrder, AccountCode")

    def get_bank_balances(self, account_code: str) -> List[Dict[str, Any]]:
        return _query_all(
            self.conn_str,
            """
            SELECT 
                YearNumber,
                CAST(MONTH(StartOn) AS tinyint) AS MonthNumber,
                CAST(Balance AS decimal(18,5)) AS Balance
            FROM Cash.fnFlowBankBalances(?)
            ORDER BY YearNumber, StartOn
            """,
            (account_code,)
        )

    # Balance sheet
    def get_balance_sheet(self) -> List[Dict[str, Any]]:
        return _query_all(self.conn_str, "SELECT AssetCode, AssetName, YearNumber, MonthNumber, Balance FROM Cash.vwBalanceSheet ORDER BY EntryNumber")