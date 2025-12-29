# pip install pyodbc
from typing import Any, Dict, Iterable, List, Optional
import pyodbc

def _query_all(conn_str: str, sql: str, params: Iterable[Any] = ()) -> List[Dict[str, Any]]:
    with pyodbc.connect(conn_str) as conn:
        cur = conn.cursor()
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]

def _query_one(conn_str: str, sql: str, params: Iterable[Any] = ()) -> Optional[Dict[str, Any]]:
    rows = _query_all(conn_str, sql, params); return rows[0] if rows else None

class SqlServerRepository:
    def __init__(self, conn_str: str): self.conn_str = conn_str
    def get_active_period(self): return _query_one(self.conn_str, "SELECT YearNumber, MonthNumber, StartOn, MonthName, Description FROM App.vwActivePeriod")
    def get_active_years(self): return _query_all(self.conn_str, "SELECT YearNumber, Description, CashStatus FROM App.vwActiveYears ORDER BY YearNumber")
    def get_months(self): return _query_all(self.conn_str, "SELECT MonthNumber, MonthName, StartOn FROM App.vwMonths ORDER BY StartOn")
    def get_company_name(self): r = _query_one(self.conn_str, "SELECT TOP (1) SubjectName FROM App.vwHomeAccount"); return r["SubjectName"] if r else ""
    def get_categories(self, cash_type: str): return _query_all(self.conn_str, "SELECT CategoryCode, Category, CashPolarityCode FROM Cash.vwFlowCategories WHERE CashType = ? ORDER BY CategoryOrder", (cash_type,))
    def get_cash_codes(self, category_code: str): return _query_all(self.conn_str, "SELECT CashCode, CashDescription FROM Cash.vwFlowCategoryCodes WHERE CategoryCode = ? ORDER BY CodeOrder", (category_code,))
    def get_cash_code_values(self, cash_code: str, year_number: int, include_active: bool, include_orderbook: bool, include_tax_accruals: bool):
        return _query_all(self.conn_str, "SELECT MonthNumber, InvoiceValue FROM Cash.vwCashCodeValues WHERE CashCode = ? AND YearNumber = ? ORDER BY MonthNumber", (cash_code, year_number))
    def get_category_totals(self): return _query_all(self.conn_str, "SELECT CategoryCode FROM Cash.vwCategoryTotals ORDER BY TotalOrder")
    def get_category_total_codes(self, category_code: str): return _query_all(self.conn_str, "SELECT SourceCategoryCode FROM Cash.vwCategoryTotalCodes WHERE CategoryCode = ? ORDER BY SourceOrder", (category_code,))
    def get_category_expressions(self): return _query_all(self.conn_str, "SELECT Category, Expression, Format FROM Cash.vwCategoryExpressions ORDER BY ExpressionOrder")
    def get_category_code_from_name(self, name: str): r = _query_one(self.conn_str, "SELECT CategoryCode FROM Cash.vwFlowCategories WHERE Category = ?", (name,)); return r["CategoryCode"] if r else None
    def get_vat_recurrence_type(self): r = _query_one(self.conn_str, "SELECT VatType FROM Cash.vwVatRecurrenceType"); return r["VatType"] if r else ""
    def get_vat_recurrence(self): return _query_all(self.conn_str, "SELECT * FROM Cash.vwVatRecurrence ORDER BY YearNumber, StartOn")
    def get_vat_recurrence_accruals(self): return _query_all(self.conn_str, "SELECT * FROM Cash.vwVatRecurrenceAccruals ORDER BY YearNumber, StartOn")
    def get_vat_period_totals(self): return _query_all(self.conn_str, "SELECT * FROM Cash.vwVatPeriodTotals ORDER BY YearNumber, StartOn")
    def get_vat_period_accruals(self): return _query_all(self.conn_str, "SELECT * FROM Cash.vwVatPeriodAccruals ORDER BY YearNumber, StartOn")
    def get_bank_accounts(self): return _query_all(self.conn_str, "SELECT AccountCode, AccountName FROM Cash.vwBankAccounts ORDER BY AccountOrder")
    def get_bank_balances(self, account_code: str): return _query_all(self.conn_str, "SELECT YearNumber, MonthNumber, Balance FROM Cash.vwBankBalances WHERE AccountCode = ? ORDER BY YearNumber, MonthNumber", (account_code,))
    def get_balance_sheet(self): return _query_all(self.conn_str, "SELECT AssetCode, AssetName, YearNumber, MonthNumber, Balance FROM Cash.vwBalanceSheet ORDER BY AssetOrder, YearNumber, MonthNumber")