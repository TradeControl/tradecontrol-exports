# Stub: Postgres repository not yet supported.
# This module provides a placeholder implementation to satisfy imports
# without requiring psycopg2 or any Postgres client.

from typing import Any, Dict, List, Optional

class PostgresRepository:
    def __init__(self, conn_str: str):
        self.conn_str = conn_str

    def _not_supported(self) -> None:
        raise NotImplementedError("PostgresRepository is not implemented. Please configure dbKind='sqlserver' or provide a Postgres adapter.")

    # Core periods/company
    def get_active_period(self) -> Optional[Dict[str, Any]]: self._not_supported()
    def get_active_years(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_months(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_company_name(self) -> str: self._not_supported()

    # Categories/codes/values
    def get_categories(self, cash_type: str) -> List[Dict[str, Any]]: self._not_supported()
    def get_cash_codes(self, category_code: str) -> List[Dict[str, Any]]: self._not_supported()
    def get_cash_code_values(self, cash_code: str, year_number: int,
                             include_active: bool, include_orderbook: bool, include_tax_accruals: bool
                             ) -> List[Dict[str, Any]]: self._not_supported()

    # Totals/expressions
    def get_category_totals(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_category_total_codes(self, category_code: str) -> List[Dict[str, Any]]: self._not_supported()
    def get_category_expressions(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_category_code_from_name(self, name: str) -> Optional[str]: self._not_supported()

    # VAT
    def get_vat_recurrence_type(self) -> str: self._not_supported()
    def get_vat_recurrence(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_vat_recurrence_accruals(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_vat_period_totals(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_vat_period_accruals(self) -> List[Dict[str, Any]]: self._not_supported()

    # Bank
    def get_bank_accounts(self) -> List[Dict[str, Any]]: self._not_supported()
    def get_bank_balances(self, account_code: str) -> List[Dict[str, Any]]: self._not_supported()

    # Balance sheet
    def get_balance_sheet(self) -> List[Dict[str, Any]]: self._not_supported()