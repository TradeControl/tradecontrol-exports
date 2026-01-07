from enum import IntEnum

class EventLogType(IntEnum):
    Error = 0
    Warning = 1
    Information = 2

class CashType(IntEnum):
    Trade = 0
    Tax = 1
    Money = 2

class ReportMode(IntEnum):
    CashFlow = 0
    Budget = 1

class CashPolarity(IntEnum):
    Expense = 0
    Income = 1
    Neutral = 2

class CategoryType(IntEnum):
    CashCode = 0
    Total = 1
    Expression = 2

class TaxType(IntEnum):
    CorporationTax = 0
    Vat = 1
    NI = 2
    General = 3

class SyntaxType(IntEnum):
    Both = 0
    Libre = 1
    Excel = 2