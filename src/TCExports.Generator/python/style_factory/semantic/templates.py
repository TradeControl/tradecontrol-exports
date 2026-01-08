from __future__ import annotations
from typing import Tuple

def cell_style_from_template_code(template_code: str, negative: bool = False) -> str:
    """
    Convert a TemplateCode (e.g., 'Num2', 'Pct1', 'Cash2') into a canonical cell style name.
    Use negative=True for explicit negative cash cells when needed.
    """
    u = (template_code or "").strip().upper()
    if u.startswith("NUM"):
        return f"{u}_CELL"
    if u.startswith("PCT"):
        return f"{u}_CELL"
    if u.startswith("CASH"):
        return f"{u}_{'NEG' if negative else 'POS'}_CELL"
    if u == "TEXT":
        return "TEXT_CELL"
    # Default passthrough
    return f"{u}_CELL"

def cash_pair_from_template_code(template_code: str) -> Tuple[str, str]:
    """
    For a 'CashX' template, returns (POS_CELL_NAME, NEG_CELL_NAME)
    """
    u = (template_code or "").strip().upper()
    if not u.startswith("CASH"):
        raise ValueError("TemplateCode is not cash")
    return (f"{u}_POS_CELL", f"{u}_NEG_CELL")