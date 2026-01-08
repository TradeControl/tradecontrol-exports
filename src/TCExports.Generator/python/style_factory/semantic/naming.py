from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

@dataclass(frozen=True)
class StyleKey:
    kind: str                 # "number", "percentage", "cash"
    decimals: int
    polarity: Optional[str]   # "pos" | "neg" | None
    cell_style_name: str      # original cell style name (canonicalized)

def parse_style_name(name: str) -> Optional[StyleKey]:
    """
    Parse semantic cell style names like:
      - NUM2_CELL
      - PCT1_CELL
      - CASH2_POS_CELL
      - CASH2_NEG_CELL
    Returns None for unknown or non-semantic names (e.g., TEXT_CELL).
    """
    if not name:
        return None
    u = name.strip().upper()

    if u == "TEXT_CELL":
        return None

    if u.startswith("NUM") and u.endswith("_CELL"):
        try:
            dp = int(u[3])
        except Exception:
            dp = 0
        return StyleKey(kind="number", decimals=dp, polarity=None, cell_style_name=u)

    if u.startswith("PCT") and u.endswith("_CELL"):
        try:
            dp = int(u[3])
        except Exception:
            dp = 0
        return StyleKey(kind="percentage", decimals=dp, polarity=None, cell_style_name=u)

    if u.startswith("CASH") and u.endswith("_CELL"):
        try:
            dp = int(u[4])
        except Exception:
            dp = 0
        if u.endswith("_POS_CELL"):
            return StyleKey(kind="cash", decimals=dp, polarity="pos", cell_style_name=u)
        if u.endswith("_NEG_CELL"):
            return StyleKey(kind="cash", decimals=dp, polarity="neg", cell_style_name=u)

    return None

def data_style_name_for_cell(cell_style_name: str) -> Optional[str]:
    """
    Map a cell style name to its data style name per convention:
      *_CELL -> *_DS
      *_POS_CELL -> *_POS_DS
      *_NEG_CELL -> *_NEG_DS
    """
    u = cell_style_name.strip().upper()
    if u.endswith("_POS_CELL"):
        return u.replace("_POS_CELL", "_POS_DS")
    if u.endswith("_NEG_CELL"):
        return u.replace("_NEG_CELL", "_NEG_DS")
    if u.endswith("_CELL"):
        return u.replace("_CELL", "_DS")
    return None