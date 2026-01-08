from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, Iterable, Iterator, List, Optional, Tuple

from ..semantic.naming import StyleKey, parse_style_name, data_style_name_for_cell

@dataclass(frozen=True)
class DataStyleSpec:
    name: str
    kind: str       # "number", "percentage", "cash_neg"
    decimals: int

@dataclass(frozen=True)
class CellStyleSpec:
    name: str
    data_style_name: str
    neg_red: bool

class StyleRegistry:
    """
    Collects StyleKeys discovered in content.xml and produces the
    required DataStyleSpec and CellStyleSpec.
    """
    def __init__(self) -> None:
        self._keys: Dict[str, StyleKey] = {}

    def add_from_cell_style_name(self, name: Optional[str]) -> None:
        if not name:
            return
        key = parse_style_name(name)
        if key:
            self._keys[key.cell_style_name] = key

    def __len__(self) -> int:
        return len(self._keys)

    def build_specs(self) -> Tuple[List[DataStyleSpec], List[CellStyleSpec]]:
        data_specs: Dict[str, DataStyleSpec] = {}
        cell_specs: List[CellStyleSpec] = []

        for key in self._keys.values():
            ds_name = data_style_name_for_cell(key.cell_style_name)
            if not ds_name:
                continue

            if key.kind == "number":
                ds = DataStyleSpec(name=ds_name, kind="number", decimals=key.decimals)
                data_specs.setdefault(ds.name, ds)
                cell_specs.append(CellStyleSpec(name=key.cell_style_name, data_style_name=ds.name, neg_red=False))

            elif key.kind == "percentage":
                ds = DataStyleSpec(name=ds_name, kind="percentage", decimals=key.decimals)
                data_specs.setdefault(ds.name, ds)
                cell_specs.append(CellStyleSpec(name=key.cell_style_name, data_style_name=ds.name, neg_red=False))

            elif key.kind == "cash":
                if key.polarity == "pos":
                    ds = DataStyleSpec(name=ds_name, kind="number", decimals=key.decimals)
                    data_specs.setdefault(ds.name, ds)
                    cell_specs.append(CellStyleSpec(name=key.cell_style_name, data_style_name=ds.name, neg_red=False))
                elif key.polarity == "neg":
                    ds = DataStyleSpec(name=ds_name, kind="cash_neg", decimals=key.decimals)
                    data_specs.setdefault(ds.name, ds)
                    cell_specs.append(CellStyleSpec(name=key.cell_style_name, data_style_name=ds.name, neg_red=True))

        return list(data_specs.values()), cell_specs