from __future__ import annotations
from typing import Optional
from lxml import etree as ET

from .xml_utils import (
    OFFICE_NS, STYLE_NS, NUMBER_NS, TABLE_NS, TEXT_NS, FO_NS, q
)
from ..mapping.registry import StyleRegistry, DataStyleSpec, CellStyleSpec

def inject_content_styles(content_xml: bytes, strip_defaults: bool = True) -> bytes:
    parser = ET.XMLParser(remove_blank_text=False)
    root = ET.fromstring(content_xml, parser=parser)

    auto = root.find(q(OFFICE_NS, "automatic-styles"))
    if auto is None:
        auto = ET.Element(q(OFFICE_NS, "automatic-styles"))
        root.insert(0, auto) if len(root) else root.append(auto)

    if strip_defaults:
        body = root.find(q(OFFICE_NS, "body"))
        if body is not None:
            ss = body.find(q(OFFICE_NS, "spreadsheet"))
            if ss is not None:
                for tbl in list(ss.findall(q(TABLE_NS, "table"))):
                    if tbl.get(q(TABLE_NS, "name")) == "Feuille1":
                        ss.remove(tbl)

    reg = StyleRegistry()
    for cell in root.findall(f".//{q(TABLE_NS, 'table-cell')}"):
        sname = cell.get(q(TABLE_NS, "style-name"))
        reg.add_from_cell_style_name(sname)
        if cell.find(q(TEXT_NS, "p")) is None:
            ET.SubElement(cell, q(TEXT_NS, "p"))

    data_specs, cell_specs = reg.build_specs()

    def find_style(tag_ns: str, tag_local: str, name_attr_ns: str, name: str):
        for el in auto.findall(q(tag_ns, tag_local)):
            if el.get(q(name_attr_ns, "name")) == name:
                return el
        return None

    def ensure_number_ds(spec: DataStyleSpec):
        ds = find_style(NUMBER_NS, "number-style", STYLE_NS, spec.name)
        if ds is None:
            ds = ET.SubElement(auto, q(NUMBER_NS, "number-style"))
            ds.set(q(STYLE_NS, "name"), spec.name)
            num = ET.SubElement(ds, q(NUMBER_NS, "number"))
            num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            num.set(q(NUMBER_NS, "grouping"), "true")
        else:
            num = ds.find(q(NUMBER_NS, "number")) or ET.SubElement(ds, q(NUMBER_NS, "number"))
            if num.get(q(NUMBER_NS, "decimal-places")) is None:
                num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-decimal-places")) is None:
                num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-integer-digits")) is None:
                num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            if num.get(q(NUMBER_NS, "grouping")) is None:
                num.set(q(NUMBER_NS, "grouping"), "true")
        return ds

    def ensure_percent_ds(spec: DataStyleSpec):
        ds = find_style(NUMBER_NS, "percentage-style", STYLE_NS, spec.name)
        if ds is None:
            ds = ET.SubElement(auto, q(NUMBER_NS, "percentage-style"))
            ds.set(q(STYLE_NS, "name"), spec.name)
            num = ET.SubElement(ds, q(NUMBER_NS, "number"))
            num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            ET.SubElement(ds, q(NUMBER_NS, "text")).text = "%"
        else:
            num = ds.find(q(NUMBER_NS, "number")) or ET.SubElement(ds, q(NUMBER_NS, "number"))
            if num.get(q(NUMBER_NS, "decimal-places")) is None:
                num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-decimal-places")) is None:
                num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-integer-digits")) is None:
                num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            if ds.find(q(NUMBER_NS, "text")) is None:
                ET.SubElement(ds, q(NUMBER_NS, "text")).text = "%"
        return ds

    def ensure_cash_neg_ds(spec: DataStyleSpec):
        ds = find_style(NUMBER_NS, "number-style", STYLE_NS, spec.name)
        if ds is None:
            ds = ET.SubElement(auto, q(NUMBER_NS, "number-style"))
            ds.set(q(STYLE_NS, "name"), spec.name)
            ET.SubElement(ds, q(NUMBER_NS, "text")).text = "("
            num = ET.SubElement(ds, q(NUMBER_NS, "number"))
            num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            num.set(q(NUMBER_NS, "grouping"), "true")
            num.set(q(NUMBER_NS, "display-factor"), "-1")
            ET.SubElement(ds, q(NUMBER_NS, "text")).text = ")"
        else:
            num = ds.find(q(NUMBER_NS, "number")) or ET.SubElement(ds, q(NUMBER_NS, "number"))
            if num.get(q(NUMBER_NS, "decimal-places")) is None:
                num.set(q(NUMBER_NS, "decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-decimal-places")) is None:
                num.set(q(NUMBER_NS, "min-decimal-places"), str(spec.decimals))
            if num.get(q(NUMBER_NS, "min-integer-digits")) is None:
                num.set(q(NUMBER_NS, "min-integer-digits"), "1")
            if num.get(q(NUMBER_NS, "grouping")) is None:
                num.set(q(NUMBER_NS, "grouping"), "true")
            num.set(q(NUMBER_NS, "display-factor"), "-1")
            has_open = any(e.tag == q(NUMBER_NS, "text") and (e.text or "") == "(" for e in ds)
            has_close = any(e.tag == q(NUMBER_NS, "text") and (e.text or "") == ")" for e in ds)
            if not has_open:
                ds.insert(0, ET.Element(q(NUMBER_NS, "text"))); ds[0].text = "("
            if not has_close:
                ds.append(ET.Element(q(NUMBER_NS, "text"))); ds[-1].text = ")"
        return ds

    def ensure_cell_style(spec: CellStyleSpec):
        cs = find_style(STYLE_NS, "style", STYLE_NS, spec.name)
        if cs is None:
            cs = ET.SubElement(auto, q(STYLE_NS, "style"))
            cs.set(q(STYLE_NS, "name"), spec.name)
            cs.set(q(STYLE_NS, "family"), "table-cell")
            cs.set(q(STYLE_NS, "parent-style-name"), "Default")
            ET.SubElement(cs, q(STYLE_NS, "table-cell-properties"))
        cs.set(q(STYLE_NS, "data-style-name"), spec.data_style_name)
        if spec.neg_red:
            tp = cs.find(q(STYLE_NS, "text-properties")) or ET.SubElement(cs, q(STYLE_NS, "text-properties"))
            tp.set(q(FO_NS, "color"), "#FF0000")
        return cs

    def ensure_cash_base_cell_style_with_maps(base_name: str, pos_cell_name: str, neg_cell_name: str):
        """
        Create a base cell style with style:map entries to route values to POS/NEG cell styles.
        Calc honors this for accounting conventions and avoids rendering a leading minus.
        """
        base = find_style(STYLE_NS, "style", STYLE_NS, base_name)
        if base is None:
            base = ET.SubElement(auto, q(STYLE_NS, "style"))
            base.set(q(STYLE_NS, "name"), base_name)
            base.set(q(STYLE_NS, "family"), "table-cell")
            base.set(q(STYLE_NS, "parent-style-name"), "Default")
            ET.SubElement(base, q(STYLE_NS, "table-cell-properties"))
        # Remove any existing maps to avoid duplicates
        for m in list(base.findall(q(STYLE_NS, "map"))):
            base.remove(m)
        # Map negative values to NEG cell style
        map_neg = ET.SubElement(base, q(STYLE_NS, "map"))
        map_neg.set(q(STYLE_NS, "condition"), "value() < 0")
        map_neg.set(q(STYLE_NS, "apply-style-name"), neg_cell_name)
        # Map non-negative values to POS cell style
        map_pos = ET.SubElement(base, q(STYLE_NS, "map"))
        map_pos.set(q(STYLE_NS, "condition"), "value() >= 0")
        map_pos.set(q(STYLE_NS, "apply-style-name"), pos_cell_name)
        return base

    # Ensure all required data-styles
    for ds in data_specs:
        if ds.kind == "number":
            ensure_number_ds(ds)
        elif ds.kind == "percentage":
            ensure_percent_ds(ds)
        elif ds.kind == "cash_neg":
            ensure_cash_neg_ds(ds)

    # Ensure all required cell-styles
    for cs in cell_specs:
        ensure_cell_style(cs)

    # Create base CASHx cell styles with conditional maps when both POS/NEG exist
    # Infer base names from POS/NEG cell style names: CASH2_POS_CELL -> CASH2_CELL
    pos_cells = {c.name for c in cell_specs if c.name.endswith("_POS_CELL")}
    neg_cells = {c.name for c in cell_specs if c.name.endswith("_NEG_CELL")}
    for pos in pos_cells:
        base = pos.replace("_POS_CELL", "_CELL")
        neg = pos.replace("_POS_CELL", "_NEG_CELL")
        if neg in neg_cells:
            ensure_cash_base_cell_style_with_maps(base, pos, neg)

    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")

def apply_default_language_to_styles(
    styles_xml: Optional[bytes],
    lang: str = "en",
    country: str = "GB",
) -> bytes:
    if styles_xml is None:
        s_root = ET.Element(q(OFFICE_NS, "document-styles"))
        ET.SubElement(s_root, q(OFFICE_NS, "styles"))
    else:
        s_root = ET.fromstring(styles_xml)

    office_styles = s_root.find(q(OFFICE_NS, "styles")) or ET.SubElement(s_root, q(OFFICE_NS, "styles"))

    def ensure_default_style(family: str):
        for ds in office_styles.findall(q(STYLE_NS, "default-style")):
            if ds.get(q(STYLE_NS, "family")) == family:
                return ds
        ds = ET.SubElement(office_styles, q(STYLE_NS, "default-style"))
        ds.set(q(STYLE_NS, "family"), family)
        return ds

    def ensure_text_props(parent):
        tp = parent.find(q(STYLE_NS, "text-properties")) or ET.SubElement(parent, q(STYLE_NS, "text-properties"))
        tp.set(q(FO_NS, "language"), lang)
        tp.set(q(FO_NS, "country"), country)
        tp.set(q(STYLE_NS, "language-asian"), lang)
        tp.set(q(STYLE_NS, "country-asian"), country)
        tp.set(q(STYLE_NS, "language-complex"), lang)
        tp.set(q(STYLE_NS, "country-complex"), country)
        return tp

    for family in ("paragraph", "text", "table-cell"):
        ds = ensure_default_style(family)
        ensure_text_props(ds)

    return ET.tostring(s_root, xml_declaration=True, encoding="UTF-8")