from __future__ import annotations
from pathlib import Path
from typing import Optional, Tuple, Union
import io
import zipfile

from .rendering.injector import inject_content_styles, apply_default_language_to_styles
from .rendering.ods_repack import repack_with_replacements
from lxml import etree as ET
from .rendering.xml_utils import OFFICE_NS, q

Locale = Tuple[str, str]

def _apply_meta_locale(meta_xml: Optional[bytes], lang: str, country: str) -> bytes:
    if meta_xml is None:
        m_root = ET.Element(q(OFFICE_NS, "document-meta"))
        ET.SubElement(m_root, q(OFFICE_NS, "meta"))
    else:
        m_root = ET.fromstring(meta_xml)
    meta = m_root.find(q(OFFICE_NS, "meta"))
    if meta is None:
        meta = ET.SubElement(m_root, q(OFFICE_NS, "meta"))
    # Rewrite dc:language (QName with dc ns might not be bound here; use literal)
    dc_lang = None
    for child in meta:
        if child.tag.endswith("language"):
            dc_lang = child
            break
    if dc_lang is None:
        dc_lang = ET.SubElement(meta, ET.QName("http://purl.org/dc/elements/1.1/", "language"))
    dc_lang.text = f"{lang}-{country}"
    return ET.tostring(m_root, xml_declaration=True, encoding="UTF-8")

def apply_styles_bytes(
    ods_bytes: bytes,
    locale: Locale = ("en", "GB"),
    strip_defaults: bool = True,
) -> bytes:
    with zipfile.ZipFile(io.BytesIO(ods_bytes), "r") as zin:
        content_xml = zin.read("content.xml")
        try:
            styles_xml = zin.read("styles.xml")
        except KeyError:
            styles_xml = None
        try:
            meta_xml = zin.read("meta.xml")
        except KeyError:
            meta_xml = None

    lang, country = locale
    new_content_xml = inject_content_styles(content_xml, strip_defaults=strip_defaults, lang=lang, country=country)
    new_styles_xml = apply_default_language_to_styles(styles_xml, lang=lang, country=country)
    new_meta_xml = _apply_meta_locale(meta_xml, lang=lang, country=country)

    return repack_with_replacements(
        ods_bytes,
        replacements={
            "content.xml": new_content_xml,
            "styles.xml": new_styles_xml,
            "meta.xml": new_meta_xml,
        },
    )