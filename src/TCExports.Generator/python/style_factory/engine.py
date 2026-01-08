from __future__ import annotations
from pathlib import Path
from typing import Optional, Tuple, Union
import io
import zipfile

from .rendering.injector import inject_content_styles, apply_default_language_to_styles
from .rendering.ods_repack import repack_with_replacements

Locale = Tuple[str, str]

def apply_styles(
    ods_path: Union[str, Path],
    locale: Locale = ("en", "GB"),
    strip_defaults: bool = True,
) -> bytes:
    p = Path(ods_path)
    content = p.read_bytes()
    updated = apply_styles_bytes(content, locale=locale, strip_defaults=strip_defaults)
    p.write_bytes(updated)
    return updated

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

    new_content_xml = inject_content_styles(content_xml, strip_defaults=strip_defaults)
    new_styles_xml = apply_default_language_to_styles(styles_xml, lang=locale[0], country=locale[1])

    return repack_with_replacements(
        ods_bytes,
        replacements={
            "content.xml": new_content_xml,
            "styles.xml": new_styles_xml,
        },
    )