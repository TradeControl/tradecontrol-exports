import io
import zipfile
from typing import Dict

def repack_with_replacements(ods_bytes: bytes, replacements: Dict[str, bytes]) -> bytes:
    in_mem = io.BytesIO(ods_bytes)
    out_mem = io.BytesIO()

    with zipfile.ZipFile(in_mem, "r") as zin, zipfile.ZipFile(out_mem, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        existing = {i.filename for i in zin.infolist()}
        for item in zin.infolist():
            name = item.filename
            if name in replacements:
                zout.writestr(name, replacements[name])
            else:
                zout.writestr(name, zin.read(name))
        for name, data in replacements.items():
            if name not in existing:
                zout.writestr(name, data)

    return out_mem.getvalue()