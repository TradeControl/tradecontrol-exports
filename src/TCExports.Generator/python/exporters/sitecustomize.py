from pathlib import Path
import sys

# Ensure parent 'python' folder is on sys.path so 'style_factory', 'data', 'i18n' resolve
_here = Path(__file__).resolve().parent        # .../python/exporters
_python_root = _here.parent                    # .../python
if str(_python_root) not in sys.path:
    sys.path.insert(0, str(_python_root))