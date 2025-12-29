import json
from pathlib import Path
from typing import Dict

class ResourceManager:
    """Simple locale resource loader: resources are JSON files named <locale>.json."""
    def __init__(self, locale: str = "en-GB", base_dir: Path | None = None):
        self.base_dir = base_dir or Path(__file__).parent / "locales"
        self._cache: Dict[str, str] = {}
        self.set_locale(locale)

    def set_locale(self, locale: str) -> None:
        path = self.base_dir / f"{locale}.json"
        if not path.exists():
            path = self.base_dir / "en-GB.json"
        if path.exists():
            # Read BOM-tolerant
            text = path.read_text(encoding="utf-8-sig")
            try:
                self._cache = json.loads(text)
            except json.JSONDecodeError:
                # Fallback: strip leading BOM if present and retry
                self._cache = json.loads(text.lstrip("\ufeff"))
        else:
            self._cache = {}

    def t(self, key: str) -> str:
        """Translate a key, falling back to the key name if missing."""
        return self._cache.get(key, key)