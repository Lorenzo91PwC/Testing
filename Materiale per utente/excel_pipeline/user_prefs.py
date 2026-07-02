"""Tiny persistent key-value store for UI preferences.

Backed by a single JSON file at the project root (``user_prefs.json``)
so values survive across app restarts. The file is gitignored, since
the values are per-machine (output paths, mostly).
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_PREFS_PATH = Path(__file__).parent.parent / "user_prefs.json"


def _read_all() -> dict[str, Any]:
    if not _PREFS_PATH.exists():
        return {}
    try:
        return json.loads(_PREFS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def _write_all(prefs: dict[str, Any]) -> None:
    try:
        _PREFS_PATH.write_text(
            json.dumps(prefs, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        # Silent: a missing-permissions failure should not crash the UI.
        pass


def load_pref(key: str, default: Any = None) -> Any:
    """Return the stored value for ``key`` or ``default`` if not set."""
    return _read_all().get(key, default)


def save_pref(key: str, value: Any) -> None:
    """Persist ``value`` under ``key``. No-op on filesystem errors."""
    prefs = _read_all()
    prefs[key] = value
    _write_all(prefs)
