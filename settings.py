"""Persistent application preferences."""

import json
from pathlib import Path


SETTINGS_FILE = Path(__file__).with_name("settings.json")
DEFAULTS = {"language": "ru", "currency": "UZS"}


def load_settings():
    """Load preferences, falling back to defaults if the file is missing or invalid."""
    try:
        with SETTINGS_FILE.open("r", encoding="utf-8") as file:
            saved = json.load(file)
    except (OSError, json.JSONDecodeError):
        saved = {}

    result = DEFAULTS.copy()
    if isinstance(saved, dict):
        result.update(saved)
    if result["language"] not in ("ru", "en_US"):
        result["language"] = DEFAULTS["language"]
    if result["currency"] not in ("UZS", "USD"):
        result["currency"] = DEFAULTS["currency"]
    return result


def save_settings(settings):
    """Save preferences as UTF-8 JSON."""
    with SETTINGS_FILE.open("w", encoding="utf-8") as file:
        json.dump(settings, file, ensure_ascii=False, indent=2)

