from __future__ import annotations

from typing import Any


def normalize_split_value(value: Any) -> str | None:
    if value is None:
        return None
    normalized = str(value).strip()
    return normalized or None
