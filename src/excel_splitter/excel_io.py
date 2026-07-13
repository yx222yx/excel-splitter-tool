from __future__ import annotations

from pathlib import Path
import warnings

from openpyxl import load_workbook


def load_workbook_with_warnings(path: Path, **kwargs):
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        workbook = load_workbook(path, **kwargs)
    messages = list(dict.fromkeys(str(item.message) for item in captured))
    return workbook, messages

