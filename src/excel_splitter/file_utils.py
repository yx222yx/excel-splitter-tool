from __future__ import annotations

from pathlib import Path
import re


_INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}


def sanitize_filename(value: str) -> str:
    sanitized = _INVALID_WINDOWS_CHARS.sub("_", value).strip().rstrip(". ")
    if not sanitized:
        sanitized = "未命名"
    if sanitized.upper() in _RESERVED_NAMES:
        sanitized = f"_{sanitized}"
    return sanitized


def render_filename(
    template: str,
    *,
    original_name: str,
    split_value: str,
    output_type: str | None = None,
) -> str:
    rendered = template.format(
        original_name=sanitize_filename(original_name),
        split_value=sanitize_filename(split_value),
        output_type=sanitize_filename(output_type or ""),
    )
    rendered = sanitize_filename(rendered)
    if output_type and "{output_type" not in template:
        if rendered.lower().endswith(".xlsx"):
            rendered = f"{rendered[:-5]}_{sanitize_filename(output_type)}.xlsx"
        else:
            rendered = f"{rendered}_{sanitize_filename(output_type)}"
    if not rendered.lower().endswith(".xlsx"):
        rendered += ".xlsx"
    return rendered


def resolve_output_path(path: Path, *, overwrite: bool) -> Path:
    if overwrite or not path.exists():
        return path
    index = 1
    while True:
        candidate = path.with_name(f"{path.stem}({index}){path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1
