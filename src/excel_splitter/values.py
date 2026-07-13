from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

from .excel_io import load_workbook_with_warnings
from .models import SheetConfig


def normalize_split_value(value: Any) -> str | None:
    if value is None:
        return None
    normalized = str(value).strip()
    return normalized or None


def extract_split_values(
    input_file: Path, sheet_configs: Iterable[SheetConfig]
) -> tuple[list[str], dict[str, int], list[str]]:
    workbook, warning_messages = load_workbook_with_warnings(
        input_file, data_only=True, read_only=True
    )
    values: list[str] = []
    seen: set[str] = set()
    empty_rows: dict[str, int] = {}
    try:
        for config in sheet_configs:
            if config.sheet_name not in workbook.sheetnames:
                raise ValueError(f"找不到 sheet：{config.sheet_name}")
            sheet = workbook[config.sheet_name]
            if config.header_row > sheet.max_row:
                raise ValueError(f"{config.sheet_name} 的表头行超出有效范围")
            if config.split_column_idx > sheet.max_column:
                raise ValueError(f"{config.sheet_name} 的拆分列超出有效范围")

            empty_count = 0
            for row in sheet.iter_rows(
                min_row=config.header_row + 1,
                min_col=config.split_column_idx,
                max_col=config.split_column_idx,
            ):
                normalized = normalize_split_value(row[0].value)
                if normalized is None:
                    empty_count += 1
                elif normalized not in seen:
                    seen.add(normalized)
                    values.append(normalized)
            empty_rows[config.sheet_name] = empty_count
    finally:
        workbook.close()
    return values, empty_rows, warning_messages
