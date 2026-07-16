from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from .excel_io import load_workbook_with_warnings
from .models import SheetConfig
from .values import normalize_split_value


COMPLETE_COPY_VALUE = "完整表"


@dataclass(slots=True)
class SplitPlan:
    values: list[str]
    keys_by_value: dict[str, set[str]] = field(default_factory=dict)
    all_keys: set[str] = field(default_factory=set)
    empty_rows: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def build_split_plan(
    input_file: Path,
    sheet_configs: Iterable[SheetConfig],
    *,
    workbook=None,
) -> SplitPlan:
    configs = tuple(sheet_configs)
    for config in configs:
        config.validate()
    reference_count = sum(config.mode == "reference" for config in configs)
    if reference_count > 1:
        raise ValueError("一次任务只能选择一个基准 Sheet")
    if any(config.mode == "linked" for config in configs) and not reference_count:
        raise ValueError("按关联键匹配时必须选择一个基准 Sheet")

    should_close = False
    if workbook is None:
        workbook, warning_messages = load_workbook_with_warnings(
            input_file, data_only=True, read_only=True
        )
        should_close = True
    else:
        warning_messages = []
    try:
        for config in configs:
            _validate_sheet_config(workbook, config)

        reference = next(
            (config for config in configs if config.mode == "reference"), None
        )
        if reference is not None:
            return _reference_plan(workbook[reference.sheet_name], reference, warning_messages)

        direct_configs = [config for config in configs if config.mode == "direct"]
        if not direct_configs:
            return SplitPlan(values=[COMPLETE_COPY_VALUE], warnings=warning_messages)
        return _direct_plan(workbook, direct_configs, warning_messages)
    finally:
        if should_close:
            workbook.close()


def _validate_sheet_config(workbook, config: SheetConfig) -> None:
    if config.sheet_name not in workbook.sheetnames:
        raise ValueError(f"找不到 sheet：{config.sheet_name}")
    sheet = workbook[config.sheet_name]
    if config.header_row > sheet.max_row:
        raise ValueError(f"{config.sheet_name} 的表头行超出有效范围")
    for column_index, column_name in (
        (config.split_column_idx, "拆分列"),
        (config.key_column_idx, "关联键列"),
    ):
        if column_index is not None and column_index > sheet.max_column:
            raise ValueError(f"{config.sheet_name} 的{column_name}超出有效范围")


def _reference_plan(sheet, config: SheetConfig, warnings: list[str]) -> SplitPlan:
    values: list[str] = []
    keys_by_value: dict[str, set[str]] = {}
    key_to_value: dict[str, str] = {}
    empty_groups = 0

    for row_index in range(config.header_row + 1, sheet.max_row + 1):
        split_value = normalize_split_value(
            sheet.cell(row=row_index, column=config.split_column_idx).value
        )
        if split_value is None:
            empty_groups += 1
            continue
        if split_value not in keys_by_value:
            values.append(split_value)
            keys_by_value[split_value] = set()

        key = normalize_split_value(
            sheet.cell(row=row_index, column=config.key_column_idx).value
        )
        if key is None:
            continue
        previous = key_to_value.get(key)
        if previous is not None and previous != split_value:
            raise ValueError(
                f"关联键冲突：{key} 同时对应 {previous} 和 {split_value}"
            )
        key_to_value[key] = split_value
        keys_by_value[split_value].add(key)

    return SplitPlan(
        values=values,
        keys_by_value=keys_by_value,
        all_keys=set(key_to_value),
        empty_rows={config.sheet_name: empty_groups},
        warnings=warnings,
    )


def _direct_plan(workbook, configs: list[SheetConfig], warnings: list[str]) -> SplitPlan:
    values: list[str] = []
    seen: set[str] = set()
    empty_rows: dict[str, int] = {}
    for config in configs:
        sheet = workbook[config.sheet_name]
        empty_count = 0
        for row_index in range(config.header_row + 1, sheet.max_row + 1):
            value = normalize_split_value(
                sheet.cell(row=row_index, column=config.split_column_idx).value
            )
            if value is None:
                empty_count += 1
            elif value not in seen:
                seen.add(value)
                values.append(value)
        empty_rows[config.sheet_name] = empty_count
    return SplitPlan(values=values, empty_rows=empty_rows, warnings=warnings)
