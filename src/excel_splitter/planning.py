from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from .excel_io import load_workbook_with_warnings
from .models import SheetConfig
from .values import normalize_split_value


COMPLETE_COPY_VALUE = "完整表"

ProgressCallback = Callable[[int, str], None]


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
    progress_callback: ProgressCallback | None = None,
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
        return _direct_plan(workbook, direct_configs, warning_messages, progress_callback)
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


def _direct_plan(
    workbook,
    configs: list[SheetConfig],
    warnings: list[str],
    progress_callback: ProgressCallback | None = None,
) -> SplitPlan:
    values: list[str] = []
    seen: set[str] = set()
    empty_rows: dict[str, int] = {}
    MAX_EMPTY_ROWS = 10000  # 连续空行上限，超出则视为后续无数据
    for config in configs:
        sheet = workbook[config.sheet_name]
        total_rows = sheet.max_row - config.header_row

        if total_rows <= 0:
            empty_rows[config.sheet_name] = 0
            continue

        if total_rows > 50000:
            warnings.append(
                f"{config.sheet_name} 行数较多（{sheet.max_row} 行），"
                f"扫描可能需要较长时间，请耐心等待"
            )

        empty_count = 0
        consecutive_empty = 0
        stopped_early = False
        if progress_callback:
            progress_callback(0, f"正在分析 {config.sheet_name}...")

        for row_index, row in enumerate(
            sheet.iter_rows(
                min_row=config.header_row + 1,
                max_row=sheet.max_row,
                min_col=config.split_column_idx,
                max_col=config.split_column_idx,
                values_only=True,
            ),
            start=1,
        ):
            value = normalize_split_value(row[0])
            if value is None:
                empty_count += 1
                consecutive_empty += 1
                if consecutive_empty >= MAX_EMPTY_ROWS:
                    actual_end = config.header_row + row_index
                    stopped_early = True
                    break
            else:
                consecutive_empty = 0
                if value not in seen:
                    seen.add(value)
                    values.append(value)
            if progress_callback and row_index % 2000 == 0:
                progress_callback(
                    int(min(row_index / total_rows * 100, 99)),
                    f"正在分析 {config.sheet_name} ({row_index}/{total_rows})",
                )

        if stopped_early:
            warnings.append(
                f"{config.sheet_name} 在行 {actual_end} 之后连续 {MAX_EMPTY_ROWS} 行为空，"
                f"已提前结束扫描，后续空行将被忽略"
            )
        empty_rows[config.sheet_name] = empty_count
    return SplitPlan(values=values, empty_rows=empty_rows, warnings=warnings)
