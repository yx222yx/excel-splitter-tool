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
    # 每个 sheet 的表区结构：[(块表头行, 块数据首行, 块数据末行), ...]，块 1 为主表
    blocks: dict[str, list[tuple[int, int, int]]] = field(default_factory=dict)
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

        blocks = {
            config.sheet_name: _detect_blocks(workbook[config.sheet_name], config.header_row)
            for config in configs
        }

        reference = next(
            (config for config in configs if config.mode == "reference"), None
        )
        if reference is not None:
            plan = _reference_plan(workbook[reference.sheet_name], reference, warning_messages)
            plan.blocks = blocks
            return plan

        direct_configs = [config for config in configs if config.mode == "direct"]
        if not direct_configs:
            return SplitPlan(values=[COMPLETE_COPY_VALUE], blocks=blocks, warnings=warning_messages)
        plan = _direct_plan(workbook, direct_configs, warning_messages, progress_callback)
        plan.blocks = blocks
        return plan
    finally:
        if should_close:
            workbook.close()


def _detect_blocks(sheet, header_row: int) -> list[tuple[int, int, int]]:
    """把表头行之下的数据区按「整行全空」切成若干块。

    块 1 = 主表（表头行 + 其下连续非空行）；每个空行之后若还有非空行则开始
    新块，块首行视为块表头。返回 [(块表头行, 块数据首行, 块数据末行), ...]，
    单表区 sheet 只有块 1。检测用读到的值（调用方保证 data_only 口径一致）。
    """
    max_row = sheet.max_row or header_row
    segments: list[tuple[int, int]] = []
    segment_start: int | None = None
    last_non_empty = header_row
    consecutive_empty = 0
    EMPTY_GAP_STOP = 10000  # 连续空行上限，超出视为后续无数据
    for row_index, row in enumerate(
        sheet.iter_rows(min_row=header_row + 1, max_row=max_row, values_only=True),
        start=header_row + 1,
    ):
        if all(normalize_split_value(value) is None for value in row):
            consecutive_empty += 1
            if consecutive_empty >= EMPTY_GAP_STOP:
                break
            if segment_start is not None:
                segments.append((segment_start, last_non_empty))
                segment_start = None
        else:
            consecutive_empty = 0
            if segment_start is None:
                segment_start = row_index
            last_non_empty = row_index
    if segment_start is not None:
        segments.append((segment_start, last_non_empty))
    if not segments:
        return [(header_row, header_row + 1, header_row)]
    blocks = [(header_row, header_row + 1, segments[0][1])]
    for start, end in segments[1:]:
        blocks.append((start, start + 1, end))
    return blocks


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
