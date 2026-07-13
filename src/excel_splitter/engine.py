from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable

from .excel_io import load_workbook_with_warnings
from .file_utils import render_filename, resolve_output_path
from .models import OutputArtifact, SheetConfig, SplitJob, SplitResult, SplitSummary
from .planning import SplitPlan, build_split_plan
from .values import normalize_split_value


ProgressCallback = Callable[[int, str], None]


class SplitEngine:
    def execute(
        self,
        job: SplitJob,
        progress_callback: ProgressCallback | None = None,
    ) -> SplitSummary:
        progress = _ProgressReporter(progress_callback)
        progress.update(0, "正在分析工作簿")
        job.validate()
        plan = build_split_plan(job.input_file, job.sheet_configs)
        target_values = self._target_values(job, plan.values)
        if not target_values:
            raise ValueError("没有可用于拆分的非空值")
        progress.update(5, f"已识别 {len(target_values)} 个拆分值")

        job.output_dir.mkdir(parents=True, exist_ok=True)
        results: list[SplitResult] = []
        errors: list[str] = []
        all_warnings = list(plan.warnings)

        for value_index, split_value in enumerate(target_values):
            range_start = 5 + (value_index * 93 / len(target_values))
            range_end = 5 + ((value_index + 1) * 93 / len(target_values))

            def report_value(fraction: float, message: str) -> None:
                absolute = range_start + (range_end - range_start) * fraction
                progress.update(round(absolute), message)

            try:
                result = self._export_value(job, split_value, plan, report_value)
            except Exception as exc:  # Continue other values and report the failed one.
                errors.append(f"{split_value}: {exc}")
                continue
            results.append(result)
            all_warnings.extend(result.warnings)

        unique_warnings = tuple(dict.fromkeys(all_warnings))
        summary = SplitSummary(
            results=results,
            total_files=sum(len(result.output_files) for result in results),
            total_discarded=sum(
                sum(result.discarded_empty_rows.values()) for result in results
            ),
            total_unmatched=sum(
                sum(result.unmatched_key_rows.values()) for result in results
            ),
            warnings=list(unique_warnings),
            errors=errors,
        )
        progress.update(100, "拆分完成")
        return summary

    def _target_values(self, job: SplitJob, available_values: list[str]) -> list[str]:
        if job.split_mode == "all":
            return available_values
        selected = _unique_normalized(job.selected_split_values)
        unavailable = [value for value in selected if value not in available_values]
        if unavailable:
            raise ValueError(f"所选拆分值不存在：{', '.join(unavailable)}")
        return selected

    def _export_value(
        self,
        job: SplitJob,
        split_value: str,
        plan: SplitPlan,
        progress: Callable[[float, str], None],
    ) -> SplitResult:
        progress(0.02, f"正在加载 {split_value}")
        formula_workbook = None
        warning_messages: list[str] = []
        if "formula" in job.output_types:
            formula_workbook, formula_warnings = load_workbook_with_warnings(
                job.input_file, data_only=False
            )
            warning_messages.extend(formula_warnings)
        value_workbook = None
        try:
            value_workbook, value_warnings = load_workbook_with_warnings(
                job.input_file, data_only=True
            )
            warning_messages = list(
                dict.fromkeys([*warning_messages, *value_warnings])
            )
            selected = set(job.selected_sheets)
            if formula_workbook is not None:
                for sheet_name in list(formula_workbook.sheetnames):
                    if sheet_name not in selected:
                        formula_workbook.remove(formula_workbook[sheet_name])
            for sheet_name in list(value_workbook.sheetnames):
                if sheet_name not in selected:
                    value_workbook.remove(value_workbook[sheet_name])

            sheet_rows: dict[str, int] = {}
            discarded_empty_rows: dict[str, int] = {}
            unmatched_key_rows: dict[str, int] = {}
            config_count = len(job.sheet_configs)
            for config_index, config in enumerate(job.sheet_configs):
                sheet = (
                    formula_workbook[config.sheet_name]
                    if formula_workbook is not None
                    else None
                )
                value_sheet = value_workbook[config.sheet_name]
                filter_start = 0.12 + (config_index * 0.58 / config_count)
                filter_end = 0.12 + ((config_index + 1) * 0.58 / config_count)

                def report_rows(current: int, total: int) -> None:
                    ratio = current / total if total else 1
                    progress(
                        filter_start + (filter_end - filter_start) * ratio,
                        f"正在筛选 {split_value} / {config.sheet_name} ({current}/{total})",
                    )

                if config.mode == "full":
                    kept = max(0, value_sheet.max_row - config.header_row)
                    discarded = 0
                    unmatched = 0
                    report_rows(kept, kept)
                elif config.mode == "linked":
                    kept, discarded, unmatched = _filter_linked_sheet(
                        sheet,
                        value_sheet,
                        config,
                        plan.keys_by_value.get(split_value, set()),
                        plan.all_keys,
                        report_rows,
                    )
                else:
                    kept, discarded = _filter_sheet(
                        sheet, value_sheet, config, split_value, report_rows
                    )
                    unmatched = 0
                sheet_rows[config.sheet_name] = kept
                discarded_empty_rows[config.sheet_name] = discarded
                unmatched_key_rows[config.sheet_name] = unmatched

            available_outputs = {
                "formula": ("公式版", formula_workbook),
                "values": ("结果值版", value_workbook),
            }
            output_files: list[OutputArtifact] = []
            for output_index, output_type in enumerate(job.output_types):
                output_label, output_workbook = available_outputs[output_type]
                progress(
                    0.75 + output_index * 0.23 / len(job.output_types),
                    f"正在保存 {split_value} / {output_label}",
                )
                filename = render_filename(
                    job.filename_template,
                    original_name=job.original_name or job.input_file.stem,
                    split_value=split_value,
                    output_type=output_label,
                )
                output_file = resolve_output_path(
                    job.output_dir / filename, overwrite=job.overwrite
                )
                output_workbook.save(output_file)
                output_files.append(
                    OutputArtifact(
                        output_type=output_type,
                        output_file=output_file,
                    )
                )
                progress(
                    0.75 + (output_index + 1) * 0.23 / len(job.output_types),
                    f"已保存 {split_value} / {output_label}",
                )
        finally:
            if formula_workbook is not None:
                formula_workbook.close()
            if value_workbook is not None:
                value_workbook.close()

        return SplitResult(
            split_value=split_value,
            output_files=output_files,
            sheet_rows=sheet_rows,
            discarded_empty_rows=discarded_empty_rows,
            unmatched_key_rows=unmatched_key_rows,
            warnings=warning_messages,
        )


def _filter_sheet(
    sheet,
    value_sheet,
    config: SheetConfig,
    split_value: str,
    progress: Callable[[int, int], None] | None = None,
) -> tuple[int, int]:
    kept = 0
    discarded_empty = 0
    total = max(0, value_sheet.max_row - config.header_row)
    for processed, row_index in enumerate(
        range(value_sheet.max_row, config.header_row, -1), start=1
    ):
        normalized = normalize_split_value(
            value_sheet.cell(row=row_index, column=config.split_column_idx).value
        )
        if normalized is None:
            discarded_empty += 1
            if sheet is not None:
                sheet.delete_rows(row_index, 1)
            value_sheet.delete_rows(row_index, 1)
        elif normalized != split_value:
            if sheet is not None:
                sheet.delete_rows(row_index, 1)
            value_sheet.delete_rows(row_index, 1)
        else:
            kept += 1
        if progress is not None and (processed % 25 == 0 or processed == total):
            progress(processed, total)
    return kept, discarded_empty


def _filter_linked_sheet(
    sheet,
    value_sheet,
    config: SheetConfig,
    target_keys: set[str],
    all_keys: set[str],
    progress: Callable[[int, int], None] | None = None,
) -> tuple[int, int, int]:
    kept = 0
    discarded_empty = 0
    unmatched = 0
    total = max(0, value_sheet.max_row - config.header_row)
    for processed, row_index in enumerate(
        range(value_sheet.max_row, config.header_row, -1), start=1
    ):
        key = normalize_split_value(
            value_sheet.cell(row=row_index, column=config.key_column_idx).value
        )
        if key is None:
            discarded_empty += 1
            if sheet is not None:
                sheet.delete_rows(row_index, 1)
            value_sheet.delete_rows(row_index, 1)
        elif key in target_keys:
            kept += 1
        else:
            if key not in all_keys:
                unmatched += 1
            if sheet is not None:
                sheet.delete_rows(row_index, 1)
            value_sheet.delete_rows(row_index, 1)
        if progress is not None and (processed % 25 == 0 or processed == total):
            progress(processed, total)
    return kept, discarded_empty, unmatched


class _ProgressReporter:
    def __init__(self, callback: ProgressCallback | None) -> None:
        self.callback = callback
        self.last_percent = -1

    def update(self, percent: int, message: str) -> None:
        normalized = max(self.last_percent, min(100, max(0, percent)))
        if self.callback is not None and (
            normalized != self.last_percent or normalized in (0, 100)
        ):
            self.callback(normalized, message)
        self.last_percent = normalized


def _unique_normalized(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        normalized = normalize_split_value(value)
        if normalized is not None and normalized not in seen:
            seen.add(normalized)
            result.append(normalized)
    return result
