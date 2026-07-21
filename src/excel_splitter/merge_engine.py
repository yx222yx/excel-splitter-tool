from __future__ import annotations

import hashlib
from copy import copy
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from .excel_io import load_workbook_with_warnings
from .merge_models import MergeJob, MergeSheetResult, MergeSummary
from .merge_planning import MergePlan, MergeSheetPlan, build_merge_plan
from .values import normalize_split_value


EMPTY_ROW_STOP_THRESHOLD = 10000
ProgressCallback = Callable[[int, str], None]


class MergeEngine:
    def execute(
        self,
        job: MergeJob,
        progress_callback: ProgressCallback | None = None,
        plan: MergePlan | None = None,
    ) -> MergeSummary:
        progress = _ProgressReporter(progress_callback)
        progress.update(0, "正在分析输入文件")
        job.validate()

        if plan is None:
            plan = build_merge_plan(job.input_files, job.sheet_configs)
        plans_by_sheet = {sheet_plan.sheet_name: sheet_plan for sheet_plan in plan.sheets}
        progress.update(8, "开始合并")

        output_workbook = Workbook(write_only=True)
        output_sheets = {
            config.sheet_name: output_workbook.create_sheet(title=config.sheet_name)
            for config in job.sheet_configs
        }
        # 已写入标题区域与表头的 sheet（由第一个包含该 sheet 的文件负责写入）
        initialized_sheets: set[str] = set()

        job.output_file.parent.mkdir(parents=True, exist_ok=True)

        source_rows: dict[str, dict[str, int]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }
        # 每个 sheet 已见过的内容指纹（指纹 -> 首个文件名），用于跳过完全重复的 sheet
        fingerprints: dict[str, dict[str, str]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }
        skipped_duplicates: dict[str, dict[str, str]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }
        errors: list[str] = []
        all_warnings = list(plan.warnings)

        total_units = max(1, len(job.input_files) * len(job.sheet_configs))
        done_units = 0

        for input_file in job.input_files:
            workbook, load_warnings = load_workbook_with_warnings(
                input_file, data_only=True, read_only=True
            )
            all_warnings.extend(load_warnings)
            try:
                for config in job.sheet_configs:
                    sheet_plan = plans_by_sheet[config.sheet_name]
                    unit_base = done_units / total_units
                    done_units += 1

                    def report_rows(processed: int, total: int | None) -> None:
                        fraction = processed / total if total else 0.0
                        percent = 8 + int((unit_base + min(fraction, 1.0) / total_units) * 89)
                        if total:
                            message = (
                                f"正在合并 {input_file.name} / {config.sheet_name} "
                                f"({processed}/{total})"
                            )
                        else:
                            message = f"正在合并 {input_file.name} / {config.sheet_name}"
                        progress.update(percent, message)

                    if input_file.name not in sheet_plan.headers_by_file:
                        continue
                    sheet = workbook[config.sheet_name]
                    if job.skip_duplicate_sheets:
                        fingerprint = _sheet_fingerprint(sheet)
                        seen = fingerprints[config.sheet_name]
                        if fingerprint in seen:
                            original = seen[fingerprint]
                            skipped_duplicates[config.sheet_name][input_file.name] = original
                            all_warnings.append(
                                f"文件 {input_file.name} 的 sheet「{config.sheet_name}」"
                                f"与文件 {original} 内容完全相同，已跳过"
                            )
                            continue
                        seen[fingerprint] = input_file.name
                    try:
                        added, unit_warnings = self._merge_file_sheet(
                            output_sheets[config.sheet_name],
                            sheet,
                            input_file,
                            sheet_plan,
                            job,
                            initialized=config.sheet_name in initialized_sheets,
                            report_rows=report_rows,
                        )
                        initialized_sheets.add(config.sheet_name)
                        source_rows[config.sheet_name][input_file.name] = added
                        all_warnings.extend(unit_warnings)
                    except Exception as exc:
                        errors.append(f"{input_file.name} / {config.sheet_name}: {exc}")
            finally:
                workbook.close()

        results: list[MergeSheetResult] = []
        for config in job.sheet_configs:
            sheet_plan = plans_by_sheet[config.sheet_name]
            sheet_source = source_rows[config.sheet_name]
            sheet_warnings: list[str] = []
            for file_name, fields in sheet_plan.missing_fields.items():
                if file_name in sheet_source:
                    sheet_warnings.append(
                        f"文件 {file_name} 缺少字段：{'、'.join(fields)}"
                    )
            for file_name, fields in sheet_plan.extra_fields.items():
                if file_name in sheet_source:
                    sheet_warnings.append(
                        f"文件 {file_name} 多出字段：{'、'.join(fields)}，已追加到表头末尾"
                    )
            results.append(
                MergeSheetResult(
                    sheet_name=config.sheet_name,
                    merged_rows=sum(sheet_source.values()),
                    source_rows=sheet_source,
                    missing_fields=dict(sheet_plan.missing_fields),
                    extra_fields=dict(sheet_plan.extra_fields),
                    skipped_duplicates=dict(skipped_duplicates[config.sheet_name]),
                    warnings=sheet_warnings,
                )
            )
            all_warnings.extend(sheet_warnings)

        progress.update(97, "正在保存输出文件")
        output_workbook.save(job.output_file)

        summary = MergeSummary(
            results=results,
            output_file=job.output_file,
            total_rows=sum(result.merged_rows for result in results),
            warnings=list(dict.fromkeys(all_warnings)),
            errors=errors,
        )
        progress.update(100, "合并完成")
        return summary

    def _merge_file_sheet(
        self,
        out_sheet,
        sheet,
        input_file: Path,
        sheet_plan: MergeSheetPlan,
        job: MergeJob,
        *,
        initialized: bool,
        report_rows: Callable[[int, int | None], None],
    ) -> tuple[int, list[str]]:
        """把单个文件的单个 sheet 合并进输出 sheet，返回（写入行数, 警告）。"""
        warnings: list[str] = []
        header_row = sheet_plan.header_row
        union_headers = sheet_plan.union_headers
        file_headers = sheet_plan.headers_by_file[input_file.name]
        mapping = [
            union_headers.index(header) if header is not None else None
            for header in file_headers
        ]
        column_count = len(union_headers) + (1 if job.include_source_column else 0)
        total_rows = (
            sheet.max_row - header_row if sheet.max_row is not None else None
        )
        added = 0
        consecutive_empty = 0

        if not initialized:
            # write_only 模式下 <cols>/行高在首次 append 时写出，布局必须先于任何行写入
            _apply_sheet_layout(out_sheet, input_file, sheet_plan)

        if initialized:
            data_rows = sheet.iter_rows(min_row=header_row + 1)
            for row in data_rows:
                values, is_empty = _map_row(
                    row, mapping, column_count, job, input_file, out_sheet
                )
                added, consecutive_empty, stopped = _append_data_row(
                    out_sheet, values, is_empty, added, consecutive_empty
                )
                if stopped:
                    warnings.append(
                        f"文件 {input_file.name} 的 sheet「{sheet_plan.sheet_name}」"
                        f"连续 {EMPTY_ROW_STOP_THRESHOLD} 行为空，已提前结束读取，"
                        f"后续空行将被忽略"
                    )
                    break
                if added % 2000 == 0:
                    report_rows(added, total_rows)
        else:
            # 第一个包含该 sheet 的文件：先写标题区域和表头，再写数据
            for row_index, row in enumerate(sheet.iter_rows(min_row=1), start=1):
                if row_index < header_row:
                    out_sheet.append(
                        [_styled_cell(out_sheet, cell) for cell in row]
                    )
                elif row_index == header_row:
                    _write_union_header(out_sheet, sheet_plan, job, row)
                else:
                    values, is_empty = _map_row(
                        row, mapping, column_count, job, input_file, out_sheet
                    )
                    added, consecutive_empty, stopped = _append_data_row(
                        out_sheet, values, is_empty, added, consecutive_empty
                    )
                    if stopped:
                        warnings.append(
                            f"文件 {input_file.name} 的 sheet「{sheet_plan.sheet_name}」"
                            f"连续 {EMPTY_ROW_STOP_THRESHOLD} 行为空，已提前结束读取，"
                            f"后续空行将被忽略"
                        )
                        break
                    if added % 2000 == 0:
                        report_rows(added, total_rows)
        return added, warnings


def _sheet_fingerprint(sheet) -> str:
    """计算一个 sheet 的内容指纹：全部行（标题区+表头行+数据行）的原始值哈希。

    基于读取到的原始行值，列顺序不同不算重复（保守语义，仍按表头名对齐合并）。
    """
    digest = hashlib.sha256()
    for row in sheet.iter_rows(values_only=True):
        digest.update(b"\x1e")  # 行分隔
        for value in row:
            digest.update(_fingerprint_token(value))
    return digest.hexdigest()


def _fingerprint_token(value) -> bytes:
    if value is None:
        return b"\x00"
    # 带上类型名，避免数字 1 与字符串 "1" 被当成相同内容
    return b"\x01" + f"{type(value).__name__}:{value}".encode("utf-8", "replace") + b"\x1f"


def _map_row(cells, mapping, column_count, job: MergeJob, input_file: Path, out_sheet):
    """按表头名把一行数据映射到并集列序，返回（输出行, 是否整行为空）。

    单元格样式（字体/填充/边框/对齐/数字格式/保护）跟随源单元格复制；
    带样式的空单元格也会写入（保留边框等视觉效果）。
    """
    values = [None] * column_count
    is_empty = True
    for index, cell in enumerate(cells):
        if index >= len(mapping):
            break
        union_index = mapping[index]
        if union_index is None:
            continue
        value = cell.value
        if normalize_split_value(value) is not None:
            is_empty = False
        if value is None and not getattr(cell, "has_style", False):
            continue
        values[union_index] = _styled_cell(out_sheet, cell)
    if not is_empty and job.include_source_column:
        values[-1] = input_file.stem
    return values, is_empty


def _styled_cell(out_sheet, source_cell):
    """把源单元格转成输出单元格：无样式时只写值，有样式时连样式一起复制。"""
    if not getattr(source_cell, "has_style", False):
        return source_cell.value
    cell = WriteOnlyCell(out_sheet, value=source_cell.value)
    _copy_cell_style(source_cell, cell)
    return cell


def _append_data_row(out_sheet, values, is_empty, added, consecutive_empty):
    """写入一行数据，返回（累计行数, 连续空行数, 是否触发提前结束）。空行跳过。"""
    if is_empty:
        consecutive_empty += 1
        return added, consecutive_empty, consecutive_empty >= EMPTY_ROW_STOP_THRESHOLD
    out_sheet.append(values)
    return added + 1, 0, False


def _write_union_header(out_sheet, sheet_plan: MergeSheetPlan, job: MergeJob, header_cells) -> None:
    """写并集字段表头行，并尽量复制第一个文件的表头样式。"""
    base_headers = sheet_plan.headers_by_file[sheet_plan.base_file]
    style_source: dict[str, object] = {}
    for index, header in enumerate(base_headers):
        if header is not None and header not in style_source and index < len(header_cells):
            style_source[header] = header_cells[index]
    fallback_style = next(iter(style_source.values()), None)

    headers = list(sheet_plan.union_headers)
    if job.include_source_column:
        headers.append(job.source_column_name)
    cells = []
    for header in headers:
        cell = WriteOnlyCell(out_sheet, value=header)
        source = style_source.get(header, fallback_style)
        if source is not None:
            _copy_cell_style(source, cell)
        cells.append(cell)
    out_sheet.append(cells)


def _copy_cell_style(source_cell, target_cell) -> None:
    try:
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.border = copy(source_cell.border)
        target_cell.protection = copy(source_cell.protection)
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format
    except Exception:
        pass  # 复制不了的样式项直接放弃，不影响合并结果


def _apply_sheet_layout(out_sheet, input_file: Path, sheet_plan: MergeSheetPlan) -> None:
    """从基准文件复制 sheet 布局：列宽、行高、冻结窗格、标题区合并单元格。"""
    layout = _read_sheet_layout(input_file, sheet_plan.sheet_name, sheet_plan.header_row)
    base_headers = sheet_plan.headers_by_file[sheet_plan.base_file]
    for out_index, header in enumerate(sheet_plan.union_headers, start=1):
        try:
            base_index = base_headers.index(header)
        except ValueError:
            continue
        width = layout["widths"].get(get_column_letter(base_index + 1))
        if width:
            out_sheet.column_dimensions[get_column_letter(out_index)].width = width
    for row_number, height in layout["row_heights"].items():
        out_sheet.row_dimensions[row_number].height = height
    if layout["freeze_panes"]:
        out_sheet.freeze_panes = layout["freeze_panes"]
    for ref in layout["merges"]:
        try:
            _min_col, _min_row, _max_col, max_row = range_boundaries(ref)
        except ValueError:
            continue
        if max_row < sheet_plan.header_row:  # 只保留标题区（表头行之上）的合并
            # WriteOnlyWorksheet 没有 merge_cells 方法，但 merged_cells.add 可用（已实测）
            out_sheet.merged_cells.add(ref)


def _sheet_archive_entry(archive: ZipFile, sheet_name: str) -> str | None:
    """在 xlsx 压缩包里定位指定 sheet 的 XML 条目。"""
    main_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    rel_id = None
    for sheet_node in workbook_xml.iter(f"{main_ns}sheet"):
        if sheet_node.get("name") == sheet_name:
            rel_id = sheet_node.get(f"{rel_ns}id")
            break
    if rel_id is None:
        return None
    rels_xml = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for rel_node in rels_xml:
        if rel_node.get("Id") == rel_id:
            target = rel_node.get("Target")
            if target:
                return target.lstrip("/") if target.startswith("/") else f"xl/{target}"
    return None


def _read_sheet_layout(path: Path, sheet_name: str, header_row: int) -> dict:
    """直接从 xlsx 压缩包解析 sheet 布局（read_only 模式不提供这些属性）。

    解析列宽、冻结窗格、前 header_row 行的行高、标题区合并单元格。
    解析失败时返回空布局，对应项直接放弃，不影响合并结果。
    """
    main_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    layout: dict = {"widths": {}, "freeze_panes": None, "row_heights": {}, "merges": []}
    try:
        with ZipFile(path, "r") as archive:
            sheet_entry = _sheet_archive_entry(archive, sheet_name)
            if sheet_entry is None:
                return layout
            with archive.open(sheet_entry) as stream:
                for event, node in ElementTree.iterparse(stream, events=("start", "end")):
                    tag = node.tag
                    if event == "start" and tag == f"{main_ns}pane":
                        if node.get("state") == "frozen":
                            layout["freeze_panes"] = node.get("topLeftCell")
                    elif event == "end" and tag == f"{main_ns}col":
                        width = float(node.get("width", 0))
                        if width:
                            for column in range(
                                int(node.get("min")), int(node.get("max")) + 1
                            ):
                                layout["widths"][get_column_letter(column)] = width
                    elif event == "start" and tag == f"{main_ns}row":
                        height = node.get("ht")
                        row_number = int(node.get("r", 0))
                        if height and row_number <= header_row:
                            layout["row_heights"][row_number] = float(height)
                    elif event == "end" and tag == f"{main_ns}mergeCell":
                        ref = node.get("ref")
                        if ref:
                            layout["merges"].append(ref)
    except Exception:
        pass
    return layout


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
