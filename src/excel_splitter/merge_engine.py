from __future__ import annotations

import hashlib
from copy import copy
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from .excel_io import load_workbook_with_warnings
from .merge_models import MergeJob, MergeSheetConfig, MergeSheetResult, MergeSummary
from .merge_planning import MergePlan, MergeSheetPlan, build_merge_plan
from .values import normalize_split_value


EMPTY_ROW_STOP_THRESHOLD = 10000
EXTRA_COLUMN_WIDTH = 12  # 并集新增列的默认列宽
ProgressCallback = Callable[[int, str], None]


class MergeEngine:
    """模板照抄式合并引擎。

    以第一个输入文件为输出模板（普通模式加载，公式/样式/合并/冻结等全部
    保留），只删除未选中的 sheet；选中 sheet 的表头行及之前内容一个字节不动，
    其余文件的数据行按表头名对齐后流式追加到末尾。
    """

    def execute(
        self,
        job: MergeJob,
        progress_callback: ProgressCallback | None = None,
        plan: MergePlan | None = None,
    ) -> MergeSummary:
        progress = _ProgressReporter(progress_callback)
        progress.update(0, "正在分析输入文件")
        job.validate()

        # identical 的 sheet 不需要字段扫描，其余 sheet 扫描表头生成并集与映射
        scan_configs = tuple(c for c in job.sheet_configs if not c.identical)
        if plan is None:
            plan = (
                build_merge_plan(job.input_files, scan_configs)
                if scan_configs
                else MergePlan(sheets=[])
            )
        plans_by_sheet = {sheet.sheet_name: sheet for sheet in plan.sheets}
        progress.update(5, "正在加载模板工作簿")

        first_file = job.input_files[0]
        workbook, load_warnings = load_workbook_with_warnings(first_file, data_only=False)
        all_warnings = [*plan.warnings, *load_warnings]
        selected = set(job.selected_sheets)
        for name in list(workbook.sheetnames):
            if name not in selected:
                workbook.remove(workbook[name])

        errors: list[str] = []
        source_rows: dict[str, dict[str, int]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }
        skipped_duplicates: dict[str, dict[str, str]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }
        fingerprints: dict[str, dict[str, str]] = {
            config.sheet_name: {} for config in job.sheet_configs
        }

        total_units = max(1, len(job.input_files) * len(job.sheet_configs))
        done_units = 0

        try:
            for config in job.sheet_configs:
                unit_base = done_units / total_units
                done_units += len(job.input_files)

                def report(fraction: float, message: str) -> None:
                    percent = 5 + int(
                        min(1.0, unit_base + fraction * len(job.input_files) / total_units) * 90
                    )
                    progress.update(percent, message)

                if config.identical:
                    self._keep_identical_sheet(
                        workbook, config, job, source_rows[config.sheet_name], all_warnings
                    )
                    report(1.0, f"已保留 {config.sheet_name}（内容一致，不合并）")
                    continue
                sheet_plan = plans_by_sheet.get(config.sheet_name)
                if sheet_plan is None:
                    raise ValueError(f"找不到 sheet：{config.sheet_name}")
                try:
                    self._merge_sheet(
                        workbook,
                        config,
                        sheet_plan,
                        job,
                        source_rows[config.sheet_name],
                        skipped_duplicates[config.sheet_name],
                        fingerprints[config.sheet_name],
                        all_warnings,
                        report,
                    )
                except Exception as exc:
                    errors.append(f"{config.sheet_name}: {exc}")

            ordered = [name for name in job.selected_sheets if name in workbook.sheetnames]
            if not ordered:
                raise ValueError("所选 sheet 在所有输入文件中都不存在")
            workbook._sheets = [workbook[name] for name in ordered]

            progress.update(97, "正在保存输出文件")
            job.output_file.parent.mkdir(parents=True, exist_ok=True)
            workbook.save(job.output_file)
        finally:
            workbook.close()

        results: list[MergeSheetResult] = []
        for config in job.sheet_configs:
            sheet_plan = plans_by_sheet.get(config.sheet_name)
            sheet_source = source_rows[config.sheet_name]
            sheet_warnings: list[str] = []
            if sheet_plan is not None:
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
                    missing_fields=dict(sheet_plan.missing_fields) if sheet_plan else {},
                    extra_fields=dict(sheet_plan.extra_fields) if sheet_plan else {},
                    skipped_duplicates=dict(skipped_duplicates[config.sheet_name]),
                    warnings=sheet_warnings,
                )
            )
            all_warnings.extend(sheet_warnings)

        summary = MergeSummary(
            results=results,
            output_file=job.output_file,
            total_rows=sum(result.merged_rows for result in results),
            warnings=list(dict.fromkeys(all_warnings)),
            errors=errors,
        )
        progress.update(100, "合并完成")
        return summary

    def _keep_identical_sheet(
        self,
        workbook,
        config: MergeSheetConfig,
        job: MergeJob,
        source_rows: dict[str, int],
        warnings: list[str],
    ) -> None:
        """identical sheet：模板里有就原样保留；没有则从第一个包含的文件整体拷贝。

        其余文件的同名 sheet 完全不读取。
        """
        sheet_name = config.sheet_name
        if sheet_name in workbook.sheetnames:
            source_rows[job.input_files[0].name] = _count_data_rows(
                workbook[sheet_name], config.header_row
            )
            return
        for input_file in job.input_files[1:]:
            probe, _ = load_workbook_with_warnings(
                input_file, data_only=False, read_only=True
            )
            contains = sheet_name in probe.sheetnames
            probe.close()
            if not contains:
                continue
            ws = _copy_sheet_into_template(workbook, input_file, sheet_name)
            source_rows[input_file.name] = _count_data_rows(ws, config.header_row)
            return
        warnings.append(f"所有输入文件都不包含 sheet：{sheet_name}")

    def _merge_sheet(
        self,
        workbook,
        config: MergeSheetConfig,
        sheet_plan: MergeSheetPlan,
        job: MergeJob,
        source_rows: dict[str, int],
        skipped_duplicates: dict[str, str],
        fingerprints: dict[str, str],
        warnings: list[str],
        report: Callable[[float, str], None],
    ) -> None:
        sheet_name = config.sheet_name
        ws = workbook[sheet_name] if sheet_name in workbook.sheetnames else None
        files = [
            path for path in job.input_files if path.name in sheet_plan.headers_by_file
        ]
        if not files:
            warnings.append(f"所有输入文件都不包含 sheet：{sheet_name}")
            return
        total = len(files)
        for index, input_file in enumerate(files):
            if ws is None or input_file == job.input_files[0]:
                # 基准内容就位：模板自带则一个字节不动；模板缺该 sheet 时整体拷贝
                if ws is None:
                    ws = _copy_sheet_into_template(workbook, input_file, sheet_name)
                source_rows[input_file.name] = _count_data_rows(ws, config.header_row)
                if job.skip_duplicate_sheets:
                    fingerprints[_sheet_fingerprint(ws)] = input_file.name
                self._prepare_union_columns(ws, config, sheet_plan, job, input_file)
                report((index + 1) / total, f"已就位 {input_file.name} / {sheet_name}")
                continue
            if job.skip_duplicate_sheets:
                fingerprint = _read_sheet_fingerprint(input_file, sheet_name)
                if fingerprint in fingerprints:
                    original = fingerprints[fingerprint]
                    skipped_duplicates[input_file.name] = original
                    warnings.append(
                        f"文件 {input_file.name} 的 sheet「{sheet_name}」"
                        f"与文件 {original} 内容完全相同，已跳过"
                    )
                    report((index + 1) / total, f"跳过重复 {input_file.name} / {sheet_name}")
                    continue
                fingerprints[fingerprint] = input_file.name
            added = self._append_file_rows(
                ws,
                input_file,
                config,
                sheet_plan,
                job,
                warnings,
                lambda fraction, message: report((index + fraction) / total, message),
            )
            source_rows[input_file.name] = added
            report((index + 1) / total, f"已合并 {input_file.name} / {sheet_name}")

    def _prepare_union_columns(
        self,
        ws,
        config: MergeSheetConfig,
        sheet_plan: MergeSheetPlan,
        job: MergeJob,
        base_file: Path,
    ) -> None:
        """在模板表头行右侧补并集新增列和（可选的）来源列，并为模板自有数据行填来源。"""
        base_headers = sheet_plan.headers_by_file[base_file.name]
        base_set = {header for header in base_headers if header is not None}
        extra_headers = [
            header for header in sheet_plan.union_headers if header not in base_set
        ]
        style_source = ws.cell(row=config.header_row, column=max(1, len(base_headers)))
        column = len(base_headers) + 1
        for header in extra_headers:
            cell = ws.cell(row=config.header_row, column=column, value=header)
            _copy_cell_style(style_source, cell)
            ws.column_dimensions[get_column_letter(column)].width = EXTRA_COLUMN_WIDTH
            column += 1
        if job.include_source_column:
            source_column = len(sheet_plan.union_headers) + 1
            cell = ws.cell(
                row=config.header_row, column=source_column, value=job.source_column_name
            )
            _copy_cell_style(style_source, cell)
            for row in ws.iter_rows(min_row=config.header_row + 1):
                if any(normalize_split_value(cell.value) is not None for cell in row):
                    ws.cell(row=row[0].row, column=source_column, value=base_file.stem)

    def _append_file_rows(
        self,
        ws,
        input_file: Path,
        config: MergeSheetConfig,
        sheet_plan: MergeSheetPlan,
        job: MergeJob,
        warnings: list[str],
        report: Callable[[float, str], None],
    ) -> int:
        """流式读取一个文件的数据行，按表头名映射追加到模板 sheet 末尾。"""
        sheet_name = config.sheet_name
        union_headers = sheet_plan.union_headers
        file_headers = sheet_plan.headers_by_file[input_file.name]
        mapping = [
            union_headers.index(header) if header is not None else None
            for header in file_headers
        ]
        column_count = len(union_headers) + (1 if job.include_source_column else 0)
        layout = _read_sheet_layout(input_file, sheet_name, None)
        row_heights = layout["row_heights"]
        source_workbook, load_warnings = load_workbook_with_warnings(
            input_file, data_only=False, read_only=True
        )
        warnings.extend(load_warnings)
        added = 0
        consecutive_empty = 0
        next_row = ws.max_row + 1
        try:
            sheet = source_workbook[sheet_name]
            total_rows = (
                sheet.max_row - config.header_row if sheet.max_row is not None else None
            )
            for row in sheet.iter_rows(min_row=config.header_row + 1):
                is_empty = True
                pending: list[tuple[int, object, object]] = []  # (目标列, 值, 源单元格)
                for index, cell in enumerate(row):
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
                    pending.append((union_index + 1, value, cell))
                if is_empty:
                    consecutive_empty += 1
                    if consecutive_empty >= EMPTY_ROW_STOP_THRESHOLD:
                        warnings.append(
                            f"文件 {input_file.name} 的 sheet「{sheet_name}」"
                            f"连续 {EMPTY_ROW_STOP_THRESHOLD} 行为空，已提前结束读取，"
                            f"后续空行将被忽略"
                        )
                        break
                    continue
                consecutive_empty = 0
                for target_column, value, source_cell in pending:
                    if isinstance(value, str) and value.startswith("="):
                        value = _translate_formula(
                            value, source_cell.coordinate, target_column, next_row
                        )
                    out_cell = ws.cell(row=next_row, column=target_column, value=value)
                    if getattr(source_cell, "has_style", False):
                        _copy_cell_style(source_cell, out_cell)
                if job.include_source_column:
                    ws.cell(row=next_row, column=column_count, value=input_file.stem)
                source_row_number = row[0].row if row else None
                if source_row_number in row_heights:
                    ws.row_dimensions[next_row].height = row_heights[source_row_number]
                next_row += 1
                added += 1
                if added % 2000 == 0:
                    fraction = added / total_rows if total_rows else 0.0
                    message = (
                        f"正在合并 {input_file.name} / {sheet_name} ({added}/{total_rows})"
                        if total_rows
                        else f"正在合并 {input_file.name} / {sheet_name}"
                    )
                    report(fraction, message)
        finally:
            source_workbook.close()
        return added


def _translate_formula(formula: str, origin: str, target_column: int, target_row: int) -> str:
    """把公式从源坐标平移到目标坐标；平移失败时保留原公式。"""
    try:
        return Translator(formula, origin=origin).translate_formula(
            f"{get_column_letter(target_column)}{target_row}"
        )
    except Exception:
        return formula


def _count_data_rows(worksheet, header_row: int) -> int:
    count = 0
    for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
        if any(normalize_split_value(value) is not None for value in row):
            count += 1
    return count


def _copy_sheet_into_template(workbook, input_file: Path, sheet_name: str):
    """openpyxl 不能跨工作簿复制 sheet：逐格复制值+样式，再补列宽/行高/冻结/合并。"""
    layout = _read_sheet_layout(input_file, sheet_name, None)
    source_workbook, _ = load_workbook_with_warnings(
        input_file, data_only=False, read_only=True
    )
    try:
        source = source_workbook[sheet_name]
        target = workbook.create_sheet(sheet_name)
        for row in source.iter_rows():
            for cell in row:
                if cell.value is None and not getattr(cell, "has_style", False):
                    continue
                out_cell = target.cell(row=cell.row, column=cell.column, value=cell.value)
                if getattr(cell, "has_style", False):
                    _copy_cell_style(cell, out_cell)
    finally:
        source_workbook.close()
    for letter, width in layout["widths"].items():
        target.column_dimensions[letter].width = width
    for row_number, height in layout["row_heights"].items():
        target.row_dimensions[row_number].height = height
    if layout["freeze_panes"]:
        target.freeze_panes = layout["freeze_panes"]
    for ref in layout["merges"]:
        target.merge_cells(ref)
    return target


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


def _sheet_fingerprint(worksheet) -> str:
    digest = hashlib.sha256()
    for row in worksheet.iter_rows(values_only=True):
        digest.update(b"\x1e")  # 行分隔
        for value in row:
            digest.update(_fingerprint_token(value))
    return digest.hexdigest()


def _read_sheet_fingerprint(path: Path, sheet_name: str) -> str:
    """流式读取一个文件的指定 sheet 并计算内容指纹。

    基于读取到的原始行值，列顺序不同不算重复（保守语义，仍按表头名对齐合并）。
    """
    workbook, _ = load_workbook_with_warnings(path, data_only=False, read_only=True)
    try:
        return _sheet_fingerprint(workbook[sheet_name])
    finally:
        workbook.close()


def _fingerprint_token(value) -> bytes:
    if value is None:
        return b"\x00"
    # 带上类型名，避免数字 1 与字符串 "1" 被当成相同内容
    return b"\x01" + f"{type(value).__name__}:{value}".encode("utf-8", "replace") + b"\x1f"


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


def _read_sheet_layout(path: Path, sheet_name: str, max_height_row: int | None) -> dict:
    """直接从 xlsx 压缩包解析 sheet 布局（read_only 模式不提供这些属性）。

    解析列宽、冻结窗格、行高（max_height_row 为 None 时全部行，否则只到该
    行）、合并单元格。解析失败时返回空布局，对应项直接放弃，不影响合并结果。
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
                        if height and (
                            max_height_row is None or row_number <= max_height_row
                        ):
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
