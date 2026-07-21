from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from .encryption import decrypt_file
from .excel_io import load_workbook_with_warnings
from .merge_models import MergeSheetConfig


@dataclass(slots=True)
class MergeSheetPlan:
    sheet_name: str
    header_row: int
    union_headers: list[str]
    base_file: str | None = None
    headers_by_file: dict[str, list[str | None]] = field(default_factory=dict)
    missing_files: list[str] = field(default_factory=list)
    missing_fields: dict[str, list[str]] = field(default_factory=dict)
    extra_fields: dict[str, list[str]] = field(default_factory=dict)


@dataclass(slots=True)
class MergePlan:
    sheets: list[MergeSheetPlan]
    warnings: list[str] = field(default_factory=list)


def build_merge_plan(
    input_files: Iterable[Path],
    sheet_configs: Iterable[MergeSheetConfig],
    passwords: dict | None = None,
) -> MergePlan:
    """流式扫描各输入文件的表头行，生成字段并集与缺失/多余报告。

    并集字段以第一个包含该 sheet 的文件的字段顺序为基准，
    其他文件多出的字段追加在末尾。表头匹配前仅做首尾空白规范化。
    """
    files = tuple(input_files)
    configs = tuple(sheet_configs)
    for config in configs:
        config.validate()
    names = [path.name for path in files]
    if len(set(names)) != len(names):
        raise ValueError("输入文件中存在同名文件，无法按文件名区分")

    plans: list[MergeSheetPlan] = []
    warnings: list[str] = []
    # 逐文件流式扫描，只读取各目标 sheet 的表头行
    headers_by_sheet: dict[str, dict[str, list[str | None]]] = {
        config.sheet_name: {} for config in configs
    }
    missing_by_sheet: dict[str, list[str]] = {
        config.sheet_name: [] for config in configs
    }
    for path in files:
        workbook, load_warnings = _open_readable_workbook(path, passwords)
        warnings.extend(load_warnings)
        try:
            for config in configs:
                if config.sheet_name not in workbook.sheetnames:
                    missing_by_sheet[config.sheet_name].append(path.name)
                    warnings.append(f"文件 {path.name} 缺少 sheet：{config.sheet_name}")
                    continue
                sheet = workbook[config.sheet_name]
                if sheet.max_row is None or config.header_row > sheet.max_row:
                    missing_by_sheet[config.sheet_name].append(path.name)
                    warnings.append(
                        f"文件 {path.name} 的 sheet「{config.sheet_name}」"
                        f"行数不足，缺少表头行，已跳过"
                    )
                    continue
                header_row_values = next(
                    sheet.iter_rows(
                        min_row=config.header_row,
                        max_row=config.header_row,
                        values_only=True,
                    )
                )
                headers_by_sheet[config.sheet_name][path.name] = [
                    _normalize_header(value) for value in header_row_values
                ]
        finally:
            workbook.close()

    for config in configs:
        file_headers = headers_by_sheet[config.sheet_name]
        missing_files = missing_by_sheet[config.sheet_name]
        base_file = next(
            (path.name for path in files if path.name in file_headers), None
        )
        union_headers: list[str] = []
        if base_file is not None:
            union_headers.extend(
                header for header in file_headers[base_file] if header is not None
            )
        for path in files:
            headers = file_headers.get(path.name)
            if headers is None or path.name == base_file:
                continue
            for header in headers:
                if header is not None and header not in union_headers:
                    union_headers.append(header)

        missing_fields: dict[str, list[str]] = {}
        extra_fields: dict[str, list[str]] = {}
        base_fields = (
            {header for header in file_headers[base_file] if header is not None}
            if base_file is not None
            else set()
        )
        for file_name, headers in file_headers.items():
            present = {header for header in headers if header is not None}
            missing = [header for header in union_headers if header not in present]
            extra = [
                header
                for header in headers
                if header is not None and header not in base_fields
            ]
            if missing:
                missing_fields[file_name] = missing
            if file_name != base_file and extra:
                extra_fields[file_name] = list(dict.fromkeys(extra))

        plans.append(
            MergeSheetPlan(
                sheet_name=config.sheet_name,
                header_row=config.header_row,
                union_headers=union_headers,
                base_file=base_file,
                headers_by_file=file_headers,
                missing_files=missing_files,
                missing_fields=missing_fields,
                extra_fields=extra_fields,
            )
        )
    return MergePlan(sheets=plans, warnings=list(dict.fromkeys(warnings)))


def _normalize_header(value) -> str | None:
    if value is None:
        return None
    normalized = str(value).strip()
    return normalized or None


def _open_readable_workbook(path: Path, passwords: dict | None):
    if passwords and path in passwords:
        stream = decrypt_file(path, passwords[path])
        return load_workbook_with_warnings(stream, data_only=True, read_only=True)
    return load_workbook_with_warnings(path, data_only=True, read_only=True)
