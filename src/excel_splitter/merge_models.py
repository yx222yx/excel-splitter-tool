from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .models import MAX_HEADER_ROW


@dataclass(frozen=True, slots=True)
class MergeSheetConfig:
    sheet_name: str
    header_row: int = 1
    identical: bool = False  # 用户标记各文件中该 sheet 内容完全一致，只保留一份不合并

    def validate(self) -> None:
        if not self.sheet_name.strip():
            raise ValueError("sheet 名不能为空")
        if self.header_row < 1:
            raise ValueError(f"{self.sheet_name} 的表头行必须大于等于 1")
        if self.header_row > MAX_HEADER_ROW:
            raise ValueError(
                f"{self.sheet_name} 的表头行必须位于前 {MAX_HEADER_ROW} 行"
            )


@dataclass(frozen=True, slots=True)
class MergeJob:
    input_files: tuple[Path, ...]
    output_file: Path
    sheet_configs: tuple[MergeSheetConfig, ...]
    include_source_column: bool = False
    source_column_name: str = "来源文件"
    skip_duplicate_sheets: bool = True
    overwrite: bool = False

    @property
    def selected_sheets(self) -> tuple[str, ...]:
        return tuple(config.sheet_name for config in self.sheet_configs)

    def validate(self) -> None:
        if len(self.input_files) < 2:
            raise ValueError("至少需要两个输入文件")
        for input_file in self.input_files:
            if input_file.suffix.lower() != ".xlsx":
                raise ValueError(f"仅支持 .xlsx 文件：{input_file}")
        if self.output_file.suffix.lower() != ".xlsx":
            raise ValueError("输出文件必须是 .xlsx 文件")
        if not self.sheet_configs:
            raise ValueError("至少选择一个 sheet")
        if len(set(self.selected_sheets)) != len(self.selected_sheets):
            raise ValueError("sheet 配置不能重复")
        for config in self.sheet_configs:
            config.validate()
        resolved_output = self.output_file.resolve()
        for input_file in self.input_files:
            if input_file.resolve() == resolved_output:
                raise ValueError(f"输出文件不能覆盖输入文件：{input_file}")
        for input_file in self.input_files:
            if not input_file.is_file():
                raise ValueError(f"输入文件不存在：{input_file}")
        if self.output_file.exists() and not self.overwrite:
            raise ValueError(f"输出文件已存在：{self.output_file}")
        if self.include_source_column and not self.source_column_name.strip():
            raise ValueError("来源列名称不能为空")


@dataclass(frozen=True, slots=True)
class MergeSheetResult:
    sheet_name: str
    merged_rows: int
    source_rows: dict[str, int] = field(default_factory=dict)
    missing_fields: dict[str, list[str]] = field(default_factory=dict)
    extra_fields: dict[str, list[str]] = field(default_factory=dict)
    skipped_duplicates: dict[str, str] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class MergeSummary:
    results: list[MergeSheetResult]
    output_file: Path
    total_rows: int
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
