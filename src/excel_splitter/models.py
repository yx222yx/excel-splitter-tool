from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


SplitMode = Literal["all", "selected"]
OutputType = Literal["formula", "values"]
SheetMode = Literal["direct", "reference", "linked", "full"]
MAX_HEADER_ROW = 15


@dataclass(frozen=True, slots=True)
class SheetConfig:
    sheet_name: str
    header_row: int
    split_column_idx: int | None
    split_column_label: str = ""
    mode: SheetMode = "direct"
    key_column_idx: int | None = None
    key_column_label: str = ""

    def validate(self) -> None:
        if not self.sheet_name.strip():
            raise ValueError("sheet 名不能为空")
        if self.header_row < 1:
            raise ValueError(f"{self.sheet_name} 的表头行必须大于等于 1")
        if self.header_row > MAX_HEADER_ROW:
            raise ValueError(
                f"{self.sheet_name} 的表头行必须位于前 {MAX_HEADER_ROW} 行"
            )
        if self.mode not in ("direct", "reference", "linked", "full"):
            raise ValueError(f"{self.sheet_name} 的处理方式无效")
        if self.mode in ("direct", "reference") and (
            self.split_column_idx is None or self.split_column_idx < 1
        ):
            raise ValueError(f"{self.sheet_name} 的拆分列必须大于等于 1")
        if self.mode in ("reference", "linked") and (
            self.key_column_idx is None or self.key_column_idx < 1
        ):
            raise ValueError(f"{self.sheet_name} 的关联键列必须大于等于 1")


@dataclass(frozen=True, slots=True)
class SplitJob:
    input_file: Path
    output_dir: Path
    sheet_configs: tuple[SheetConfig, ...]
    split_mode: SplitMode = "all"
    selected_split_values: tuple[str, ...] = ()
    filename_template: str = "{original_name}_{split_value}_{output_type}"
    output_types: tuple[OutputType, ...] = ("formula", "values")
    overwrite: bool = False
    original_name: str | None = None

    @property
    def selected_sheets(self) -> tuple[str, ...]:
        return tuple(config.sheet_name for config in self.sheet_configs)

    def validate(self, *, check_input_file: bool = True) -> None:
        if self.input_file.suffix.lower() != ".xlsx":
            raise ValueError("第一版仅支持 .xlsx 文件")
        if not self.sheet_configs:
            raise ValueError("至少选择一个 sheet")
        if len(set(self.selected_sheets)) != len(self.selected_sheets):
            raise ValueError("sheet 配置不能重复")
        for config in self.sheet_configs:
            config.validate()
        reference_count = sum(
            config.mode == "reference" for config in self.sheet_configs
        )
        if reference_count > 1:
            raise ValueError("一次任务只能选择一个基准 Sheet")
        if any(config.mode == "linked" for config in self.sheet_configs) and not reference_count:
            raise ValueError("按关联键匹配时必须选择一个基准 Sheet")
        if self.split_mode not in ("all", "selected"):
            raise ValueError("拆分模式必须是 all 或 selected")
        if self.split_mode == "selected" and not any(
            str(value).strip() for value in self.selected_split_values
        ):
            raise ValueError("手动模式下至少选择一个拆分值")
        if not self.output_types:
            raise ValueError("至少选择一种输出版本")
        if len(set(self.output_types)) != len(self.output_types) or any(
            output_type not in ("formula", "values")
            for output_type in self.output_types
        ):
            raise ValueError("输出版本只能是 formula 或 values，且不能重复")
        if not self.filename_template.strip():
            raise ValueError("文件名模板不能为空")
        try:
            self.filename_template.format(
                original_name="input", split_value="value", output_type="公式版"
            )
        except (KeyError, ValueError) as exc:
            raise ValueError(
                "文件名模板仅支持 {original_name}、{split_value} 和 {output_type}"
            ) from exc
        if check_input_file and not self.input_file.is_file():
            raise ValueError(f"输入文件不存在：{self.input_file}")


@dataclass(frozen=True, slots=True)
class OutputArtifact:
    output_type: OutputType
    output_file: Path


@dataclass(frozen=True, slots=True)
class SplitResult:
    split_value: str
    output_files: list[OutputArtifact]
    sheet_rows: dict[str, int]
    discarded_empty_rows: dict[str, int]
    unmatched_key_rows: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class SplitSummary:
    results: list[SplitResult]
    total_files: int
    total_discarded: int
    total_unmatched: int = 0
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
