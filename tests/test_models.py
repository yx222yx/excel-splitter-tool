from pathlib import Path

import pytest

from excel_splitter.models import SheetConfig, SplitJob


def test_selected_mode_requires_at_least_one_value(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(SheetConfig("人员", 2, 1, "A - 部门"),),
        split_mode="selected",
        selected_split_values=(),
    )

    with pytest.raises(ValueError, match="至少选择一个拆分值"):
        job.validate()


def test_all_mode_does_not_require_selected_values(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(SheetConfig("人员", 2, 1, "A - 部门"),),
        split_mode="all",
    )

    job.validate(check_input_file=False)


def test_header_row_must_be_within_first_fifteen_rows():
    SheetConfig("人员", 15, 1, "A - 部门").validate()

    with pytest.raises(ValueError, match="前 15 行"):
        SheetConfig("人员", 16, 1, "A - 部门").validate()


def test_filename_template_accepts_output_type_placeholder(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(SheetConfig("人员", 2, 1, "A - 部门"),),
        filename_template="{original_name}_{split_value}_{output_type}",
    )

    job.validate(check_input_file=False)


def test_sheet_config_supports_full_reference_and_linked_modes():
    SheetConfig("说明", 1, None, mode="full").validate()
    SheetConfig(
        "人员归属", 1, 1, "A - 二级团队", mode="reference", key_column_idx=2
    ).validate()
    SheetConfig("工时明细", 1, None, mode="linked", key_column_idx=1).validate()


def test_linked_mode_requires_one_reference_sheet(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(
            SheetConfig("工时明细", 1, None, mode="linked", key_column_idx=1),
        ),
    )

    with pytest.raises(ValueError, match="基准 Sheet"):
        job.validate(check_input_file=False)


def test_only_one_reference_sheet_is_allowed(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(
            SheetConfig("人员归属", 1, 1, mode="reference", key_column_idx=2),
            SheetConfig("另一归属", 1, 1, mode="reference", key_column_idx=2),
        ),
    )

    with pytest.raises(ValueError, match="只能选择一个基准 Sheet"):
        job.validate(check_input_file=False)


def test_job_requires_at_least_one_output_type(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(SheetConfig("人员", 1, 1),),
        output_types=(),
    )

    with pytest.raises(ValueError, match="至少选择一种输出版本"):
        job.validate(check_input_file=False)


def test_job_rejects_unknown_output_type(tmp_path: Path):
    job = SplitJob(
        input_file=tmp_path / "input.xlsx",
        output_dir=tmp_path / "out",
        sheet_configs=(SheetConfig("人员", 1, 1),),
        output_types=("unknown",),
    )

    with pytest.raises(ValueError, match="输出版本"):
        job.validate(check_input_file=False)
