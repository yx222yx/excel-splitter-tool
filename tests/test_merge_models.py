from pathlib import Path

import pytest

from excel_splitter.merge_models import MergeJob, MergeSheetConfig
from excel_splitter.models import MAX_HEADER_ROW


def _touch(path: Path) -> Path:
    path.write_bytes(b"x")
    return path


def _valid_job(tmp_path: Path, **overrides) -> MergeJob:
    file_a = _touch(tmp_path / "部门A.xlsx")
    file_b = _touch(tmp_path / "部门B.xlsx")
    kwargs = {
        "input_files": (file_a, file_b),
        "output_file": tmp_path / "合并结果.xlsx",
        "sheet_configs": (MergeSheetConfig("工时"),),
    }
    kwargs.update(overrides)
    return MergeJob(**kwargs)


def test_sheet_config_valid_defaults():
    MergeSheetConfig("工时").validate()


def test_sheet_config_rejects_blank_name():
    with pytest.raises(ValueError, match="sheet 名不能为空"):
        MergeSheetConfig("  ").validate()


def test_sheet_config_rejects_header_row_below_one():
    with pytest.raises(ValueError, match="表头行必须大于等于 1"):
        MergeSheetConfig("工时", header_row=0).validate()


def test_sheet_config_rejects_header_row_beyond_limit():
    with pytest.raises(ValueError, match=f"前 {MAX_HEADER_ROW} 行"):
        MergeSheetConfig("工时", header_row=MAX_HEADER_ROW + 1).validate()


def test_job_valid_passes(tmp_path):
    _valid_job(tmp_path).validate()


def test_job_requires_at_least_two_input_files(tmp_path):
    file_a = _touch(tmp_path / "部门A.xlsx")
    job = _valid_job(tmp_path, input_files=(file_a,))
    with pytest.raises(ValueError, match="至少需要两个输入文件"):
        job.validate()


def test_job_rejects_non_xlsx_input(tmp_path):
    file_a = _touch(tmp_path / "部门A.xlsx")
    file_b = _touch(tmp_path / "部门B.xls")
    job = _valid_job(tmp_path, input_files=(file_a, file_b))
    with pytest.raises(ValueError, match="仅支持 .xlsx 文件"):
        job.validate()


def test_job_rejects_non_xlsx_output(tmp_path):
    job = _valid_job(tmp_path, output_file=tmp_path / "合并结果.xls")
    with pytest.raises(ValueError, match="输出文件必须是 .xlsx 文件"):
        job.validate()


def test_job_requires_sheet_configs(tmp_path):
    job = _valid_job(tmp_path, sheet_configs=())
    with pytest.raises(ValueError, match="至少选择一个 sheet"):
        job.validate()


def test_job_rejects_duplicate_sheet_configs(tmp_path):
    job = _valid_job(
        tmp_path,
        sheet_configs=(MergeSheetConfig("工时"), MergeSheetConfig("工时")),
    )
    with pytest.raises(ValueError, match="sheet 配置不能重复"):
        job.validate()


def test_job_rejects_output_overwriting_input(tmp_path):
    file_a = _touch(tmp_path / "部门A.xlsx")
    job = _valid_job(tmp_path, output_file=file_a, overwrite=True)
    with pytest.raises(ValueError, match="输出文件不能覆盖输入文件"):
        job.validate()


def test_job_rejects_missing_input_file(tmp_path):
    file_a = _touch(tmp_path / "部门A.xlsx")
    missing = tmp_path / "不存在.xlsx"
    job = _valid_job(tmp_path, input_files=(file_a, missing))
    with pytest.raises(ValueError, match="输入文件不存在"):
        job.validate()


def test_job_rejects_existing_output_without_overwrite(tmp_path):
    output = _touch(tmp_path / "合并结果.xlsx")
    job = _valid_job(tmp_path, output_file=output)
    with pytest.raises(ValueError, match="输出文件已存在"):
        job.validate()


def test_job_allows_existing_output_with_overwrite(tmp_path):
    output = _touch(tmp_path / "合并结果.xlsx")
    job = _valid_job(tmp_path, output_file=output, overwrite=True)
    job.validate()


def test_job_rejects_blank_source_column_name(tmp_path):
    job = _valid_job(
        tmp_path, include_source_column=True, source_column_name="  "
    )
    with pytest.raises(ValueError, match="来源列名称不能为空"):
        job.validate()
