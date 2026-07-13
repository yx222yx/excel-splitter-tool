from pathlib import Path

from openpyxl import load_workbook

from excel_splitter.engine import SplitEngine
from excel_splitter.models import SheetConfig, SplitJob


def _output_file(result, output_type: str):
    return next(
        item.output_file for item in result.output_files if item.output_type == output_type
    )


def _job(sample_workbook: Path, output_dir: Path, mode: str = "all") -> SplitJob:
    return SplitJob(
        input_file=sample_workbook,
        output_dir=output_dir,
        sheet_configs=(
            SheetConfig("人员", 2, 1, "A - 部门"),
            SheetConfig("项目", 1, 1, "A - 所属部门"),
        ),
        split_mode=mode,
        selected_split_values=("临床部",) if mode == "selected" else (),
    )


def test_engine_exports_formula_and_values_workbooks_per_union_value(sample_workbook, tmp_path):
    summary = SplitEngine().execute(_job(sample_workbook, tmp_path / "输出"))

    assert summary.total_files == 6
    assert [result.split_value for result in summary.results] == [
        "临床部",
        "研发部",
        "市场部",
    ]
    assert summary.errors == []

    assert [item.output_type for item in summary.results[0].output_files] == [
        "formula",
        "values",
    ]

    clinical = load_workbook(
        _output_file(summary.results[0], "formula"), data_only=False
    )
    assert clinical.sheetnames == ["人员", "项目"]
    assert clinical["人员"].max_row == 3
    assert clinical["人员"]["A3"].value == "临床部"
    assert clinical["人员"]["D3"].value == "=C3*2"
    assert clinical["项目"].max_row == 2
    assert clinical["项目"]["A2"].value == "临床部"
    assert clinical["人员"].freeze_panes == "B3"
    assert clinical["人员"].column_dimensions["A"].width == 18
    assert "A1:D1" in {str(item) for item in clinical["人员"].merged_cells.ranges}
    clinical.close()

    clinical_values = load_workbook(
        _output_file(summary.results[0], "values"), data_only=False
    )
    assert clinical_values["人员"]["A3"].value == "临床部"
    assert clinical_values["人员"]["D3"].value is None
    clinical_values.close()


def test_union_value_missing_from_sheet_keeps_titles_and_header(sample_workbook, tmp_path):
    summary = SplitEngine().execute(_job(sample_workbook, tmp_path / "输出"))
    market_result = next(item for item in summary.results if item.split_value == "市场部")

    market = load_workbook(_output_file(market_result, "formula"))
    assert market["人员"].max_row == 2
    assert market["人员"]["A1"].value == "人员奖金明细"
    assert market["人员"]["A2"].value == "部门"
    assert market["项目"]["A2"].value == "市场部"
    market.close()


def test_selected_mode_exports_only_selected_values_and_reports_empty_rows(
    sample_workbook, tmp_path
):
    summary = SplitEngine().execute(
        _job(sample_workbook, tmp_path / "输出", mode="selected")
    )

    assert summary.total_files == 2
    assert summary.results[0].split_value == "临床部"
    assert summary.results[0].discarded_empty_rows == {"人员": 1, "项目": 0}
    assert summary.total_discarded == 1


def test_engine_filters_by_cached_result_but_preserves_formula(formula_workbook, tmp_path):
    job = SplitJob(
        input_file=formula_workbook,
        output_dir=tmp_path / "输出",
        sheet_configs=(SheetConfig("公式数据", 1, 1, "A - 部门公式"),),
        split_mode="selected",
        selected_split_values=("临床部",),
    )

    summary = SplitEngine().execute(job)

    formula_output = load_workbook(
        _output_file(summary.results[0], "formula"), data_only=False
    )
    assert formula_output["公式数据"].max_row == 2
    assert formula_output["公式数据"]["A2"].value == '=IF(1=1,"临床部","")'
    assert formula_output["公式数据"]["B2"].value == "张三"
    formula_output.close()

    values_output = load_workbook(
        _output_file(summary.results[0], "values"), data_only=False
    )
    assert values_output["公式数据"].max_row == 2
    assert values_output["公式数据"]["A2"].value == "临床部"
    assert values_output["公式数据"]["C2"].value == 2
    values_output.close()


def test_engine_uses_reference_keys_to_filter_linked_sheets_and_keeps_full_sheets(
    linked_workbook, tmp_path
):
    job = SplitJob(
        input_file=linked_workbook,
        output_dir=tmp_path / "输出",
        sheet_configs=(
            SheetConfig(
                "人员归属", 1, 1, mode="reference", key_column_idx=2
            ),
            SheetConfig("工时明细", 1, None, mode="linked", key_column_idx=1),
            SheetConfig("团队汇总", 1, 1, mode="direct"),
            SheetConfig("说明", 1, None, mode="full"),
        ),
        split_mode="selected",
        selected_split_values=("团队甲",),
    )

    summary = SplitEngine().execute(job)

    result = summary.results[0]
    assert result.sheet_rows == {
        "人员归属": 2,
        "工时明细": 2,
        "团队汇总": 1,
        "说明": 2,
    }
    assert result.discarded_empty_rows == {
        "人员归属": 0,
        "工时明细": 1,
        "团队汇总": 0,
        "说明": 0,
    }
    assert result.unmatched_key_rows == {
        "人员归属": 0,
        "工时明细": 1,
        "团队汇总": 0,
        "说明": 0,
    }
    assert summary.total_unmatched == 1

    output = load_workbook(_output_file(result, "formula"), data_only=False)
    assert [output["工时明细"].cell(row=row, column=1).value for row in range(2, 4)] == [
        "张三",
        "李四",
    ]
    assert output["说明"].max_row == 3
    output.close()


def test_engine_exports_one_complete_copy_when_all_sheets_are_full(
    linked_workbook, tmp_path
):
    job = SplitJob(
        input_file=linked_workbook,
        output_dir=tmp_path / "输出",
        sheet_configs=(
            SheetConfig("人员归属", 1, None, mode="full"),
            SheetConfig("说明", 1, None, mode="full"),
        ),
    )

    summary = SplitEngine().execute(job)

    assert summary.total_files == 2
    assert summary.results[0].split_value == "完整表"
    output = load_workbook(_output_file(summary.results[0], "formula"))
    assert output["人员归属"].max_row == 4
    assert output["说明"].max_row == 3
    output.close()


def test_engine_can_export_formula_version_only(sample_workbook, tmp_path):
    job = _job(sample_workbook, tmp_path / "输出", mode="selected")
    job = SplitJob(
        input_file=job.input_file,
        output_dir=job.output_dir,
        sheet_configs=job.sheet_configs,
        split_mode=job.split_mode,
        selected_split_values=job.selected_split_values,
        output_types=("formula",),
    )

    summary = SplitEngine().execute(job)

    assert summary.total_files == 1
    assert [item.output_type for item in summary.results[0].output_files] == [
        "formula"
    ]
    workbook = load_workbook(_output_file(summary.results[0], "formula"))
    assert workbook["人员"]["D3"].value == "=C3*2"
    workbook.close()


def test_engine_can_export_values_version_only(sample_workbook, tmp_path):
    job = _job(sample_workbook, tmp_path / "输出", mode="selected")
    job = SplitJob(
        input_file=job.input_file,
        output_dir=job.output_dir,
        sheet_configs=job.sheet_configs,
        split_mode=job.split_mode,
        selected_split_values=job.selected_split_values,
        output_types=("values",),
    )

    summary = SplitEngine().execute(job)

    assert summary.total_files == 1
    assert [item.output_type for item in summary.results[0].output_files] == [
        "values"
    ]
    workbook = load_workbook(_output_file(summary.results[0], "values"))
    assert workbook["人员"]["D3"].value is None
    workbook.close()


def test_engine_reports_monotonic_progress_to_completion(sample_workbook, tmp_path):
    events = []

    SplitEngine().execute(
        _job(sample_workbook, tmp_path / "输出", mode="selected"),
        progress_callback=lambda percent, message: events.append((percent, message)),
    )

    percentages = [percent for percent, _ in events]
    assert percentages[0] == 0
    assert percentages[-1] == 100
    assert percentages == sorted(percentages)
    assert all(message for _, message in events)
