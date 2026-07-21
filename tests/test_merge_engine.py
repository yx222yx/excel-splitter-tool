from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from excel_splitter.merge_engine import MergeEngine
from excel_splitter.merge_models import MergeJob, MergeSheetConfig
from excel_splitter.merge_planning import build_merge_plan


def _make_workbook(path: Path, sheets: dict[str, list[list]]) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name, rows in sheets.items():
        sheet = workbook.create_sheet(sheet_name)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    workbook.close()
    return path


def _read_rows(path: Path, sheet_name: str) -> list[list]:
    workbook = load_workbook(path)
    rows = [list(row) for row in workbook[sheet_name].iter_rows(values_only=True)]
    workbook.close()
    return rows


def test_basic_merge_two_files(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "工时"], ["张三", 8], ["李四", 7]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名", "工时"], ["王五", 6]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    summary = MergeEngine().execute(job)

    assert summary.errors == []
    assert summary.total_rows == 3
    result = summary.results[0]
    assert result.sheet_name == "工时"
    assert result.merged_rows == 3
    assert result.source_rows == {"部门A.xlsx": 2, "部门B.xlsx": 1}
    assert _read_rows(output, "工时") == [
        ["姓名", "工时"],
        ["张三", 8],
        ["李四", 7],
        ["王五", 6],
    ]


def test_merge_aligns_by_header_name_when_column_order_differs(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "部门", "工时"], ["张三", "临床部", 8]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["工时", "姓名", "部门"], [6, "王五", "研发部"]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    summary = MergeEngine().execute(job)

    assert summary.errors == []
    assert _read_rows(output, "工时") == [
        ["姓名", "部门", "工时"],
        ["张三", "临床部", 8],
        ["王五", "研发部", 6],
    ]


def test_merge_unions_fields_and_reports_missing_and_extra(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "工时"], ["张三", 8]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名", "项目"], ["王五", "项目X"]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    summary = MergeEngine().execute(job)

    result = summary.results[0]
    # 缺失字段相对并集计算：B 缺「工时」，A 也缺 B 多出的「项目」
    assert result.missing_fields == {
        "部门A.xlsx": ["项目"],
        "部门B.xlsx": ["工时"],
    }
    assert result.extra_fields == {"部门B.xlsx": ["项目"]}
    assert any("部门B.xlsx" in w and "工时" in w for w in summary.warnings)
    assert _read_rows(output, "工时") == [
        ["姓名", "工时", "项目"],
        ["张三", 8, None],
        ["王五", None, "项目X"],
    ]


def test_merge_skips_file_missing_sheet_with_warning(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {
            "工时": [["姓名"], ["张三"]],
            "汇总": [["指标"], [10]],
        },
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名"], ["王五"]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"), MergeSheetConfig("汇总")),
    )

    summary = MergeEngine().execute(job)

    assert summary.errors == []
    assert any("部门B.xlsx" in w and "汇总" in w for w in summary.warnings)
    assert _read_rows(output, "汇总") == [["指标"], [10]]
    assert _read_rows(output, "工时") == [["姓名"], ["张三"], ["王五"]]


def test_merge_keeps_title_area_from_first_file_when_header_row_late(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["2026 年 7 月工时汇总", None], ["姓名", "工时"], ["张三", 8]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["部门B 自己的标题", None], ["姓名", "工时"], ["王五", 6]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时", header_row=2),),
    )

    summary = MergeEngine().execute(job)

    assert summary.errors == []
    assert _read_rows(output, "工时") == [
        ["2026 年 7 月工时汇总", None],
        ["姓名", "工时"],
        ["张三", 8],
        ["王五", 6],
    ]


def test_merge_skips_blank_rows(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "工时"], ["张三", 8], [None, None], [" ", None], ["李四", 7]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名", "工时"], ["王五", 6]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    summary = MergeEngine().execute(job)

    assert summary.total_rows == 3
    assert _read_rows(output, "工时") == [
        ["姓名", "工时"],
        ["张三", 8],
        ["李四", 7],
        ["王五", 6],
    ]


def test_merge_adds_source_column_when_enabled(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名"], ["张三"]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名"], ["王五"]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
        include_source_column=True,
    )

    summary = MergeEngine().execute(job)

    assert summary.errors == []
    assert _read_rows(output, "工时") == [
        ["姓名", "来源文件"],
        ["张三", "部门A"],
        ["王五", "部门B"],
    ]


def test_merge_reuses_provided_plan(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "工时"], ["张三", 8]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["工时", "姓名"], [6, "王五"]]},
    )
    configs = (MergeSheetConfig("工时"),)
    plan = build_merge_plan((file_a, file_b), configs)
    assert plan.sheets[0].union_headers == ["姓名", "工时"]
    assert plan.sheets[0].base_file == "部门A.xlsx"
    assert plan.sheets[0].missing_files == []

    output_with_plan = tmp_path / "带plan.xlsx"
    output_without_plan = tmp_path / "不带plan.xlsx"
    summary_with_plan = MergeEngine().execute(
        MergeJob(
            input_files=(file_a, file_b),
            output_file=output_with_plan,
            sheet_configs=configs,
        ),
        plan=plan,
    )
    summary_without_plan = MergeEngine().execute(
        MergeJob(
            input_files=(file_a, file_b),
            output_file=output_without_plan,
            sheet_configs=configs,
        )
    )

    assert summary_with_plan.total_rows == summary_without_plan.total_rows == 2
    assert _read_rows(output_with_plan, "工时") == _read_rows(
        output_without_plan, "工时"
    )


def test_merge_reports_progress_in_order(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名"], ["张三"]]},
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名"], ["王五"]]},
    )
    events: list[tuple[int, str]] = []
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=tmp_path / "合并结果.xlsx",
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    MergeEngine().execute(job, progress_callback=lambda p, m: events.append((p, m)))

    percents = [percent for percent, _ in events]
    assert percents[0] == 0
    assert percents[-1] == 100
    assert percents == sorted(percents)


def test_merge_copies_header_style_and_column_width_from_first_file(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {"工时": [["姓名", "工时"], ["张三", 8]]},
    )
    workbook = load_workbook(file_a)
    sheet = workbook["工时"]
    sheet["A1"].font = Font(bold=True)
    sheet["B1"].font = Font(bold=True)
    sheet.column_dimensions["A"].width = 21
    workbook.save(file_a)
    workbook.close()
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {"工时": [["姓名", "工时"], ["王五", 6]]},
    )
    output = tmp_path / "合并结果.xlsx"
    job = MergeJob(
        input_files=(file_a, file_b),
        output_file=output,
        sheet_configs=(MergeSheetConfig("工时"),),
    )

    MergeEngine().execute(job)

    workbook = load_workbook(output)
    merged = workbook["工时"]
    assert merged["A1"].font.bold
    assert merged.column_dimensions["A"].width == 21
    workbook.close()


def _merge_job(files, output, sheet_names=("汇总",), **overrides):
    kwargs = {
        "input_files": tuple(files),
        "output_file": output,
        "sheet_configs": tuple(MergeSheetConfig(name) for name in sheet_names),
    }
    kwargs.update(overrides)
    return MergeJob(**kwargs)


def test_merge_skips_identical_sheet_by_default(tmp_path):
    identical = {"汇总": [["指标", "说明"], [100, "相同内容"], [200, "完全一致"]]}
    file_a = _make_workbook(tmp_path / "部门A.xlsx", identical)
    file_b = _make_workbook(tmp_path / "部门B.xlsx", identical)
    output = tmp_path / "合并结果.xlsx"

    summary = MergeEngine().execute(_merge_job([file_a, file_b], output))

    result = summary.results[0]
    assert result.merged_rows == 2
    assert result.source_rows == {"部门A.xlsx": 2}
    assert result.skipped_duplicates == {"部门B.xlsx": "部门A.xlsx"}
    assert any("部门B.xlsx" in w and "完全相同" in w for w in summary.warnings)
    assert _read_rows(output, "汇总") == [
        ["指标", "说明"],
        [100, "相同内容"],
        [200, "完全一致"],
    ]


def test_merge_writes_duplicates_when_skip_disabled(tmp_path):
    identical = {"汇总": [["指标"], [100], [200]]}
    file_a = _make_workbook(tmp_path / "部门A.xlsx", identical)
    file_b = _make_workbook(tmp_path / "部门B.xlsx", identical)
    output = tmp_path / "合并结果.xlsx"

    summary = MergeEngine().execute(
        _merge_job([file_a, file_b], output, skip_duplicate_sheets=False)
    )

    result = summary.results[0]
    assert result.merged_rows == 4
    assert result.skipped_duplicates == {}
    assert _read_rows(output, "汇总") == [["指标"], [100], [200], [100], [200]]


def test_merge_skips_only_the_identical_sheet(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx",
        {
            "汇总": [["指标"], [100]],
            "工时": [["姓名"], ["张三"]],
        },
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx",
        {
            "汇总": [["指标"], [100]],
            "工时": [["姓名"], ["王五"]],
        },
    )
    output = tmp_path / "合并结果.xlsx"

    summary = MergeEngine().execute(
        _merge_job([file_a, file_b], output, sheet_names=("汇总", "工时"))
    )

    results = {r.sheet_name: r for r in summary.results}
    assert results["汇总"].skipped_duplicates == {"部门B.xlsx": "部门A.xlsx"}
    assert results["汇总"].merged_rows == 1
    assert results["工时"].skipped_duplicates == {}
    assert results["工时"].merged_rows == 2


def test_merge_keeps_first_of_three_files_when_two_identical(tmp_path):
    file_a = _make_workbook(tmp_path / "部门A.xlsx", {"汇总": [["指标"], [100]]})
    file_b = _make_workbook(tmp_path / "部门B.xlsx", {"汇总": [["指标"], [200]]})
    file_c = _make_workbook(tmp_path / "部门C.xlsx", {"汇总": [["指标"], [200]]})
    output = tmp_path / "合并结果.xlsx"

    summary = MergeEngine().execute(_merge_job([file_a, file_b, file_c], output))

    result = summary.results[0]
    assert result.merged_rows == 2
    assert result.source_rows == {"部门A.xlsx": 1, "部门B.xlsx": 1}
    assert result.skipped_duplicates == {"部门C.xlsx": "部门B.xlsx"}
    assert _read_rows(output, "汇总") == [["指标"], [100], [200]]


def test_merge_does_not_treat_different_column_order_as_duplicate(tmp_path):
    file_a = _make_workbook(
        tmp_path / "部门A.xlsx", {"工时": [["姓名", "工时"], ["张三", 8]]}
    )
    file_b = _make_workbook(
        tmp_path / "部门B.xlsx", {"工时": [["工时", "姓名"], [8, "张三"]]}
    )
    output = tmp_path / "合并结果.xlsx"

    summary = MergeEngine().execute(_merge_job([file_a, file_b], output, sheet_names=("工时",)))

    result = summary.results[0]
    assert result.skipped_duplicates == {}
    assert result.merged_rows == 2
