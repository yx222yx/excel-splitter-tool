from excel_splitter.preview import list_sheet_names, preview_sheet


def test_list_sheet_names_returns_workbook_order(sample_workbook):
    assert list_sheet_names(sample_workbook) == ["人员", "项目", "说明"]


def test_preview_suggests_header_and_builds_unique_column_options(sample_workbook):
    preview = preview_sheet(sample_workbook, "人员", max_rows=4)

    assert preview.suggested_header_row == 2
    assert preview.rows[0][0] == "人员奖金明细"
    assert [option.label for option in preview.columns_for_header(2)] == [
        "A - 部门",
        "B - 姓名",
        "C - 金额",
        "D - 计算值",
    ]


def test_preview_displays_cached_formula_results(formula_workbook):
    preview = preview_sheet(formula_workbook, "公式数据", max_rows=3)

    assert preview.rows[1] == ("临床部", "张三", 2)
    assert preview.rows[2] == ("研发部", "李四", 4)


def test_preview_supports_paged_rows(formula_workbook):
    preview = preview_sheet(formula_workbook, "公式数据", start_row=2, max_rows=2)

    assert preview.start_row == 2
    assert preview.end_row == 3
    assert preview.total_rows == 4
    assert preview.has_more is True
    assert preview.rows == (
        ("临床部", "张三", 2),
        ("研发部", "李四", 4),
    )


def test_preview_suggests_header_within_first_ten_rows(late_header_workbook):
    preview = preview_sheet(late_header_workbook, "晚表头", max_rows=10)

    assert preview.suggested_header_row == 8
    assert [option.label for option in preview.columns_for_header(8)] == [
        "A - 部门",
        "B - 姓名",
        "C - 金额",
    ]


def test_preview_suggests_header_within_first_fifteen_rows(
    fourteenth_header_workbook,
):
    preview = preview_sheet(fourteenth_header_workbook, "更晚表头", max_rows=15)

    assert preview.suggested_header_row == 14
    assert [option.label for option in preview.columns_for_header(14)] == [
        "A - 部门",
        "B - 姓名",
        "C - 金额",
    ]
