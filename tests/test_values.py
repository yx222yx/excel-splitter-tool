from excel_splitter.models import SheetConfig
from excel_splitter.values import extract_split_values, normalize_split_value


def test_normalize_split_value_strips_ascii_and_full_width_spaces():
    assert normalize_split_value("  \u3000临床部\u3000 ") == "临床部"
    assert normalize_split_value(None) is None
    assert normalize_split_value("   ") is None


def test_extract_split_values_returns_normalized_union_in_source_order(sample_workbook):
    configs = (
        SheetConfig("人员", 2, 1, "A - 部门"),
        SheetConfig("项目", 1, 1, "A - 所属部门"),
    )

    values, empty_rows, warnings = extract_split_values(sample_workbook, configs)

    assert values == ["临床部", "研发部", "市场部"]
    assert empty_rows == {"人员": 1, "项目": 0}
    assert warnings == []


def test_extract_split_values_uses_cached_formula_results(formula_workbook):
    configs = (SheetConfig("公式数据", 1, 1, "A - 部门公式"),)

    values, empty_rows, _ = extract_split_values(formula_workbook, configs)

    assert values == ["临床部", "研发部"]
    assert empty_rows == {"公式数据": 1}
