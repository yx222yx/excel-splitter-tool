from openpyxl import load_workbook
import pytest

from excel_splitter.models import SheetConfig
from excel_splitter.planning import build_split_plan


def _configs():
    return (
        SheetConfig(
            "人员归属",
            1,
            1,
            "A - 二级团队",
            mode="reference",
            key_column_idx=2,
            key_column_label="B - 姓名",
        ),
        SheetConfig(
            "工时明细",
            1,
            None,
            mode="linked",
            key_column_idx=1,
            key_column_label="A - 员工姓名",
        ),
        SheetConfig("说明", 1, None, mode="full"),
    )


def test_build_split_plan_maps_reference_values_to_keys(linked_workbook):
    plan = build_split_plan(linked_workbook, _configs())

    assert plan.values == ["团队甲", "团队乙"]
    assert plan.keys_by_value == {
        "团队甲": {"张三", "李四"},
        "团队乙": {"王五"},
    }
    assert plan.all_keys == {"张三", "李四", "王五"}


def test_build_split_plan_rejects_key_assigned_to_multiple_values(linked_workbook):
    workbook = load_workbook(linked_workbook)
    workbook["人员归属"].append(["团队乙", "张三", 400])
    workbook.save(linked_workbook)
    workbook.close()

    with pytest.raises(ValueError, match="张三.*团队甲.*团队乙"):
        build_split_plan(linked_workbook, _configs())


def test_build_split_plan_returns_complete_copy_value_when_all_sheets_are_full(
    linked_workbook,
):
    configs = (
        SheetConfig("人员归属", 1, None, mode="full"),
        SheetConfig("说明", 1, None, mode="full"),
    )

    plan = build_split_plan(linked_workbook, configs)

    assert plan.values == ["完整表"]
    assert plan.keys_by_value == {}


def test_build_split_plan_requires_reference_for_linked_sheet(linked_workbook):
    configs = (
        SheetConfig("工时明细", 1, None, mode="linked", key_column_idx=1),
    )

    with pytest.raises(ValueError, match="基准 Sheet"):
        build_split_plan(linked_workbook, configs)
