import json

from excel_splitter.cli import main


def test_cli_inspect_prints_sheet_names(sample_workbook, capsys):
    exit_code = main(["inspect", str(sample_workbook)])

    assert exit_code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["sheets"] == ["人员", "项目", "说明"]


def test_cli_split_exports_selected_value(sample_workbook, tmp_path, capsys):
    exit_code = main(
        [
            "split",
            str(sample_workbook),
            "--sheet",
            "人员:2:1",
            "--sheet",
            "项目:1:1",
            "--value",
            "临床部",
            "--output-dir",
            str(tmp_path / "out"),
        ]
    )

    assert exit_code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["total_files"] == 2
    assert (tmp_path / "out" / "人员数据_临床部_公式版.xlsx").exists()
    assert (tmp_path / "out" / "人员数据_临床部_结果值版.xlsx").exists()
