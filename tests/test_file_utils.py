from pathlib import Path

from excel_splitter.file_utils import render_filename, resolve_output_path


def test_render_filename_replaces_invalid_windows_characters():
    filename = render_filename(
        "{original_name}_{split_value}",
        original_name="人员:数据",
        split_value="临床/一部",
    )

    assert filename == "人员_数据_临床_一部.xlsx"


def test_resolve_output_path_adds_sequence_when_file_exists(tmp_path: Path):
    first = tmp_path / "结果.xlsx"
    first.touch()

    assert resolve_output_path(first, overwrite=False) == tmp_path / "结果(1).xlsx"
    assert resolve_output_path(first, overwrite=True) == first


def test_render_filename_appends_output_type_when_template_omits_it():
    assert render_filename(
        "{original_name}_{split_value}",
        original_name="人员数据",
        split_value="临床部",
        output_type="公式版",
    ) == "人员数据_临床部_公式版.xlsx"


def test_render_filename_supports_output_type_placeholder():
    assert render_filename(
        "{original_name}-{output_type}-{split_value}",
        original_name="人员数据",
        split_value="临床部",
        output_type="结果值版",
    ) == "人员数据-结果值版-临床部.xlsx"
