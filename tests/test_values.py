from excel_splitter.values import normalize_split_value


def test_normalize_split_value_strips_ascii_and_full_width_spaces():
    assert normalize_split_value("  \u3000临床部\u3000 ") == "临床部"
    assert normalize_split_value(None) is None
    assert normalize_split_value("   ") is None
