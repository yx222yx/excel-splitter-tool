from __future__ import annotations

import argparse
from dataclasses import asdict
import json
from pathlib import Path
import sys

from .engine import SplitEngine
from .models import SheetConfig, SplitJob
from .preview import list_sheet_names


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excel-splitter")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="列出工作簿中的 sheet")
    inspect_parser.add_argument("input_file", type=Path)

    split_parser = subparsers.add_parser("split", help="按字段值拆分工作簿")
    split_parser.add_argument("input_file", type=Path)
    split_parser.add_argument(
        "--sheet",
        action="append",
        required=True,
        metavar="SHEET:HEADER:COLUMN",
        help="可重复指定，例如 人员:2:1",
    )
    split_parser.add_argument("--value", action="append", default=[])
    split_parser.add_argument("--output-dir", type=Path, required=True)
    split_parser.add_argument(
        "--filename-template", default="{original_name}_{split_value}_{output_type}"
    )
    split_parser.add_argument("--overwrite", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "inspect":
            payload = {"input_file": str(args.input_file), "sheets": list_sheet_names(args.input_file)}
        else:
            configs = tuple(_parse_sheet_config(value) for value in args.sheet)
            job = SplitJob(
                input_file=args.input_file,
                output_dir=args.output_dir,
                sheet_configs=configs,
                split_mode="selected" if args.value else "all",
                selected_split_values=tuple(args.value),
                filename_template=args.filename_template,
                overwrite=args.overwrite,
            )
            payload = asdict(SplitEngine().execute(job))
        print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))
        return 0
    except (OSError, ValueError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


def _parse_sheet_config(value: str) -> SheetConfig:
    try:
        sheet_name, header_row, column_index = value.rsplit(":", 2)
        return SheetConfig(
            sheet_name=sheet_name,
            header_row=int(header_row),
            split_column_idx=int(column_index),
        )
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"无效 sheet 配置 {value!r}，应为 SHEET:HEADER:COLUMN"
        ) from exc


if __name__ == "__main__":
    raise SystemExit(main())
