"""Command line entry point: ``excel-audit file.xlsx -o report.md``.

Standard-library argparse only, so installing the tool needs just
openpyxl (see README). No Excel installation is required.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl

from .checks import audit_workbook
from .report import render_markdown


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="excel-audit",
        description="Check an Excel workbook for broken refs, inconsistent "
        "column formulas and volatile functions, and write a Markdown report.",
    )
    parser.add_argument("workbook", help="Path to a .xlsx workbook")
    parser.add_argument(
        "-o",
        "--output",
        default="report.md",
        help="Where to write the Markdown report (default: report.md)",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print the report instead of writing a file",
    )
    args = parser.parse_args(argv)

    path = Path(args.workbook)
    if not path.is_file():
        print(f"error: no such file: {path}", file=sys.stderr)
        return 2
    try:
        workbook = openpyxl.load_workbook(path, data_only=False)
    except Exception as exc:  # corrupt file, wrong format, unreadable
        print(f"error: could not open {path}: {exc}", file=sys.stderr)
        return 2

    result = audit_workbook(workbook)
    text = render_markdown(str(path), result)
    if args.stdout:
        print(text)
    else:
        Path(args.output).write_text(text, encoding="utf-8")
        print(
            f"{len(result['broken_refs'])} broken ref(s), "
            f"{len(result['inconsistent_columns'])} inconsistent column(s), "
            f"{result['volatile_total']} volatile call(s) "
            f"-> {args.output}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
