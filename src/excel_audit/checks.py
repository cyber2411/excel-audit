"""The four v0.1 checks.

Everything here works on formulas as stored in the file (openpyxl,
``data_only=False``), so no Excel installation is needed. A check never
crashes the whole audit: one bad sheet is reported, the rest still run.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict

# Functions that force Excel to recalculate on every change. Counting them
# is a cheap proxy for "this workbook may feel slow".
VOLATILE_FUNCTIONS = {
    "CELL",
    "INDIRECT",
    "INFO",
    "NOW",
    "OFFSET",
    "RAND",
    "RANDBETWEEN",
    "TODAY",
}

_SHEET_REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_][A-Za-z0-9_.]*))!")
_CELL_REF_RE = re.compile(r"\$?[A-Z]{1,3}\$?\d+")
_FUNCTION_RE = re.compile(r"([A-Z][A-Z0-9.]*)\s*\(")


def iter_formula_cells(worksheet):
    """Yield (coordinate, formula) for every formula cell in a sheet."""
    for row in worksheet.iter_rows():
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                yield cell.coordinate, value


def find_broken_refs(workbook):
    """Broken references: ``#REF!`` and references to missing sheets.

    Returns a list of dicts with sheet / cell / formula / reason.
    External-workbook refs (``[Other.xlsx]Sheet!A1``) are skipped: the
    other file is not available to check, so flagging them would be a
    false positive.
    """
    sheet_names = set(workbook.sheetnames)
    issues = []
    for ws in workbook.worksheets:
        for coord, formula in iter_formula_cells(ws):
            if "#REF!" in formula:
                issues.append(
                    {
                        "sheet": ws.title,
                        "cell": coord,
                        "formula": formula,
                        "reason": "contains #REF! (a deleted row, column or sheet)",
                    }
                )
                continue
            if formula.startswith("=["):
                continue
            for quoted, plain in _SHEET_REF_RE.findall(formula):
                name = quoted or plain
                if name not in sheet_names:
                    issues.append(
                        {
                            "sheet": ws.title,
                            "cell": coord,
                            "formula": formula,
                            "reason": f"references missing sheet '{name}'",
                        }
                    )
                    break
    return issues


def normalize_formula(formula):
    """Replace every cell reference with ``CELL``.

    ``=B2*C2`` and ``=B3*C3`` both become ``=CELL*CELL``, so a column
    whose rows were copied correctly collapses to one pattern, while a
    row typed by hand (``=B5+C5``) stands out.
    """
    return _CELL_REF_RE.sub("CELL", formula)


def find_inconsistent_columns(workbook):
    """Columns where one formula pattern dominates but a row differs.

    A column is flagged only when one pattern appears at least twice
    and another pattern also appears — that is the copy-paste shape
    (most rows copied correctly, one row typed by hand). A column of
    unrelated one-off formulas (e.g. a summary/value column where every
    row is a different calculation) is NOT flagged: flagging it would be
    a false positive.
    """
    issues = []
    for ws in workbook.worksheets:
        by_column = defaultdict(Counter)
        examples = {}
        for coord, formula in iter_formula_cells(ws):
            column = "".join(ch for ch in coord if ch.isalpha())
            pattern = normalize_formula(formula)
            by_column[column][pattern] += 1
            examples.setdefault((column, pattern), (coord, formula))
        for column, patterns in sorted(by_column.items()):
            if len(patterns) > 1 and patterns.most_common(1)[0][1] >= 2:
                detail = "; ".join(
                    f"{pattern} (x{count}, e.g. {examples[(column, pattern)][0]})"
                    for pattern, count in patterns.most_common()
                )
                issues.append(
                    {
                        "sheet": ws.title,
                        "cell": f"column {column}",
                        "formula": detail,
                        "reason": "inconsistent formulas within one column",
                    }
                )
    return issues


def count_volatile_functions(workbook):
    """Count volatile-function calls per function name."""
    counts = Counter()
    for ws in workbook.worksheets:
        for _coord, formula in iter_formula_cells(ws):
            for name in _FUNCTION_RE.findall(formula.upper()):
                if name in VOLATILE_FUNCTIONS:
                    counts[name] += 1
    return counts


def sheet_stats(workbook):
    """Per-sheet stats: used size, formula count, volatile count."""
    stats = []
    for ws in workbook.worksheets:
        formulas = list(iter_formula_cells(ws))
        volatile = sum(
            1
            for _coord, formula in formulas
            for name in _FUNCTION_RE.findall(formula.upper())
            if name in VOLATILE_FUNCTIONS
        )
        stats.append(
            {
                "sheet": ws.title,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "formula_count": len(formulas),
                "volatile_count": volatile,
            }
        )
    return stats


def audit_workbook(workbook):
    """Run all four checks and return one plain-dict result."""
    broken = find_broken_refs(workbook)
    inconsistent = find_inconsistent_columns(workbook)
    volatile = count_volatile_functions(workbook)
    stats = sheet_stats(workbook)
    return {
        "broken_refs": broken,
        "inconsistent_columns": inconsistent,
        "volatile_counts": dict(volatile),
        "volatile_total": sum(volatile.values()),
        "sheet_stats": stats,
        "issue_total": len(broken) + len(inconsistent),
    }
