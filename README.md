# excel-audit

**Check an Excel workbook for the bugs that hide in plain sight — then get a Markdown report you can actually act on.**

Small business spreadsheets grow by copy-paste. One row typed by hand, one deleted sheet, one `INDIRECT` too many — and suddenly totals quietly go wrong. `excel-audit` finds four things, fast, with no Excel installation needed:

| Check | What it catches |
| ----- | --------------- |
| Broken references | `#REF!` and formulas pointing at sheets that no longer exist |
| Inconsistent column formulas | A column where one row's formula differs from the rest (the classic copy-paste bug) |
| Volatile functions | `INDIRECT`, `OFFSET`, `TODAY`, `NOW`, `RAND` and friends — the functions that make big workbooks feel slow |
| Sheet stats | Per-sheet size, formula count and volatile-call count |

## Install

Requires Python 3.9+.

```bash
pip install openpyxl
pip install -e .          # from this repo, gives you the `excel-audit` command
```

Windows, step by step:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install openpyxl
pip install -e .
```

No other dependencies, no paid tools, no Excel needed. (The original plan used `click` for the CLI; this build uses the standard-library `argparse` instead, so installation is one package, not two.)

## Usage

```bash
excel-audit samples/bakery-demo.xlsx -o report.md
# 2 broken ref(s), 1 inconsistent column(s), 2 volatile call(s) -> report.md
```

Or print to the terminal instead of a file:

```bash
excel-audit my-workbook.xlsx --stdout
```

### Example output (from the included sample)

```markdown
# Excel audit: samples/bakery-demo.xlsx

**Summary:** 2 broken reference(s), 1 inconsistent column(s), 2 volatile function call(s).

## Broken references

| Sheet | Cell | Problem | Formula |
| ----- | ---- | ------- | ------- |
| Summary | B2 | references missing sheet 'Stock List' | `='Stock List'!A1` |
| Summary | B3 | contains #REF! (a deleted row, column or sheet) | `=SUM(#REF!)` |

## Inconsistent columns

| Sheet | Cell | Problem | Formula |
| ----- | ---- | ------- | ------- |
| Sales | column D | inconsistent formulas within one column | `=CELL*CELL (x3, e.g. D2); =CELL+CELL (x1, e.g. D5)` |

## Volatile functions

| Function | Count |
| -------- | ----- |
| INDIRECT | 1 |
| TODAY | 1 |
```

The sample workbook (`samples/bakery-demo.xlsx`, rebuilt any time with `python samples/make_sample.py`) has one deliberate fault per check, so you can see every check fire before trusting it on a real file.

## What it does NOT do (v0.1 honesty box)

- It reads formulas as stored in the file — it does not calculate values, so it cannot catch a formula that is *valid but logically wrong* (e.g. pointing at the wrong, existing column).
- External-workbook references (`[Other.xlsx]Sheet!A1`) are skipped, not flagged — the other file is not available to verify.
- `max_row` / `max_column` reflect the used range, which formatting alone can inflate.
- Not on PyPI yet. That comes after the tool passes zero-false-positive tests on real workbooks.

## Contributing

Good first issues are labelled on GitHub — start with the five v0.1 issues (broken refs, inconsistent columns, volatile counts, stats table, README examples). Run the tests before opening a pull request:

```bash
python -m unittest discover -s tests -v
```

## License

MIT — see [LICENSE](LICENSE).
