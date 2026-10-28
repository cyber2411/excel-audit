# First 5 issues — ready to create on GitHub (one approval card each or batch)

Repo: github.com/cyber2411/excel-audit · Status: skeleton code already implements a v0.1 of each check; these issues are the public roadmap / good-first-issues.

1. **feat: detect #REF! and missing-sheet references**
   Body: Flag any formula containing `#REF!` and any formula that references a sheet name not present in the workbook. External-workbook refs (`[Other.xlsx]Sheet!A1`) must be skipped to avoid false positives. Acceptance: broken refs appear in the Markdown report with sheet + cell + formula. Sample: Summary!B2/B3 in samples/bakery-demo.xlsx.

2. **feat: flag inconsistent formulas within a column**
   Body: Normalize each formula (cell refs → CELL) and group by column. Flag a column only when one pattern repeats (≥2 cells) and another pattern also occurs — the copy-paste-bug shape. A column of unrelated one-off formulas must NOT be flagged. Acceptance: Sales column D in the sample flags exactly once.

3. **feat: count volatile functions (INDIRECT/OFFSET/TODAY etc.)**
   Body: Count calls to CELL, INDIRECT, INFO, NOW, OFFSET, RAND, RANDBETWEEN, TODAY across all sheets; show per-function counts in the report. Acceptance: sample shows INDIRECT 1, TODAY 1.

4. **feat: sheet-level stats table in Markdown report**
   Body: Per sheet: rows used, columns used, formula count, volatile-call count, rendered as a Markdown table. Acceptance: sample report stats match openpyxl max_row/max_column.

5. **docs: README with install + 3 example runs**
   Body: README must have one-line pitch, install steps (incl. Windows venv steps), 3 example commands with real sample output, the checks table, an honesty box of limitations, contributing pointer and license. Acceptance: a new user can install and run the sample in under 5 minutes.
