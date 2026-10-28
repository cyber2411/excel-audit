# Contributing to excel-audit

Small, focused pull requests beat big ones. One check, one fix, one docs
improvement at a time.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .\.venv\Scripts\Activate.ps1
pip install openpyxl
```

No other dependencies.

## Before you open a PR

1. Add or update a test in `tests/` for the behaviour you changed.
2. Run `python -m unittest discover -s tests -v` — all tests must pass.
3. Run the tool on `samples/bakery-demo.xlsx` and check the report still
   reads sensibly.
4. Keep the "honesty box" in the README truthful: if your change adds a
   limitation, document it there.

## Good first issues

The five v0.1 issues on GitHub are the roadmap. If you are unsure which
one to take, open an issue and ask — a maintainer will point you at one
that fits.
