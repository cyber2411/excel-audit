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

## Sheet stats

| Sheet | Rows used | Columns used | Formulas | Volatile calls |
| ----- | --------- | ------------ | -------- | -------------- |
| Sales | 5 | 4 | 4 | 0 |
| Summary | 5 | 2 | 4 | 2 |
