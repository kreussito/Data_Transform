# Data_Transform

Excel data to be filtered, extracted and finally transformed to standard format; each step to be logged and documented.

Reinsurance treaty submission workbooks arrive as a collage of sheets copied from
different original files. This tool extracts a declared subset of each sheet, transforms
it in two documented steps, and writes the result back beneath the original data so a
reviewer can follow every step without leaving Excel.

## The idea

**The workbook describes itself.** Sheet `00` declares *what* to extract, as a
vocabulary humans copy from. Markers in **column A** of each data sheet declare *where*
it lives. Nothing depends on cell position.

```
Column A of 01. History          Meaning
─────────────────────────────    ─────────────────────────────────────────────
Currency = USD                   an attribute of the sheet
H_Year basis = UW                the same, declared as a hypothesis
Header_1                         this row is the extraction row
Info_1 = L                       column L selects which rows are extracted
```

Records are marked with `=ROW()`, which is self-validating: if the value at row *r* is
not *r*, the sheet has been manipulated and the run stops. Totals and footnotes are
excluded simply by not being marked — no heuristics guess at intent.

## Usage

```bash
pip install openpyxl
python -m datatransform Intake_v1.xlsx -o output/Intake_v1_transformed.xlsx
```

The source workbook is never modified; it is the audit baseline. Output goes to a copy,
with two logs in `logs/` — one technical, one a plain-language record of every decision
and interpretation.

```bash
pytest tests/
```

## What's here

| | |
|---|---|
| `specification_v1.md` | The rules, in full |
| `Intake_v1.xlsx` | Reference workbook — sheet 00 plus dataset 01 in both orientations |
| `datatransform/` | Implementation |
| `tests/` | Test suite covering the specification's rules |

Dataset `01. History` is implemented end to end. Datasets `02`–`10` are not yet
specified; see §13 of the specification.
