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

## Getting it onto a machine

Everything the tool needs is in the repository plus **one** third-party package. There is
no database, no service, no build step and no configuration file — a clone and an install.

**Prerequisites:** Python 3.10 or newer (developed and tested on 3.11) and git.

```bash
git clone https://github.com/kreussito/Data_Transform.git
cd Data_Transform

# macOS / Linux
python3 -m venv .venv && source .venv/bin/activate

# Windows PowerShell
#   py -m venv .venv
#   .venv\Scripts\Activate.ps1

pip install -e ".[dev]"
```

Then prove it works before trusting it with real data:

```bash
pytest                               # 303 tests — the specification's rules
python -m datatransform Intake_v1.xlsx -o output/Intake_v1_transformed.xlsx
```

The second command writes the transformed workbook and two logs, and prints a per-sheet
summary. If both succeed, the machine is set up correctly.

Without the dev extra, `pip install -e .` (or just `pip install openpyxl`) is enough to
run the tool; the extra only adds `pytest` and `ruff`.

## Usage

```bash
python -m datatransform <workbook.xlsx> -o <output.xlsx> [--log-dir logs]
```

`pip install -e .` also puts a `datatransform` command on the path, so `datatransform
<workbook.xlsx>` works from any directory.

The source workbook is **never modified**; it is the audit baseline. Output goes to a
copy, with two logs in `logs/` — one technical, one a plain-language record of every
decision and interpretation.

### Regenerating the reference workbooks

```bash
python tools/build_v1.py             # Intake_v1.xlsx
python tools/build_intake.py         # the other five
python tools/build_mexico.py         # Intake_Mexico_v1.xlsx — the cat aggregates
python tools/demo_profile_shapes.py  # the three risk-profile presentations
```

Excel is not required at any point: the tool reads and writes `.xlsx` directly and
injects cached values for every formula it emits, so the figures are readable without a
recalculation.

## What's here

| | |
|---|---|
| `specification_v1.md` | The rules, in full |
| `Intake_v1.xlsx` | Fire treaty — datasets 00–05, both orientations |
| `Intake_Engineering_v1.xlsx` | Engineering — one per-risk section carrying both `03` and `04` |
| `Intake_EngineeringCombined_v1.xlsx` | The same treaty with both loss datasets in **one list** |
| `Intake_FireCat_v1.xlsx` | Fire + Nat Cat — the two cat sections share one sheet |
| `Intake_FireEQWind_v1.xlsx` | Fire + Earthquake + Hurricane — a sheet per section |
| `Intake_FireCatLosses_v1.xlsx` | Fire + Nat Cat carrying only `01`, `02` and `04`, split by EQ and Wind |
| `Intake_Mexico_v1.xlsx` | Mexican cat aggregates — datasets `06` / `07`, three versions per sheet |
| `datatransform/` | Implementation |
| `tools/` | Generators for the reference workbooks |
| `tests/` | Test suite covering the specification's rules |

## Sections

A treaty is one or more sections, declared in `⟦SECTIONS⟧` of sheet 00:

| Section | Kind | Datasets |
|---|---|---|
| `Fire` | per risk | 01, 02, 03, 05 |
| `Earthquake` | cat | 01, 02, 04, 05 |

**A section is a block.** Whether the blocks sit in one sheet or three is immaterial —
the two Nat Cat workbooks carry identical figures in the two layouts and are verified
to produce identical crosschecks. Rules are scoped by section kind and evaluated once
per applicable section, so a Fire treaty reports the cat rule as *not applicable* rather
than complaining about a sheet that correctly does not exist.

Datasets `00`–`05` are implemented. `06`–`10` are not yet specified; see §13 of the
specification, and §13.1 for improvements parked pending a decision.
