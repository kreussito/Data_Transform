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

Everything the tool needs is in the repository plus **one** third-party package. No
database, no service, no build step, no configuration file — a clone and an install.

**Prerequisites:** Python 3.10 or newer and git. Nothing else — in particular **Excel is
not required**, on any platform: the tool reads and writes `.xlsx` directly and caches a
value for every formula it emits, so the figures are readable without a recalculation.

### Windows

Install Python from [python.org](https://www.python.org/downloads/windows/) or the
Microsoft Store, ticking **"Add python.exe to PATH"**. Then, in PowerShell:

```powershell
git clone https://github.com/kreussito/Data_Transform.git
cd Data_Transform

py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
```

If PowerShell refuses to run the activation script — *"running scripts is disabled on
this system"* — that is Windows' execution policy, not a problem with this repository:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Or skip activation entirely and call the interpreter in the environment directly:
`.venv\Scripts\python.exe -m datatransform ...`.

**On a locked-down machine with no git and no internet**, copy the folder across and use
the bundled dependency instead:

```powershell
py -m venv .venv
.venv\Scripts\python.exe -m pip install --no-index --find-links vendor openpyxl
.venv\Scripts\python.exe -m datatransform "C:\path\to\Intake.xlsx"
```

(`pip download openpyxl -d vendor` on a connected machine first, then carry `vendor/`
along.) Everything else the tool uses is the standard library.

### macOS / Linux

```bash
git clone https://github.com/kreussito/Data_Transform.git
cd Data_Transform

python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
```

### Proving it works before trusting it with real data

```bash
pytest                                # 440 tests — the specification's rules
python -m datatransform Intake_FireCatFull_v1.xlsx -o out.xlsx
```

The second command writes the transformed workbook and two logs and prints a per-sheet
summary. If both succeed, the machine is set up correctly. `pytest` includes a
`test_portability.py` that runs the CLI in a subprocess under a Windows ANSI code page,
so a machine that would have failed on encoding fails the test suite first.

Without the dev extra, `pip install -e .` (or just `pip install openpyxl`) is enough to
run the tool; the extra only adds `pytest` and `ruff`.

### Things that differ on Windows, and are handled

| | |
|---|---|
| **Redirected output** | `stdout` falls back to cp1252 when piped to a file, and neither `⟦` nor `→` exists there. The CLI puts its streams into UTF-8 before printing anything; where a console genuinely cannot, characters are shown as `\u27e6` rather than replaced by `?` — ugly, but it says something was there |
| **Log files** | Written UTF-8 explicitly, not in the machine's code page, so a log written on Windows reads the same everywhere |
| **Paths with spaces** | `C:\Users\...\My Documents\` works; quote the argument as usual. Covered by a test |
| **Where output lands** | Beside the *source* workbook when `-o` is omitted, not in the current directory — which on a mapped network drive may not be writable |
| **Excel** | Not needed. Nor is LibreOffice |

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
python tools/build_intake.py         # the other six
python tools/build_mexico.py         # Intake_Mexico_v1.xlsx — the cat aggregates
python tools/transpose_pack.py       # the complete pack, flipped on its side
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
| `Intake_FireCatFull_v1.xlsx` | **The complete treaty** — Fire + EQ + Wind, every dataset `01`–`09` |
| `Intake_FireCatFull_Transposed_v1.xlsx` | The same pack **on its side** — every dataset transposed |
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
