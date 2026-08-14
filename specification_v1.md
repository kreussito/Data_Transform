# Data_Transform — Specification v1

Extraction, filtering and transformation of reinsurance treaty submission data from
Excel, driven by metadata declared in the workbook itself.

| | |
|---|---|
| **Version** | 1 |
| **Status** | Design agreed for `00` and `01`; sheets `02`–`10` not yet specified |
| **Cadence** | Once per treaty, per year |
| **Reference workbook** | `Intake_v1.xlsx` |

---

## 1 · Purpose and principles

The tool reads a hand-assembled Excel workbook, extracts a declared subset of each
sheet, transforms it in two documented steps, and writes the result back into the same
sheet beneath the original data so that a human reviewer can follow every step without
leaving Excel.

Four principles govern the design.

**1. The workbook describes itself.** What to extract is declared in sheet `00`; where it
lives is declared by markers in column A of each data sheet. Nothing is inferred from
position.

**2. The human declares, the tool obeys.** Row selection, field selection and
interpretation are marked by the underwriter in the sheet. The tool does not guess which
rows are data and which are totals — it is told.

**3. Silence is never an outcome.** A declared field that is missing, an uncalculated
formula, an error cell, or an undeclared assumption produces an error or a logged
question. Nothing is skipped quietly.

**4. The evidence survives.** The source workbook is never modified. Every transformation
is written as a visible block with control sums that tie back to the step before it.

---

## 2 · Workbook inventory

| Sheet | Role | Status |
|---|---|---|
| `00. Nomenclature & Interdependencies` | Vocabulary and dataset register | Authored |
| `01. History` | Premium and loss history | Authored |
| `02. EPI Projections` | Estimated premium income | Authored |
| `03. Large Losses` | Individual large claims | Authored |
| `04. Cat Losses` | Catastrophe events | Authored |
| `05. Risk Profiles` | Banded exposure | Authored |
| `06. EQ Aggs` | Earthquake aggregates | Authored |
| `07. Wind Aggs` | Windstorm aggregates | Authored |
| `08 …` | Fire splits **or** Engineering — see §2.1 | Authored |
| `09. Rate Development` | Rate change history | Authored, optional |
| `10. Triangles` | Development triangles | Authored |
| `Exchange rates` | FX rates for conversion | Authored |
| `19. Assumptions & Queries` | Hypothesis register | **Generated** |
| `20. Summary` | Collected step-2 blocks | **Generated** |

The workbook is assembled by copying sheets in from original submission files. Expect
external links, stale formula caches, heterogeneous layouts and imported defined names.

### 2.1 Sheet 08 — two exclusive variants

Sheet `08` carries exactly one of:

- **Fire** — split by Res / Com / Ind **and** by B / C / BI
- **Engineering** (Projects & Renewables) — split by Res / Com / Ind

The variant is declared, not sniffed. A declaration that disagrees with the sheet
content is an interdependency failure.

### 2.2 Sheet matching

Sheets are matched by **numeric prefix first**, text second. Both sides are normalised
for whitespace, non-breaking spaces, quote style and case. `10. Triangles` and
`10. Triangels` therefore resolve to the same sheet.

---

## 3 · Sheet 00 — the nomenclature

Sheet `00` is a **dictionary for humans**, not an enforcement layer. It tells whoever
prepares a data sheet which names to write. It does not declare sheets required.

### 3.1 Layout

- Column **A** is empty. No markers live in `00`.
- **One row per dataset.**
- Column **B** — the exact sheet name.
- Column **C** — a short key, used in block IDs and log entries.
- Columns **D … M** — the **header labels**: the minimum information to extract.
- Columns **N** onward — the **attribute names** that sheet should declare.
- Row **4** declares the section boundaries (`Headers` above D, `Attributes` above N).

The fixed section boundary is necessary because datasets have ragged header counts; a
per-row boundary could not be located unambiguously.

Names are written **plain**, with no prefixes, because humans copy them verbatim into
their sheets. Any prefix in `00` would invite the very mismatch the vocabulary exists to
prevent.

### 3.2 Row 5 — dataset 01

| Cell | Value |
|---|---|
| B5 | `01. History` |
| C5 | `01 History` |
| D5 | `Year` |
| E5 | `Premium` |
| F5 | `Incurred Losses` |
| N5 | `Currency` |
| O5 | `Scale` |
| P5 | `Year basis` |
| Q5 | `Premium basis` |
| R5 | `Loss basis` |
| S5 | `PF transfer` |
| T5 | `As at` |

`Year` is deliberately unqualified. Whether it means underwriting or occurrence year is
an **attribute**, not part of the field name — see §5.

### 3.3 Vocabulary block

A `⟦VOCABULARY⟧` block lower in the sheet lists the permitted values per attribute. An
attribute value outside its vocabulary is an error, so that a typo becomes a failure
rather than a silently created new category.

| Attribute | Permitted values |
|---|---|
| `Year basis` | `UW` \| `Occurrence` |
| `Premium basis` | `GWP` \| `GNPI` \| `Written` \| `Earned` \| `Signed` |
| `Loss basis` | `Paid` \| `Incurred` |
| `PF transfer` | `with clean cut` \| `without clean cut` \| `none` |
| `Exposure basis` | `Sum Insured` \| `EML` \| `PML` \| `MPL` |
| `Scale` | `1` \| `1,000` \| `1,000,000` |
| `Currency` | ISO 4217 |

---

## 4 · Marker grammar

Every data sheet declares its own structure through markers in **column A**.

| Rule | |
|---|---|
| M1 | All markers live in column A of the data sheet |
| M2 | Markers are **found, not positioned** — any row, any order |
| M3 | Form is `Token` or `Token = value` |
| M4 | The `_i` suffix binds a marker to block *i* |
| M5 | An attribute without a suffix applies sheet-wide; with a suffix it overrides for that block |
| M6 | The `H_` prefix marks the value as a hypothesis rather than a given |
| M7 | **No `Header_i` anywhere in column A → nothing is extracted from that sheet** |

### 4.1 Marker reference

| Marker | Meaning |
|---|---|
| `Header_i` | Row-wise: this row is block *i*'s extraction row |
| `Header_i = <col>` | Transposed: this column is block *i*'s extraction column |
| `Info_i = <col>` | Row-wise: this column holds block *i*'s record selectors |
| `Info_i = <row>` | Transposed: this row holds block *i*'s record selectors |
| `Transpose_i` | Block *i* is transposed |
| `<Attribute> = <value>` | An attribute of the sheet or block |
| `H_<Attribute> = <value>` | The same, declared as a hypothesis |

---

## 5 · Attributes and hypotheses

### 5.1 Form

Attributes take the form `Name = Value`. The name must appear in `00`'s attribute
section for that dataset; the value must appear in the vocabulary block.

### 5.2 The `H_` prefix

`H_` marks a value the underwriter assumed rather than received. The tool records the
value either way and **auto-creates a register entry** for the hypothesis. Assumptions
are therefore documented where they are made, and the query list assembles itself.

```
Premium basis = GNPI              ← confirmed: taken from the submission
H_PF transfer = with clean cut    ← hypothesis: our reading, unverified
```

### 5.3 Confidence

Confidence is **derived from the markers**, never judged.

| State | Condition | Meaning |
|---|---|---|
| **Confirmed** | attribute present, no `H_` | given or verified |
| **Assumed** | attribute present with `H_` | a human considered it and made a call |
| **Open** | attribute listed in `00` but absent from the sheet | nobody decided; a default was applied |

`Open` ranks **worse than** `Assumed`: an undeclared attribute means the question was
never asked, which is more dangerous than a documented guess.

**A dataset's confidence is the worst confidence of any hypothesis touching it.**
Mechanical, conservative, and not open to averaging.

### 5.4 The register — sheet `19`

Entries come from three sources:

1. `H_` markers in the sheets — generated
2. Tool findings (missing attributes, failed interdependencies, undeclared columns) — generated
3. Narrative judgments that are not attribute-shaped — entered by hand

Filtering the register to `status ∈ {open, queried}`, sorted by impact, yields **the
query list for the broker** as a by-product of extraction.

### 5.5 Hypotheses to anticipate

Recurring assumptions that materially move numbers:

- **Year basis** — underwriting, occurrence or calendar
- **Premium basis** — GWP vs GNPI vs written vs earned vs signed
- **Loss basis** — paid vs incurred; gross or net of other reinsurance; ALAE/LAE; IBNR
- **Double counting** — whether large losses (`03`) and cat losses (`04`) are already inside history claims (`01`), or additional. The most common material misunderstanding in a submission.
- **Scope** — whole account vs single LOB; run-off portfolios; portfolio transfers with or without clean cut
- **Thresholds** — large loss threshold from ground up or excess of deductible; cat event definition and hours clause
- **Exposure** — sum insured vs EML/PML/MPL; CRESTA zone version
- **Dates** — as-at consistency across sheets
- **Triangles** — cumulative vs incremental; paid vs incurred; origin period definition
- **FX** — which rate, as at which date, applied to which sheets

---

## 6 · Blocks and orientation

A sheet may contain several independent blocks, and their orientations may differ. The
`_i` suffix binds `Header_i`, `Info_i` and `Transpose_i` together.

### 6.1 Row-wise

`Header_i` is a bare token in column A. **Its row is the extraction row.**

### 6.2 Transposed

`Transpose_i` is present, and `Header_i = <col>` names the extraction column.

Column A cannot encode a column by position, so in transposed mode `Header_i` and
`Info_i` take values instead of relying on their own position. This is the only
asymmetry between the two orientations.

### 6.3 The extraction row / column

The extraction row contains **only the declared labels**, each placed in the column its
data occupies. Blank cells elsewhere in that row are meaningful — they say "not
extracted".

This has three consequences:

- **The source's own header is never read.** It may use any naming whatever.
- **Ambiguity is resolved visually by a human.** If the source has two columns headed
  `Premium`, the underwriter writes `Premium` in the extraction row under the one they
  mean.
- **The sheet is self-evidencing.** Anyone opening it sees exactly which columns feed the
  extraction.

The extraction row is never itself a record; its selector cell stays empty.

---

## 7 · Record selection

| | Row-wise | Transposed |
|---|---|---|
| Selector location | column named by `Info_i` | row named by `Info_i` |
| Selector value | `=ROW()` | `=COLUMN()` |
| Integrity check | value must equal its own row | value must equal its own column |
| Provenance field | `Source row` | `Source column` |

- Selector filled → **extract**. Selector empty → **skip**.
- `Info_i` is a pure marker. It never appears as data, but is carried into step 1 as
  provenance.

### 7.1 Why `=ROW()` rather than a tick

A formula returning its own coordinate is **self-validating**. If the value at row *r*
is not *r*, the sheet has been manipulated in a way that broke the markers — rows
inserted or deleted after a paste-as-values, or a block copied from another sheet
carrying its old numbers. Given that this workbook is assembled by copying sheets
between files, that is not hypothetical. A plain `x` could never catch it.

It also renumbers itself when the extraction row is inserted, which a hard-coded number
would not.

### 7.2 The uncalculated-formula trap

Reading with `data_only=True` returns the *cached* result. If the workbook has not been
recalculated and saved since the sheets were copied in, there is no cache and the read
returns `None` — **indistinguishable from an empty cell**, which under the rules above
means "skip this row". An entire block could vanish with no error raised.

**Rule:** consult both handles. A cell that holds a formula in the formula handle but
returns `None` in the value handle is *"workbook not calculated"* — a hard stop, never a
skipped record.

---

## 8 · Field resolution

| Rule | |
|---|---|
| F1 | Fields resolve by **label match against `00`**, never by position |
| F2 | Match is strict, after `.strip()` on both sides, including non-breaking spaces |
| F3 | A duplicate label within the extraction row is an error |
| F4 | A label declared in `00` but absent from the extraction row is an error |
| F5 | Every column/row containing data but not extracted is **logged**, and raises a question |
| F6 | Merged cells are resolved before matching |
| F7 | **Only resolved addresses are read** |

F7 is the substance of the filter: the block is never read as a rectangle and then
trimmed. Resolution produces an address map, and only those addresses are ever touched.

F5 preserves the "a new column appeared and nobody noticed" catch that would otherwise
be lost by ignoring the source header.

### 8.1 Worked resolution — `01. History`, row-wise

| Declared in `00` | Found at | Resolves to |
|---|---|---|
| `Year` | B8 | column **B** |
| `Premium` | D8 | column **D** |
| `Incurred Losses` | G8 | column **G** |

Records are rows 9–13. Cells read: `B`, `D`, `G` × rows 9–13 = **15 cells, and no
others.**

### 8.2 Worked resolution — `01. History`, transposed

| Declared in `00` | Found at | Resolves to |
|---|---|---|
| `Year` | C8 | row **8** |
| `Premium` | C10 | row **10** |
| `Incurred Losses` | C13 | row **13** |

Records are columns D–H. Cells read: rows `8`, `10`, `13` × columns D–H = **the same 15
cells, rotated.**

Both orientations produce an identical step-1 block, differing only in whether
provenance reads `Source row` or `Source column`. That equivalence is the test that the
rule set holds.

---

## 9 · Output

### 9.1 Placement

| Rule | |
|---|---|
| O1 | The source workbook is never modified; all writing happens on a copy |
| O2 | The **step 1** block begins 3 blank rows below the **last non-empty row of the entire sheet** |
| O3 | The **step 2** block begins 3 blank rows below step 1 |
| O4 | Output is always written **row-wise**, whatever the source orientation |
| O5 | Every block carries a machine anchor so re-runs **replace** rather than stack |
| O6 | `20. Summary` collects **step-2 blocks only**, block by block, regenerated wholesale |

O2 uses the last non-empty row of the whole sheet, not the last record, so that
footnotes and totals below the data are never overwritten.

O5 is necessary because three blank rows are a *visual* separator, not a machine one —
real data blocks contain blank rows. Each generated block is tagged:

```
⟦DT:01 History:STEP1:v1⟧
```

### 9.2 Block content

**Step 1 — filtered and interpreted.** Assumptions in plain language, the count of
candidate and excluded records with reasons, the extracted data, provenance, and control
sums.

**Step 2 — normalised.** Sorted, columns reordered, calculations applied, with control
sums tying back to step 1.

### 9.3 Example — step 1 for `01. History`

| | Source row | Year | Premium | Incurred Losses |
|---|---|---|---|---|
| | 9 | 2021 | 15,900 | 14,930 |
| | 10 | 2022 | 16,740 | 10,030 |
| | 11 | 2023 | 17,520 | 12,660 |
| | 12 | 2024 | 18,390 | 11,700 |
| | 13 | 2025 | 19,200 | 10,250 |
| **Control** | *n = 5* | | **87,750** | **59,570** |
| **Check vs. source total row 14** | | | 87,750 ✓ | 59,570 ✓ |

The `Total` row deliberately excluded from extraction returns as an **independent
control**. The sheet's own arithmetic verifies our extraction, in front of the reviewer.

---

## 10 · Controls

| Rule | |
|---|---|
| C1 | Every block carries control sums: records in / out / excluded, plus per-measure totals |
| C2 | Control sums are written as **live Excel formulas** alongside the tool's expected value |
| C3 | Each step is labelled **value-preserving** (sort, reorder) or **value-changing** (filter, calculate) |
| C4 | Excluded totals are reused as independent checks |

C2 matters: a pasted number asks the reviewer to trust the tool, which is what the audit
trail exists to avoid. A live formula recomputes in front of them, and a mismatch
between formula and expected value flags a tool defect.

C3 turns control sums into a test. Sorting and column-swapping *cannot* change a total —
if they do, that is a bug. Filtering and calculation legitimately change totals, so those
require an explained delta.

### 10.1 Interdependencies

Cross-sheet reconciliation, each with an explicit tolerance because figures rounded to
thousands will never satisfy exact equality:

| Rule | Left | Right | Relation | Severity |
|---|---|---|---|---|
| R-01 | Σ `03` Large Losses, year *y* | `01` Incurred Losses, year *y* | ≤ | error |
| R-02 | Σ `04` Cat Losses, year *y* | `01` Incurred Losses, year *y* | ≤ | error |
| R-03 | `10` Triangles latest diagonal, year *y* | `01` Incurred Losses, year *y* | = | error |
| R-04 | `02` EPI first projected year | `01` Premium trend | plausibility | warning |

---

## 11 · Logging

Two logs, linked to the sheets by block ID.

**Debug log** — technical and verbose: resolved address maps, cell ranges, coercion
failures, timings.

**Process log** — decisions and interpretations in plain language, for a reviewer:

> `01 History` / STEP 1 — header row 8, selector column L. Column `Premium` resolved to
> D, `Incurred Losses` to G. 6 candidate rows, 5 extracted, 1 excluded (row 14, selector
> empty). Currency USD, scale 1,000. Year basis assumed UW (hypothesis H-01, open).

Every run is stamped with source-file hash, spec version, timestamp and tool version.
With a yearly cadence, that stamp is what answers "why did the 2026 numbers look like
that" in 2028.

---

## 12 · Technical constraints

| Rule | |
|---|---|
| T1 | The workbook is opened **twice** — `data_only=True` for values, `data_only=False` for formulas. One pass cannot give both |
| T2 | Error cells (`#REF!`, `#N/A`, `#VALUE!`) inside an extraction range are **fatal**, never coerced to null or zero |
| T3 | A pre-flight pass scans for external links, error cells and merged ranges before extraction |
| T4 | Pivot tables and charts do not survive an `openpyxl` round-trip — **open question**, see §13 |

T2 exists because a silently zeroed error cell reaches an underwriter under a
clean-looking control sum.

T3 earns its keep on a hand-assembled workbook: copied sheets carry formulas pointing at
absent source files, which resolve to stale caches or `#REF!`.

---

## 13 · Open questions

1. **Pivots in the output.** Do reviewers need pivot tables in the *written* file? If the
   source is preserved untouched, losing them in the derived audit copy may cost nothing —
   which would make plain `openpyxl` sufficient. If they are needed, the options are
   `xlwings`/COM (perfect fidelity, requires Excel) or surgical OOXML editing (headless,
   considerably more work).
2. **Pivot source ranges.** If a pivot's source is a whole column or an auto-expanding
   table, appended blocks will be drawn into it on refresh.
3. **Mandatory attributes.** Does the attribute list in `00` row *n* constitute the
   checklist of what *must* be declared, or are some attributes genuinely optional?
4. **Sheet 08, Engineering variant.** Does it carry a second split axis analogous to
   Fire's B/C/BI?
5. **Datasets `02`–`10`.** Header labels and attributes not yet specified.

---

## 14 · Status of the reference workbook

`Intake_v1.xlsx` implements this specification for:

- `00. NC+Interdep` — dataset register, vocabulary, marker legend
- `01. History` — row-wise, 5 records
- `01. History_Transposed` — transposed, 5 records

Both `01` sheets carry identical data and must produce identical step-1 output. The
dataset rows for `02`–`05` in `00` are provisional sketches, marked as such, pending
specification.
