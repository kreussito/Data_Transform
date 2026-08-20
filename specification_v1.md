# Data_Transform — Specification v1

Extraction, filtering and transformation of reinsurance treaty submission data from
Excel, driven by metadata declared in the workbook itself.

| | |
|---|---|
| **Version** | 1 |
| **Status** | `00`–`07` implemented, per-risk and cat sections; `08`–`10` outstanding |
| **Cadence** | Once per treaty, per year |
| **Reference workbooks** | `Intake_v1.xlsx` · `Intake_Engineering_v1.xlsx` · `Intake_EngineeringCombined_v1.xlsx` · `Intake_FireCat_v1.xlsx` · `Intake_FireEQWind_v1.xlsx` · `Intake_FireCatLosses_v1.xlsx` · `Intake_Mexico_v1.xlsx` |

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

| Sheet | Holds | State |
|---|---|---|
| `00. Nomenclature & Interdependencies` | Vocabulary, dataset register, inventory | — |
| `01. History` | Premium and loss history per section | **implemented** |
| `02. EPI Projections` | Estimated premium income, N and N+1 | **implemented** |
| `03. Large Losses` | Individual large claims, per-risk sections | **implemented** |
| `04. Cat Losses` | Catastrophe events, cat sections | **implemented** |
| `05. Risk Profiles` | Banded exposure — the per-risk rating basis, §2.4 | **implemented** |
| `06. EQ Aggs` | Earthquake sums insured per cat zone, §2.5 | **implemented** |
| `07. Wind Aggs` | Windstorm sums insured per cat zone, §2.5 | **implemented** |
| `08. Splits` | The cedent's book split — sums insured, one block per section; the ratio source for `06`/`07`, §2.6 | **implemented** |
| `09. Rate Development` | Rate change history — optional | outstanding |
| `10. Triangles` | Loss development triangles | outstanding |
| `11. Exchange rates` | FX rates for conversion between currencies | outstanding |
| `20. Summary` | Collected step-2 blocks | outstanding, **generated** |

Sheet `00` carries this table as its `⟦INVENTORY⟧` block (§3.3), so a reader of any pack
can tell a sheet that is *absent from this treaty* from one that was *never specified*.
The register above `⟦INVENTORY⟧` lists what the pack actually carries; `⟦SECTIONS⟧` says
which roles each section expects. Three different questions, three blocks.

`Exchange rates` is numbered `11` so that every dataset sorts by its role (§2.2); it was
the one sheet without a number, which made it the one sheet the ordering could not place.

The reference workbook carries `01` and `02` in both orientations, as
`…_Transposed` sheets, so every rule is exercised both ways.

There is no separate register sheet. Assumptions, hypotheses and control sums are
documented **in the sheet they belong to**, beneath the original data, so a reviewer
never has to cross-reference another tab to see how a figure was arrived at.

The workbook is assembled by copying sheets in from original submission files. Expect
external links, stale formula caches, heterogeneous layouts and imported defined names.

### 2.1 Sheet 08 — two exclusive variants

Sheet `08` carries exactly one of two variants, each split on **two axes**:

| Variant | Axis 1 — occupancy | Axis 2 — cover |
|---|---|---|
| **Fire** | Residential · Commercial · Industrial | B · C · BI |
| **Engineering** | Residential · Commercial · Industrial | Projects · Renewables |

**The occupancy axis is common to both.** That is the useful part: Res / Com / Industrial
is how the *portfolio* is divided, and it divides an Engineering book exactly as it
divides a Fire one. What distinguishes the variants is only the second axis — what is
being covered. So one split mechanism serves both, and the variant selects the second
axis rather than the whole layout.

The variant is declared, not sniffed. A declaration that disagrees with the sheet
content is an interdependency failure.

Both variants are the same machinery as `06` and `07`: the buckets are the product of the
declared axes (§2.6), which gives Fire nine and Engineering six. Nothing in `08` needs a
split mechanism of its own.

Still to be specified for `08`: the **measures** at each intersection — a risk count,
a sum insured, a premium, or some combination — and therefore the header labels sheet
00 will declare. The axes are settled; what is counted along them is not.

### 2.2 Sheet matching

Sheets are matched by **numeric prefix first**, text second. Both sides are normalised
for whitespace, non-breaking spaces, quote style and case. `10. Triangles` and
`10. Triangels` therefore resolve to the same sheet.

### 2.3 Sections — what a treaty is made of

A treaty is one or more **sections**, declared in `⟦SECTIONS⟧` of sheet 00:

| Section | Kind | Datasets |
|---|---|---|
| `Fire` | per risk | 01, 02, 03, 05 |
| `Earthquake` | cat | 01, 02, 04, 05 |
| `Windstorm` | cat | 01, 02, 04, 05 |

A Fire-only or Engineering treaty lists one per-risk section. A Nat Cat treaty lists
cat sections and **no `03` anywhere** — the absence of large losses is *declared*, not
inferred from a missing sheet.

**A section is a block.** The `Header_i` mechanism already built for multi-block sheets
is the section mechanism, so where the blocks sit does not matter:

```
three sheets                        two sheets, one combined
01. History Fire   Section = Fire   01. History Fire  Section   = Fire
01. History EQ     Section = EQ     01. History Cat   Section_1 = Earthquake
01. History Wind   Section = Wind                     Section_2 = Windstorm
```

Both shapes yield the same section-blocks and the same results. `Intake_FireCat_v1.xlsx`
and `Intake_FireEQWind_v1.xlsx` carry identical figures in the two layouts, and are
verified to produce identical crosschecks.

**Roles.** The leading number of a dataset key is its **role**: `01 History Fire` and
`01 History EQ` both play role `01`. Rules reference the role, so one declaration
serves every section. Where a role and section resolve to several blocks, the row-wise
one wins — a transposed twin is the same data in a different shape — and anything still
ambiguous resolves to nothing rather than a guess.

**Block boundaries.** A stacked block's records stop where the next block's `Header_j`
begins. Without that, the first block of a stacked sheet would scan to the end and
swallow the records below it. Blocks that *overlay* one another are bounded differently —
see §6.5.

### 2.4 Sheet 05 — the risk profile

A banded exposure profile: the portfolio cut by risk size, which is what a per-risk
excess-of-loss treaty is rated on. Declared in `00` in **the order a profile is read** —
what the band is, where it sits, what it earns, then what it is made of. Step 2 writes
the columns in exactly that order (§9.2 S2), so the declaration *is* the layout:

| Field | Type | |
|---|---|---|
| `Band` | text | the cedent's own label, verbatim — **optional** |
| `Band from` | number | lower bound — **optional** |
| `Band to` | number | upper bound — **optional** |
| `Premium` | number | |
| `Number of Risks` | number | |
| `Exposure` | number | on the declared basis |

Attributes: `Section · Currency · Scale · Share basis · Exposure basis · Includes fac ·
Layered business · As at`.

**Why the bounds are numeric and separate from the label.** Natural alphanumeric sorting
reads `"1,000 – 5,000"` as the digits `1` then `000`, so it would sort *before*
`"500 – 1,000"`. And exposure rating a layer of 2,000 xs 1,000 has to allocate it across
bands, which cannot be done against bounds that exist only as prose. The label is kept
for identity — never parsed, exactly as `Year` is (§8.4) — and the numbers are what the
machine works with.

**Why all three band fields are optional.** A band arrives in three shapes, and
**step 2 always ends up with both bounds**:

| The source gives | Step 1 shows | Step 2 does |
|---|---|---|
| two numeric columns | both bounds | nothing to work out |
| one label — `1-1,000,000` | the label | reads the bounds off it, and reports what it read |
| the two columns, no label | both bounds | names each band by its own bounds |

What must *not* happen is neither: a block extracting no bounds and no label cannot
produce a profile, and says so rather than sorting on nothing.

The label is optional because **a band's identity is its bounds**. `1-1,000,000` and
`1 – 1.000.000` are the same band written two ways; the two numbers are not. Where a
label exists it is kept verbatim and never parsed for identity — the same treatment
`Year` gets (§8.4) — but nothing downstream depends on it.

Where the columns exist they are extracted. Where they do not, **step 2 reads the bounds
off the label** and writes what it read. That division is not a convenience: step 1
carries only what the sheet says (S11), and reading `1` and `10,000` out of `"1-10,000"`
is interpretation.

The parse follows the same discipline as decimals (§8.4): **unambiguous forms are read,
ambiguous ones are fatal.** A label nobody can read without guessing stops the run and
names itself, and the escape hatch is always to add the two numeric columns.

An open bound is an *absent* bound, not a zero: `> 1,000,000` gives `Band from` and
leaves `Band to` empty; `< 10,000` does the reverse. Empty stays empty (§8.4).

**Continuity** is checked but not over-specified, because both conventions are in use:

```
0 – 10,000 · 10,000 – 20,000      shared boundary
1 – 10,000 · 10,001 – 20,000      gapless integers
```

So a gap is reported only where the next band starts **more than one unit** above the
previous band's end, and an overlap only where it starts *below* it. Either convention
passes; a band genuinely missing from the middle of a profile does not.

**Why the field is `Exposure`, not `Sum Insured`.** `Exposure basis` may declare `EML`,
`PML` or `MPL`. A field name carries one meaning across the workbook (§3.3), so a column
headed `Sum Insured` holding an EML figure would be a false label in the output. The
field says what it is; the attribute says on what basis. In practice a profile is on sums
insured rather than PMLs, and that is what the attribute will usually say.

**Three attributes decide whether the figures are comparable with anything else:**

| Attribute | | |
|---|---|---|
| `Share basis` | `100%` · `ceded only` | usually gross; ceded is the exception |
| `Includes fac` | `yes` · `no` | facultative business is normally in the profile and often reinsured separately, so a profile including it overstates what the treaty sees |
| `Layered business` | `included` · `excluded` | a risk written in layers has no single band |

None of these can be inferred from the numbers, and each of them changes what a
comparison against `01` means — so each is declared, and an undeclared one ranks `Open`.

**Several profiles in one sheet need no new machinery.** A profile is a snapshot at a
date, so a pack showing two or three years carries one **block per profile**, each with
its own `As at_i`. Whether the blocks sit one below the other or side by side is already
handled: stacked blocks share a selector column, side-by-side blocks use different ones,
and §6.5 tells them apart without anyone declaring which arrangement it is.

**One profile per line of business.** Every section expects `05`, whatever its kind, so a
Fire + EQ + Windstorm treaty carries three profiles and an Engineering treaty one. The
same block mechanism serves that too: `Intake_FireCat_v1.xlsx` puts all three on one
sheet as `Section_1..3`, `Intake_FireEQWind_v1.xlsx` gives each its own sheet, and the two
are verified to produce identical profiles.

### 2.5 Sheets 06 and 07 — the cat aggregates

Where `05` cuts the portfolio by risk size, `06` and `07` cut it by **geography**: the
sum insured sitting in each cat zone, which is what a catastrophe model is fed and what a
cat treaty is rated on. `06` carries the earthquake zoning, `07` the windstorm zoning.
The two are the same dataset with different catalogues, and the code treats them so.

| Field | Type | |
|---|---|---|
| `Zone` | text | the zone code, verbatim — `13a`, not `13.1` |
| `Res Building` `Res Content` `Res BI` | number | residential |
| `Com Building` `Com Content` `Com BI` | number | commercial |
| `Ind Building` `Ind Content` `Ind BI` | number | industrial |
| `Total` | number | **optional** — see below |

Eleven headers, which is why the register in `00` no longer stops at ten (§3.1).

**The zone is text.** `13a` and `14c` are not numbers, and `13` is not a Mexican
earthquake zone at all — it exists only as `13a` and `13b`. Sorting is natural
alphanumeric (§9.2 S12), which puts `2` before `10` and `13b` before `14a` without anyone
declaring a rank.

#### `⟦ZONES⟧` — the catalogue is declared, not inferred

Sheet `00` carries a `⟦ZONES⟧` block: one row per scheme, holding the complete zoning.

```
Mexico EQ     1, 2, 3, …, 12, 13a, 13b, 14a, 14b, 14c, 14d, 15, …, 48     (52 zones)
Mexico Wind   1, 2, 3, …, 42                                              (42 zones)
```

A block says which catalogue it is on through the `Zone scheme` attribute. A scheme the
block names but `⟦ZONES⟧` does not list is **fatal** — the tool will not invent a zoning.

**Every declared zone appears, and a zone with nothing in it reads 0.** How that comes
about depends on the cedent, and the two cases are handled differently on purpose:

| The cedent sends | Step 1 | Step 2 |
|---|---|---|
| the complete table, zeros and all | all 52 rows — they were extracted | nothing to complete |
| only the zones it writes in | those rows, and no others | fills the rest from `⟦ZONES⟧` and says so |

The first is the normal case in Mexico, where the zoning is the regulator's and the form
has a row per zone. The second is common elsewhere, and there step 2 reports what it
added: *"46 declared zone(s) with no entry in the source, shown as 0"*.

The completion is value-preserving — adding zeros cannot move the control sum — and it is
the point of the exercise: an absent row reads as *no data*, a zero reads as *nothing
there*, and only the second is a statement about the portfolio. **Step 1 is never
completed**, in either case: it shows the rows the sheet has and no more, because step 1
carries only what was extracted (S11). Where the two differ, the difference is itself
information about how the cedent reports.

#### `Total` — derived or checked, never corrected

Where the cedent supplies a total column it is **checked** against the nine buckets and
the verdict recorded; where the cedent supplies none it is **derived** as their sum and
labelled value-adding. A total that disagrees with its parts is reported and left alone
(S19) — the tool does not know which of the ten figures is the wrong one, and correcting
a source figure would break the audit baseline (§1).

#### What the numbers actually mean

An aggregate is a single number carrying a great many decisions, and two cedents' figures
are not comparable until those decisions are known. None can be read off the numbers, so
each is a declared attribute, and an undeclared one ranks `Open` (§5.3):

| Attribute | | |
|---|---|---|
| `Coinsurance` | `deducted` · `not deducted` | whether the co-insurer's share is already out |
| `Deductible` | `deducted` · `not deducted` | policy deductibles netted off, or gross |
| `Standard deductible` | e.g. `2%` EQ, `1.5%` hurricane | the market convention this book is written on |
| `Limit basis` | `full sum insured` · `per risk limit` · `per location limit` | whether a limit has already capped the figure |
| `Multi-location` | `split by location` · `allocated to main zone` · `duplicated in each zone` | a policy covering sites in several zones |
| `Includes fac` | `yes` · `no` | as in `05` |
| `Share basis` | `100%` · `ceded only` | as everywhere |
| `Exposure basis` | `Sum Insured` · `EML` · `PML` · `MPL` | as in `05` |

`Multi-location` is the one that can double-count: *duplicated in each zone* means the
column totals exceed the portfolio, which is defensible for zone-level modelling and
indefensible as a portfolio figure. It has to be visible.

#### `⟦SPLITS⟧` — declared, currently empty

Cedents often send fewer than nine buckets — a single figure per zone, or building and
content without BI, or no occupancy split at all. The ratios that expand what arrived
into the nine belong in a `⟦SPLITS⟧` block in `00`, keyed by scheme and occupancy. The
block exists and is **deliberately empty**: a split is an assumption, and an assumption
belongs where a reviewer can see it and where it is versioned with the pack, not inside
the code. Until it is filled, a block sending fewer buckets carries what it sent.

#### Three versions, one sheet

A cedent sends the same aggregate three times, and the set is fixed:

| `Period` | `As at` | |
|---|---|---|
| `{N} 9 months` | 30.09.N | the nine-month estimate of the expiring year |
| `{N+1} at inception` | 01.01.N+1 | projection for the first day of the renewal year |
| `{N+1} at expiry` | 31.12.N+1 | projection for its last day |

They sit on **one sheet as three stacked blocks** sharing a selector column, which §6.5
already recognises without anyone declaring the arrangement. `at inception` and `at
expiry` are added to `⟦PERIOD ORDER⟧` at ranks 4 and 5, so the three sort chronologically
rather than alphabetically.

No single version is interesting. The movement between them is, and only against the
premium movement — which is §10.4.

**No rule compares `06` with `07`.** Either peril can be bought alone or both together,
and the covered books need not be the same, so a comparison between them would fail on
perfectly ordinary submissions.

### 2.6 Splits — from what the cedent reported onto the target cells

A cat model wants sums insured on a **fixed** grid: occupancy × cover for earthquake —
`Res Building` … `Ind BI` — and occupancy alone for windstorm. That target does not
change with the line of business. What changes is what arrives.

| Dataset | Target cells | |
|---|---|---|
| `06` Earthquake | 9 | Res · Com · Ind × Building · Content · BI |
| `07` Windstorm | 3 | Res · Com · Ind |

**The measure is sums insured**, however they are interpreted — the interpretation
attributes (`Coinsurance`, `Deductible`, `Limit basis`, `Multi-location`) stay on `06`/`07`
where the figures are, and say what those sums mean.

#### The bridge

The operation is not "multiply the missing axis on". It is a map from each reported
category onto a distribution over the target cells:

```
B(z, t) = Σₐ  S(z, a) · M(a → t)          with   Σₜ M(a → t) = 1
```

Every case is that one formula:

| The zone reports | `M` is | |
|---|---|---|
| one figure | the joint distribution `p(i,j)` | case a) |
| occupancy | `p(j\|i)`, zero outside the row | case b) |
| cover | `p(i\|j)`, zero outside the column | case c) |
| **Projects / Renewables** | `q(i\|a) · p(j\|a)` — no zeros | the third axis |
| occupancy **and** cover | fitted to both margins — see below | |

The zeros are why a reported figure survives untouched: the row sums to exactly what
arrived. The segment axis has no zeros, and that is precisely what distinguishes it — it
is a *translation between two descriptions of the same book*, not a refinement of one
axis into another.

Two steps, kept apart on purpose: **expand** what arrived onto the section's full
occupancy × cover grid, then **project** that grid onto the axes the dataset declares.
Earthquake declares both, so nothing is projected away. Windstorm declares occupancy
alone, so the cover axis is summed out — but the reported cover figures still shaped the
answer rather than being discarded, which is why a BI-heavy zone comes out less
residential than the book average.

#### Sheet 08 supplies `M`, one per section

`08` is the cedent's own split table: the book's sums insured on whatever grid it keeps,
**one block per section**, because an earthquake book and a hurricane book need not be
the same. It is a dataset in its own right — it gets both steps, it is sorted, its `Total`
is reconciled against its parts (S19) — *and* it is where `06`/`07` read their ratios.

Two properties fall out of taking ratios from amounts rather than from typed percentages:

- **They cannot fail to close.** A share derived from sums insured sums to 1 by
  construction; a typed `38% / 42% / 20%` can be mistyped and has to be checked.
- **`08`'s scale and currency are irrelevant.** Only ratios are read, so `08` need not
  agree with `06`/`07` on either — only with itself.

And the map knows something a single occupancy vector could not: **the occupancy mix
differs by cover**. Residential BI is nearly nil, so `p(Res|BI)` is small where
`p(Res|Building)` is large. Applying one occupancy vector to all three covers would
invent residential business interruption.

#### Two reported margins: neither may move

Where a zone reports occupancy *and* cover as two vectors, conditioning on one preserves
that one and lets the other drift. Both were reported, so neither may move. The grid from
`08` is used as a seed and scaled alternately to the row and column totals until it
satisfies both (iterative proportional fitting). **The seed decides only how the two
margins interact**; the margins themselves come out exactly as sent.

#### Changing the shape is a sheet edit

Nothing about the axes is in the code. The categories, their order and their names come
from `⟦AXES⟧`; the second axis is found by **position**, not by the word "Cover", so
renaming it — or giving an engineering book Projects/Renewables as its first axis and
something else as its second — changes no code. The segment conventions are rows in
`⟦SPLITS⟧`. An engineering aggregate is therefore a different declaration, not a
different program, which is what keeps it cheap to change while the shape is still being
argued about.

#### The one edge no cedent supplies

`q(i|segment)` — Projects and Renewables against residential/commercial/industrial.
Nobody cross-tabulates that way, so it is a declared convention in `⟦SPLITS⟧`:

| From | To | Share | Source |
|---|---|---|---|
| `Renewables` | `Ind` | 100% | House convention |
| `Projects` | `Com` | 50% | House convention |
| `Projects` | `Ind` | 50% | House convention |
| `Commercial` | `Com` · `Ind` | 75% · 25% | House convention |

Wind farms and solar fields are industrial risks; a project is half commercial and half
industrial. Residential comes out at nil, which is right — there is no residential
engineering. Every note that applies one of these says so and names the source, and the
block's confidence drops to `Assumed` (§5.3).

`⟦SPLITS⟧` therefore holds only what is *our* judgement. Everything the cedent supplies
comes from `08` and carries provenance. Where a section has no `08` at all, `⟦SPLITS⟧` may
also carry a fallback ratio — last year's, or another client's — and it is written into
the block as an assumption with its origin named.

#### Amounts or percentages — only ratios are read

A split table may arrive as sums insured or as percentages, and neither needs handling:
the grid is normalised, so both give the same distribution.

One shape does need recognising. **Percentages that close per row are not one
distribution but three** — a cover mix for each occupancy, with nothing saying how large
each occupancy is. Normalising such a grid whole would silently assert that the three
occupancies are equally big. So that shape is detected (every row summing to 100% or to
1, with differing totals) and the weight is taken from the `Total` column, which is the
only place it can be.

#### Two views of the same book — a finding, not a refusal

A cedent may send the occupancy split **and** a Projects/Renewables split. Each implies an
occupancy mix, and nothing in the numbers says which one it stands behind. Neither
choosing one quietly nor refusing the submission is right: the first loses a column
without a trace, the second throws away a good pack over a question a person answers in a
sentence.

So the block is built from the **richer** level — the order is: two margins, one margin,
segmentation, bare total — the other is bridged as well, and the two are shown side by
side under step 2:

```
TWO VIEWS OF THE SAME BOOK — THEY DO NOT AGREE
The block reports the reported grid and the Projects/Renewables segmentation. Both
describe the same portfolio, so each implies an occupancy mix. Neither is an error
and nothing here fails the run.

              from the reported grid   from the segmentation   difference
  Res                         89,569                       0      -100.0%
  Com                         78,370                 111,957       +42.9%
  Ind                         55,978                 111,957      +100.0%
  Largest difference                                              100.0%

QUESTION FOR THE UNDERWRITER: which view does the cedent stand behind — the reported
grid or the Projects/Renewables segmentation? …
Answer:  ▁▁▁▁▁▁▁▁▁▁
```

It is written **under step 2** rather than as a note, because it is not a description of
what was done — it is something a person has to answer, and an answer needs somewhere to
be written. The empty cell is the only place in the whole output where the tool asks for
input rather than reporting a result.

Where the two agree within tolerance the block still appears, saying so. A confirmation
is worth as much as a discrepancy: it means two independent descriptions of the book
line up, which is the strongest thing that can be said about a split.

#### What it refuses

**A column belonging to no declared level.** If it can be bridged it becomes the finding
above; if nothing knows what it is, it can be neither used nor compared, and dropping a
column that carries money is not something this step does in silence.

**Splitting happens inside a zone, never across zones.** An occupancy mix is a property
of the portfolio and transfers plausibly; the geographic distribution *is* the analysis,
and assuming it would be inventing the answer. A block without `Zone` is refused.

Also refused: a block reporting no level at all; an `08` carrying nothing under a cover
the block reports, rather than spreading it evenly and calling that an answer; and a
segment with no declared mapping.

#### What is assumption and what is not

Applying a book-level ratio to each zone assumes every zone has the book's mix. That is
false — Mexico City is not a coastal resort zone — and false exactly where a model is
most sensitive. But because every zone is multiplied by the same shares, **the cell
totals over all zones reproduce the cedent's own split exactly**. The distribution
*between* zones is the assumption; the split *of the book* is not, and a reader can be
told precisely which half to trust.

A split redistributes and never creates: the zone total is unchanged, so control sums tie
and §10.4 reads the same figures whether or not a split ran. Value-adding in detail,
value-preserving in sum.

---

## 3 · Sheet 00 — the nomenclature

Sheet `00` is a **dictionary for humans**, not an enforcement layer. It tells whoever
prepares a data sheet which names to write. It does not declare sheets required.

### 3.1 Layout

- Column **A** is empty. No markers live in `00`.
- **One row per dataset.**
- Column **B** — the exact sheet name.
- Column **C** — a short key, used in block IDs and log entries.
- Columns **D** onward — the **header labels**: the minimum information to extract.
- After them, the **attribute names** that sheet should declare.
- Row **4** declares where each run begins: `Headers` over the first header column,
  `Attributes` over the first attribute column.

**The boundary is read from row 4, not fixed in the code.** Datasets have ragged header
counts, so it cannot be found per row; but nailing it to column N would cap every pack in
the world at ten headers, and `06` needs eleven (§2.5). So row 4 is the declaration and
the code follows it. A sheet that carries neither label falls back to the original
`D … M` / `N` layout, which is what every pack written before the boundary moved relies
on — those workbooks keep reading unchanged.

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

### 3.3 The other blocks of sheet 00

Sheet 00 is the **frame**: what each sheet holds, what the names mean, and what must
tie. Beneath the dataset register it carries further blocks, each found by its
anchor rather than by position.

#### `⟦GLOBAL⟧` — workbook-wide facts

| Attribute | Value |
|---|---|
| `Actual year` | `2025` |
| `Cedent` | Example Insurance SA |
| `Treaty` | Property per Risk XL |
| `Loss share warning` | `20%` |
| `Rate change warning` | `20%` |

`Actual year` is **N**, the expiring year; the renewal being underwritten is **N+1**.
Every `{N}` reference in `⟦RULES⟧` and in the step-2 figures resolves from it.

`Loss share warning` is the point above which a year's declared losses are flagged as
event-driven rather than attritional — see §10.3. It sits here because it is a matter of
underwriting judgment, not a mechanic: a different underwriter may want 15% or 30%, and
should not need the tool changed to get it.

`Rate change warning` is the same idea for §10.4: the movement in premium per unit of
exposure beyond which the growth block raises a warning. Both default to 20% where the
workbook declares nothing.

Attributes now resolve in **three tiers**, each overriding the one above:

| Tier | Where | Scope |
|---|---|---|
| Global | `⟦GLOBAL⟧` | the whole workbook |
| Sheet | column A, unsuffixed | one sheet |
| Block | column A, `_i` suffix | one block |

#### `⟦TYPES⟧` — field name → datatype

| Field | Type |
|---|---|
| `Year` | text |
| `Premium` · `Incurred Losses` · `EPI` · `Loss amount` · `Exposure` · `Band from` · `Band to` · `Number of Risks` · `Number of Claims` | number |
| `Name of Loss` · `Band` · `Event ID` · `Event Name` · `Claim Reference` | text |
| `Date of Loss` · `Event Date` · `Event End Date` | **date** |

**A field name carries one type across the whole workbook** — which is precisely what a
nomenclature is for. A declared field with no entry here is an error; nothing is
defaulted, because defaulting `Claim Reference` to a number would be silent and wrong.

#### `⟦PERIOD ORDER⟧` — how suffixes sort within one year

| Rank | Suffix |
|---|---|
| 1 | `est` |
| 2 | `9 months` |
| 3 | `re-est` |
| 4 | `at inception` |
| 5 | `at expiry` |

Plain alphanumeric would put `2025 9 months` before `2025 est`, because `9` sorts
before `e`. Chronologically that is backwards. This block states the intended order;
a bare year sorts first, and a suffix nobody declared sorts last — so the ordering
stays total whatever a pack contains. See §9.2 S13.

Ranks 4 and 5 order the three versions of a cat aggregate (§2.5). A pack that carries no
aggregates simply never uses them; a block is a declaration of what *may* appear, not of
what must.

#### `⟦ZONES⟧` — the cat zone catalogues

| Scheme | Zones |
|---|---|
| `Mexico EQ` | `1, 2, …, 12, 13a, 13b, 14a, 14b, 14c, 14d, 15, …, 48` |
| `Mexico Wind` | `1, 2, …, 42` |

One row per zoning scheme, holding the **complete** list. A `06`/`07` block names its
scheme through the `Zone scheme` attribute; step 2 fills every declared zone the cedent
did not list with 0, and a scheme not listed here is fatal. See §2.5 and S18.

#### `⟦INVENTORY⟧` — every dataset the standard defines

| Role | Sheet | Holds | State |
|---|---|---|---|
| `01` … `11`, `20` | | | see §2 |

The same table as §2, written into every pack. It answers a question the register cannot:
the register lists the sheets *this* workbook carries, so a dataset missing from it might
be absent from the treaty or might never have been specified at all. `⟦INVENTORY⟧` tells
the two apart, and carries the `State` column so a reader knows what the tool will
actually do with a sheet if one turns up.

#### `⟦AXES⟧` — the dimensions a dataset is split along

| Dataset | Axis | Categories |
|---|---|---|
| `06 EQ Aggs` | Occupancy | `Res, Com, Ind` |
| `06 EQ Aggs` | Cover | `Building, Content, BI` |
| `07 Wind Aggs` | Occupancy | `Res, Com, Ind` |

The dataset's buckets are the **product** of its axes, named by joining the categories —
`Res Building` … `Ind BI` for two axes, `Res` · `Com` · `Ind` for one. Those are the same
names the register declares, so the arithmetic and sheet 00 cannot drift apart. See §2.6.

#### `⟦SPLITS⟧` — the ratios, and where each came from

| Dataset | Axis | From | Category | Share | Source |
|---|---|---|---|---|---|
| `07 Wind Aggs` | Occupancy | | `Res` | 38% | Cedent book split, 30.09.2025 |
| `07 Wind Aggs` | Occupancy | | `Com` | 42% | Cedent book split, 30.09.2025 |
| `07 Wind Aggs` | Occupancy | | `Ind` | 20% | Cedent book split, 30.09.2025 |
| `*` | Occupancy | `Commercial` | `Com` | 75% | House convention |
| `*` | Occupancy | `Commercial` | `Ind` | 25% | House convention |

A blank `From` distributes a whole axis the block never reported. A filled one re-splits a
category that arrived merged — `Commercial` covering both commercial and industrial risks.
`*` states a house convention once; a row naming the dataset overrides it for that book.

`Source` is not decoration. A ratio from the cedent's own prior submission and one
borrowed from another book are both assumptions, but not equally good ones, and the reader
has to be able to tell them apart. It is written into every note the split produces.

#### `⟦RULES⟧` — crosschecks and interdependencies

See §10.1.

### 3.4 Vocabulary block

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
| `Dataset_i = <key>` | Which dataset block *i* **is** — defaults to the sheet-name match (§6.4) |
| `Section_i = <name>` | Which section of the treaty block *i* belongs to (§2.3) |
| `<Attribute> = <value>` | An attribute of the sheet or block |
| `H_<Attribute> = <value>` | The same, declared as a hypothesis |

`Header`, `Info`, `Transpose` and `Dataset` are **structural**: they say how to read the
block, so they never appear in the attribute list a reviewer reads. `Section` is not
structural — which part of the treaty a figure describes is a fact *about the data*, and
belongs beside `Currency` and `Year basis` where a reviewer will look for it.

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

Confidence is recorded at **dataset and block level only**, with one exception: a
**transposed** block also stamps each record, because transposition reinterprets every
record structurally rather than merely reordering it.

### 5.4 Where hypotheses are documented

Hypotheses are written into the **step 1 block of the sheet they belong to**, listing
id, attribute, value, confidence, status and the cell the `H_` marker was found at.
They come from two sources:

1. `H_` markers in the sheet — the underwriter documented the assumption where they made it
2. Tool findings — an attribute listed in sheet 00 but never declared, a column
   containing data that nobody declared, a failed interdependency

Every item at `Assumed` or `Open` is, in effect, **a question for the broker**, produced
as a by-product of extraction rather than remembered by hand between renewals.

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

The extraction row is never itself a record; its selector cell stays empty. **No block's
extraction row is ever a record of any block** — which matters once two blocks overlay
one another (§6.5), since one block's extraction row then sits among another's records.

### 6.4 A block may name its own dataset

The sheet name resolves to a dataset via `00` (§2.2). That is a **default, not a law**: a
block may override it.

```
A12:  Header_1
A13:  Dataset_1 = 03 Large
A14:  Header_2
A15:  Dataset_2 = 04 Cat
```

A key that `00` does not declare is fatal. A sheet whose name matches no dataset is
skipped as before — unless one of its blocks names a dataset, in which case the sheet is
read and every block must name one.

### 6.5 Stacked and overlaid blocks

Two arrangements, told apart by the **selector column** without anyone declaring which
is which:

| | Blocks sit | Selector columns | Boundary |
|---|---|---|---|
| **Stacked** | one below the other, different rows | **shared** | block *i* stops where block *j*'s `Header_j` begins |
| **Overlaid** | over the *same* rows | **different** | none — each selector already identifies its own rows |

Stacked is how a combined History sheet carries Earthquake and Windstorm: without the
boundary the first block would scan to the end and swallow the second's records.

Overlaid is how **one loss list carries both `03` and `04`** — the Engineering and
Miscellaneous case, where a cedent reports large losses and cat events together:

```
        B          C          D       E              F        G      H     I      J        K
11      Claim no.  Cat code   U/W Yr  Description    Amount   From   To    Claims        ← source header
12                            Year    Name of Loss   Loss amount  Date of Loss           ← Header_1
14                 Event ID   Year    Event Name     Loss amount  Event Date  …          ← Header_2
16      CL-2201               2021    Turbine …      1 420    …                 =ROW()
17                 EV-301     2021    Flood, Saxony  1 150    …                          =ROW()
```

with

```
J:  =IF($C16="",  ROW(), "")      a row with no cat code is a large loss
K:  =IF($C16<>"", ROW(), "")      a row with one is a cat event
```

**The selector column does the classifying**, and nothing is inferred: the cedent's own
cat-code column drives it, and a human wrote the formula that says so. A selector holding
`""` is a formula that decided the row is not this block's — the same as an empty cell.

A boundary here would be fatal: `Header_2` sits *above* the shared records, so block 1
would stop before reading any of them. That is why the boundary follows the selector
column rather than the header position.

Because each block sees the other's rows and columns, both are **reconciled** afterwards,
or the step-1 block would state two false things about itself:

- a column the sibling extracts is not "contained data, not declared in `00`" — it *is*
  declared, in the other block's dataset;
- a row the sibling extracts is not an *excluded* record — nothing was dropped. It is
  reported instead as `3 extracted by 04 Cat on this sheet — one list, two datasets`.

`Intake_Engineering_v1.xlsx` and `Intake_EngineeringCombined_v1.xlsx` are the same treaty
in the two layouts, verified to produce identical records, identical crosschecks and an
identical §10.3 check.

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
- **Records follow the extraction row / column.** Row-wise, candidates begin at the row
  below `Header_i`; transposed, at the column right of it. This is not position
  dependence — the extraction row's own location is discovered from the marker — but it
  keeps captions above the block, and column A's marker channel, out of the record set.

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
| F4 | A label declared in `00` but absent from the extraction row is an error, **unless `00` marks it `(optional)`** |
| F5 | Every column/row containing data but not extracted is **logged**, and raises a question |
| F6 | Merged cells are resolved before matching |
| F7 | **Only resolved addresses are read** |

F7 is the substance of the filter: the block is never read as a rectangle and then
trimmed. Resolution produces an address map, and only those addresses are ever touched.

F5 preserves the "a new column appeared and nobody noticed" catch that would otherwise
be lost by ignoring the source header.

**F4 and optional fields.** Sheet 00 declares an optional field by annotating the
header, and the annotation is stripped from the name:

```
Number of Claims (optional)   →   field "Number of Claims", optional
```

The human still writes the bare name in the extraction row. Optionality belongs in 00
because it is a statement about the *dataset* — what a cat listing may or may not carry
— not about one workbook. Where the field is absent, it is simply not a column: it is
excluded from the output, from the totals and from the column order, and step 1 says so
in one line:

```
Declared optional in sheet 00, absent from this block: Number of Claims
```

That line is the difference between *the cedent did not report it* and *the tool lost
it*, and only the first of those is acceptable in an audit trail. Everything not marked
optional stays fatal when missing.

### 8.1 Worked resolution — `01. History`, row-wise

| Declared in `00` | Found at | Resolves to |
|---|---|---|
| `Year` | B8 | column **B** |
| `Premium` | D8 | column **D** |
| `Incurred Losses` | G8 | column **G** |

Records are rows 9–14. Cells read: `B`, `D`, `G` × rows 9–14 = **18 cells, and no
others.**

### 8.2 Worked resolution — `01. History`, transposed

| Declared in `00` | Found at | Resolves to |
|---|---|---|
| `Year` | C8 | row **8** |
| `Premium` | C10 | row **10** |
| `Incurred Losses` | C13 | row **13** |

Records are columns D–I. Cells read: rows `8`, `10`, `13` × columns D–I = **the same 18
cells, rotated.**

Both orientations produce an identical step-1 block, differing only in whether
provenance reads `Source row` or `Source column`. That equivalence is the test that the
rule set holds.

### 8.4 Types

Each declared field has a type, applied **at extraction**, not in step 2. Step 1 already
carries control sums, and a sum over text silently under-counts — so typing must happen
before anything is added up. This keeps the two steps clean: **step 1 is reading, step 2
is business transformation.**

| Field of dataset 01 | Type | Written as |
|---|---|---|
| `Year` | text | `@` |
| `Premium` | number | `#,##0` |
| `Incurred Losses` | number | `#,##0` |

Types are declared **per field**, never inferred from position: on `03. Large Losses`,
`Loss Date` and `Claim Reference` occupy the positions a measure would and must never be
summed. Only fields typed `number` enter a control sum.

They are declared in **`⟦TYPES⟧` of sheet 00**, by field name, so the frame states both
what a field is called and how it is read.

**Dates** follow the same principle. `2023-06-15`, `15.06.2023` and a real Excel date
cell all read; `01/02/2025` does **not**, because it is 1 February or 2 January
depending on convention and guessing moves a loss between years. A sheet resolves it by
declaring `Date format`. Dates are written as dates, formatted `yyyy-mm-dd`.

**Two rules govern every conversion:**

1. **Empty stays empty.** `None` is an absence, not a zero. Zero is a claim about the
   data; conflating them understates a loss ratio without leaving a trace.
2. **Ambiguity is fatal.** `"1.234"` is 1234 under a German convention and 1.234 under
   an English one. Guessing is a 1000× error in a premium figure, so an ambiguous form
   stops the run and says which readings were possible.

Unambiguous forms are read: `"15 900"`, `"15'900"`, `"15,900.50"`, `"15.900,50"`,
`"1.234.567"`, `"(500)"` for a negative, and a trailing currency or unit symbol.

**Reporting.** A conversion is *routine* when the declared type is simply applied to a
well-formed cell (`2021` read as `"2021"`); it is *notable* when the source was
malformed — a number arriving as text, a date where a year was expected. Only notable
conversions appear in the step-1 block, since a block a human must read is worth keeping
legible; both kinds reach the debug log.

**Why `Year` is text.** A pack can carry `2026` and `2026 9 months` — a full year and a
partial period — as two separate records. The year alone therefore does not identify a
record; **the full label does.** Kept as text, the label *is* the identity: never parsed
back to a number, never normalised, never merged. See §9.2 S10.

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
| O7 | **Both steps or neither.** A dataset with no step-2 spec is not written at all, and the run reports an **error** |
| O8 | A group check is written last, 3 blank rows below everything else — §10.3 on sheet `01`, §10.4 on the aggregate sheet |

O2 uses the last non-empty row of the whole sheet, not the last record, so that
footnotes and totals below the data are never overwritten.

**O7.** The alternative — writing step 1 alone with a note that step 2 is undefined —
was rejected. A step-1 block sitting in an output workbook *looks finished*: it has its
records, its control sums and its tie-back, and a reader who did not write the tool has
no reason to treat it differently from any other block. The note is one line among
twenty. Leaving the sheet exactly as the source had it cannot mislead anyone, and the
`error` status makes the gap impossible to miss in the run report.

The cost is real and accepted: hypotheses the tool raised correctly about that sheet are
computed and then discarded. A missing spec is a gap in the *tool*, and until it is
filled the tool has nothing to say about that sheet.

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

### 9.2.1 Step-2 rules

| | Rule | Kind |
|---|---|---|
| **S1** | Sort by the key field ascending, **natural alphanumeric** | value-preserving |
| **S2** | Column order **exactly as declared in sheet 00** | value-preserving |
| **S3** | Derived columns appended after the declared ones | value-adding |
| **S4** | Derived measures written as **live Excel formulas** | value-adding |
| **S5** | Provenance stays the leftmost column | — |
| **S6** | Value-preserving steps must not move a numeric total; if one does, the run fails | check |
| **S7** | Control sums tie back to step 1 and must be unchanged | check |
| **S8** | Confidence carried forward from step 1 | — |
| **S9** | Every rule applied is written into the block in plain language | — |
| **S10** | Key labels preserved verbatim — never parsed, merged or de-duplicated. A repeated leading number raises an overlap hypothesis | — |
| **S11** | **Step 1 carries only extracted information.** Every computed figure belongs to step 2 | — |
| **S12** | Figures relating *two records* are written as **block-level derived figures** beneath the data, not as columns | value-adding |
| **S13** | Within one leading number, suffixes sort by the rank declared in `⟦PERIOD ORDER⟧`, not alphabetically | value-preserving |
| **S14** | A dataset may declare an **aggregate table**: a second step-2 table grouping the detail and summing its measures | value-preserving |
| **S15** | A dataset may declare **several** aggregate tables. Each must reach the same total as the detail, and an attribute may redirect a grouping | value-preserving |
| **S16** | A declared field the source does not supply may be **derived in step 2** from one it does — never in step 1 | value-adding |
| **S17** | A dataset may declare **cumulative** columns: a running share of a measure's total | value-adding |
| **S18** | A dataset may declare its key **complete** against a catalogue in `00`. Keys the source omits are written with every measure at 0 | value-preserving |
| **S19** | A dataset may declare a field as the **identity** of others: supplied, it is checked against them and the verdict recorded; absent, it is derived as their sum | check / value-adding |
| **S20** | A dataset may declare **target cells**. A source reporting on any other axis is bridged onto them using ratios read from `08`, or declared in `⟦SPLITS⟧` where no cedent can supply them | value-adding in detail, value-preserving in sum |

**S14.** `03. Large Losses` emits the claim detail and then, beneath it, the annual sum:

```
Annual sum of large losses
Years with no record are shown as 0: 2022
  Year   Loss amount
  2021         2,770
  2022             0      ← no large loss that year
  2023         5,330
  2024         2,600
  2025         2,060
  Control (n = 5)  12,760  — must equal the detail total
```

The zero row matters: **an absent year reads as "no data", a zero reads as "nothing
happened"**, and only one of those is true. The window comes from `01. History`, so
every year of the history has a counterpart here. Grouping is value-preserving, so the
aggregate total must equal the detail total or the run fails.

The aggregate also makes R-01 expressible — see §10.1.

**S15 — one set of losses, three questions.** `04. Cat Losses` declares:

| Field | Type | |
|---|---|---|
| `Year` | text | on the treaty's year basis |
| `Event ID` | text | **the grouping key** — a PERILS/PCS code or the cedent's own |
| `Event Name` | text | the human label |
| `Loss amount` | number | this event's share of this year |
| `Event Date` | date | when the event began |
| `Event End Date` | date | when it ended |
| `Number of Claims` | number | **optional** — see §8 F4 |

A cat event is not confined to one underwriting year: risks written in 2022 and in
2023 were both on cover when the Aegean earthquake struck, so the cedent reports the
event **once per underwriting year it touches**. `Event ID` — not the name, not the year
— is what says those two rows are one event.

That single fact makes three different tables all correct at once:

```
Windstorm — 5 records, 3 events, 14,800 in total

Annual sum of cat losses      By occurrence year          By event
  group by Year                 group by year(Event Date)   group by Event ID
  2021    1,400                 2022    5,400               EV-201   5,400
  2022    4,000                 2024    9,400               EV-202   7,600
  2023        0                                             EV-203   1,800
  2024    2,800
  2025    6,600
  ─────────────                 ─────────────               ─────────────
         14,800                        14,800                      14,800
```

Each answers a question the others cannot:

| Table | Question | Ties to |
|---|---|---|
| Annual sum | what did this treaty year cost? | `01. History` — this is R-02's table |
| By occurrence year | when did the losses actually happen? | nothing; informational under a UW basis |
| By event | what did the event cost? | the figure a cat layer is priced against |

`EV-202` — Windstorm Bettina, 2024-12-30 to 2025-01-02 — is the reason the third table
exists. Its 7,600 appears in **no** annual row: 1,000 landed in underwriting year 2024
and 6,600 in 2025. A reviewer looking only at annual sums would price the layer against
6,600 and be wrong by a fifth.

Bettina is also why `Event End Date` is extracted. The end date is not used to
reconstruct or reallocate anything — the client reports the split, and the tool does not
second-guess it — but a reviewer must be able to see *why* one event sits in two years,
and the pair of dates is what shows it.

**Which date defines the occurrence year is the cedent's convention, not the tool's.**
The default is the event's start; a pack that counts an event by when it ended declares
so in column A, and the table follows without any code changing:

```
Occurrence year from = Event End Date       →   Bettina's 7,600 moves to 2025
```

The attribute must name a field the block actually extracts; naming anything else is
fatal, not silently ignored. And S6 still binds every table: all three groupings of the
same records must reach the same total, or the run fails.

**S16 and S17 — the risk profile.** `05` is where both appear. A band's numeric bounds
are declared optional (§2.4) because the source supplies them either as two columns or
only inside the label. Where they are absent, step 2 reads them off the label and names
what it read:

```
· Band from and Band to read off Band — the source declares no such column
  (value-adding): '0 – 1 000' → 0 … 1,000; '1 001 – 5 000' → 1,001 … 5,000 …
```

The division is exactly S11's: step 1 shows four columns because the sheet has four;
step 2 shows six because it worked two of them out, and says so. The derived columns
still appear in **the order sheet 00 declares** (S2), not appended at the end — they are
declared fields that happened to arrive inside another one.

Sorting then happens on the bound as a *number*, which is the whole reason the bounds
exist: natural alphanumeric would place `10 001 – 25 000` before `5 001 – 10 000`.

**S17.** A profile is read cumulatively — "where does the book sit?" is a question no
single band answers. So `05` emits three running shares, each a live formula over the
range above it:

```
=IFERROR(SUM(G$58:G60)/SUM(G$58:G$62),"")
```

A reviewer can see both the running sum and the total it is divided by, which a bare
percentage would not show.

**S18 — the zone that isn't there.** `06` and `07` declare their key complete against
`⟦ZONES⟧` (§2.5). Step 2 writes every declared zone, and reports the ones it added:

```
· 46 declared zone(s) with no entry in the source, shown as 0: 3, 4, 5, 6, … (value-preserving)
```

It is value-preserving in the strict sense S6 requires — adding zeros cannot move a
total, and the control sum proves it did not. What it changes is what a reader sees: a
missing row is silence, and a zero is an answer.

The catalogue lives in the workbook rather than in the code because zonings are a matter
of the territory, not of the tool. A pack for another country declares its own `⟦ZONES⟧`
and nothing else changes.

**S19 — a total is a claim, and claims get checked.** `06` declares `Total` as the
identity of its nine buckets. Which way that runs depends on the source:

| The source gives | Step 2 does | |
|---|---|---|
| nine buckets and a total | checks the sum against it, per record | `Total checked against 9 part(s): all agree` |
| nine buckets, no total | derives it as their sum | value-adding, and labelled so |
| a disagreement | **reports it** | `Total checked against 9 part(s): 2 record(s) DISAGREE` |

The third row is the one that matters. The tool does not know whether the total or one
of the nine is wrong, so it does not touch either: it names the records and leaves the
figures as the cedent sent them. Correcting a source figure would break the audit
baseline (§1), and a quietly corrected total is worse than a visible contradiction —
the contradiction is a question for the cedent, and it should reach them as one.

**S20 — the split.** Where a cedent reports at a coarser level than the dataset's buckets,
step 2 multiplies the missing axes onto what came. The mechanics and the reasoning are in
§2.6; what belongs here is the arithmetic's one guarantee: **the record total does not
move.** A split is the rare step that is value-adding in detail and value-preserving in
sum, and S6's control check proves it on every run — which is what keeps an assumed
occupancy mix from leaking into the growth analysis (§10.4).

The block's confidence follows the weakest thing in it. A block extracted cleanly reads
`Confirmed`; once three of its five columns exist only because a declared ratio was
applied, it reads `Assumed`, and says how many columns and on whose ratio:

```
Confidence: Assumed — step 1 read Confirmed, but 3 column(s) rest on a declared
ratio rather than on a reported figure
```

**Bounds are never summed.** A band bound is a number but not a quantity: totalling the
lower edges of a profile produces a figure that means nothing and would sit in the
control row inviting interpretation. So the control sums cover the **measures** —
`Premium`, `Number of Risks`, `Exposure` — and the bounds are left out of them, in both
steps and in the reference sheets' own total rows.

**S11.** A reviewer must be able to compare step 1 against the source cell by cell with
nothing interposed. Step 1's control sums are not an exception: they *verify the
extraction* rather than deriving a business measure, and they are what makes the
cell-by-cell comparison checkable.

**S13.** With the declared order in place, sheet 02 reads chronologically:

```
2025 est · 2025 9 months · 2025 re-est · 2026
```

Without it, alphanumeric ordering would give `2025 9 months · 2025 est · 2025 re-est`,
which invites the reader to compare the wrong pair. The fallback is unchanged: no
declared order means plain natural sort.

**S12.** `Loss Ratio %` describes one record, so it is a column. `Estimation error` and
`Implied growth` each relate two records, so they have no per-row meaning:

```
Derived figures
  Estimation error   2025 re-est 19,200 ÷ 2025 est 18,500 − 1   =  +3.8%
  Implied growth     2026 20,900 ÷ 2025 re-est 19,200 − 1       =  +8.9%
```

Where a record is missing, the figure reports *why* rather than being omitted.

**S1 — natural sort.** Each label splits into runs of digits and non-digits; digit runs
compare as numbers, the rest as text. This keeps same-year variants adjacent *and*
orders correctly by magnitude, which neither a plain string sort nor a numeric sort
manages alone:

```
2021 · 2022 · 2023 · 2025 · 2025 9 months
999  · 2020 · 2021 · 2022 · 10000
```

**S2 — order from sheet 00.** The target order is the order of D5, E5, F5 …, not a
separate setting. This makes 00 the single source of truth for both *what* is extracted
and *in what order* it appears, and it is what "columns swapped" means in practice: in
the source, `Premium` sits in D and `Incurred Losses` in G with unrelated columns
between them; step 2 puts them in declared order, adjacent.

**S10 — overlapping periods.** When two extracted records share a leading number, the
column total counts that year more than once:

```
H-04  Period overlap 2025  2025 · 2025 9 months  Assumed  open
      Year 2025 appears 2 times as overlapping periods; the control sum counts
      each of them, so it verifies extraction rather than being a portfolio total.
```

The tool **flags and changes nothing**. Which record to use is an underwriting
judgment, not the tool's to make. The control sums remain valid for what they are:
**they verify extraction fidelity, not business meaning.**

A per-record ratio such as `Loss Ratio %` stays sound on a partial period, because
premium and losses come from the same months. Only *comparing* it to a full year, or
summing across overlapping periods, is not — which is what the flag exists to say.

### 9.3 Example — step 1 for `01. History`

| | Source row | Year | Premium | Incurred Losses |
|---|---|---|---|---|
| | 9 | 2021 | 15,900 | 14,930 |
| | 10 | 2022 | 16,740 | 10,030 |
| | 11 | 2023 | 17,520 | 12,660 |
| | 12 | 2024 | 18,390 | 11,700 |
| | 13 | 2025 | 19,200 | 10,250 |
| | 14 | 2025 9 months | 14,400 | 7,100 |
| **Control** | *n = 6* | | **102,150** | **66,670** |
| **Check vs. source total row 15** | | | 102,150 ✓ | 66,670 ✓ |

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

Declared in `⟦RULES⟧` of sheet 00, in condensed form:

Each side takes one of three forms, distinguished by shape alone:

| Form | Meaning |
|---|---|
| `01 History.Premium@{N}` | a field on the record matching `{N}` |
| `01 History.Year basis` | an **attribute** of the sheet — no record selector |
| `SUM(03 Large.Loss amount@{Y})` | the field **summed** over every matching record |

`{N}` and `{N+1}` resolve from `⟦GLOBAL⟧`. `{Y}` is a **wildcard**: the rule is declared
once and evaluated once per year present in the data, reported as `R-01/2023`.

| ID | Left | Rel | Right | Tol | Scope |
|---|---|---|---|---|---|
| R-05 | `01.Premium@{N} 9 months` | `=` | `02.EPI@{N} 9 months` | 1 | *all* |
| R-06 | `01.Premium@{N}` | `=` | `02.EPI@{N} re-est` | 1 | *all* |
| R-07 | `02.EPI@{N} re-est` | `>=` | `02.EPI@{N} 9 months` | 0 | *all* |
| R-01 | `SUM(03.Loss amount@{Y})` | `<=` | `01.Incurred Losses@{Y}` | 1 | **per risk** |
| R-02 | `SUM(04.Loss amount@{Y})` | `<=` | `01.Incurred Losses@{Y}` | 1 | **cat** |
| R-09 | `01.Year basis` | `=` | `03.Year basis` | 0 | **per risk** |
| R-10 | `01.Year basis` | `=` | `04.Year basis` | 0 | **cat** |

**Scope selects the sections a rule runs over**, and the rule is evaluated once per
applicable section — reported as `R-01/Fire/2023`. On Fire + EQ + Wind, R-05 runs three
times, R-01 only on Fire, R-02 only on the two cat sections. Adding a section to
`⟦SECTIONS⟧` brings its checks with it.

R-05 is the strongest tie: the same nine months of the same year, reported in two
sheets. R-06 catches an estimate masquerading as an actual — a full-year figure for N in
`History` cannot be an actual, because N is not over. R-07 is arithmetic: premium
accrues, so a re-estimate cannot fall below what is already booked.

Every tolerance is explicit, because figures rounded to thousands never satisfy exact
equality and a check that always fails gets ignored.

**Records resolve by leading number, then suffix.** `{N+1}` matches a lone record for
that year whatever its own suffix, so both `2026` and `2026 est` work; where several
records share a year, the suffix must match exactly. No label is ever rewritten.

#### Attribute compatibility is a precondition, not part of the comparison

Before any rule is evaluated, `Premium basis`, `Loss basis`, `Share basis` and
`Currency` must agree on both sides. `Share basis` is `100%` (cession plus retention)
or `ceded only`; comparing a 100% loss against a ceded premium is the same class of
error as GWP against GNPI. `Scale` is normalised — the comparison happens in the left block's scale.

Where a basis or currency differs, **the rule is not evaluated**. It is reported as
*skipped*, with the reason, and raises a question. Quietly comparing GWP against GNPI
would be worse than not checking at all.

A rule that fails or is skipped becomes a hypothesis in the step-1 block of **both**
sheets it names, so a reviewer reading either one sees the tie.

#### Four outcomes, not two

| | Meaning |
|---|---|
| `passed` / `failed` | evaluated |
| `not applicable` | the dataset is absent from this pack — structure, not a problem. Raises nothing |
| `skipped` | it should apply but could not be evaluated — **a question** |

The distinction matters, and `⟦SECTIONS⟧` is what states it. A Nat Cat treaty declares
no per-risk section, so R-01 and R-09 are *not applicable* and stay silent, rather than
filling the log with complaints about a sheet that correctly does not exist. Equally, a
Fire treaty reports R-02 as not applicable.

### 10.2 Occurrence-year consistency

Where `Year basis = Occurrence`, the year must equal the year of the loss date — a loss
dated `2023-06-15` cannot sit in the `2022` row. Under an **underwriting** basis it
legitimately can, because a policy incepting in one year produces losses in the next.

So the check runs only where the declared basis makes it meaningful, and stays silent
otherwise. A tool that flagged this under a UW basis would be wrong, not thorough.

**Exactly one date defines the year.** A cat record carries two — `Event Date` and
`Event End Date` — and only the one the cedent counts from is compared. An event running
`2024-12-30 → 2025-01-02` is not misfiled because its *end* falls in the next year, and
checking every date field would report it as an error. The field is the one named by
`Occurrence year from`, defaulting to the first date field declared; naming a field the
block does not extract as a date is fatal.

#### Still to be declared

| Rule | Left | Right | Relation |
|---|---|---|---|
| R-03 | `10` Triangles latest diagonal, year *y* | `01` Incurred Losses, year *y* | = |
| R-08 | `02` EPI `{N} est` | last year's pack, same label | = |

R-08 needs a stored prior run.

### 10.3 Declared losses against total incurred

`⟦RULES⟧` compares one figure with one other figure. This check cannot be written that
way, because **both sides are sums over several blocks**: a treaty may carry large
losses and cat losses at once, and a Nat Cat treaty carries one cat sheet per peril. So
it is made over a **section group** — every section of one kind — and asks two questions
of each year:

| | Question | Outcome |
|---|---|---|
| **Is it possible?** | Σ declared losses ≤ total incurred, same year | `EXCEEDS` — an **error** |
| **Is it ordinary?** | Σ declared losses ÷ total incurred ≤ threshold | `WARNING` — not an error |

Grouping by section *kind* covers every case with one rule:

| Treaty | Group | Declared losses | Held against |
|---|---|---|---|
| Engineering / Miscellaneous | one per-risk section | `03` **+** `04`, whichever exist | that section's `01` |
| Fire only | one per-risk section | `03` alone | that section's `01` |
| Fire + Nat Cat | *per risk* | `03` (Fire) | Fire's `01` |
| | *cat* | `04` (EQ) **+** `04` (Wind) | EQ's `01` **+** Wind's `01` |

**Why the sum matters.** R-01 and R-02 already check each loss dataset on its own. Only
the sum catches large *and* cat losses that are each plausible and jointly are not — the
Engineering case, where a year's large losses and its hailstorm are separately unremarkable
and together exceed what the portfolio incurred.

**Why the share matters.** A year whose declared losses make up more than the declared
share of total incurred was driven by a handful of events rather than by attrition, which
changes how it is rated: those events are normally stripped out and rated separately. That
is a *fact about the portfolio*, not a mistake in the pack, so it is a warning and never
sets the run to failed.

The threshold is declared in `⟦GLOBAL⟧` as `Loss share warning`, so it is the underwriter's
number rather than the tool's. `20%`, `20` and `0.2` are all read as 20% — a share above 1
can only have been meant as a percentage. Where it is absent, 20% applies and the written
table says so.

```
DECLARED LOSSES AGAINST TOTAL INCURRED — PER RISK SECTIONS
Sections summed: Engineering · scale 1,000
Declared losses = Large losses (03) + Cat losses (04); total incurred = Incurred Losses in 01.
Two questions per year: declared ≤ incurred (tolerance 1), and declared ÷ incurred ≤ 20%.

  Year   Large (03)   Cat (04)   Declared   Incurred (01)   Share   Status
  2021        1,420      1,150      2,570           5,100   50.4%   WARNING
  2022            0          0          0           6,350    0.0%   within limits
  2023        3,170        900      4,070           4,820   84.4%   WARNING
  2024        1,975      2,600      4,575           7,240   63.2%   WARNING
  2025        1,180          0      1,180           5,680   20.8%   WARNING
  Control (n = 5)
WARNING: 2021, 2023, 2024, 2025. Not an error: a fact about the portfolio, raised so
it is priced knowingly.
```

**What it refuses to do.** The same precondition as §10.1, narrowed to what a loss
comparison needs — `Loss basis`, `Share basis` and `Currency` must agree across every
block summed. A 100% figure added to a ceded-only one is a number that means nothing, so
where they differ the group is **not evaluated** and says why. Scale is normalised to the
group's `01` blocks.

A section contributes **once**. `Intake_v1.xlsx` carries `01. History` and its transposed
twin, which are the same portfolio in two shapes; summing both would double it.

Where only some sections of a group report a history row for a year, the incurred total
for that year covers fewer sections than the declared losses do. The comparison still
runs — it is conservative, since a smaller basis can only make a finding more likely,
never hide one — and the years affected are **named** in the table rather than the gap
being closed by assumption.

**Where it is written.** At the end of sheet `01`, three blank rows below everything else
(§9.1 O8) — it is about the treaty rather than about one block. Where a group spans
several `01` sheets, as EQ and Wind do when each has its own, the same table is written at
the end of each and each copy names the others, so a second copy does not read as a second
finding. The per-role columns appear only where there is more than one role to split; with
one they would merely repeat the total.

The status word is `within limits`, never `OK`: step 1 and step 2 already write an
`OK`/`MISMATCH` control check, and one word must not mean two things on one sheet.

### 10.4 Exposure growth against premium growth

Like §10.3 this is a group check rather than a `⟦RULES⟧` line, because it spans three
blocks and reaches into two other datasets. It exists because **no version of a cat
aggregate is interesting on its own** — the movement between them is, and only against
what the premium did over the same span.

The figure the block produces is the **implied rate change**:

```
(1 + premium growth) ÷ (1 + exposure growth) − 1
```

Premium up 9% carried on 12.8% more exposure is a rate *cut* of about 3.4%, however the
premium line reads by itself. That is the number a renewal turns on, and it is the bridge
to `09. Rate Development`: if the cedent claims +4% and this says −3.4%, one of the two
is wrong and the difference is worth a conversation.

| Side | Where it comes from |
|---|---|
| Exposure, expiring | first version's `Total` summed over all zones — the `{N} 9 months` estimate |
| Exposure, renewal | last version's `Total` summed over all zones — the `{N+1} at expiry` projection |
| Premium, expiring | `01` `Premium`, year N |
| Premium, renewal | `02` `EPI`, year N+1 |

Premium for the renewal year comes from `02` because **`01` has no forward figure**. A
history sheet ends at the expiring year by definition, so the EPI re-estimate is the only
statement about N+1 there is.

The exposure figures are taken from the **step-2** records, not step 1: completing the
zone list only adds zeros and cannot move the sum, but a *derived* `Total` (S19) exists
nowhere else.

```
EXPOSURE AND PREMIUM GROWTH — EARTHQUAKE
Exposure is the sum of the zone totals of each version; premium is the expiring year
from 01 and the renewal year from 02, since 01 carries no forward figure.
Implied rate change = (1 + premium growth) ÷ (1 + exposure growth) − 1. Beyond ±20% it
is flagged — a warning, never an error: a book may shrink or grow for good reasons.

  Version             As at         Exposure   Change
  2025 9 months       30.09.2025     311,508
  2026 at inception   01.01.2026     330,819    +6.2%
  2026 at expiry      31.12.2026     351,379    +6.2%

  Premium, expiring year   2025          1,512
  Premium, renewal year    2026 EPI      1,648

  Exposure growth                       +12.8%
  Premium growth                         +9.0%
  Implied rate change                     -3.4%
  Status                            within limits
```

**Nothing here can fail a run.** Portfolios shrink for good reasons — a cedent drops a
segment, a currency moves, a large scheme leaves — and they grow for good reasons too. An
earlier draft of these rules held that exposure should track premium and that a book
should not shrink; both were wrong as stated, and neither survives. What remains is a
threshold, `Rate change warning` in `⟦GLOBAL⟧`, beyond which the movement is **raised**
so that the renewal is priced knowingly. Exactly as §10.3 treats an unusual loss share,
and with the same status word, `within limits`.

**One table per cat section per aggregate role**, so a treaty covering both perils gets an
earthquake table and a hurricane table, each written at the end of its own aggregate sheet
(§9.1 O8) rather than on `01` — it is a statement about that aggregate. Where a version
carries no `Period`, the versions cannot be ordered and the table says `NOT EVALUATED`
naming the blocks, rather than guessing a sequence.

---

## 11 · Logging

Two logs, linked to the sheets by block ID.

**Debug log** — technical and verbose: resolved address maps, cell ranges, coercion
failures, timings.

**Process log** — decisions and interpretations in plain language, for a reviewer:

> `01 History` / STEP 1 — header row 8, selector column L. Column `Premium` resolved to
> D, `Incurred Losses` to G. 7 candidate rows, 6 extracted, 1 excluded (row 15, selector
> empty). Currency USD, scale 1,000. Year basis assumed UW (hypothesis H-01, open).
> 2025 appears twice, as a full year and as its first nine months (H-04, open).

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
| T4 | Pivot tables are **not carried into the output**. Reviewers do not need them there, and the source workbook is preserved untouched, so plain `openpyxl` is sufficient throughout |
| T5 | A run may not write over its own source: the source is the audit baseline, and an output path equal to it is refused |

T2 exists because a silently zeroed error cell reaches an underwriter under a
clean-looking control sum.

T3 earns its keep on a hand-assembled workbook: copied sheets carry formulas pointing at
absent source files, which resolve to stale caches or `#REF!`.

---

## 13 · Decisions taken, and what remains

Resolved:

| | Decision |
|---|---|
| Pivots in the output | **No.** Plain `openpyxl` throughout |
| Pivot source ranges | Not a concern, since pivots are not carried forward |
| Mandatory attributes | **Mandatory.** The attribute list in `00` is the checklist; an undeclared attribute ranks `Open` |
| Sheet 08, split axes | **Occupancy (Res / Com / Industrial) is common to both variants**; the second axis is B / C / BI for Fire, Projects / Renewables for Engineering — §2.1 |
| Hypothesis register | **No separate sheet.** Documented in-sheet beneath the original data |
| Row-level confidence | Off, **except for transposed blocks** |
| Hours clause | **Not modelled.** The cedent reports the event split; the tool shows it, and does not redo it |
| Number of claims on `04` | **Optional** — declared `(optional)` in sheet 00 |
| Dataset with no step-2 spec | **Nothing written**, run reports an error — §9.1 O7 |
| Loss-share threshold | **20%**, declared in `⟦GLOBAL⟧` so an underwriter can change it |
| Rate-change threshold | **20%**, likewise — §10.4 |
| Header register width | **Read from row 4**, not fixed at ten columns; older packs fall back to `D … M` — §3.1 |
| `06` vs `07` | **No rule compares them.** Either peril can be bought alone, and the covered books need not match — §2.5 |
| Zones with no exposure | **Written as 0.** Where the cedent already sends them, step 1 shows them; where it does not, step 2 completes from `⟦ZONES⟧` — an absent row is silence, a zero is an answer |
| A `Total` disagreeing with its parts | **Reported, never corrected** — §9.2.1 S19 |
| Occupancy / cover splits | Buckets are the **product of declared axes** (`⟦AXES⟧`); a coarser report is expanded with ratios from `⟦SPLITS⟧` — §2.6 |
| Where the ratios come from | The **cedent**, for the whole book — so the ratio is a reported figure, and `⟦SPLITS⟧` records its `Source`. An underwriter may substitute a prior year's or another client's split; it is then marked as an assumption |
| Applying a book ratio per zone | **Accepted knowingly.** The zone-level mix is an assumption; the book-level split is not, and the bucket totals reproduce it exactly |
| Merged `Commercial` | **75% Com / 25% Ind**, a house convention in `⟦SPLITS⟧` |
| Splitting across zones | **Never.** The geography must be reported; a `06`/`07` block without `Zone` is refused |

### 13.1 Parked — decided against for now, or awaiting a decision

Each of these is a real gap with a known shape. None is a defect: the tool refuses
cleanly in every case, and none is worth building before the decision behind it is made.

| | What | Why it is parked |
|---|---|---|
| **P1** | **`05` ↔ `01` premium comparison.** Σ profile premium against the portfolio premium for the same period | `Includes fac` means the two *legitimately* differ, so it cannot be a pass/fail rule. It wants a §10.3-style comparison block naming the attributes that explain the gap — and which premium it ties to (`01` actual or `02` EPI re-estimate) is not yet decided |
| **P2** | **A declared number format.** `⟦GLOBAL⟧ Number format = 1,000.00` \| `1.000,00`, mirroring `Date format` | `1,000` is genuinely ambiguous — one separator, exactly three trailing digits — so it is refused (§8.4). `1,000,000` and `1.000.000` both group unambiguously and are read. A declared convention would settle the four-digit case everywhere at once, not just in band labels |
| **P3** | **`@*` — every record.** `SUM(05.Premium@*)` | Every record selector today is year-shaped, so a dataset whose records are bands cannot be summed in a rule. No rule needs it until `08` exists and two non-year tables must tie |

Remaining:

1. **Datasets `09`–`11`.** Header labels and attributes not yet specified. `08`'s axes
   are settled (§2.1) and its split machinery is shared with `06`/`07` (§2.6).
2. **Sheet `08`'s measures.** What is *counted* at each intersection — risk count,
   sum insured, premium — is the last open question on `08`.
3. **Sheet `20. Summary`.** Specified in §9.1 O6 but not yet implemented; it needs at
   least two datasets to be meaningful.

---

## 14 · Implementation

```
datatransform/
    nomenclature.py   sheet 00: dataset register and vocabulary   (§3)
    markers.py        column A marker grammar                     (§4)
    extract.py        field resolution and record selection       (§7, §8)
    coerce.py         type coercion                                (§8.4)
    crosschecks.py    interdependencies between sheets             (§10.1)
    lossshare.py      declared losses against total incurred       (§10.3)
    transform.py      step 1 and step 2                           (§9.2, §10)
    specs.py          step-2 mechanics, per dataset
    writer.py         block layout, control sums, anchors         (§9, §10)
    recalc.py         cached formula results                      (§7.2)
    runner.py         orchestration and the two logs              (§11)
```

```
tools/build_intake.py   regenerates the reference workbook
```

Run it:

```bash
python -m datatransform Intake_v1.xlsx -o output/Intake_v1_transformed.xlsx
pytest tests/
```

The pipeline makes **two passes**: every sheet is extracted first, so `⟦RULES⟧` has both
sides of each crosscheck available, and only then is anything transformed and written.

### 14.1 Step 2 for datasets 01 and 02

Held in `specs.py`, since sort order and derived measures are mechanics rather than
facts a human declares:

| | |
|---|---|
| Sort | `Year` ascending, on the numeric reading — value-preserving |
| Column order | `Year` \| `Premium` \| `Incurred Losses` — value-preserving |
| Calculation | `Loss Ratio % = Incurred Losses / Premium` — value-adding, written as a live formula |
| Types | `Year` text · `Premium`, `Incurred Losses` number — see §8.4 |

The value-preserving steps are checked: if sorting or reordering moves a measure total,
the run fails rather than reporting a plausible wrong number.

### 14.2 The reference workbooks

Seven treaty shapes, generated by `tools/build_v1.py`, `tools/build_intake.py` and
`tools/build_mexico.py`:

| Workbook | Sections | Demonstrates |
|---|---|---|
| `Intake_v1.xlsx` | Fire | both orientations; R-02 not applicable |
| `Intake_Engineering_v1.xlsx` | Engineering | **one per-risk section carrying both `03` and `04`** — the §10.3 sum |
| `Intake_EngineeringCombined_v1.xlsx` | Engineering | the same treaty with **both loss datasets in one list** — §6.5 |
| `Intake_FireCat_v1.xlsx` | Fire, Earthquake, Windstorm | the two cat sections **share one sheet** |
| `Intake_FireEQWind_v1.xlsx` | Fire, Earthquake, Windstorm | **a sheet per section** |
| `Intake_FireCatLosses_v1.xlsx` | Fire, Earthquake, Windstorm | **only `01`, `02` and `04`** — the money and the events. Fire declares no loss dataset, so R-01/R-09 are *not applicable* rather than failed |
| `Intake_Mexico_v1.xlsx` | Earthquake, Hurricane | the **cat aggregates** — eleven headers, `⟦ZONES⟧`, three versions per sheet, §10.4 |

`Intake_FireCat_v1.xlsx` and `Intake_FireEQWind_v1.xlsx` carry identical figures and are
verified to produce identical crosschecks — the test that a section really is just a
block.

Their `04` blocks are built to exercise S15 rather than to look tidy:

| | |
|---|---|
| `EV-101` Aegean earthquake | one event, **two underwriting years** (2022 and 2023) |
| `EV-202` Windstorm Bettina | 2024-12-30 → 2025-01-02: **spans the renewal date** |
| Earthquake | reports `Number of Claims` |
| Windstorm | **does not** — the optional field is simply absent |
| Earthquake | no losses in 2024 or 2025, so the annual table shows the zero rows |

`Intake_v1.xlsx` implements this specification for:

- `00. NC+Interdep` — dataset register, vocabulary, marker legend
- `01. History` — row-wise, 6 records
- `01. History_Transposed` — transposed, 6 records
- `02. EPI Projections` — row-wise, 4 records: `N est`, `N 9 months`, `N re-est`, `N+1`
- `02. EPI Projections_Transposed` — transposed, the same 4 records
- `03. Large Losses` — row-wise, 8 claims across 2021, 2023, 2024 and 2025; **2022
  deliberately has none**, so the annual table exercises the zero row

Each dataset is present in both orientations and verified to produce identical output —
same records, same totals, same derived figures — differing only in whether provenance
reads `Source row` or `Source column`.

Both `01` sheets carry identical data and are verified to produce identical output,
differing only in whether provenance reads `Source row` or `Source column`.

The six records are the full years 2021–2025 plus `2025 9 months`, so the reference
workbook exercises S10: two records share the year 2025, the overlap hypothesis is
raised, and the control sum counts that year twice — deliberately, since it verifies
extraction and not business meaning.

`Intake_FireCatLosses_v1.xlsx` is the shape a cedent sends when the submission is the
money and the cat events and nothing else. It exists for two things the fuller packs
cannot show:

| | |
|---|---|
| A rule whose dataset is absent | Fire carries no `03`, so R-01 and R-09 report **not applicable** and the run stays clean — §10.1's four outcomes, exercised on a real pack |
| A group with nothing to compare | §10.3's per-risk group finds no loss dataset at all, so **no block is written** on `01. History Fire`. The cat group still sums EQ + Windstorm against both histories |

`Intake_Mexico_v1.xlsx` is built to exercise §2.5 and §10.4 rather than to look like a
full submission — it carries `01`, `02`, `06` and `07` and nothing else:

| | |
|---|---|
| `06. EQ Aggs` | 11 headers, **three stacked versions** on one sheet, all 52 zones — three of them at 0 |
| `07. Wind Aggs` | the **other** reporting level — one figure per zone, occupancy added from `⟦SPLITS⟧` (§2.6); three buckets, not nine |
| `⟦ZONES⟧` · `⟦AXES⟧` · `⟦SPLITS⟧` | the catalogues, the axes of each dataset, and one cedent ratio plus the 75/25 house convention |
| Zone order | the source lists zones in catalogue order; the sort is exercised on `13a`/`13b` and `14a`–`14d` |
| `01` / `02` | one section each, carrying only what §10.4 reads: `Premium` for N and `EPI` for N+1 |

A note on the sandbox this was built in: LibreOffice was unavailable, so formula results
are cached by `recalc.py` writing `<v>` alongside `<f>` directly. Where LibreOffice or
Excel is available, either will recalculate the same values on open.
