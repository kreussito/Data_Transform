# Schlachtplan — Refaktorierung Blatt für Blatt

Stand: nach Commit `34e2bad` (Step 2 als deklarierte Pipeline).
Ausgangsgröße: **14.282 Zeilen** in 41 Dateien.

---

## 0 · Das Ziel

Heute ist der Code nach **Phasen** geschnitten — Lesen, Transformieren, Schreiben,
Prüfen —, und jedes Blatt läuft durch alle vier hindurch. Wer 06 ändern will, muss
alle vier verstehen.

Der Endzustand dreht das um: **eine Datei pro Datensatz.**

```
datatransform/
  datasets/
    __init__.py          Registry: Rolle → Datensatz
    _01_history.py       Pipeline + Layout + Kreuzprüfungen von 01
    _02_epi.py
    _03_large.py
    _04_cat.py
    _05_profile.py
    _06_eq_aggs.py
    _07_wind_aggs.py
    _08_splits.py
    _09_rate.py
  extract/               WIE gelesen wird     (Geometrie, beide Orientierungen)
  operations/            WIE gerechnet wird   (fertig — Etappe 1 der letzten Runde)
  render/                WIE geschrieben wird (Vokabular statt Zellkoordinaten)
  crosscheck/            WIE geprüft wird     (eine Form, vier Deklarationen)
  declarations.py        die Sprache von Blatt 00
  model.py               Block, Record, Dataset
  runner.py              der Ablauf, sonst nichts
```

Danach ist „Blatt für Blatt" wörtlich wahr: eine Änderung an 06 berührt genau
`_06_eq_aggs.py`, und die vier Schichten darunter bleiben unangetastet.

---

## 1 · Die Spielregeln

| | |
|---|---|
| **P1** | **Zellgleichheit ist das Abnahmekriterium**, nicht „die Tests sind grün". Bei der letzten Refaktorierung haben 452 grüne Tests zwei Fehler durchgelassen, die der Zellvergleich gefunden hat. |
| **P2** | **Eine Refaktorierung ändert kein Verhalten.** Findet man dabei einen Fehler, wird er notiert und in einem **eigenen** Commit behoben — nie im selben. |
| **P3** | **Jede Etappe endet grün, committet und gepusht.** Nie zwei Etappen gleichzeitig offen. |
| **P4** | **Die Quellmappe wird nie verändert.** Gilt unverändert weiter; jede Etappe muss das beweisen (die SHA-256 im Log). |
| **P5** | **Jede Etappe hat ihren Copilot-Prompt** und läuft auf dem Windows-Rechner. |

---

## 2 · Etappe 0 — Das Sicherheitsnetz

**Ohne dies keine einzige weitere Zeile.** Der Zellvergleich, den wir bisher von Hand
gemacht haben, wird ein committetes Werkzeug.

**Warum zuerst:** Jede der folgenden acht Etappen wird mit demselben Satz abgenommen:
*„alle neun Mappen, Zelle für Zelle identisch"*. Solange das Handarbeit ist, wird es
irgendwann übersprungen — und genau dann geht etwas kaputt.

**Was entsteht:** `tools/golden.py`

- rechnet alle neun Referenzmappen in ein Verzeichnis,
- schreibt pro Mappe eine Textdatei mit **jeder** Zelle: Blatt, Koordinate, Wert,
  Zahlenformat, Füllfarbe, Fett/Kursiv, und bei Formeln sowohl Formel als auch
  zwischengespeicherter Wert,
- schreibt die Logdateien mit hinein (die Etappe 5 braucht das),
- deterministisch sortiert, damit `diff` brauchbar ist.

**Abnahme:** `python tools/golden.py --out golden/before` und ein zweiter Lauf in
`golden/after` ergeben ohne Codeänderung ein leeres `diff -r`.

**Umfang:** ~120 Zeilen neu. **Risiko:** keines.

### Copilot-Prompt

```
Create tools/golden.py — a golden-master harness for refactoring.

Context: this project reads reinsurance submission workbooks and writes two
transformed blocks beneath the original data. Nine reference workbooks live in the
repository root (Intake_*.xlsx). The CLI is `python -m datatransform <source>
-o <output> --log-dir <dir>`. The source workbook is never modified.

Requirements:
1. Discover every Intake_*.xlsx in the repository root, sorted by name.
2. For each, run datatransform.runner.run() into a temporary output workbook and a
   per-pack log directory under the target directory.
3. Dump the OUTPUT workbook to a plain-text file, one line per non-empty cell:
       <sheet>!<coordinate>\t<repr(value)>\t<number_format>\t<fill rgb or ->\t<B|I|->
   Read the workbook twice with openpyxl (data_only=False for formulas,
   data_only=True for cached values) and emit BOTH for any cell holding a formula.
   Sheets in workbook order; cells sorted by row then column.
4. Copy each pack's log files into the target directory as well.
5. Also write summary.txt: per pack the source sha256, the number of formula values
   cached, and each sheet outcome — exactly what the CLI prints.
6. CLI: `python tools/golden.py --out golden/before`. Create the directory; refuse to
   overwrite a non-empty one unless --force is given.
7. Print one line per pack so a long run shows progress.

The point is that `diff -r golden/before golden/after` is empty for any change that is
a pure refactoring. Nothing may depend on dict ordering, temporary paths or timestamps
— filter those out of the logs if the log lines contain them.

Match the style of the existing tools/: module docstring explaining WHY the file
exists, comments only where the reason is not obvious from the code, no type-checking
ceremony. Do not modify anything under datatransform/.
```

---

## 3 · Die acht Etappen

| # | Etappe | heute | danach | Risiko | Nutzen |
|---|---|---|---|---|---|
| 1 | `writer.py` → Rendering-Vokabular | 774 | ~550 | niedrig | **sehr hoch** |
| 2 | Vier Kreuzprüfungen → eine Form | 1.179 | ~450 | mittel | **sehr hoch** |
| 3 | Blatt-00-Leser deklarativ | 435 | ~220 | niedrig | hoch |
| 4 | `model.py` aufteilen | 395 | 2 × ~200 | keines | mittel |
| 5 | `datasets/` — eine Datei pro Blatt | — | ~450 | keines | **das Ziel** |
| 6 | `extract.py` — eine Geometrie | 643 | ~400 | **hoch** | hoch |
| 7 | `runner.py` — Ablauf ≠ Protokoll | 446 | 250 + 160 | niedrig | mittel |
| 8 | `tools/build_*.py` — Mappen als Daten | 2.441 | ~900 | niedrig | mittel |
| 9 | Tests pro Blatt + Dokumentation | 3.600 | — | keines | hoch |

Reihenfolge: **0 → 1 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9.**
Etappe 4 ist so billig, dass sie jederzeit dazwischenpasst. Etappe 6 ist die heikelste
und steht bewusst hinter 5: erst soll die Struktur stehen, bevor wir dort anfassen, wo
schon zweimal ein Fehler saß.

---

### Etappe 1 · `writer.py` → Rendering-Vokabular

**Das Problem.** Jede Schreibfunktion kennt gleichzeitig Spaltennummern, Schriften,
Füllungen, Zahlenformate und Formeltexte. `write_loss_share` ist 113 Zeilen,
`write_growth` 89, `write_rate_claim` 69 — und alle drei haben **dasselbe Skelett**:

```
Anker → Titel → erklärende Prosa → Kopfzeile → Datenzeilen →
Kontrollzeile mit SUMME-Formeln → Befundzeile
```

Dreimal von Hand gebaut. Genau das Problem, das `transform.py` hatte, bevor es zur
Pipeline wurde.

**Das Ziel.** Ein Paket `render/` mit einem kleinen Vokabular:

| Element | was es ist |
|---|---|
| `Anchor(key, step)` | die `⟦DT:…⟧`-Marke, an der ein Neulauf erkennt, was er ersetzen darf |
| `Title(text, fill)` | eine Überschriftszeile |
| `Prose(text)` | eine erklärende Zeile in Kursiv |
| `Table(columns)` | Kopfzeile + Datenzeilen; jede `Column` weiß Kopf, Wert, Format, Formel |
| `ControlRow(measures)` | die SUMME-Zeile samt zwischengespeichertem Wert |
| `TieBack(...)` | die Gegenprobe gegen Step 1 (S6) |
| `Verdict(text)` | die Befundzeile am Ende |
| `Question(text)` | die leere, blau hinterlegte Zelle für den UW (§2.6) |

Der Cursor (`self.row`), das Setzen von Schrift/Format/Füllung und das Mitschreiben der
Formelwerte in den Cache bleiben **an einer** Stelle. Danach ist `write_loss_share` eine
Deklaration von ~25 Zeilen.

**Fallstricke.**
- Der Formelwert-Cache (`formula_values`) muss weiterhin für **jede** geschriebene
  Formel gefüllt werden — sonst ist die Mappe ohne Excel nicht lesbar (T-Anforderung:
  Excel wird nicht vorausgesetzt).
- Die laufenden Anteile (`running_of`) schreiben absolute Bereiche über zwei
  Zeilennummern, die erst nach dem Schreiben der Datenzeilen bekannt sind. Das
  Vokabular muss zweiphasig sein: erst Zeilen belegen, dann Formeln füllen.
- Spalte A bleibt der Markerkanal; geschrieben wird ab Spalte B (`FIRST_COL`).
- Schriften und Füllungen sind Teil der Abnahme, nicht Kosmetik — `golden.py` vergleicht
  sie mit.

**Abnahme.** `diff -r golden/before golden/after` leer; 452 Tests grün.

#### Copilot-Prompt

```
Refactor datatransform/writer.py into a small rendering vocabulary, without changing a
single written cell.

Read first: datatransform/writer.py in full, and datatransform/operations/__init__.py
for the style of the pipeline refactoring that was done to transform.py — this is the
same move applied to writing.

Observation to act on: write_loss_share (113 lines), write_growth (89) and
write_rate_claim (69) have an identical skeleton — anchor, title, prose lines, header
row, data rows, a control row of SUM formulas with cached values, a verdict line. That
skeleton is written out three times by hand.

Create a package datatransform/render/ containing:
- cursor.py   the cell cursor: current row, _put/_line/_formula, font/fill/format
              handling, and the formula-value cache (sheet, coordinate, expected value)
              that keeps the workbook readable without Excel.
- elements.py frozen dataclasses: Anchor, Title, Prose, Table, Column, ControlRow,
              TieBack, Verdict, Question. Each has render(cursor) -> None.
- report.py   a Report: an ordered sequence of elements, rendered in one pass.

Then rewrite write_loss_share, write_growth and write_rate_claim as Report
declarations, and move the step-1/step-2 block writing onto the same cursor.

Hard constraints:
1. Every cell written must be byte-identical afterwards: value, number format, fill
   colour, font (name, size, bold, italic, colour), and for formulas both the formula
   text and the cached value.
2. Two-phase rendering: the cumulative-share formulas span an absolute range whose
   first and last row are only known after the data rows have been placed. Keep that
   working; do not pre-compute row numbers by guessing.
3. Column A stays untouched — it is the marker channel. Writing starts at FIRST_COL.
4. No behaviour change of any kind. If you find a bug, leave it and note it in your
   final message.
5. Keep the existing docstrings' reasoning where it still applies — they explain WHY a
   block is written where it is, which is the part a reader needs.

Verify with: python tools/golden.py --out golden/after && diff -r golden/before
golden/after   (must be empty), and python -m pytest -q (452 passed).
```

---

### Etappe 2 · Die vier Kreuzprüfungen

**Das Problem.** `crosschecks.py` (425), `growth.py` (230), `lossshare.py` (289) und
`ratechange.py` (235) — zusammen **1.179 Zeilen** — tun im Kern dasselbe:

```
Blöcke finden (Rolle, Sektion) → Periode(n) wählen → eine Größe rechnen →
gegen eine ⟦GLOBAL⟧-Schwelle einstufen → Tabelle mit Titel, Zeilen, Status, Gesamturteil
```

Sichtbarster Beleg: `threshold_of()` steht dreimal fast wortgleich in drei Modulen, und
`_incompatible()` zweimal.

**Das Ziel.** Ein Paket `crosscheck/`:

- `base.py` — die gemeinsame Form: `Check`-Protokoll, `Finding`, die vier
  Ergebnisworte (bestanden / nicht bestanden / nicht anwendbar / übersprungen, §10.1),
  Schwellenauflösung aus `⟦GLOBAL⟧`, Sektionsgruppierung, Skalenprüfung.
- `rules.py` — die in Blatt 00 deklarierten Regeln (das heutige `crosschecks.py`).
- `loss_share.py`, `growth.py`, `rate.py` — je ~80 Zeilen **Deklaration**.

**Was dabei *nicht* zusammengefasst werden darf.** Die drei sind sich ähnlich, aber
nicht gleich, und die Unterschiede sind fachlich:

| | was verglichen wird | über was summiert wird |
|---|---|---|
| §10.3 Loss Share | deklarierte Schäden gegen Incurred aus 01 | mehrere Rollen, pro Jahr |
| §10.4 Growth | Exposure gegen Prämie | zwei Fassungen **derselben** Größe |
| §10.5 Rate Claim | behauptete gegen implizite Rate | Sektionen — Prämie addiert, Exposure **nicht** |

Die gemeinsame Form ist *finden → paaren → messen → einstufen → berichten*; die
**Messung** bleibt pro Prüfung deklariert. Wer das zusammenzieht, baut die
Exposure-Addition wieder ein, die wir in §10.5 bewusst ausgeschlossen haben.

**Abnahme.** Zellgleich; zusätzlich müssen die Logzeilen identisch bleiben
(`golden.py` nimmt sie mit auf).

#### Copilot-Prompt

```
Unify the four cross-check modules behind one shape, without changing any result.

Read first, in full: datatransform/crosschecks.py, growth.py, lossshare.py,
ratechange.py, and specification_v1.md sections 10.1 to 10.5.

What they share: find blocks by role and section, pick the period(s), compute a
measure, classify it against a threshold declared in the sheet-00 block GLOBAL, and
produce a table object with a title, rows, per-row status and an overall verdict.
Evidence of the duplication: threshold_of() appears three times almost verbatim, and
_incompatible() twice.

Create datatransform/crosscheck/:
- base.py       the shared shape: a Check protocol, a Finding, the four outcome words
                from section 10.1 (passed / failed / not applicable / skipped),
                threshold resolution from GLOBAL via constants.read_share, section
                grouping, and the scale-compatibility guard.
- rules.py      today's crosschecks.py — the rules declared in sheet 00.
- loss_share.py, growth.py, rate.py — one declaration each.

Do NOT merge the measures themselves. The three checks differ in ways that are
deliberate:
- 10.3 sums several roles per year and holds them against Incurred Losses in 01;
- 10.4 compares two versions of the SAME measure across periods;
- 10.5 combines sections, where premium adds but exposure does NOT — the growth is an
  exposure-weighted mean. Collapsing that would silently re-introduce the double count
  we explicitly excluded.

Keep every public name the writer and runner import today, or update both call sites.
Keep the reasoning in the docstrings — it is the record of why each check exists.

Verify with golden.py (empty diff, logs included) and pytest (452 passed).
```

---

### Etappe 3 · Blatt-00-Leser deklarativ

**Das Problem.** `nomenclature.py` hat zehn `_read_*`-Methoden — `_read_axes`,
`_read_splits`, `_read_zones`, `_read_globals`, `_read_types`, `_read_sections`,
`_read_period_order`, `_read_rules`, `_read_vocabulary` plus das dynamische Register.
Jede macht dasselbe: Anker finden, *n* Spalten lesen, Zeilen in Objekte übersetzen.

**Das Ziel.** Eine Tabelle von Blockbeschreibungen:

```python
BLOCKS = (
    BlockSpec("⟦AXES⟧",     columns=("dataset", "axis", "categories"), build=Axis),
    BlockSpec("⟦SPLITS⟧",   columns=(...),                             build=SplitRule),
    ...
)
```

und **eine** Leseschleife.

**Der eigentliche Gewinn** liegt hinter der Zeilenersparnis: ein neuer `⟦…⟧`-Block ist
danach *eine Zeile*. Damit wird der Vorschlag aus der Spalte-A-Diskussion — Attribute
und Annahmen aus Spalte A in einen eigenen, lesbaren Notizblock zu verlagern — von
teuer zu billig.

**Fallstricke.** `⟦REGISTER⟧` ist dynamisch (die Attributspalten stehen hinter der
breitesten Kopfzeilenliste) und passt nicht in dieselbe Form — es bleibt ein Sonderfall
und soll einer bleiben. `_read_rules` hat eine eigene Referenzsprache (`parse_reference`).

**Abnahme.** Zellgleich. Zusätzlich: eine Mappe mit **fehlendem** Block muss weiterhin
dieselbe Fehlermeldung erzeugen — Blatt 00 ist die Stelle, an der ein Lauf laut
scheitern soll.

#### Copilot-Prompt

```
Make the sheet-00 readers in datatransform/nomenclature.py declarative.

Read first: datatransform/nomenclature.py in full and specification_v1.md section 3.

There are ten _read_* helpers. Each finds an anchor token in column A, reads a fixed
number of columns, and turns rows into objects. Replace them with a table of block
descriptions plus one reading loop:

    BlockSpec(anchor, columns, build, skip_label_row=True, required=False)

Two blocks must stay special and should stay obviously special:
- the REGISTER block, whose attribute columns start after the widest header list, so
  its width is not fixed;
- the RULES block, which carries its own reference grammar (parse_reference).

Requirements:
1. Adding a new sheet-00 block must become a one-line change. State in the module
   docstring that this is the point.
2. A missing required block must fail exactly as loudly as today, with the same
   message. Sheet 00 is where a run is meant to stop, not to guess.
3. No behaviour change; every reference workbook must read identically.

Verify with golden.py (empty diff) and pytest (452 passed).
```

---

### Etappe 4 · `model.py` aufteilen

**Das Problem.** `model.py` (395) enthält zwei Dinge, die nichts miteinander zu tun
haben:

- das **Datenmodell** der gelesenen Daten: `Block`, `Record`, `Dataset`, `Confidence`,
  `Hypothesis`, `Orientation`, `FieldType`;
- die **Sprache von Blatt 00**: `Axis`, `SplitRule`, `Section`, `Rule`, `Reference`,
  `RuleResult`, `parse_reference`.

Wer wissen will, was ein `Block` ist, liest dabei die Regelgrammatik mit.

**Das Ziel.** `model.py` (Daten) und `declarations.py` (Blatt-00-Sprache).
Reine Umzugs- und Importarbeit, kein Risiko. Gute Etappe für einen kurzen Abend.

---

### Etappe 5 · `datasets/` — eine Datei pro Blatt

**Das ist das eigentliche Ziel des ganzen Plans**, und es ist nach 1–4 fast geschenkt:
alle drei Bestandteile eines Blattes existieren dann bereits als Deklaration.

Aus `specs.py` (Pipeline), aus `render/` (Layout) und aus `crosscheck/` (Prüfungen)
wird je Datensatz **eine** Datei:

```python
# datatransform/datasets/_06_eq_aggs.py
"""06 · Earthquake aggregates — spec §2.5, §2.6."""

KEY  = "06 EQ Aggs"
ROLE = "06"

STEP2 = Pipeline(
    Split(total=F_TOTAL),
    Identity(total=F_TOTAL),
    Complete(key=F_ZONE, catalogue_attribute="Zone scheme"),
    SortBy((F_ZONE,)),
    Cumulative("Cumulative exposure %", F_TOTAL),
)

LAYOUT = Report(Anchor(...), Title(...), Table(...), ControlRow(...), TieBack(...))

CHECKS = (ExposureGrowth(), )
```

Plus eine Registry in `datasets/__init__.py`, die `step2_for()` ersetzt (exakter
Schlüssel zuerst, dann die Rolle — §2.3, unverändert).

**Abnahme.** Zellgleich; `specs.py` verschwindet, `step2_for` behält sein Verhalten
für mehrsektionale Packs (`01 History Fire`, `01 History EQ` → Rolle `01`).

---

### Etappe 6 · `extract.py` — eine Geometrie

**Die heikelste Etappe.** Hier saßen beide Fehler, die uns transponierte Blätter
gekostet haben.

**Das Problem.** `_select_rows` (66 Zeilen) und `_select_columns` (59) sind
Spiegelbilder; `_block_boundary` und `_label_band` ebenso. Jede Änderung muss zweimal
gemacht werden, und wenn sie nur einmal gemacht wird, merkt es niemand — außer man
baut wie damals die vollständig transponierte Referenzmappe.

**Das Ziel.** Eine Geometrie: ein Block hat eine **Hauptachse** (entlang der die
Datensätze laufen) und eine **Nebenachse** (entlang der die Felder liegen). Zeilenweise
ist Haupt = Zeile, transponiert ist Haupt = Spalte. Alles andere ist identisch. Damit
gibt es eine Auswahlroutine statt zweier, und die Orientierung ist ein Paar von
Zugriffsfunktionen.

**Vorbedingung.** Vor der ersten Änderung: einen Test, der für **jeden** der neun Packs
die zeilenweise gegen die transponierte Fassung datensatzweise vergleicht. Die Mappe
`Intake_FireCatFull_Transposed_v1.xlsx` ist dafür schon da — sie wurde mechanisch
gekippt, also gehört jeder Unterschied dem Werkzeug.

**Was unangetastet bleibt.** Die `=ROW()`/`=COLUMN()`-Selbstvalidierung, die
Zwei-Pass-Lesung (Werte und Formeln getrennt) und die Regel, dass eine unberechnete
Formel den Lauf anhält. Das sind die drei Zusagen, auf denen die Prüfbarkeit steht.

---

### Etappe 7 · `runner.py` — Ablauf ≠ Protokoll

Fünf `_log_*`-Funktionen (~150 Zeilen Formatierung) stecken mitten im Ablauf. Sie
gehören neben den Writer, nicht in die Orchestrierung. Danach liest sich `runner.py`
als das, was es ist:

```
Vorprüfung → für jedes Blatt: extrahieren, transformieren, schreiben, prüfen → Bericht
```

**Abnahme.** Die Logdateien müssen **zeichengleich** bleiben — `golden.py` nimmt sie
deshalb von Anfang an mit auf.

---

### Etappe 8 · `tools/build_*.py` — Mappen als Daten

**2.441 Zeilen** in `build_intake.py` (1.338), `build_v1.py` (699) und
`build_mexico.py` (404). Die komplexeste Datei des Projekts ist ausgerechnet die, die
niemand verstehen muss: sie baut nur die Testmappen.

**Das Ziel.** Jeder Pack wird eine Deklaration über die gemeinsamen Bausteine in
`sheets.py`: welche Sektionen, welche Datensätze, welche Zonen, welche Absicht. Die
sechs Packs in `build_intake.py` unterscheiden sich in wenigen Zeilen Inhalt und
wiederholen hunderte Zeilen Mechanik.

**Abnahme (anders als sonst).** Nicht die Ausgabe, sondern die **Eingabe** muss gleich
bleiben: die neun `.xlsx` müssen nach dem Neubau denselben Inhalt haben. Byte-Vergleich
geht nicht (Zip-Zeitstempel), also vergleicht man sie mit demselben Zellenauszug, den
`golden.py` schon kann — dann ist auch die Ausgabe automatisch gleich.

---

### Etappe 9 · Tests pro Blatt, und die Dokumentation

**Tests.** Heute nach Themen geschnitten (`test_pipeline`, `test_aggregates`,
`test_variants` …). Eine Änderung an 06 zeigt sich in vier Dateien. Ziel: pro Datensatz
eine Datei, die dem Blatt folgt, plus je eine pro Schicht (`test_extract`,
`test_render`, `test_crosscheck`). 3.600 Zeilen umzusortieren ist viel Arbeit und null
Risiko — gute Etappe, wenn wenig Zeit ist.

**Dokumentation.** Zum Schluss, nicht vorher — sonst schreiben wir sie zweimal.

- `README.md`: fünf veraltete Aussagen, die schlimmste ist *„Datasets 00–05 are
  implemented. 06–10 are not yet specified"* (Zeile 169). Ebenfalls falsch: die
  Rollenliste in der Sections-Tabelle, `Intake_FireEQWind_v1.xlsx` als „Hurricane"
  (der Pack deklariert *Windstorm*) und `Intake_v1.xlsx` als „00–05" (er hat kein 04).
- `specification_v1.md`: §14.2 auf die neue Verzeichnisstruktur ziehen, §1974
  (Modulliste) nachführen.
- `datatransform.egg-info/` aus der Versionsverwaltung nehmen — sechs Dateien liegen
  fälschlich auf GitHub.

---

## 4 · Was wir *nicht* anfassen

Damit der Plan eine Grenze hat:

| | warum |
|---|---|
| Die Marker-Grammatik in Spalte A | Sie ist der Grund, warum nichts von einer Zellposition abhängt. Optik ist kein Grund, Prüfbarkeit aufzugeben. |
| `=ROW()`/`=COLUMN()`-Selbstvalidierung | Der einzige Schutz gegen eine nachträglich manipulierte Mappe. |
| Die Zwei-Pass-Lesung | Ohne sie liest man irgendwann eine unberechnete Formel als Zahl. |
| Der Formelwert-Cache | Ohne ihn braucht der Leser Excel. Das war eine Zusage. |
| „Nie raten" | Mehrdeutige Dezimaltrennung und mehrdeutige Datumsangaben bleiben tödlich. |
| Die Quellmappe | Wird nie verändert. |

Offene fachliche Punkte (Datensatz `10 Triangles`, `11 Exchange rates`, `20 Summary`,
die geparkten Punkte aus §13.1) sind **kein** Teil dieses Plans. Neue Funktion und
Refaktorierung nie im selben Commit.

---

## 5 · Der Rhythmus je Etappe

```
1  git checkout claude/excel-metadata-filtering-30g0ue && git pull
2  python tools/golden.py --out golden/before
3  … die Etappe umsetzen (Copilot-Prompt oben) …
4  python -m pytest -q                       →  452 passed
5  python tools/golden.py --out golden/after
6  diff -r golden/before golden/after        →  leer
7  git commit  (eine Etappe = ein Commit, Betreff nennt die Etappe)
8  git push -u origin claude/excel-metadata-filtering-30g0ue
```

Schritt 6 ist der, der zählt. Bei der Pipeline-Umstellung war er es, der zwei Fehler
gefunden hat, die Schritt 4 durchgelassen hatte.

Unter Windows steht in Schritt 6 statt `diff -r`:

```
fc /L golden\before\<datei> golden\after\<datei>
```

oder in der PowerShell für alle auf einmal:

```powershell
Compare-Object (Get-ChildItem -Recurse golden\before) (Get-ChildItem -Recurse golden\after) -Property Name, Length
```

— besser ist, `golden.py` gibt am Ende selbst einen Vergleich aus; das ist in seinem
Prompt oben nicht gefordert und wäre eine sinnvolle kleine Erweiterung, sobald Etappe 0
steht.

---

## 6 · Erwartetes Ergebnis

| | vorher | nachher |
|---|---|---|
| Zeilen gesamt | 14.282 | ~10.500 |
| davon Werkzeuge/Testmappen | 3.140 | ~900 |
| größte Datei | 1.338 | ~400 |
| Dateien, die man für *ein* Blatt lesen muss | 6–8 | **1** |

Die Zeilenzahl ist dabei das unwichtigste Maß. Das Ziel der letzten Zeile ist das
eigentliche: eine Änderung an einem Blatt soll an genau einer Stelle stattfinden.
