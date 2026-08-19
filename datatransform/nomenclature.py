"""Sheet 00 — the nomenclature. Specification_v1.md §3."""

from __future__ import annotations

import re
import unicodedata

from .model import Axis, Dataset, ExtractionError, FieldType, Rule, Section, SplitRule

HEADER_COL_FIRST = 4     # D
HEADER_COL_LAST = 13     # M
ATTR_COL_FIRST = 14      # N
ATTR_COL_LAST = 40
SECTION_ROW = 4
FIRST_DATASET_ROW = 5

GLOBAL_ANCHOR = "⟦GLOBAL⟧"
TYPES_ANCHOR = "⟦TYPES⟧"
VOCAB_ANCHOR = "⟦VOCABULARY⟧"
RULES_ANCHOR = "⟦RULES⟧"
PERIOD_ANCHOR = "⟦PERIOD ORDER⟧"
SECTIONS_ANCHOR = "⟦SECTIONS⟧"
ZONES_ANCHOR = "⟦ZONES⟧"
AXES_ANCHOR = "⟦AXES⟧"
SPLITS_ANCHOR = "⟦SPLITS⟧"

ANY_DATASET = "*"


def norm(value) -> str:
    """Trim, collapse whitespace, normalise NBSP and quote style — spec §8 F2, §2.2."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value))
    s = s.replace(" ", " ").replace("„", '"').replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", s).strip()


def sheet_sort_key(name: str) -> tuple[str, str]:
    """Numeric prefix first, text second — spec §2.2."""
    n = norm(name)
    m = re.match(r"^(\d+)\s*[.\-]?\s*(.*)$", n)
    return (m.group(1).zfill(3), m.group(2).casefold()) if m else ("999", n.casefold())


def match_sheet(declared: str, available) -> str | None:
    """Numeric prefix first, text second — spec §2.2.

    The prefix is the stable identifier; the words after it are cosmetic, so
    ``10. Triangles`` resolves against a sheet actually spelled ``10. Triangels``.
    """
    exact = [s for s in available if norm(s).casefold() == norm(declared).casefold()]
    if exact:
        return exact[0]

    prefix, text = sheet_sort_key(declared)
    if prefix == "999":                       # no numeric prefix: nothing safe to fall back on
        return None
    candidates = [s for s in available if sheet_sort_key(s)[0] == prefix]
    if len(candidates) <= 1:
        return candidates[0] if candidates else None
    # Several sheets share the prefix — pick the closest text rather than guessing.
    return min(candidates, key=lambda s: _distance(sheet_sort_key(s)[1], text))


def _distance(a: str, b: str) -> int:
    """Levenshtein distance, used only to disambiguate a shared numeric prefix."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _find_anchor(ws, anchor: str) -> int | None:
    for row in range(1, ws.max_row + 1):
        if norm(ws.cell(row=row, column=2).value) == anchor:
            return row
    return None


def _read_block(ws, anchor: str, columns: int, skip_label_row: bool = True):
    """Rows of an ⟦ANCHOR⟧ block, stopping at a blank row or the next anchor."""
    start = _find_anchor(ws, anchor)
    if start is None:
        return []

    rows = []
    for row in range(start + (2 if skip_label_row else 1), ws.max_row + 1):
        values = [norm(ws.cell(row=row, column=2 + i).value) for i in range(columns)]
        if not values[0]:
            break
        if values[0].startswith("⟦"):
            break
        rows.append((row, values))
    return rows


class Nomenclature:
    """Sheet 00: the frame — datasets, globals, types, vocabulary and rules."""

    def __init__(self, datasets, vocabulary, globals_=None, types=None,
                 rules=None, period_order=None, sections=None, zones=None,
                 axes=None, splits=None):
        self.datasets = datasets
        self.vocabulary = vocabulary
        self.globals = globals_ or {}
        self.types = types or {}
        self.rules = rules or []
        self.period_order = period_order or {}
        self.sections = sections or []
        self.zones = zones or {}
        self.axes = axes or []
        self.splits = splits or []

    def axes_for(self, dataset_key: str) -> list[Axis]:
        """The dimensions this dataset is split along, in declared order — spec §2.6."""
        wanted = norm(dataset_key).casefold()
        return [a for a in self.axes if norm(a.dataset).casefold() == wanted]

    def buckets_for(self, dataset_key: str) -> tuple[str, ...]:
        """The product of the axes, named by joining the categories — spec §2.6.

        Occupancy × cover gives ``Res Building`` … ``Ind BI``; the occupancy axis alone
        gives ``Res`` · ``Com`` · ``Ind``. The names this produces are the field names in
        sheet 00, so the register and the arithmetic cannot drift apart.
        """
        names: tuple[str, ...] = ()
        for axis in self.axes_for(dataset_key):
            names = (axis.categories if not names
                     else tuple(f"{a} {b}" for a in names for b in axis.categories))
        return names

    def splits_for(self, dataset_key: str, axis: str, source_category: str = ""):
        """Declared ratios for one axis, most specific first — spec §2.6.

        A rule naming the dataset beats a rule naming ``*``, so a house convention can be
        stated once and overridden where a particular book differs.
        """
        wanted, want_from = norm(axis).casefold(), norm(source_category).casefold()
        matching = [r for r in self.splits
                    if r.applies_to(dataset_key)
                    and norm(r.axis).casefold() == wanted
                    and norm(r.source_category).casefold() == want_from]
        named = [r for r in matching if r.dataset != "*"]
        return named or matching

    def zones_for(self, scheme: str) -> list[str]:
        """Every zone of a scheme, in declared order — spec §2.5.

        A cat aggregate lists only the zones the cedent has exposure in. The zones with
        none are the ones worth seeing: absent reads as "no data", zero reads as "nothing
        there", and only one of those is true. So the full list is declared.
        """
        wanted = norm(scheme).casefold()
        for name, codes in self.zones.items():
            if norm(name).casefold() == wanted:
                return list(codes)
        return []

    @property
    def actual_year(self) -> int | None:
        """N — the expiring year. The renewal being underwritten is N+1."""
        raw = self.globals.get("Actual year")
        if raw is None:
            return None
        try:
            return int(float(str(raw).replace(",", "")))
        except ValueError:
            raise ExtractionError(
                f"sheet 00 ⟦GLOBAL⟧: 'Actual year' is {raw!r}, which is not a year"
            ) from None

    def type_of(self, field: str) -> FieldType:
        declared = self.types.get(norm(field).casefold())
        if declared is None:
            raise ExtractionError(
                f"field {field!r} has no datatype in sheet 00 ⟦TYPES⟧ — "
                "every declared field needs one"
            )
        return declared

    def field_types(self, headers) -> dict[str, FieldType]:
        return {h: self.type_of(h) for h in headers}

    @staticmethod
    def register_columns(ws) -> tuple[int, int, int]:
        """Where the headers end and the attributes begin — read from row 4, spec §3.1.

        Row 4 has always *said* where the boundary is (``Headers →`` above D,
        ``Attributes →`` above N); the code used to hard-code it, which capped every
        dataset at ten headers. ``06. EQ Aggs`` needs eleven — a zone and the nine
        occupancy × cover buckets and a total — so the label is now what decides, and
        widening the register is a matter of moving it.

        A sheet that declares neither label keeps the original layout, so workbooks
        written before this change still read correctly.
        """
        first, attrs = None, None
        for col in range(2, ATTR_COL_LAST + 1):
            text = norm(ws.cell(row=SECTION_ROW, column=col).value).casefold()
            if first is None and text.startswith("header"):
                first = col
            elif text.startswith("attribute"):
                attrs = col
                break
        first = first or HEADER_COL_FIRST
        attrs = attrs or ATTR_COL_FIRST
        return first, attrs - 1, attrs

    @classmethod
    def read(cls, ws) -> "Nomenclature":
        header_first, header_last, attr_first = cls.register_columns(ws)
        datasets: dict[str, Dataset] = {}
        for row in range(FIRST_DATASET_ROW, ws.max_row + 1):
            sheet_name = norm(ws.cell(row=row, column=2).value)
            key = norm(ws.cell(row=row, column=3).value)
            if sheet_name.startswith("⟦"):
                break                       # the register ends where the next block begins
            if not sheet_name or not key:
                continue
            declared = [
                h for h in (
                    norm(ws.cell(row=row, column=c).value)
                    for c in range(header_first, header_last + 1)
                ) if h
            ]
            headers, optional = [], []
            for label in declared:
                # "Number of Claims (optional)" — the name a human writes is the name
                # without the annotation; the annotation is metadata about the name.
                bare = re.sub(r"\s*\(optional\)\s*$", "", label, flags=re.I)
                headers.append(bare)
                if bare != label:
                    optional.append(bare)
            attributes = tuple(
                a for a in (
                    norm(ws.cell(row=row, column=c).value)
                    for c in range(attr_first, ATTR_COL_LAST + 1)
                ) if a
            )
            if not headers:
                continue
            datasets[sheet_name] = Dataset(sheet_name, key, tuple(headers),
                                           attributes, tuple(optional))

        return cls(
            datasets,
            cls._read_vocabulary(ws),
            cls._read_globals(ws),
            cls._read_types(ws),
            cls._read_rules(ws),
            cls._read_period_order(ws),
            cls._read_sections(ws),
            cls._read_zones(ws),
            cls._read_axes(ws),
            cls._read_splits(ws),
        )

    @staticmethod
    def _read_axes(ws) -> list[Axis]:
        """⟦AXES⟧ — the dimensions each dataset is split along — spec §2.6."""
        out = []
        for row, (dataset, name, categories) in _read_block(ws, AXES_ANCHOR, 3):
            listed = tuple(c.strip() for c in categories.replace(";", ",").split(",")
                           if c.strip())
            if not name or not listed:
                raise ExtractionError(
                    f"sheet 00 ⟦AXES⟧ row {row}: {dataset!r} declares an axis with no "
                    f"{'name' if not name else 'categories'}"
                )
            out.append(Axis(dataset, name, listed))
        return out

    @staticmethod
    def _read_splits(ws) -> list[SplitRule]:
        """⟦SPLITS⟧ — declared ratios, with where they came from — spec §2.6."""
        out = []
        for row, cells in _read_block(ws, SPLITS_ANCHOR, 6):
            dataset, axis, frm, category, share, source = cells
            if not category:
                continue
            try:
                value = float(str(share).replace("%", "").replace(",", "."))
            except (TypeError, ValueError):
                raise ExtractionError(
                    f"sheet 00 ⟦SPLITS⟧ row {row}: share {share!r} is not a number"
                )
            if str(share).strip().endswith("%") or value > 1:
                value /= 100.0
            if value < 0:
                raise ExtractionError(
                    f"sheet 00 ⟦SPLITS⟧ row {row}: share {share!r} is negative"
                )
            out.append(SplitRule(dataset or ANY_DATASET, axis, frm, category,
                                 value, source))
        return out

    @staticmethod
    def _read_zones(ws) -> dict[str, list[str]]:
        """⟦ZONES⟧ — the complete zone list of each scheme — spec §2.5."""
        out = {}
        for row, (scheme, codes) in _read_block(ws, ZONES_ANCHOR, 2):
            listed = [c.strip() for c in codes.replace(";", ",").split(",") if c.strip()]
            if not listed:
                raise ExtractionError(
                    f"sheet 00 ⟦ZONES⟧ row {row}: scheme {scheme!r} lists no zones"
                )
            out[scheme] = listed
        return out

    @staticmethod
    def _read_globals(ws) -> dict[str, str]:
        return {name: value for _, (name, value) in _read_block(ws, GLOBAL_ANCHOR, 2)}

    @staticmethod
    def _read_types(ws) -> dict[str, FieldType]:
        return {
            name.casefold(): FieldType.parse(kind)
            for _, (name, kind) in _read_block(ws, TYPES_ANCHOR, 2)
        }

    @staticmethod
    def _read_sections(ws) -> list[Section]:
        """⟦SECTIONS⟧ — what this treaty is made of — spec §2.3."""
        out = []
        for row, (name, kind, roles) in _read_block(ws, SECTIONS_ANCHOR, 3):
            if not kind:
                raise ExtractionError(
                    f"sheet 00 ⟦SECTIONS⟧ row {row}: section {name!r} declares no kind"
                )
            declared = tuple(
                part.strip().zfill(2) for part in roles.replace(";", ",").split(",")
                if part.strip()
            )
            out.append(Section(name, kind.strip().casefold(), declared))
        return out

    def sections_of_kind(self, kind: str) -> list[Section]:
        return [s for s in self.sections if s.kind == kind.strip().casefold()]

    def section_named(self, name: str) -> Section | None:
        wanted = norm(name).casefold()
        return next((s for s in self.sections if norm(s.name).casefold() == wanted), None)

    @staticmethod
    def _read_period_order(ws) -> dict[str, int]:
        """Suffix → rank, for sorting within one leading number — spec §9.2 S13."""
        order = {}
        for row, (rank, suffix) in _read_block(ws, PERIOD_ANCHOR, 2):
            if not suffix:
                continue
            try:
                order[suffix.casefold()] = int(float(rank))
            except ValueError:
                raise ExtractionError(
                    f"sheet 00 ⟦PERIOD ORDER⟧ row {row}: rank {rank!r} is not a number"
                ) from None
        return order

    @staticmethod
    def _read_rules(ws) -> list[Rule]:
        rules = []
        for row, values in _read_block(ws, RULES_ANCHOR, 8):
            rid, left, rel, right, tol, sev, scope, note = values
            if not (left and rel and right):
                raise ExtractionError(
                    f"sheet 00 ⟦RULES⟧ row {row}: incomplete rule {rid!r}"
                )
            try:
                tolerance = float(str(tol).replace(",", "")) if tol else 0.0
            except ValueError:
                raise ExtractionError(
                    f"sheet 00 ⟦RULES⟧ row {row}: tolerance {tol!r} is not a number"
                ) from None
            rules.append(Rule(rid, left, rel, right, tolerance,
                              sev or "error", note, scope))
        return rules

    @staticmethod
    def _read_vocabulary(ws) -> dict[str, list[str]]:
        return {
            name: [v.strip() for v in values.split("|") if v.strip()]
            for _, (name, values) in _read_block(ws, VOCAB_ANCHOR, 2)
        }

    def dataset_for(self, sheet_name: str) -> Dataset | None:
        target = match_sheet(sheet_name, list(self.datasets))
        return self.datasets.get(target) if target else None

    def validate_value(self, attribute: str, value) -> bool:
        """An attribute with no vocabulary entry is unconstrained — spec §3.3."""
        allowed = self.vocabulary.get(attribute)
        if not allowed:
            return True
        if attribute == "Currency":
            return bool(re.fullmatch(r"[A-Z]{3}", norm(value).upper()))
        return norm(value).casefold() in {a.casefold() for a in allowed}


def read_nomenclature(wb) -> Nomenclature:
    name = match_sheet("00. Nomenclature & Interdependencies", wb.sheetnames)
    if name is None:
        name = next((s for s in wb.sheetnames if sheet_sort_key(s)[0] == "000"), None)
    if name is None:
        raise ExtractionError("sheet 00 not found — the nomenclature is required")
    return Nomenclature.read(wb[name])
