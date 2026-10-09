"""Carton dimensions: parsing, unit resolution and box comparison.

Everything is normalized to millimetres and converted to inches only for display.
Millimetres are the hard data (vendors measure in mm); inches are a rounded derivative.

Unit labels in the source files are not trusted, because they are wrong in practice:

* The ERP free-text field mixes units inside one export: "460 x 310 x 310, Confirmed"
  (mm), "45 x 32 x 34 Accurate" (cm), "920 x 310 x 160 - Confirmed 08/07/2026" (mm).
* Vendor packing lists and our own instruction files title columns "LENGTH (cm)" while
  the values are millimetres.

So the unit is always *verified* against an independent signal: the ERP's numeric
inch fields, the inch columns of the same file, or the vendor's declared volume (CBM).
Only when no signal exists does magnitude decide (no real carton is under 150 mm long).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from .models import Box, Number
from .sheet_utils import normalize_header

MM_PER_INCH = 25.4

TOLERANCE_MM = 5.0
"""Max per-side difference still considered the same carton (vendor rounding)."""

_MIN_REAL_CARTON_MM = 150
_NUMBER_RE = re.compile(r"\d+(?:[.,]\d+)?")
_DATE_RE = re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}")
_PO_RE = re.compile(r"(?i)\bpo\s*#?\s*\d+")

_UNIT_FACTORS = {"mm": 1.0, "cm": 10.0}


class Verdict(StrEnum):
    EXACT = "match"
    TOLERANCE = "match (within tolerance)"
    ROTATED = "same box, sides in another order"
    DIFFERENT = "NO match"

    @property
    def is_match(self) -> bool:
        return self in (Verdict.EXACT, Verdict.TOLERANCE)


@dataclass(frozen=True)
class ParsedDims:
    box_mm: Box | None
    unit: str | None
    note: str | None = None


def extract_numbers(text: Any) -> list[float]:
    """Numbers in a free-text size, after removing dates and PO references.

    "Confirmed 08/07/2026" or "PO 13" would otherwise be read as carton sides.
    """
    if text is None or isinstance(text, (int, float)):
        return []
    raw = _PO_RE.sub(" ", _DATE_RE.sub(" ", str(text)))
    return [float(m.group().replace(",", ".")) for m in _NUMBER_RE.finditer(raw)]


def _relative_error_vs_inches(box_mm: list[float], inches: tuple) -> float | None:
    usable = [
        (mm, inch)
        for mm, inch in zip(box_mm, inches, strict=True)
        if isinstance(inch, (int, float)) and inch > 0
    ]
    if not usable:
        return None
    errors = [abs(mm - inch * MM_PER_INCH) / (inch * MM_PER_INCH) for mm, inch in usable]
    return sum(errors) / len(errors)


def resolve_metric_unit(
    nums: list[float],
    inches: tuple = (None, None, None),
    cartons: Number | None = None,
    declared_cbm: Number | None = None,
) -> tuple[Box, str]:
    """Decide whether three numbers are mm or cm; return (box in mm, unit).

    Signals, strongest first: inches for the same box, then the vendor's declared
    volume, then magnitude.
    """
    candidates = {unit: [n * f for n in nums[:3]] for unit, f in _UNIT_FACTORS.items()}

    errors = {u: _relative_error_vs_inches(box, inches) for u, box in candidates.items()}
    if all(e is not None for e in errors.values()) and min(errors.values()) < 0.10:
        unit = min(errors, key=errors.get)
        return tuple(candidates[unit]), unit

    if declared_cbm and declared_cbm > 0:
        unit = _unit_from_volume(candidates, cartons, float(declared_cbm))
        if unit:
            return tuple(candidates[unit]), unit

    unit = "cm" if max(nums[:3]) < _MIN_REAL_CARTON_MM else "mm"
    return tuple(candidates[unit]), unit


def _unit_from_volume(candidates: dict, cartons: Number | None, cbm: float) -> str | None:
    """Unit whose carton volume reproduces the declared CBM (total or per carton)."""
    best_unit, best_err = None, math.inf
    for unit, box in candidates.items():
        per_carton = math.prod(box) / 1e9
        targets = [per_carton]
        if isinstance(cartons, (int, float)) and cartons > 0:
            targets.append(per_carton * cartons)
        err = min(abs(t - cbm) / cbm for t in targets)
        if err < best_err:
            best_unit, best_err = unit, err
    return best_unit if best_err < 0.15 else None


def parse_erp_dims(text: Any, case_inches: tuple = (None, None, None)) -> ParsedDims:
    """Interpret the ERP free-text carton size. Never invents a missing side."""
    nums = extract_numbers(text)
    if not nums:
        return ParsedDims(None, None, "ERP has no carton size text")
    if len(nums) < 3:
        return ParsedDims(
            None, None, f"ERP carton size text has only {len(nums)} side(s) ('{text}')"
        )
    box, unit = resolve_metric_unit(nums, case_inches)
    return ParsedDims(box, unit)


def inches_to_box(inches: tuple) -> Box | None:
    """ERP numeric inch fields -> mm box, or None when any side is blank or 0.

    A 0 is how the ERP stores "not filled in", not a 0 mm side; comparing against it
    would create false mismatches.
    """
    if not all(isinstance(v, (int, float)) and v > 0 for v in inches):
        return None
    return tuple(v * MM_PER_INCH for v in inches)


def mm_to_inches(mm: float) -> float:
    return round(mm / MM_PER_INCH, 2)


def fmt_box(box: Box | None) -> str | None:
    if box is None:
        return None

    def one(v: float) -> str:
        return str(int(round(v))) if abs(v - round(v)) < 0.05 else f"{v:.1f}"

    return "*".join(one(v) for v in box)


def compare_boxes(a: Box, b: Box, tolerance: float = TOLERANCE_MM) -> tuple[Verdict, str | None]:
    """Compare two cartons in mm; `b` is the packing list. Returns (verdict, note).

    Side order does not make it a different carton (460x310x460 == 310x460x460), but it
    gets its own verdict: the ERP should store axes in the vendor's order.
    """
    diffs = [abs(x - y) for x, y in zip(a, b, strict=True)]
    if max(diffs) < 0.05:
        return Verdict.EXACT, None
    if max(diffs) <= tolerance:
        return Verdict.TOLERANCE, (
            f"within tolerance, up to {max(diffs):.0f} mm apart ({fmt_box(a)} vs {fmt_box(b)})"
        )
    sorted_diffs = [abs(x - y) for x, y in zip(sorted(a), sorted(b), strict=True)]
    if max(sorted_diffs) <= tolerance:
        return Verdict.ROTATED, (
            f"same carton, sides in another order ({fmt_box(a)} vs {fmt_box(b)})"
        )
    return Verdict.DIFFERENT, (
        f"different carton: {fmt_box(a)} mm vs packing list {fmt_box(b)} mm "
        f"(largest side difference {max(diffs):.0f} mm)"
    )


# --------------------------------------------------------------------------- headers

_UNIT_SUFFIXES = (("inches", "in"), ("inch", "in"), ("mm", "mm"), ("cm", "cm"), ("in", "in"))
_AXES = {"l": "L", "length": "L", "long": "L", "w": "W", "width": "W",
         "h": "H", "height": "H", "hight": "H"}


def _axis_and_unit(header: Any) -> tuple[str | None, str | None]:
    h = normalize_header(header)
    unit = None
    for suffix, canonical in _UNIT_SUFFIXES:
        if h.endswith(suffix) and len(h) > len(suffix):
            unit, h = canonical, h[: -len(suffix)]
            break
    return _AXES.get(h), unit


@dataclass(frozen=True)
class LwhColumns:
    cols: tuple[int, int, int]
    unit: str | None  # unit the HEADER claims; verified before use


def find_lwh_columns(headers: list[Any]) -> list[LwhColumns]:
    """Consecutive Length/Width/Height header triples, e.g. 'L (Inch)'..'HEIGHT (cm)'."""
    tagged = [
        (idx, axis, unit)
        for idx, header in enumerate(headers, start=1)
        for axis, unit in [_axis_and_unit(header)]
        if axis
    ]
    triples, i = [], 0
    while i <= len(tagged) - 3:
        a, b, c = tagged[i : i + 3]
        if (a[1], b[1], c[1]) == ("L", "W", "H"):
            triples.append(LwhColumns((a[0], b[0], c[0]), a[2] or b[2] or c[2]))
            i += 3
        else:
            i += 1
    return triples


def split_metric_and_inch(headers: list[Any]) -> tuple[LwhColumns | None, LwhColumns | None]:
    triples = find_lwh_columns(headers)
    metric = next((t for t in triples if t.unit != "in"), None)
    inch = next((t for t in triples if t.unit == "in"), None)
    return metric, inch
