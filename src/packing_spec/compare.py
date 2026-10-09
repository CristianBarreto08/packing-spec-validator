"""The two comparisons. Kept separate because they answer different questions.

**Case qty - how many bags per carton?** Four sources of the same number:

    D  ERP QtyPerCase
    E  deprecated custom column (reads "Ignore" today)
    F  packing instructions we sent
    G  vendor packing list

The packing list is the reference: it is what physically arrived. A value is *correct*
when G agrees with any other source, checked in this order (the order of the original
spreadsheet template the buyers already used):

    J  all four equal     -> D
    K  E = F = G          -> E
    L  E = G              -> E
    M  F = G              -> F   (instructions)
    N  D = G              -> D   (ERP)
    I  "Correct case qty" = first non-empty of J..N

Text never equals a number, so while E says "Ignore", J/K/L can never fire and M or N
decide. That is deliberate: the column is shown untouched but cannot influence results
until it is removed from the ERP.

**Dimensions - is it the same carton?** Three sources against the packing list: the
ERP free-text size, the ERP inch fields, and the carton we asked for. One agreeing
source is enough to know the right value; any ERP field that disagrees is stale.

ERP fields and instructions are fixed differently, so they are tracked separately: a
stale ERP field is an UPDATE (amber, load the known value); a vendor that did not pack
as instructed is a REVIEW (red) even when the ERP already matches the vendor, because
accepting their packing or pushing back is a decision, not a data-entry task.
"""

from __future__ import annotations

from typing import Any

from .catalog import SkuCatalog
from .dims import TOLERANCE_MM, Verdict, compare_boxes, fmt_box, inches_to_box, parse_erp_dims
from .models import (
    CaseQtyRow,
    DimensionRow,
    ErpProduct,
    MergedInstruction,
    Number,
    ShippedSku,
    Status,
)

CHECK_ORDER = ("J", "K", "L", "M", "N")


def _eq(a: Any, b: Any) -> bool:
    """Numeric equality only: text such as 'Ignore' never equals anything."""
    if isinstance(a, bool) or isinstance(b, bool):
        return False
    if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
        return False
    return abs(a - b) < 1e-9


def resolve_case_qty(d: Any, e: Any, f: Any, g: Any) -> tuple[Number | None, dict]:
    """Return (correct value, {J..N}) following the template's precedence."""
    checks = {
        "J": d if _eq(d, e) and _eq(d, f) and _eq(d, g) else None,
        "K": e if _eq(e, f) and _eq(e, g) else None,
        "L": e if _eq(e, g) else None,
        "M": f if _eq(f, g) else None,
        "N": d if _eq(d, g) else None,
    }
    correct = next((checks[k] for k in CHECK_ORDER if checks[k] is not None), None)
    return correct, checks


_WHY_OK = {
    "J": "all four sources agree",
    "K": "custom column, instructions and packing list agree",
    "L": "custom column agrees with the packing list",
    "M": "packing list matches the instructions we sent",
    "N": "packing list matches the ERP QtyPerCase",
}


def _fmt(v: Any) -> str:
    if v is None:
        return "(blank)"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def build_case_qty_rows(
    products: list[ErpProduct],
    shipped: dict[str, ShippedSku],
    instructions: dict[str, MergedInstruction],
    catalog: SkuCatalog,
) -> list[CaseQtyRow]:
    rows = []
    for p in products:
        pl = shipped.get(p.sku)
        instr = instructions.get(p.sku) or MergedInstruction()
        d, e, f = p.qty_per_case, p.supplier_case_qty, instr.bags_per_carton
        g = pl.case_qty if pl else None
        correct, checks = resolve_case_qty(d, e, f, g)
        notes: list[str] = []

        if pl is None:
            status = Status.NOT_SHIPPED
            notes.append("not on this packing list")
        elif g is None:
            status = Status.REVIEW
            notes.append("vendor did not state bags per carton on the packing list")
        elif correct is not None:
            why = _WHY_OK[next(k for k in CHECK_ORDER if checks[k] is not None)]
            if _eq(d, g):
                status = Status.OK
                notes.append(why)
            else:
                status = Status.UPDATE
                notes.append(f"{why}, but ERP QtyPerCase is stale: it says {_fmt(d)}, "
                             f"correct value is {_fmt(g)}")
        else:
            status = Status.REVIEW
            parts = [f"packing list={_fmt(g)}", f"instructions={_fmt(f)}", f"ERP={_fmt(d)}"]
            if isinstance(e, str) and e.strip():
                parts.append(f"custom column='{e.strip()}'")
            notes.append("matches no source: " + ", ".join(parts))

        vendor_vs_instr = None
        if f is not None and g is not None:
            if _eq(f, g):
                vendor_vs_instr = "match"
            else:
                vendor_vs_instr = f"NO match (asked {_fmt(f)}, shipped {_fmt(g)})"
                status = Status.REVIEW
                notes.append(f"vendor did NOT pack as instructed [{instr.bags_source}]: asked "
                             f"for {_fmt(f)} bags/carton, shipped {_fmt(g)}")
        elif g is not None:
            vendor_vs_instr = "(no instructions for this SKU)"

        if pl and pl.case_conflict:
            notes.append(pl.case_conflict)
            status, correct = Status.REVIEW, None
        if pl:
            notes += [f"packing list check: {a}" for a in pl.alerts]

        rows.append(CaseQtyRow(
            sku=p.sku,
            name=p.name or catalog.names.get(p.sku),
            pack_size=catalog.pack_sizes.get(p.sku),
            erp_qty_per_case=d,
            supplier_case_qty=e,
            instructions=f,
            packing_list=g,
            correct=correct,
            checks=checks,
            fba_case_qty=p.fba_case_qty,
            instructions_source=instr.bags_source,
            vendor_vs_instructions=vendor_vs_instr,
            status=status,
            note=" | ".join(notes),
            pl_rows=pl.rows if pl else [],
        ))
    return rows


_FIELD = {"text": "SUPPLIERCASEDIMENSIONS", "inches": "CaseLength/Width/Height",
          "instructions": "the packing instructions"}


def build_dimension_rows(
    products: list[ErpProduct],
    shipped: dict[str, ShippedSku],
    instructions: dict[str, MergedInstruction],
    catalog: SkuCatalog,
    tolerance: float = TOLERANCE_MM,
) -> list[DimensionRow]:
    rows = []
    for p in products:
        pl = shipped.get(p.sku)
        instr = instructions.get(p.sku) or MergedInstruction()
        parsed = parse_erp_dims(p.dims_text, p.case_inches)
        inches_box = inches_to_box(p.case_inches)
        pl_box = pl.box_mm if pl else None
        notes: list[str] = []
        if parsed.note and pl is not None:
            notes.append(parsed.note)
        if parsed.unit == "cm":
            notes.append(f"ERP size text was in CENTIMETRES ('{str(p.dims_text).strip()}'); "
                         f"converted to mm")
        if inches_box is None and pl is not None:
            notes.append("ERP CaseLength/Width/Height incomplete (a side is blank or 0)")

        sources = {"text": parsed.box_mm, "inches": inches_box, "instructions": instr.box_mm}
        verdicts: dict[str, Verdict | None] = {}
        for label, box in sources.items():
            if box is None or pl_box is None:
                verdicts[label] = None
                continue
            verdicts[label], note = compare_boxes(box, pl_box, tolerance)
            if note:
                notes.append(f"[{label}] {note}")

        erp_available = [k for k in ("text", "inches") if sources[k] is not None]
        erp_matched = [k for k in erp_available if verdicts[k] and verdicts[k].is_match]
        erp_stale = [k for k in erp_available if k not in erp_matched]
        instr_ok = bool(verdicts["instructions"] and verdicts["instructions"].is_match)
        matched_any = erp_matched + (["instructions"] if instr_ok else [])

        if pl is None:
            status, correct = Status.NOT_SHIPPED, None
            notes.append("not on this packing list")
        elif pl_box is None:
            status, correct = Status.REVIEW, None
            notes.append("vendor did not state the carton size on the packing list")
        elif not matched_any:
            status, correct = Status.REVIEW, None
            if erp_available:
                notes.append("nothing matches the packing list (neither ERP nor instructions); "
                             "confirm with the vendor whether the carton changed")
            else:
                notes.append(f"ERP has no usable carton size; vendor says {fmt_box(pl_box)} mm "
                             f"- confirm with the vendor, then load it")
        elif erp_stale or not erp_available:
            status, correct = Status.UPDATE, pl_box
            stale = erp_stale or ["text", "inches"]
            confirmed = {"text": "the ERP size text", "inches": "the ERP inch fields",
                         "instructions": "the carton we asked for"}[matched_any[0]]
            notes.append(f"carton confirmed by {confirmed}; UPDATE "
                         f"{' and '.join(_FIELD[k] for k in stale)} to {fmt_box(pl_box)} mm")
        else:
            status, correct = Status.OK, pl_box
            notes.append("both ERP fields match the packing list" if len(erp_available) == 2
                         else f"ERP {_FIELD[erp_available[0]]} matches the packing list "
                              f"(the other field is empty)")

        if instr.box_mm is not None and pl_box is not None and not instr_ok:
            status = Status.REVIEW
            notes.append(f"vendor did NOT use the carton we asked for [{instr.box_source}]: "
                         f"asked {fmt_box(instr.box_mm)} mm, shipped {fmt_box(pl_box)} mm")
        if pl and pl.box_conflict:
            notes.append(pl.box_conflict)
            status, correct = Status.REVIEW, None

        rows.append(DimensionRow(
            sku=p.sku,
            name=p.name or catalog.names.get(p.sku),
            pack_size=catalog.pack_sizes.get(p.sku),
            fba_case_qty=p.fba_case_qty,
            erp_text=p.dims_text,
            erp_text_mm=parsed.box_mm,
            erp_inches=p.case_inches,
            instructions_mm=instr.box_mm,
            instructions_source=instr.box_source,
            packing_list_mm=pl_box,
            correct_mm=correct,
            text_vs_pl=_label(verdicts["text"]),
            inches_vs_pl=_label(verdicts["inches"]),
            instructions_vs_pl=_label(verdicts["instructions"]),
            status=status,
            note=" | ".join(notes),
            pl_rows=pl.rows if pl else [],
        ))
    return rows


def _label(verdict: Verdict | None) -> str | None:
    return verdict.value if verdict else None
