"""Plain data containers shared by the readers, the comparison engine and the writers."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any

Box = tuple[float, float, float]
"""Carton size in millimetres, in the axis order the source wrote it (L, W, H)."""

Number = int | float


class Status(StrEnum):
    """Row outcome. The same four values drive both reports and their colors."""

    OK = "OK"  # green: every source agrees, nothing to change in the ERP
    UPDATE = "UPDATE"  # amber: correct value is known, an ERP field is stale
    REVIEW = "REVIEW"  # red: no agreement, missing data, or vendor ignored instructions
    NOT_SHIPPED = "NOT SHIPPED"  # grey: SKU is not on this packing list


@dataclass
class ErpProduct:
    """One row of the ERP product export: what the system *thinks* the packing is."""

    sku: str
    name: str | None
    fba_case_qty: Number | None  # how WE pack for the marketplace; reference only
    qty_per_case: Number | None  # how the VENDOR should pack; compared
    supplier_case_qty: Any  # deprecated custom column, carried through verbatim
    dims_text: Any  # free-text carton size, mm or cm
    case_inches: tuple[Number | None, Number | None, Number | None]
    row: int


@dataclass
class InstructionEntry:
    """What one packing-instructions file asked for, for one internal SKU."""

    bags_per_carton: Number | None
    box_mm: Box | None
    row: int


@dataclass
class InstructionsFile:
    """A parsed packing-instructions workbook (one per purchase order)."""

    path: Any
    label: str
    sheet: str
    header_row: int
    entries: dict[str, InstructionEntry]
    internal_sku_added: bool
    row_skus: list[tuple[int, str | None]]
    warnings: list[str] = field(default_factory=list)


@dataclass
class MergedInstruction:
    """Instructions for one SKU after combining several POs, with provenance per field."""

    bags_per_carton: Number | None = None
    bags_source: str | None = None
    box_mm: Box | None = None
    box_source: str | None = None
    sources: list[str] = field(default_factory=list)


@dataclass
class PackingLine:
    """One physical row of the vendor packing list."""

    row: int
    parent: str | None
    pcs_per_bag: Number | None
    bags_per_carton: Number | None
    cartons: Number | None
    quantity: Number | None
    box_mm: Box | None
    marketplace_sku: str | None = None
    internal_sku: str | None = None


@dataclass
class ShippedSku:
    """All packing-list rows that resolved to the same internal SKU."""

    sku: str
    case_qty: Number | None
    case_qty_values: list[Number]
    box_mm: Box | None
    cartons: Number
    quantity: Number
    parent: str | None
    rows: list[int]
    case_conflict: str | None = None
    box_conflict: str | None = None
    alerts: list[str] = field(default_factory=list)


@dataclass
class CaseQtyRow:
    sku: str
    name: str | None
    pack_size: Number | None
    erp_qty_per_case: Number | None
    supplier_case_qty: Any
    instructions: Number | None
    packing_list: Number | None
    correct: Number | None
    checks: dict[str, Number | None]
    fba_case_qty: Number | None
    instructions_source: str | None
    vendor_vs_instructions: str | None
    status: Status
    note: str
    pl_rows: list[int]


@dataclass
class DimensionRow:
    sku: str
    name: str | None
    pack_size: Number | None
    fba_case_qty: Number | None
    erp_text: Any
    erp_text_mm: Box | None
    erp_inches: tuple[Number | None, Number | None, Number | None]
    instructions_mm: Box | None
    instructions_source: str | None
    packing_list_mm: Box | None
    correct_mm: Box | None
    text_vs_pl: str | None
    inches_vs_pl: str | None
    instructions_vs_pl: str | None
    status: Status
    note: str
    pl_rows: list[int]
