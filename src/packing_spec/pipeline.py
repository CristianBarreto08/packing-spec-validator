"""End-to-end run: discover inputs, read, compare, write, summarize.

Nothing is fixed silently: anything that cannot be resolved is red in the workbook AND
listed as a warning on the console.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .catalog import load_catalog
from .compare import build_case_qty_rows, build_dimension_rows
from .dims import TOLERANCE_MM
from .discovery import DiscoveredInputs, classify_inputs, log_discovery
from .erp import load_erp_export
from .instructions import describe_merge, load_instructions, merge_instructions
from .models import CaseQtyRow, DimensionRow, MergedInstruction, ShippedSku, Status
from .packing_list import aggregate_by_sku, read_packing_list
from .po_tags import output_name, po_tag, source_label
from .report import (
    write_case_qty_report,
    write_dimension_report,
    write_instructions_with_internal_sku,
)

log = logging.getLogger(__name__)


@dataclass
class RunConfig:
    catalog: Path
    out_dir: Path
    input_dir: Path | None = None
    erp: Path | None = None
    packing_list: Path | None = None
    instructions: list[Path] = field(default_factory=list)  # explicit = priority order
    pl_sheet: str | None = None
    vendor: str | None = None
    po: str | None = None
    tolerance_mm: float = TOLERANCE_MM
    run_date: date = field(default_factory=date.today)


@dataclass
class RunResult:
    case_qty_path: Path
    dimensions_path: Path
    case_rows: list[CaseQtyRow]
    dimension_rows: list[DimensionRow]
    warnings: list[str]
    missing_instructions: bool
    instruction_order: list[str]
    extra_outputs: list[Path] = field(default_factory=list)


def _resolve_inputs(cfg: RunConfig) -> DiscoveredInputs:
    if cfg.erp:
        # Explicit files: the order given IS the priority; it is never re-sorted,
        # because whoever typed it knows which PO is the newest.
        found = DiscoveredInputs(erp=cfg.erp, packing_list=cfg.packing_list,
                                 instructions=list(cfg.instructions))
        log_discovery(found)
        return found
    if cfg.input_dir is None:
        raise ValueError("Pass either an input folder or explicit file paths (--erp ...).")
    return classify_inputs(cfg.input_dir)


def run(cfg: RunConfig) -> RunResult:
    log.info("\n== 1. Inputs ==")
    inputs = _resolve_inputs(cfg)
    cfg.out_dir.mkdir(parents=True, exist_ok=True)
    warnings: list[str] = []
    extra_outputs: list[Path] = []

    log.info("\n== 2. ERP export and SKU catalog ==")
    products, erp_warnings = load_erp_export(inputs.erp, inputs.sheets.get(inputs.erp))
    warnings += erp_warnings
    known = {p.sku for p in products}
    catalog = load_catalog(cfg.catalog)

    log.info("\n== 3. Vendor packing list ==")
    shipped: dict[str, ShippedSku] = {}
    if inputs.packing_list:
        sheet = cfg.pl_sheet or inputs.sheets.get(inputs.packing_list)
        pl = read_packing_list(inputs.packing_list, known, catalog, sheet)
        warnings += pl.warnings
        shipped = aggregate_by_sku(pl.lines)
        unknown = sorted(set(shipped) - known)
        if unknown:
            warnings.append(f"{len(unknown)} packing-list SKU(s) are not in the ERP export and "
                            f"are NOT in the reports: {', '.join(unknown)}. Add them to the "
                            f"export and re-run.")
    else:
        warnings.append("No vendor packing list: nothing to compare against, every row is "
                        "NOT SHIPPED.")

    log.info("\n== 4. Packing instructions ==")
    merged: dict[str, MergedInstruction] = {}
    loaded = []
    for path in inputs.instructions:
        info = load_instructions(path, source_label(path), known, catalog,
                                 inputs.sheets.get(path))
        if info is None:
            warnings.append(f"Could not read instructions file '{path.name}'; ignored.")
            continue
        warnings += info.warnings
        loaded.append(info)
        if info.internal_sku_added:
            tag = " ".join(p for p in (cfg.vendor, info.label) if p)
            extra_outputs.append(write_instructions_with_internal_sku(info, cfg.out_dir, tag))
    if loaded:
        merged, stats = merge_instructions(loaded)
        for line in describe_merge(loaded, stats, merged):
            log.info("  %s", line)
    else:
        log.info("  no instructions file: 'Instructions we sent' stays empty")

    log.info("\n== 5. Compare and write ==")
    case_rows = build_case_qty_rows(products, shipped, merged, catalog)
    dim_rows = build_dimension_rows(products, shipped, merged, catalog, cfg.tolerance_mm)

    po = cfg.po or po_tag([inputs.packing_list, *inputs.instructions, inputs.erp])
    order = [f"{i}. {f.label}: {f.path.name}" for i, f in enumerate(loaded, start=1)]
    run_info = [
        f"ERP export: {inputs.erp.name}",
        f"Packing list: {inputs.packing_list.name if inputs.packing_list else '(none)'}",
        "Packing instructions, priority order (first wins, others fill gaps):"
        if order else "Packing instructions: (none)",
        *[f"    {line}" for line in order],
        f"Dimension tolerance: {cfg.tolerance_mm:g} mm",
        *[f"WARNING: {w}" for w in warnings],
    ]
    case_path = write_case_qty_report(
        case_rows, cfg.out_dir / output_name("PackingCaseQty", cfg.vendor, po, cfg.run_date),
        run_info)
    dim_path = write_dimension_report(
        dim_rows, cfg.out_dir / output_name("PackingDimensions", cfg.vendor, po, cfg.run_date),
        run_info)

    return RunResult(case_path, dim_path, case_rows, dim_rows, warnings,
                     missing_instructions=not loaded, instruction_order=order,
                     extra_outputs=extra_outputs)


def summary_lines(result: RunResult) -> list[str]:
    """Human summary, including the operator prompt when instructions were missing."""
    lines = []
    for title, rows in (("CASE QTY (bags per carton)", result.case_rows),
                        ("CARTON DIMENSIONS", result.dimension_rows)):
        lines.append(f"\n  {title}: {len(rows)} SKUs")
        for status in Status:
            lines.append(f"    {status.value:<12} {sum(1 for r in rows if r.status is status)}")
    if result.warnings:
        lines.append(f"\n--- Warnings ({len(result.warnings)}) ---")
        lines += [f"  ! {w}" for w in result.warnings]
    if result.missing_instructions:
        # Its own block, not one more warning: without instructions the only case-qty
        # source left is the ERP QtyPerCase, which is exactly the field known to go stale.
        bar = "=" * 72
        lines += [
            "", bar, "ASK THE OPERATOR", bar,
            "  No packing-instructions file was provided for this shipment.",
            "  'Instructions we sent' is EMPTY and case qty was checked against the",
            "  ERP QtyPerCase only - the field most likely to be out of date.",
            "  -> Confirm whether instructions exist for this PO. If they do, add the",
            "     file to the input folder and re-run; if not, these results stand.",
            bar,
        ]
    lines.append(f"\nDone:\n  {result.case_qty_path}\n  {result.dimensions_path}")
    lines += [f"  {p}" for p in result.extra_outputs]
    return lines
