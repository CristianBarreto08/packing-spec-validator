from pathlib import Path

import openpyxl

from packing_spec.instructions import merge_instructions
from packing_spec.models import InstructionEntry, InstructionsFile
from packing_spec.po_tags import order_by_priority, po_numbers, po_tag
from packing_spec.sheet_utils import clean_sku, find_sku_column_by_content


def test_dates_in_file_names_are_not_po_numbers():
    assert po_numbers("Packing list WITH CONTAINER 2026-09-02 PO 12.xlsx") == ["12"]
    assert po_tag([Path("PL PO 154-155 2026-09-03.xlsx")]) == "PO 154-155"


def test_newest_po_has_priority_then_newest_date():
    paths = [Path("Instr PO 154 2026-08-01.xlsx"), Path("Instr PO 155 2026-07-01.xlsx"),
             Path("Instr PO 154 2026-09-01.xlsx")]
    assert [p.name for p in order_by_priority(paths)] == [
        "Instr PO 155 2026-07-01.xlsx", "Instr PO 154 2026-09-01.xlsx",
        "Instr PO 154 2026-08-01.xlsx"]


def _file(label, entries):
    return InstructionsFile(Path(label), label, "Box Information", 1, entries, False, [])


def test_merge_is_field_by_field_with_provenance():
    newer = _file("PO 2", {"A": InstructionEntry(15, None, 5)})
    older = _file("PO 1", {"A": InstructionEntry(12, (410, 310, 300), 9),
                           "B": InstructionEntry(8, None, 10)})
    merged, stats = merge_instructions([newer, older])
    assert merged["A"].bags_per_carton == 15 and merged["A"].bags_source == "PO 2"
    assert merged["A"].box_mm == (410, 310, 300) and merged["A"].box_source == "PO 1"
    assert merged["B"].bags_source == "PO 1"
    assert stats["PO 1"] == {"bags": 1, "box": 1}


def test_sku_column_found_by_content_not_by_header():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ITEM NO.", "FBA SKU", None, "QUANTITY"])
    known = {f"FBA-X-{i}-P5" for i in range(5)}
    for i in range(5):
        ws.append([f"X-{i}", f"Q{i}Z-ABC-123", f"FBA-X-{i}-P5", 100])
    ws.append(["X-9", "Q9Z-ABC-123", "#N/A", 100])
    col, ratio = find_sku_column_by_content(ws, 1, known)
    assert col == 3 and ratio == 1.0
    assert clean_sku("#N/A") is None and clean_sku(" FBA-X-1-P5 ") == "FBA-X-1-P5"
