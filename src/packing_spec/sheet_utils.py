"""Header-driven worksheet helpers.

Columns are located by normalized header text, or by *content*, never by a fixed
letter: ERP exports and vendor spreadsheets reorder, rename and blank out columns
from one shipment to the next.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterable
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .models import Number

# An Excel error is not a SKU. Vendor files built with lookups carry "#N/A" where
# the lookup failed; treating it as a SKU would silently drop the row instead of
# letting the fallback resolver (parent + pack size) recover it.
EXCEL_ERRORS = frozenset(
    {"#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!", "N/A"}
)


class ColumnNotFoundError(ValueError):
    """A required column is missing. Raised instead of guessing a column."""


def normalize_header(value: Any) -> str:
    """Lowercase alphanumerics only: 'BAGS (BOXES)/\\nCARTON' -> 'bagsboxescarton'."""
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def clean_sku(value: Any) -> str | None:
    """SKU text, or None when the cell is blank or holds an Excel error."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() in EXCEL_ERRORS:
        return None
    return text


def to_number(value: Any) -> Number | None:
    """Numeric cell value; ERP exports often store numbers as text. Whole floats become int."""
    if isinstance(value, bool):
        return None
    if isinstance(value, str):
        try:
            value = float(value.strip().replace(",", ""))
        except ValueError:
            return None
    if isinstance(value, (int, float)):
        return int(round(value)) if abs(value - round(value)) < 1e-9 else value
    return None


def find_header_row(ws: Worksheet, max_scan: int = 15, min_text_cells: int = 3) -> int:
    """First row with several text cells. Vendor files put letterheads above the table."""
    for r in range(1, min(max_scan, ws.max_row) + 1):
        texts = [v for v in row_values(ws, r) if isinstance(v, str) and v.strip()]
        if len(texts) >= min_text_cells:
            return r
    return 1


def row_values(ws: Worksheet, row: int) -> list[Any]:
    return [ws.cell(row=row, column=c).value for c in range(1, ws.max_column + 1)]


def find_column(headers: list[Any], predicate: Callable[[str], bool]) -> int | None:
    """1-based index of the first header whose normalized text satisfies `predicate`."""
    for idx, header in enumerate(headers, start=1):
        if predicate(normalize_header(header)):
            return idx
    return None


def require_column(headers: list[Any], predicate: Callable[[str], bool], label: str) -> int:
    col = find_column(headers, predicate)
    if col is None:
        raise ColumnNotFoundError(f"No column found for '{label}'. Headers seen: {headers}")
    return col


def find_sku_column_by_content(
    ws: Worksheet,
    header_row: int,
    known_skus: set[str],
    min_ratio: float = 0.6,
    min_hits: int = 3,
) -> tuple[int | None, float]:
    """Locate the internal-SKU column by what it contains, not by what it is called.

    Vendors routinely leave the internal-SKU column without a header while a column
    literally named "FBA SKU" holds marketplace SKUs; matching by name picks the wrong
    one. Here the winner is the column with the most values that exist in the ERP.

    It ranks by *number* of hits, using the ratio only as a floor: an instructions file
    lists every SKU on the PO, not just the ones on this container, so its correct
    column can sit near 50% while a wrong column scores a lower ratio but is never
    ahead on hits.
    """
    best_col, best_hits, best_ratio = None, 0, 0.0
    for c in range(1, ws.max_column + 1):
        hits = total = 0
        for r in range(header_row + 1, ws.max_row + 1):
            value = clean_sku(ws.cell(row=r, column=c).value)
            if value is None:
                continue
            total += 1
            hits += value in known_skus
        if total < 3 or hits < min_hits or hits / total < min_ratio:
            continue
        if hits > best_hits:
            best_col, best_hits, best_ratio = c, hits, hits / total
    return best_col, best_ratio


def merged_anchor_map(ws: Worksheet, columns: Iterable[int]) -> dict[tuple[int, int], int]:
    """{(row, col): anchor_row} for vertically merged cells in the given columns.

    When one SKU ships in two packing configurations, vendors merge the SKU cell over
    both rows. openpyxl only returns the value on the anchor row, so without this the
    second configuration would vanish from the comparison without any warning.
    """
    wanted = set(columns)
    anchors: dict[tuple[int, int], int] = {}
    for rng in ws.merged_cells.ranges:
        if rng.max_row == rng.min_row:
            continue
        for col in wanted:
            if rng.min_col <= col <= rng.max_col:
                for r in range(rng.min_row, rng.max_row + 1):
                    anchors[(r, col)] = rng.min_row
    return anchors


def cell_through_merge(
    ws: Worksheet, row: int, col: int | None, anchors: dict[tuple[int, int], int]
) -> Any:
    if col is None:
        return None
    return ws.cell(row=anchors.get((row, col), row), column=col).value
