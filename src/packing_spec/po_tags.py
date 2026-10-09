"""PO numbers and dates read from file names: used to label sources and order priority.

Dates are stripped BEFORE looking for PO numbers: "... PO 1041 2026-09-02.xlsx" would
otherwise yield "PO 1041-2026-09-02". Only digits that follow the word "PO" count; any
other number in a file name is a date, a version or an internal ID.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

DATE_RE = re.compile(r"\d{4}[-_.]\d{1,2}[-_.]\d{1,2}|\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}")
PO_RE = re.compile(r"(?i)\bpo\s*#?\s*(\d{1,6}(?:\s*[-,+]\s*\d{1,6})*)")


def po_numbers(name: str) -> list[str]:
    numbers: list[str] = []
    for match in PO_RE.finditer(DATE_RE.sub(" ", name)):
        for num in re.findall(r"\d{1,6}", match.group(1)):
            if num not in numbers:
                numbers.append(num)
    return numbers


def source_label(path: Path) -> str:
    """Short provenance label for one instructions file, e.g. 'PO 1042'."""
    numbers = po_numbers(path.name)
    return "PO " + "-".join(numbers[:2]) if numbers else path.stem[:24]


def order_by_priority(paths: list[Path]) -> list[Path]:
    """Newest PO first: highest PO number, then newest date in the name, then name.

    This is an inference, not a fact, which is why the order is always printed and can
    be overridden by passing the files explicitly on the command line.
    """
    def newest_date(path: Path) -> str:
        found = [re.sub(r"\D", "", d) for d in DATE_RE.findall(path.name)]
        return max(found) if found else ""

    def key(path: Path) -> tuple:
        numbers = [int(n) for n in po_numbers(path.name)]
        return (-(max(numbers) if numbers else 0), _descending(newest_date(path)), path.name)

    return sorted(paths, key=key)


def _descending(text: str) -> tuple[int, ...]:
    return tuple(-ord(c) for c in text)


def po_tag(paths: list[Path | None]) -> str | None:
    """'PO 1041-1042' from all input names; None rather than a guess when absent."""
    numbers: list[str] = []
    for path in paths:
        if path is None:
            continue
        for num in po_numbers(path.name):
            if num not in numbers:
                numbers.append(num)
    return "PO " + "-".join(numbers[:3]) if numbers else None


def output_name(prefix: str, vendor: str | None, po: str | None, run_date: date) -> str:
    return " ".join(p for p in (prefix, vendor, po, run_date.isoformat()) if p) + ".xlsx"
