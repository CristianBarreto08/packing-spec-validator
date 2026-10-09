from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from packing_spec.catalog import SkuCatalog
from packing_spec.models import ErpProduct, MergedInstruction, ShippedSku
from packing_spec.pipeline import RunConfig

SAMPLE = Path(__file__).resolve().parents[1] / "examples" / "sample_data"


def erp(sku: str = "FBA-AB-1-P5", qty: object = 10, text: object = "450 x 320 x 210",
        inches: tuple = (17.72, 12.6, 8.27)) -> ErpProduct:
    return ErpProduct(sku, "Test item", 4, qty, "Ignore", text, inches, 2)


def shipped(sku: str = "FBA-AB-1-P5", case_qty: int = 10,
            box: tuple = (450.0, 320.0, 210.0), **kw) -> ShippedSku:
    return ShippedSku(sku=sku, case_qty=case_qty, case_qty_values=[case_qty], box_mm=box,
                      cartons=5, quantity=250, parent="AB-1", rows=[7], **kw)


def instr(bags=None, box=None, label: str = "PO 2") -> MergedInstruction:
    return MergedInstruction(bags, label if bags is not None else None,
                             box, label if box is not None else None, [label])


@pytest.fixture
def catalog() -> SkuCatalog:
    return SkuCatalog()


@pytest.fixture
def sample_config(tmp_path: Path) -> RunConfig:
    return RunConfig(catalog=SAMPLE / "reference" / "sku_catalog.xlsx",
                     input_dir=SAMPLE / "input", out_dir=tmp_path, vendor="Northwind",
                     run_date=date(2026, 9, 21))
