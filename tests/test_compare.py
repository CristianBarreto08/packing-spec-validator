from conftest import erp, instr, shipped

from packing_spec.compare import build_case_qty_rows, build_dimension_rows, resolve_case_qty
from packing_spec.models import Status


def test_precedence_and_text_custom_column_never_matches():
    correct, checks = resolve_case_qty(10, "Ignore", 10, 10)
    assert correct == 10
    assert checks["J"] is None and checks["K"] is None and checks["L"] is None
    assert checks["M"] == 10 and checks["N"] == 10
    # When the custom column is numeric and agrees, the template's J wins.
    assert resolve_case_qty(10, 10, 10, 10)[1]["J"] == 10
    assert resolve_case_qty(12, "Ignore", None, 10) == (None, {k: None for k in "JKLMN"})


def test_stale_erp_qty_is_update_with_known_value(catalog):
    [row] = build_case_qty_rows([erp(qty=12)], {"FBA-AB-1-P5": shipped(case_qty=10)},
                                {"FBA-AB-1-P5": instr(bags=10)}, catalog)
    assert row.status is Status.UPDATE
    assert row.correct == 10
    assert "stale" in row.note


def test_vendor_ignoring_instructions_is_review_even_when_erp_agrees(catalog):
    [row] = build_case_qty_rows([erp(qty=30)], {"FBA-AB-1-P5": shipped(case_qty=30)},
                                {"FBA-AB-1-P5": instr(bags=24)}, catalog)
    assert row.correct == 30  # value still known...
    assert row.status is Status.REVIEW  # ...but accepting it is a decision
    assert row.vendor_vs_instructions.startswith("NO match")


def test_two_packings_for_one_sku_blank_the_correct_value(catalog):
    conflict = shipped(case_qty=60, case_conflict="MORE THAN ONE carton packing")
    [row] = build_case_qty_rows([erp(qty=60)], {"FBA-AB-1-P5": conflict}, {}, catalog)
    assert row.status is Status.REVIEW and row.correct is None


def test_stale_inch_fields_are_update(catalog):
    [row] = build_dimension_rows([erp(inches=(7.0, 16.1, 15.0))],
                                 {"FBA-AB-1-P5": shipped()}, {}, catalog)
    assert row.status is Status.UPDATE
    assert row.correct_mm == (450, 320, 210)
    assert "CaseLength/Width/Height" in row.note


def test_instructions_alone_can_confirm_the_carton(catalog):
    product = erp(text="400 x 300 x 300", inches=(12, 10, 8))
    [row] = build_dimension_rows([product], {"FBA-AB-1-P5": shipped()},
                                 {"FBA-AB-1-P5": instr(box=(450, 320, 210))}, catalog)
    assert row.status is Status.UPDATE
    assert "SUPPLIERCASEDIMENSIONS and CaseLength/Width/Height" in row.note


def test_vendor_used_a_different_carton_than_requested(catalog):
    [row] = build_dimension_rows([erp()], {"FBA-AB-1-P5": shipped()},
                                 {"FBA-AB-1-P5": instr(box=(500, 350, 300))}, catalog)
    assert row.status is Status.REVIEW
    assert row.correct_mm == (450, 320, 210)
