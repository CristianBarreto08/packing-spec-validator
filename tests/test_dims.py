import pytest

from packing_spec.dims import (
    Verdict,
    compare_boxes,
    inches_to_box,
    parse_erp_dims,
    resolve_metric_unit,
)


def test_erp_text_in_cm_is_detected_against_inch_fields():
    parsed = parse_erp_dims("45 x 32 x 34 Accurate", (17.72, 12.6, 13.39))
    assert parsed.unit == "cm"
    assert parsed.box_mm == pytest.approx((450, 320, 340))


def test_dates_and_po_references_are_not_read_as_sides():
    parsed = parse_erp_dims("920 x 310 x 160 - Confirmed 08/07/2026 PO 13")
    assert parsed.box_mm == (920, 310, 160)
    assert parsed.unit == "mm"


def test_magnitude_decides_without_inches_and_missing_sides_are_never_invented():
    assert parse_erp_dims("37 x 37 x 38").unit == "cm"
    two_sides = parse_erp_dims("45 x 32 Accurate")
    assert two_sides.box_mm is None and "only 2" in two_sides.note
    assert parse_erp_dims("-").box_mm is None


def test_zero_inch_side_means_missing_not_zero():
    assert inches_to_box((24.4, 16.1, 0)) is None
    assert inches_to_box((10, 10, 10)) == pytest.approx((254, 254, 254))


def test_header_says_cm_but_declared_cbm_proves_millimetres():
    # 450*320*210 mm = 0.03024 m3 per carton; 6 cartons declared as 0.181 m3.
    box, unit = resolve_metric_unit([450, 320, 210], cartons=6, declared_cbm=0.181)
    assert unit == "mm" and box == (450, 320, 210)
    box, unit = resolve_metric_unit([45, 32, 21], cartons=6, declared_cbm=0.181)
    assert unit == "cm" and box == (450, 320, 210)


@pytest.mark.parametrize(
    ("a", "b", "verdict"),
    [
        ((450, 320, 210), (450, 320, 210), Verdict.EXACT),
        ((450, 320, 210), (452, 321, 206), Verdict.TOLERANCE),
        ((320, 450, 210), (450, 320, 210), Verdict.ROTATED),
        ((450, 320, 210), (460, 320, 210), Verdict.DIFFERENT),
    ],
)
def test_compare_boxes(a, b, verdict):
    assert compare_boxes(a, b)[0] is verdict
