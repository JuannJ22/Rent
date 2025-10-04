import math

import pytest

from hojas.hoja01_loader import IVA_MULTIPLIER, PRICE_TOLERANCE, _coerce_float


def _diff_ratio(ventas, cantidad, expected_con_iva):
    expected_sin_iva = expected_con_iva / IVA_MULTIPLIER
    venta_unitaria = ventas / cantidad
    return abs(venta_unitaria - expected_sin_iva) / expected_sin_iva


def test_sample_row_is_within_tolerance_when_using_quantity_column():
    ventas = 947_798.32
    cantidad = 30
    expected_con_iva = 37_596.005

    diff_ratio = _diff_ratio(ventas, cantidad, expected_con_iva)

    assert diff_ratio < PRICE_TOLERANCE


def test_wrong_quantity_triggers_price_mismatch():
    ventas = 947_798.32
    cantidad = 26
    expected_con_iva = 37_596.005

    diff_ratio = _diff_ratio(ventas, cantidad, expected_con_iva)

    assert diff_ratio > PRICE_TOLERANCE


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("37.596,005", 37_596.005),
        ("37596,005", 37_596.005),
        ("1,234,567.89", 1_234_567.89),
        ("(1.234,56)", -1_234.56),
    ],
)
def test_coerce_float_supports_common_decimal_formats(raw, expected):
    parsed = _coerce_float(raw)
    assert parsed is not None
    assert math.isclose(parsed, expected, rel_tol=0, abs_tol=1e-9)
