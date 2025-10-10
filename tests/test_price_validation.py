import math

import pytest

from hojas.hoja01_loader import (
    IVA_MULTIPLIER,
    PRICE_TOLERANCE,
    _build_discount_formula,
    _coerce_float,
    _guess_map,
    _is_iva_exempt,
)


def _diff_ratio(ventas, cantidad, expected_con_iva, *, iva_exempt=False):
    expected_unit = expected_con_iva if iva_exempt else expected_con_iva / IVA_MULTIPLIER
    venta_unitaria = ventas / cantidad
    return abs(venta_unitaria - expected_unit) / expected_unit


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


def test_exempt_products_do_not_apply_iva_multiplier():
    ventas = 10_000
    cantidad = 5
    expected_con_iva = 2_000

    diff_ratio = _diff_ratio(ventas, cantidad, expected_con_iva, iva_exempt=True)

    assert diff_ratio == 0


def test_build_discount_formula_applies_iva_multiplier_for_taxed_products():
    formula = _build_discount_formula("F", "E", "J", 7, iva_exempt=False)

    assert formula == "=1-((F7*1.19)/E7/J7)"


def test_build_discount_formula_skips_iva_multiplier_for_exempt_products():
    formula = _build_discount_formula("F", "E", "J", 7, iva_exempt=True)

    assert formula == "=1-((F7)/E7/J7)"


def test_guess_map_prefers_facturada_quantity_column():
    cols = [
        "Nit",
        "Descripción",
        "Cant Pedida",
        "Cant Fact.",
        "Ventas",
    ]

    mapping = _guess_map(cols)

    assert mapping["cantidad"] == "Cant Fact."


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


@pytest.mark.parametrize(
    "description, expected",
    [
        ("Producto EXENTO de IVA", True),
        ("Servicio excluido IVA", True),
        ("Producto gravado", False),
        (None, False),
    ],
)
def test_is_iva_exempt_detects_keywords(description, expected):
    assert _is_iva_exempt(description) is expected
