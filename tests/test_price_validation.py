from hojas.hoja01_loader import IVA_MULTIPLIER, PRICE_TOLERANCE


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
