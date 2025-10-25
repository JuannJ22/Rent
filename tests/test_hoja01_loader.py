import pandas as pd

from openpyxl import Workbook

from hojas.hoja01_loader import (
    _build_price_mismatch_message,
    _build_sika_customer_message,
    _build_vendor_mismatch_message,
    _combine_reason_messages,
    _load_terceros_lookup,
    _load_vendedores_document_lookup,
    _normalize_nit_value,
    _normalize_product_key,
    _drop_full_rentability_rows,
)


def test_build_price_mismatch_message_includes_lista() -> None:
    message = _build_price_mismatch_message(
        100.0, 120.0, 5, 0.2, lista_precio=1
    )

    assert "lista 1" in message
    assert "mayor" in message


def test_build_price_mismatch_message_without_lista() -> None:
    message = _build_price_mismatch_message(100.0, 80.0, None, 0.2)

    assert "la lista" in message
    assert "menor" in message


def test_build_vendor_mismatch_message_includes_code() -> None:
    assert _build_vendor_mismatch_message("A1") == "Está creado con código A1."
    assert _build_vendor_mismatch_message(None) is None


def test_build_sika_customer_message_for_valid_lists() -> None:
    assert _build_sika_customer_message(7) == "CLIENTE CONSTRUCTORA SIKA TIPO A"
    assert _build_sika_customer_message(9) == "CLIENTE CONSTRUCTORA SIKA TIPO B"
    assert _build_sika_customer_message(1) is None


def test_load_terceros_lookup_uses_second_column_for_lista() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "TERCEROS"
    ws.append(["123", 7, "A1"])

    lookup = _load_terceros_lookup(wb)
    nit_key = _normalize_nit_value("123")

    assert lookup[nit_key]["lista"] == 7
    assert lookup[nit_key]["vendedor"] == "A1"


def test_load_vendedores_document_lookup_uses_quantity_column() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "VENDEDORES"
    ws.append(["123", "A1", "FAC", "PRF", "1001", "Producto X", 7])

    lookup = _load_vendedores_document_lookup(wb)

    key = _normalize_product_key("Producto X")
    assert lookup[key][0]["cantidad"] == 7


def test_combine_reason_messages_single_line() -> None:
    message = _combine_reason_messages([" Uno", "Dos ", "", None])

    assert message == "Uno Dos"
    assert "\n" not in message


def test_drop_full_rentability_rows_removes_100_percent_values() -> None:
    df = pd.DataFrame(
        {
            "descripcion": ["ok", "full_text", "full_numeric", "nan_value"],
            "renta": [50, "100%", 100.0, None],
        }
    )

    result = _drop_full_rentability_rows(df)

    assert list(result["descripcion"]) == ["ok", "nan_value"]


def test_drop_full_rentability_rows_removes_fractional_hundreds() -> None:
    df = pd.DataFrame(
        {
            "descripcion": ["keep", "remove_fraction"],
            "renta": [0.85, 1.0],
        }
    )

    result = _drop_full_rentability_rows(df)

    assert list(result["descripcion"]) == ["keep"]
