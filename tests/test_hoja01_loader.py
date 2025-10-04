from hojas.hoja01_loader import (
    _build_price_mismatch_message,
    _build_sika_customer_message,
    _build_vendor_mismatch_message,
    _combine_reason_messages,
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


def test_combine_reason_messages_single_line() -> None:
    message = _combine_reason_messages([" Uno", "Dos ", "", None])

    assert message == "Uno Dos"
    assert "\n" not in message
