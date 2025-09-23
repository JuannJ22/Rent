from __future__ import annotations

from typing import List, Tuple

import pytest

from rentabilidad.domain.politicas import EstrategiaSimple
from rentabilidad.domain.servicios import GeneradorInforme


class DummyBus:
    def __init__(self) -> None:
        self.events: List[Tuple[str, str]] = []

    def publish(self, topic: str, payload: str) -> None:
        self.events.append((topic, payload))


def test_generador_informe_normaliza_y_emite_eventos() -> None:
    bus = DummyBus()
    generador = GeneradorInforme(EstrategiaSimple(), bus)

    filas = [
        {
            "nit": "900123",  # strings sin espacios
            "cliente": "ACME S.A.",
            "sucursal": "Bogotá",
            "producto": "P001",
            "descripcion": "Producto estrella",
            "linea": "Línea 1",
            "grupo": "Grupo 2",
            "cantidad": 5,
            "ventas": 1000,
            "costos": 600,
            "descuento": 0.1,
            "vendedor": "Ana",
            "renta_pct": 0.2,
            "utilidad_pct": 0.18,
        }
    ]

    informe = generador.construir(filas)

    assert len(informe.filas) == 1
    linea = informe.filas[0]
    assert linea.nit == "900123"
    assert linea.ingreso == pytest.approx(900)  # 1000 * (1 - 0.1)
    assert linea.margen == pytest.approx((900 - 600) / 900)
    assert linea.precio == pytest.approx(1000 / 5)

    data = informe.to_rows()[0]
    assert data["ventas"] == pytest.approx(900)
    assert data["margen"] == pytest.approx((900 - 600) / 900)
    assert data["precio"] == pytest.approx(1000 / 5)

    topics = [topic for topic, _ in bus.events]
    assert "log" in topics
    assert any("normalizadas" in payload for topic, payload in bus.events if topic == "log")
