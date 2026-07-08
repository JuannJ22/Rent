from __future__ import annotations

from datetime import date
from pathlib import Path

from rentabilidad.app.use_cases import ejecutar_productos_script as use_case


class _Bus:
    def __init__(self) -> None:
        self.events: list[tuple[str, str]] = []

    def publish(self, channel: str, message: str) -> None:
        self.events.append((channel, message))


class _Service:
    def __init__(self, result: Path | Exception) -> None:
        self.result = result
        self.calls: list[date] = []

    def generate(self, target_date: date) -> Path:
        self.calls.append(target_date)
        if isinstance(self.result, Exception):
            raise self.result
        return self.result


class _Settings:
    def __init__(self, service: _Service) -> None:
        self.service = service

    def build_product_service(self) -> _Service:
        return self.service


def test_run_generates_products_through_application_service(monkeypatch, tmp_path: Path) -> None:
    target = tmp_path / "productos0708.xlsx"
    service = _Service(target)
    bus = _Bus()

    class _FixedStrategy:
        def default(self) -> date:
            return date(2026, 7, 8)

    monkeypatch.setattr(use_case, "TodayStrategy", _FixedStrategy)
    monkeypatch.setattr(use_case, "settings", _Settings(service))

    response = use_case.run(bus)

    assert response.ok is True
    assert response.ruta_salida == str(target)
    assert service.calls == [date(2026, 7, 8)]
    assert ("done", f"Listado de productos generado: {target}") in bus.events


def test_run_reports_service_errors(monkeypatch) -> None:
    service = _Service(RuntimeError("SIIGO no generó el archivo esperado"))
    bus = _Bus()

    monkeypatch.setattr(use_case, "settings", _Settings(service))

    response = use_case.run(bus)

    assert response.ok is False
    assert "SIIGO no generó el archivo esperado" in response.mensaje
    assert any(channel == "error" for channel, _ in bus.events)
