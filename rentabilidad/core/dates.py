"""Herramientas relacionadas con el cálculo de fechas de trabajo."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Protocol


class DateStrategy(Protocol):
    """Estrategia para obtener una fecha de referencia por defecto."""

    def default(self) -> date:
        """Retorna la fecha predeterminada."""


class TodayStrategy:
    """Estrategia que devuelve la fecha actual del sistema."""

    def default(self) -> date:
        return date.today()


class YesterdayStrategy:
    """Estrategia que devuelve la fecha del día anterior."""

    def default(self) -> date:
        return date.today() - timedelta(days=1)


@dataclass(frozen=True)
class DateResolver:
    """Resuelve fechas recibidas como texto aplicando una estrategia base."""

    strategy: DateStrategy

    def resolve(self, value: str | None) -> date:
        """Convierte ``value`` a ``date`` o usa la estrategia por defecto."""

        if value:
            try:
                return datetime.strptime(value, "%Y-%m-%d").date()
            except ValueError as exc:  # pragma: no cover - validación simple
                raise ValueError(
                    "La fecha debe tener el formato YYYY-MM-DD"
                ) from exc
        return self.strategy.default()


def ensure_previous_day(reference: date) -> date:
    """Devuelve el día anterior a ``reference``."""

    return reference - timedelta(days=1)


__all__ = [
    "DateResolver",
    "DateStrategy",
    "TodayStrategy",
    "YesterdayStrategy",
    "ensure_previous_day",
]
