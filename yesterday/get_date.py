"""Punto de entrada legacy para ``from yesterday.get_date import ...``."""

from rentabilidad.core.dates import DateResolver, YesterdayStrategy

__all__ = ["DateResolver", "YesterdayStrategy"]
