from abc import ABC, abstractmethod
from .entidades import LineaVenta


class EstrategiaRentabilidad(ABC):
    """Define el contrato para calcular el costo ajustado de una línea."""

    @abstractmethod
    def costo_ajustado(self, linea: LineaVenta) -> float:
        """Devuelve el costo a registrar para ``linea``."""


class EstrategiaSimple(EstrategiaRentabilidad):
    """Estrategia por defecto que mantiene el costo reportado."""

    def costo_ajustado(self, linea: LineaVenta) -> float:
        """Retorna el costo sin modificaciones."""

        return linea.costos  # punto de extensión: fletes, bonificaciones, etc.
