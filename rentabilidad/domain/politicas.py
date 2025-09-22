from abc import ABC, abstractmethod
from .entidades import LineaVenta


class EstrategiaRentabilidad(ABC):
    @abstractmethod
    def costo_ajustado(self, l: LineaVenta) -> float: ...


class EstrategiaSimple(EstrategiaRentabilidad):
    def costo_ajustado(self, l: LineaVenta) -> float:
        return l.costos  # punto de extensión: fletes, bonificaciones, etc.
