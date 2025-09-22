from dataclasses import dataclass
from typing import List, Dict, Any


@dataclass
class LineaVenta:
    nit: str
    cliente: str
    sucursal: str
    producto: str
    descripcion: str
    linea: str
    grupo: str
    cantidad: float
    ventas: float
    costos: float
    descuento: float = 0.0
    vendedor: str = ""
    renta_pct: float = 0.0
    utilidad_pct: float = 0.0

    @property
    def ingreso(self) -> float:
        return self.ventas * (1 - self.descuento)

    @property
    def margen(self) -> float:
        v = self.ingreso
        return 0.0 if v == 0 else (v - self.costos) / v

    @property
    def precio(self) -> float:
        return 0.0 if self.cantidad == 0 else self.ventas / self.cantidad


@dataclass
class Informe:
    filas: List[LineaVenta]

    def filtrar_bajo_margen(self, umbral: float = 0.10) -> "Informe":
        return Informe([f for f in self.filas if f.margen < umbral])

    def to_rows(self) -> List[Dict[str, Any]]:
        return [
            dict(
                nit=f.nit,
                cliente=f.cliente,
                sucursal=f.sucursal,
                producto=f.producto,
                descripcion=f.descripcion,
                cantidad=f.cantidad,
                ventas=f.ingreso,
                costos=f.costos,
                margen=f.margen,
                renta_pct=f.renta_pct,
                utilidad_pct=f.utilidad_pct,
                precio=f.precio,
                descuento=f.descuento,
                linea=f.linea,
                grupo=f.grupo,
                vendedor=f.vendedor,
            )
            for f in self.filas
        ]
