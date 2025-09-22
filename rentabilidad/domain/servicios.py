from datetime import datetime
from typing import Iterable

from .entidades import Informe, LineaVenta
from .politicas import EstrategiaRentabilidad


class GeneradorInforme:
    def __init__(self, estrategia: EstrategiaRentabilidad, event_bus):
        self.estrategia = estrategia
        self.bus = event_bus

    def _emit(self, topic: str, msg: str) -> None:
        if self.bus:
            self.bus.publish(topic, msg)

    def construir(self, rows: Iterable[dict]) -> Informe:
        self._emit("log", "Normalizando datos…")
        filas: list[LineaVenta] = []
        for r in rows:
            try:
                lv = LineaVenta(
                    nit=str(r.get("nit", "") or "").strip(),
                    cliente=str(r.get("cliente", "") or "").strip(),
                    sucursal=str(r.get("sucursal", "") or "").strip(),
                    producto=str(r.get("producto", "") or "").strip(),
                    descripcion=str(r.get("descripcion", "") or "").strip(),
                    linea=str(r.get("linea", "") or "").strip(),
                    grupo=str(r.get("grupo", "") or "").strip(),
                    cantidad=float(r.get("cantidad", 0) or 0),
                    ventas=float(r.get("ventas", 0) or 0),
                    costos=float(r.get("costos", 0) or 0),
                    descuento=float(r.get("descuento", 0) or 0),
                    vendedor=str(r.get("vendedor", "") or "").strip(),
                    renta_pct=float(r.get("renta_pct", 0) or 0),
                    utilidad_pct=float(r.get("utilidad_pct", 0) or 0),
                )
            except (TypeError, ValueError) as exc:  # pragma: no cover - datos externos
                self._emit("error", f"Fila inválida descartada: {exc}")
                continue

            lv.costos = self.estrategia.costo_ajustado(lv)
            filas.append(lv)

        self._emit("log", f"{len(filas)} líneas normalizadas.")
        return Informe(filas)

    @staticmethod
    def parse_fecha(fecha: str | None) -> datetime | None:
        if not fecha:
            return None
        try:
            return datetime.strptime(fecha, "%Y-%m-%d")
        except ValueError:
            return None
