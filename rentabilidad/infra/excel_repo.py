from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook

from rentabilidad.core.excz import ExczFileFinder


def _limpiar_texto(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0") and texto.replace(".0", "").isdigit():
        return texto[:-2]
    return texto


@dataclass
class ExcelRepo:
    """Lee fuentes EXCZ y devuelve filas normalizadas (dicts)."""

    base_dir: Path
    prefix: str = "EXCZ980"
    hoja: str = "Hoja1"

    def _resolver_fecha(self, fecha: Optional[str]) -> Optional[datetime]:
        if not fecha:
            return None
        try:
            return datetime.strptime(fecha, "%Y-%m-%d")
        except ValueError:
            return None

    def _buscar_archivo(self, fecha: Optional[datetime]) -> Optional[Path]:
        finder = ExczFileFinder(self.base_dir)
        if fecha:
            encontrado = finder.find_for_date(self.prefix, fecha.date())
            if encontrado:
                return encontrado
        return finder.find_latest(self.prefix)

    def cargar_por_fecha(self, fecha: Optional[str]) -> List[Dict]:
        objetivo = self._resolver_fecha(fecha)
        archivo = self._buscar_archivo(objetivo)
        if not archivo or not archivo.exists():
            return []

        libro = load_workbook(archivo, data_only=True, read_only=True)
        try:
            hoja = libro[self.hoja]
        except KeyError:
            hoja = libro[libro.sheetnames[0]]

        filas: List[Dict] = []
        try:
            for valores in hoja.iter_rows(min_row=8, values_only=True):
                nit, sucursal, cliente, linea, grupo, producto, descripcion, cantidad, ventas, costos, renta_pct, utilidad_pct, *_ = (
                    list(valores) + [None] * 5
                )
                texto_cliente = _limpiar_texto(cliente)
                texto_descripcion = _limpiar_texto(descripcion)
                texto_linea = _limpiar_texto(linea)

                if not texto_cliente or not texto_descripcion:
                    continue
                if texto_cliente.lower().startswith("total"):
                    continue
                if texto_descripcion.lower().startswith("total"):
                    continue
                if texto_linea.lower().startswith("total"):
                    continue

                cantidad_num = float(cantidad or 0)
                ventas_num = float(ventas or 0)
                costos_num = float(costos or 0)

                def _normalizar_pct(valor) -> float:
                    if valor is None:
                        return 0.0
                    try:
                        numero = float(valor)
                    except (TypeError, ValueError):
                        return 0.0
                    return numero / 100 if abs(numero) > 1 else numero

                filas.append(
                    {
                        "nit": _limpiar_texto(nit),
                        "sucursal": _limpiar_texto(sucursal),
                        "cliente": texto_cliente,
                        "linea": texto_linea,
                        "grupo": _limpiar_texto(grupo),
                        "producto": _limpiar_texto(producto),
                        "descripcion": texto_descripcion,
                        "cantidad": cantidad_num,
                        "ventas": ventas_num,
                        "costos": costos_num,
                        "descuento": 0.0,
                        "vendedor": "",
                        "renta_pct": _normalizar_pct(renta_pct),
                        "utilidad_pct": _normalizar_pct(utilidad_pct),
                    }
                )
        finally:
            libro.close()

        return filas
