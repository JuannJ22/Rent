from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook

from .fs import asegurar_carpeta


class ExporterExcel:
    def __init__(self, ruta_plantilla: Path):
        self.ruta = Path(ruta_plantilla)

    def volcar(
        self,
        filas: List[Dict],
        hoja_out: Optional[str] = None,
        ruta_salida: Optional[Path] = None,
        fila_inicio: int = 7,
    ) -> Path:
        libro = load_workbook(self.ruta)
        try:
            if hoja_out:
                hoja = libro[hoja_out]
            else:
                hoja = libro[libro.sheetnames[0]]

            max_row = hoja.max_row or fila_inicio
            if max_row >= fila_inicio:
                hoja.delete_rows(fila_inicio, max_row - fila_inicio + 1)

            for idx, row in enumerate(filas, start=fila_inicio):
                hoja.cell(idx, 1).value = row.get("nit")
                hoja.cell(idx, 2).value = row.get("cliente")
                hoja.cell(idx, 3).value = row.get("descripcion") or row.get("producto")
                hoja.cell(idx, 4).value = row.get("vendedor")
                hoja.cell(idx, 5).value = row.get("cantidad")
                hoja.cell(idx, 6).value = row.get("ventas")
                hoja.cell(idx, 7).value = row.get("costos")
                hoja.cell(idx, 8).value = row.get("margen")
                hoja.cell(idx, 9).value = row.get("utilidad_pct", row.get("margen"))
                hoja.cell(idx, 10).value = row.get("precio")
                hoja.cell(idx, 11).value = row.get("descuento")

            destino = ruta_salida or self.ruta.with_name(self.ruta.stem + "_OUT.xlsx")
            asegurar_carpeta(destino)
            libro.save(destino)
            return destino
        finally:
            libro.close()
