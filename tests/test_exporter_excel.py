from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

from rentabilidad.infra.exporter_excel import ExporterExcel


def test_exporter_excel_volca_datos_en_hoja(tmp_path) -> None:
    plantilla_path = tmp_path / "plantilla.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Base"
    for col in range(1, 12):
        ws.cell(row=6, column=col, value=f"col{col}")
    wb.save(plantilla_path)
    wb.close()

    filas = [
        {
            "nit": "123456",
            "cliente": "Cliente X",
            "descripcion": "Producto X",
            "producto": "PX",
            "vendedor": "Ana",
            "cantidad": 4,
            "ventas": 500.0,
            "costos": 200.0,
            "margen": 0.25,
            "utilidad_pct": 0.22,
            "precio": 125.0,
            "descuento": 0.05,
        }
    ]

    destino = tmp_path / "Informes" / "Marzo 02.xlsx"
    exporter = ExporterExcel(Path(plantilla_path))
    salida = exporter.volcar(filas, ruta_salida=destino, fila_inicio=7)

    assert salida == destino
    assert destino.exists()

    libro = load_workbook(destino)
    hoja = libro.active
    assert hoja.cell(7, 1).value == "123456"
    assert hoja.cell(7, 2).value == "Cliente X"
    assert hoja.cell(7, 3).value == "Producto X"
    assert hoja.cell(7, 5).value == 4
    assert hoja.cell(7, 6).value == 500.0
    assert hoja.cell(7, 7).value == 200.0
    assert hoja.cell(7, 8).value == 0.25
    assert hoja.cell(7, 9).value == 0.22
    assert hoja.cell(7, 10).value == 125.0
    assert hoja.cell(7, 11).value == 0.05
    libro.close()
