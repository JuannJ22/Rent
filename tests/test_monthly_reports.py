from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

from rentabilidad.services.monthly_reports import (
    MonthlyReportConfig,
    MonthlyReportService,
)


ORANGE = PatternFill(fill_type="solid", start_color="FFFCD5B4", end_color="FFFCD5B4")
YELLOW = PatternFill(fill_type="solid", start_color="FFFFFF00", end_color="FFFFFF00")


def _create_templates(base_dir: Path) -> tuple[Path, Path]:
    codigos_path = base_dir / "PLANTILLACODIGOS.xlsx"
    cobros_path = base_dir / "PLANTILLAMALCOBRO.xlsx"

    wb_codigos = Workbook()
    ws_codigos = wb_codigos.active
    ws_codigos.append(
        [
            None,
            "NIT",
            "CLIENTE",
            "DESCRIPCION",
            "VENDEDOR",
            "CANTIDAD",
            "VENTAS",
            "COSTOS",
            "RENTA",
            "UTILIDAD",
            "PRECIO",
            "DESCUENTO",
            "RAZON",
        ]
    )
    for _ in range(23):
        ws_codigos.append([None] * 13)
    ws_codigos.append([None] * 13)
    total_row = ws_codigos.max_row
    ws_codigos.cell(total_row, 2).value = "TOTAL"
    ws_codigos.cell(total_row, 6).value = "=SUM(F2:F24)"
    ws_codigos.cell(total_row, 7).value = "=SUM(G2:G24)"
    wb_codigos.save(codigos_path)
    wb_codigos.close()

    wb_cobros = Workbook()
    ws_cobros = wb_cobros.active
    ws_cobros.append(
        [
            "FECHA",
            "VENDEDOR",
            "FACTURA",
            "CANTIDAD",
            "PRODUCTO",
            "DESCUENTO AUTORIZADO",
            "DESCUENTO FACTURADO",
            "OBSERVACION",
            "SOLUCION",
            "VALOR DEL ERROR",
            "VALOR COBRADO",
        ]
    )
    wb_cobros.save(cobros_path)
    wb_cobros.close()

    return codigos_path, cobros_path


def _create_informe(base_dir: Path) -> Path:
    informes_dir = base_dir / "Informes" / "Marzo"
    informes_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "MARZO 00"
    for _ in range(5):
        ws.append([None] * 12)
    ws.append(
        [
            "NIT",
            "NIT - SUCURSAL - CLIENTE",
            "DESCRIPCION",
            "COD. VENDEDOR",
            "CANTIDAD",
            "VENTAS",
            "COSTOS",
            "% RENTA.",
            "% UTILI.",
            "PRECIO",
            "DESCUENTO",
            "RAZON",
        ]
    )
    row_codigos = [
        "123",
        "000123-000-CLIENTE UNO",
        "Producto A",
        "Vendedor Uno",
        2,
        2400,
        1500,
        0.25,
        0.35,
        1200,
        0.20,
        "Precio diferente",
    ]
    start_row = ws.max_row + 1
    for col_idx, value in enumerate(row_codigos, start=1):
        cell = ws.cell(start_row, col_idx, value)
        if col_idx == 4:
            cell.fill = ORANGE
    ws.cell(start_row, 12).fill = ORANGE
    row_cobros = [
        "456",
        "000456-000-CLIENTE DOS",
        "Producto B",
        "Vendedor Dos",
        5,
        5500,
        3200,
        0.18,
        0.27,
        1100,
        None,
        "Doc",
    ]
    start_row += 1
    for col_idx, value in enumerate(row_cobros, start=1):
        cell = ws.cell(start_row, col_idx, value)
        if col_idx == 12:
            cell.fill = YELLOW
            cell.comment = Comment("Doc: FV-123 Observación de prueba", "QA")
    ws.cell(start_row, 6).fill = YELLOW

    ws_ter = wb.create_sheet("TERCEROS")
    ws_ter.append(["NIT", "Lista"])
    ws_ter.append(["456", 10])

    ws_prec = wb.create_sheet("PRECIOS")
    ws_prec.append(["PRODUCTO", "LISTA 12", "LISTA 10"])
    ws_prec.append(["Producto A", 1200, 1000])
    ws_prec.append(["Producto B", 2000, 1500])

    informe_path = informes_dir / "INFORME_20230301.xlsx"
    wb.save(informe_path)
    wb.close()
    return informes_dir


def test_monthly_reports_generation(tmp_path):
    codigos_tpl, cobros_tpl = _create_templates(tmp_path)
    informes_dir = _create_informe(tmp_path)
    consolidados_dir = tmp_path / "Consolidados"
    config = MonthlyReportConfig(
        informes_dir=informes_dir.parent,
        plantilla_codigos=codigos_tpl,
        plantilla_malos_cobros=cobros_tpl,
        consolidados_codigos_dir=consolidados_dir / "Codigos",
        consolidados_cobros_dir=consolidados_dir / "Cobros",
    )
    service = MonthlyReportService(config)

    assert service.list_months() == ["Marzo"]

    codigos_path = service.generar_codigos_incorrectos("Marzo", bus=None)
    wb_codigos = load_workbook(codigos_path)
    ws_codigos = wb_codigos.active
    assert ws_codigos.cell(2, 1).value == "2023-03-01"
    assert ws_codigos.cell(2, 2).value == "123"
    assert ws_codigos.cell(2, 4).value == "Producto A"
    assert ws_codigos.cell(2, 13).value == "Precio diferente"
    assert ws_codigos.cell(25, 2).value == "TOTAL"
    wb_codigos.close()

    cobros_path = service.generar_malos_cobros("Marzo", bus=None)
    wb_cobros = load_workbook(cobros_path)
    ws_cobros = wb_cobros.active
    assert ws_cobros.cell(2, 1).value == "2023-03-01"
    assert ws_cobros.cell(2, 2).value == "Vendedor Dos"
    assert ws_cobros.cell(2, 3).value == "FV-123"
    assert ws_cobros.cell(2, 5).value == "Producto B"
    autorizado = ws_cobros.cell(2, 6).value
    facturado = ws_cobros.cell(2, 7).value
    assert round(autorizado, 4) == 0.25
    assert facturado == pytest.approx(0.3455, rel=1e-3)
    assert ws_cobros.cell(2, 8).value == "Observación de prueba"
    valor_error = ws_cobros.cell(2, 10).value
    assert valor_error == pytest.approx((facturado - autorizado) * 2000 * 5)
    wb_cobros.close()


def test_codigos_incorrectos_inserta_filas(tmp_path):
    codigos_tpl, cobros_tpl = _create_templates(tmp_path)
    informes_dir = tmp_path / "Informes" / "Abril"
    informes_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "ABRIL 00"
    for _ in range(5):
        ws.append([None] * 12)
    ws.append(
        [
            "NIT",
            "CLIENTE",
            "DESCRIPCION",
            "VENDEDOR",
            "CANTIDAD",
            "VENTAS",
            "COSTOS",
            "RENTA",
            "UTILIDAD",
            "PRECIO",
            "DESCUENTO",
            "RAZON",
        ]
    )
    for idx in range(24):
        start_row = ws.max_row + 1
        values = [
            f"10{idx:02d}",
            f"CLIENTE {idx}",
            f"Producto {idx}",
            f"Vendedor {idx}",
            1,
            1000 + idx,
            800 + idx,
            0.2,
            0.3,
            1200,
            0.1,
            f"Observacion {idx}",
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(start_row, col_idx, value)
            if col_idx in (4, 12):
                cell.fill = ORANGE

    informe_path = informes_dir / "INFORME_20230401.xlsx"
    wb.save(informe_path)
    wb.close()

    config = MonthlyReportConfig(
        informes_dir=informes_dir.parent,
        plantilla_codigos=codigos_tpl,
        plantilla_malos_cobros=cobros_tpl,
        consolidados_codigos_dir=tmp_path / "Consolidados" / "Codigos",
        consolidados_cobros_dir=tmp_path / "Consolidados" / "Cobros",
    )
    service = MonthlyReportService(config)

    codigos_path = service.generar_codigos_incorrectos("Abril", bus=None)
    wb_codigos = load_workbook(codigos_path)
    ws_codigos = wb_codigos.active
    # 24 filas de datos deben mover la fila TOTAL una posición hacia abajo
    assert ws_codigos.cell(2, 2).value == "1000"
    assert ws_codigos.cell(25, 2).value == "1023"
    assert ws_codigos.cell(26, 2).value == "TOTAL"
    wb_codigos.close()


def test_month_directory_missing(tmp_path):
    codigos_tpl, cobros_tpl = _create_templates(tmp_path)
    config = MonthlyReportConfig(
        informes_dir=tmp_path / "Informes",
        plantilla_codigos=codigos_tpl,
        plantilla_malos_cobros=cobros_tpl,
        consolidados_codigos_dir=tmp_path / "Codigos",
        consolidados_cobros_dir=tmp_path / "Cobros",
    )
    service = MonthlyReportService(config)

    try:
        service.generar_codigos_incorrectos("Abril", bus=None)
    except FileNotFoundError:
        pass
    else:
        raise AssertionError("Expected FileNotFoundError for missing month")
