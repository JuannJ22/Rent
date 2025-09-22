from __future__ import annotations

from datetime import datetime
from pathlib import Path

from ..dto import GenerarInformeRequest, GenerarInformeResponse
from ...config import settings
from ...domain.politicas import EstrategiaSimple
from ...domain.servicios import GeneradorInforme
from ...infra.excel_repo import ExcelRepo
from ...infra.exporter_excel import ExporterExcel
from ...infra.fs import ayer_str


def _parse_fecha(fecha: str | None) -> datetime | None:
    if not fecha:
        return None
    try:
        return datetime.strptime(fecha, "%Y-%m-%d")
    except ValueError:
        return None


def run(req: GenerarInformeRequest, bus) -> GenerarInformeResponse:
    fecha_texto = req.fecha or ayer_str()
    objetivo = _parse_fecha(fecha_texto)
    if not objetivo:
        mensaje = "La fecha debe tener el formato YYYY-MM-DD"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    try:
        bus.publish("log", f"Fecha objetivo: {fecha_texto}")
        repo = ExcelRepo(settings.excz_dir, prefix=settings.excz_prefix, hoja=settings.excz_sheet)
        bus.publish(
            "log",
            f"Buscando EXCZ en {settings.excz_dir} con prefijo {settings.excz_prefix}",
        )
        rows = repo.cargar_por_fecha(fecha_texto)
        bus.publish("log", f"{len(rows)} filas leídas del EXCZ.")

        generador = GeneradorInforme(EstrategiaSimple(), bus)
        informe = generador.construir(rows)

        plantilla = Path(req.ruta_plantilla)
        if not plantilla.exists():
            raise FileNotFoundError(f"No existe la plantilla indicada: {plantilla}")

        exporter = ExporterExcel(plantilla)
        salida = settings.context.informe_path(objetivo.date())
        hoja = settings.plantilla_hoja
        out_path = exporter.volcar(informe.to_rows(), hoja_out=hoja, ruta_salida=salida)

        bus.publish("done", f"Informe generado: {out_path}")
        return GenerarInformeResponse(ok=True, mensaje="OK", ruta_salida=str(out_path))
    except Exception as exc:  # pragma: no cover - interacción con IO
        bus.publish("error", str(exc))
        return GenerarInformeResponse(ok=False, mensaje=str(exc))
