from __future__ import annotations

from datetime import datetime
from io import StringIO
from pathlib import Path
import contextlib
import sys

from excel_base.clone_from_template import TemplateCloneService
from hojas import hoja01_loader

from ..dto import GenerarInformeRequest, GenerarInformeResponse
from ...config import settings
from ...infra.fs import ayer_str


def _parse_fecha(fecha: str | None) -> datetime | None:
    if not fecha:
        return None
    try:
        return datetime.strptime(fecha, "%Y-%m-%d")
    except ValueError:
        return None


def _clone_template(target_date: datetime, plantilla: Path, bus, *, force: bool) -> Path:
    service = TemplateCloneService(settings.context)
    destino = settings.context.informe_path(target_date.date())

    if destino.exists() and not force:
        bus.publish("log", f"Usando informe existente: {destino}")
        return destino

    bus.publish("log", f"Clonando plantilla hacia {destino}…")
    resultado = service.clone(plantilla, target_date.date())
    bus.publish("log", f"Plantilla preparada: {resultado}")
    return resultado


def _run_rentv1_loader(path: Path, fecha: datetime, bus) -> tuple[int, str]:
    args = [
        "hoja01_loader.py",
        "--excel",
        str(path),
        "--fecha",
        fecha.strftime("%Y-%m-%d"),
        "--exczdir",
        str(settings.excz_dir),
        "--excz-prefix",
        settings.excz_prefix,
    ]
    if settings.plantilla_hoja:
        args.extend(["--hoja", settings.plantilla_hoja])

    buffer = StringIO()
    exit_code = 0
    old_argv = sys.argv[:]
    try:
        sys.argv = args
        with contextlib.redirect_stdout(buffer):
            hoja01_loader.main()
    except SystemExit as exc:  # pragma: no cover - dependiente del script externo
        exit_code = exc.code if isinstance(exc.code, int) else 1
    except Exception as exc:  # pragma: no cover - defensivo
        exit_code = 1
        buffer.write(f"ERROR: {exc}\n")
    finally:
        sys.argv = old_argv

    return exit_code, buffer.getvalue()


def _emit_loader_output(output: str, bus) -> tuple[str | None, str | None]:
    last_info: str | None = None
    last_error: str | None = None
    for raw in output.splitlines():
        text = raw.strip()
        if not text:
            continue
        if text.upper().startswith("ERROR"):
            bus.publish("error", text)
            last_error = text
        else:
            bus.publish("log", text)
            last_info = text
    return last_info, last_error


def run(req: GenerarInformeRequest, bus) -> GenerarInformeResponse:
    fecha_texto = req.fecha or ayer_str()
    objetivo = _parse_fecha(fecha_texto)
    if not objetivo:
        mensaje = "La fecha debe tener el formato YYYY-MM-DD"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    plantilla = Path(req.ruta_plantilla)
    if not plantilla.exists():
        mensaje = f"No existe la plantilla indicada: {plantilla}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    try:
        bus.publish("log", f"Fecha objetivo: {fecha_texto}")
        force_clone = req.fecha is None
        informe_path = _clone_template(objetivo, plantilla, bus, force=force_clone)

        bus.publish("log", "Ejecutando motor Rentv1 para actualizar el informe…")
        exit_code, output = _run_rentv1_loader(informe_path, objetivo, bus)
        last_info, last_error = _emit_loader_output(output, bus)

        if exit_code != 0:
            mensaje = last_error or "El proceso Rentv1 finalizó con errores"
            bus.publish("error", mensaje)
            return GenerarInformeResponse(ok=False, mensaje=mensaje)

        mensaje_ok = last_info or f"Informe generado: {informe_path}"
        bus.publish("done", mensaje_ok)
        return GenerarInformeResponse(
            ok=True,
            mensaje="OK",
            ruta_salida=str(informe_path),
        )
    except Exception as exc:  # pragma: no cover - interacción con IO
        bus.publish("error", str(exc))
        return GenerarInformeResponse(ok=False, mensaje=str(exc))
