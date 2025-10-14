from __future__ import annotations

import os
import subprocess

from ..dto import GenerarInformeRequest, GenerarInformeResponse
from ...config import settings


def _emit_lines(raw_output: str, channel: str, bus) -> None:
    for raw in raw_output.splitlines():
        text = raw.strip()
        if text:
            bus.publish(channel, text)


def run(req: GenerarInformeRequest, bus) -> GenerarInformeResponse:
    del req
    if os.name != "nt":
        mensaje = "La ejecución manual solo está disponible en Windows."
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    batch_script = settings.manual_batch_script
    if batch_script is None:
        mensaje = "No se configuró el script manual a ejecutar."
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    if not batch_script.exists():
        mensaje = f"No se encontró el script manual: {batch_script}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    bus.publish("log", f"Ejecutando script manual: {batch_script}")

    command = [os.environ.get("COMSPEC", "cmd.exe"), "/c", str(batch_script)]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            cwd=str(batch_script.parent),
        )
    except FileNotFoundError as exc:  # pragma: no cover - dependiente del sistema
        mensaje = f"No se pudo iniciar el intérprete de comandos: {exc}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)
    except Exception as exc:  # pragma: no cover - defensivo
        mensaje = f"Error al ejecutar el script manual: {exc}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    if result.stdout:
        _emit_lines(result.stdout, "log", bus)
    if result.stderr:
        _emit_lines(result.stderr, "error", bus)

    if result.returncode != 0:
        mensaje = f"El script finalizó con código {result.returncode}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    bus.publish("done", "Script manual ejecutado correctamente")
    return GenerarInformeResponse(ok=True, mensaje="OK")
