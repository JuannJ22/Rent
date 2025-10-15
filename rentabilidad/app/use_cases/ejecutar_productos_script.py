"""Ejecución del script de productos desde la interfaz gráfica."""

from __future__ import annotations

import os
import subprocess

from ..dto import GenerarInformeResponse
from ...config import settings


def _emit_lines(raw_output: str, channel: str, bus) -> None:
    for raw in raw_output.splitlines():
        text = raw.strip()
        if text:
            bus.publish(channel, text)


def run(bus) -> GenerarInformeResponse:
    if os.name != "nt":
        mensaje = "La generación de productos solo está disponible en Windows."
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    script_path = settings.productos_batch_script
    if script_path is None:
        mensaje = "No se configuró el script Productos.bat a ejecutar."
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    if not script_path.exists():
        mensaje = f"No se encontró el script de productos: {script_path}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    bus.publish("log", f"Ejecutando script de productos: {script_path}")

    command = [os.environ.get("COMSPEC", "cmd.exe"), "/c", str(script_path)]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            cwd=str(script_path.parent),
        )
    except FileNotFoundError as exc:  # pragma: no cover - dependiente del sistema
        mensaje = f"No se pudo iniciar el intérprete de comandos: {exc}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)
    except Exception as exc:  # pragma: no cover - defensivo
        mensaje = f"Error al ejecutar el script de productos: {exc}"
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

    bus.publish("done", "Productos generados correctamente")
    return GenerarInformeResponse(ok=True, mensaje="OK")

