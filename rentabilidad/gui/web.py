"""Interfaz web basada en NiceGUI para el panel de rentabilidad."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable
from types import SimpleNamespace

from nicegui import ui

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContextFactory


load_env()

_CONTEXT = PathContextFactory(os.environ).create()
RUTA_PLANTILLA = str(_CONTEXT.template_path())

# --- estilos soft globales ---
ui.add_head_html(
    """
<style>
  .q-card { border-radius: 1rem !important; }
  .q-field__control, .q-btn { border-radius: .75rem !important; }
  .q-btn { box-shadow: 0 1px 2px rgba(0,0,0,.06) !important; }
</style>
"""
)

state = SimpleNamespace(empty=None, log=None, last_update=None)


def copiar_ruta() -> None:
    """Copia ``RUTA_PLANTILLA`` al portapapeles usando JavaScript."""

    ui.run_javascript(f"navigator.clipboard.writeText({json.dumps(RUTA_PLANTILLA)})")
    agregar_log("Ruta de la plantilla copiada al portapapeles.")
    touch_last_update()


def abrir_carpeta() -> None:
    """Intenta abrir la carpeta que contiene la plantilla base."""

    carpeta = Path(RUTA_PLANTILLA).parent
    if not carpeta.exists():
        agregar_log(f"No se encontró la carpeta: {carpeta}")
        touch_last_update()
        return

    if sys.platform.startswith("win"):
        os.startfile(carpeta)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(carpeta)], check=False)
    else:
        subprocess.run(["xdg-open", str(carpeta)], check=False)
    agregar_log(f"Carpeta abierta: {carpeta}")
    touch_last_update()


def _crear_manejador_log(mensaje: str) -> Callable[[], None]:
    """Genera un callback que registra ``mensaje`` en el historial."""

    def _handler() -> None:
        agregar_log(mensaje)
        touch_last_update()

    return _handler


REPORT_CARDS = [
    {
        "icon": "bolt",
        "title": "Informe automático",
        "description": "Genera el informe del día anterior en un solo paso.",
        "action_text": "Generar informe automático",
        "handler": _crear_manejador_log("Se inició la generación automática del informe."),
    },
    {
        "icon": "calendar_month",
        "title": "Informe manual",
        "description": "Permite elegir una fecha específica para regenerar el informe.",
        "action_text": "Generar informe manual",
        "handler": _crear_manejador_log("Se inició la generación manual del informe."),
    },
    {
        "icon": "inventory_2",
        "title": "Listado de productos",
        "description": "Descarga y limpia el listado de productos desde SIIGO.",
        "action_text": "Generar listado de productos",
        "handler": _crear_manejador_log("Se inició la generación del listado de productos."),
    },
]


def setup_ui() -> None:
    """Construye la interfaz web siguiendo el estilo solicitado."""

    def limpiar_log() -> None:
        if state.log:
            state.log.clear()
            state.log.classes(add="hidden")
        if state.empty:
            state.empty.classes(remove="hidden")

    with ui.column().classes("max-w-5xl mx-auto py-10 gap-6"):

        with ui.column().classes("gap-2 w-full"):
            with ui.row().classes("items-center gap-2"):
                ui.icon("folder_open").classes("text-gray-600")
                ui.label("Plantilla base").classes("font-medium")

            with ui.row().classes("items-center gap-2 w-full"):
                ui.input(value=RUTA_PLANTILLA) \
                    .props("readonly") \
                    .classes("flex-1 bg-gray-50 rounded-xl p-2 h-10 min-h-0 text-sm")
                ui.button("Copiar", on_click=copiar_ruta)
                ui.button("Abrir carpeta", on_click=abrir_carpeta)

        with ui.row().classes("gap-4 flex-wrap w-full"):
            for card in REPORT_CARDS:
                with ui.card().classes(
                    "rounded-2xl shadow-sm border border-gray-200 bg-white flex-1 min-w-[260px]"
                ):
                    with ui.row().classes("items-center gap-2 px-4 pt-4"):
                        ui.icon(card["icon"]).classes("text-violet-500")
                        ui.label(card["title"]).classes("font-medium")
                    ui.label(card["description"]).classes("px-4 pb-2 text-sm text-gray-500")
                    ui.button(card["action_text"], on_click=card["handler"]) \
                        .classes("mx-4 mb-4 w-full").props("color=primary")

        # --- Registro de Actividades ---
        with ui.card().classes(
            "rounded-2xl shadow-sm border border-gray-200 bg-white mt-6"
        ):
            with ui.row().classes("items-center gap-2 px-4 pt-4"):
                ui.icon("activity").classes("text-violet-500")
                ui.label("Registro de Actividades").classes("font-medium")

                ui.button("Limpiar", icon="delete", on_click=limpiar_log) \
                    .props("flat").classes("ml-auto")

            with ui.element("div").classes("px-4 pb-4"):
                state.empty = ui.column().classes(
                    "items-center justify-center h-56 w-full text-gray-400 bg-gray-50 rounded-xl"
                )
                with state.empty:
                    ui.icon("inbox").classes("text-4xl")
                    ui.label("El registro de actividades aparecerá aquí").classes("text-sm")

                state.log = ui.column().classes("hidden w-full gap-1 mt-3")

        # Footer
        with ui.row().classes("items-center justify-between text-xs text-gray-500 mt-2"):
            with ui.row().classes("items-center gap-1"):
                ui.icon("check_circle").classes("text-emerald-500")
                ui.label("Sistema listo")
            state.last_update = ui.label(
                f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
            )


def agregar_log(msg: str) -> None:
    """Muestra ``msg`` en la tarjeta de registro de actividades."""

    if not state.log or not state.empty:
        return

    state.empty.classes(add="hidden")
    state.log.classes(remove="hidden")
    with state.log:
        ui.label(msg).classes("text-sm text-gray-700")


def touch_last_update() -> None:
    """Actualiza la marca de tiempo del pie de página."""

    if state.last_update:
        state.last_update.text = (
            f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
        )


def main() -> None:  # pragma: no cover - punto de entrada manual
    """Ejecuta la aplicación NiceGUI dentro de una ventana nativa."""

    setup_ui()
    ui.run(
        reload=False,
        native=True,
        show=False,
        title="Rentabilidad",
        window_size=(1280, 720),
    )


__all__ = ["RUTA_PLANTILLA", "setup_ui", "agregar_log", "touch_last_update", "main"]

if __name__ in {"__main__", "__mp_main__"}:  # pragma: no cover
    main()
