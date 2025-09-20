"""Interfaz web basada en NiceGUI para el panel de rentabilidad."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Callable

from nicegui import ui
from nicegui.element import Element

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContextFactory


load_env()

_CONTEXT = PathContextFactory(os.environ).create()
RUTA_PLANTILLA = str(_CONTEXT.template_path())


empty_state: Element | None = None
log_list: Element | None = None


def _copiar_ruta() -> None:
    """Copia ``RUTA_PLANTILLA`` al portapapeles usando JavaScript."""

    ui.run_javascript(f"navigator.clipboard.writeText({json.dumps(RUTA_PLANTILLA)})")
    agregar_log("Ruta de la plantilla copiada al portapapeles.")


def _abrir_carpeta() -> None:
    """Intenta abrir la carpeta que contiene la plantilla base."""

    carpeta = Path(RUTA_PLANTILLA).parent
    if not carpeta.exists():
        agregar_log(f"No se encontró la carpeta: {carpeta}")
        return

    if sys.platform.startswith("win"):
        os.startfile(carpeta)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(carpeta)], check=False)
    else:
        subprocess.run(["xdg-open", str(carpeta)], check=False)
    agregar_log(f"Carpeta abierta: {carpeta}")


def _crear_manejador_log(mensaje: str) -> Callable[[], None]:
    """Genera un callback que registra ``mensaje`` en el historial."""

    def _handler() -> None:
        agregar_log(mensaje)

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

    global empty_state, log_list

    slots: dict[str, Element] = {}

    def limpiar_log() -> None:
        lista = slots.get("log_list")
        vacio = slots.get("empty_state")
        if not lista or not vacio:
            return
        lista.clear()
        lista.add_class("hidden")
        vacio.remove_class("hidden")

    with ui.column().classes("max-w-5xl mx-auto py-10 gap-6"):
        with ui.card().classes("rounded-2xl shadow-sm border border-gray-200 bg-white"):
            with ui.row().classes("items-center gap-2 px-4 pt-4"):
                ui.icon("description").classes("text-blue-500")
                ui.label("Plantilla base").classes("font-medium text-slate-700")
                ui.space()
                ui.button("Copiar", icon="content_copy", on_click=_copiar_ruta) \
                    .props("outline").classes("text-sm")
                ui.button("Abrir carpeta", icon="folder_open", on_click=_abrir_carpeta) \
                    .props("outline").classes("text-sm")
            with ui.row().classes("items-center gap-3 px-4 pb-4 w-full"):
                ui.input(value=RUTA_PLANTILLA, readonly=True) \
                    .classes("flex-1 bg-blue-50 rounded-xl p-2 h-10 min-h-0 text-sm")

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

        with ui.card().classes(
            "rounded-2xl shadow-sm border border-gray-200 bg-white mt-6"
        ):
            with ui.row().classes("items-center gap-2 px-4 pt-4"):
                ui.icon("activity").classes("text-violet-500")
                ui.label("Registro de Actividades").classes("font-medium")
                ui.button("Limpiar", icon="delete", on_click=limpiar_log) \
                    .props("flat").classes("ml-auto")

            with ui.element("div").classes("px-4 pb-4"):
                empty = ui.column().classes(
                    "items-center justify-center h-56 w-full text-gray-400 bg-gray-50 rounded-xl"
                )
                with empty:
                    ui.icon("inbox").classes("text-4xl")
                    ui.label("El registro de actividades aparecerá aquí").classes("text-sm")

                logs = ui.column().classes("hidden w-full gap-1 mt-3")

    slots["empty_state"] = empty
    slots["log_list"] = logs

    empty_state = empty
    log_list = logs


def agregar_log(msg: str) -> None:
    """Muestra ``msg`` en la tarjeta de registro de actividades."""

    if not log_list or not empty_state:
        return

    empty_state.add_class("hidden")
    log_list.remove_class("hidden")
    with log_list:
        ui.label(msg).classes("text-sm text-gray-700")


def main() -> None:  # pragma: no cover - punto de entrada manual
    """Ejecuta la aplicación NiceGUI."""

    setup_ui()
    ui.run()


__all__ = ["RUTA_PLANTILLA", "setup_ui", "agregar_log", "main"]


if __name__ == "__main__":  # pragma: no cover
    main()
