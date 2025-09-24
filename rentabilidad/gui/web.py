"""Interfaz web basada en NiceGUI para el panel de rentabilidad."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from datetime import datetime
from types import SimpleNamespace

from nicegui import app, ui

from rentabilidad.app.dto import GenerarInformeRequest
from rentabilidad.app.use_cases.generar_informe_automatico import run as uc_auto
from rentabilidad.app.use_cases.generar_informe_manual import run as uc_manual
from rentabilidad.app.use_cases.listar_productos import run as uc_listado
from rentabilidad.config import bus, settings
from rentabilidad.infra.fs import ayer_str

state = SimpleNamespace(
    empty=None,
    log=None,
    last_update=None,
    status=None,
    status_icon=None,
    status_icon_classes=tuple(),
)

RUTA_PLANTILLA = str(settings.ruta_plantilla)
STATIC_DIR = Path(__file__).with_name("static")
LOGO_FILE = STATIC_DIR / "logo.svg"

STATUS_UI = {
    "idle": {"icon": "check_circle", "classes": ("text-emerald-500",)},
    "running": {"icon": "autorenew", "classes": ("text-sky-500", "animate-spin")},
    "success": {"icon": "check_circle", "classes": ("text-emerald-500",)},
    "error": {"icon": "error", "classes": ("text-rose-500",)},
}

LOG_STYLES = {
    "info": {
        "icon": "info",
        "icon_class": "text-sky-500",
        "text_class": "text-gray-700",
    },
    "success": {
        "icon": "check_circle",
        "icon_class": "text-emerald-500",
        "text_class": "text-emerald-600",
    },
    "error": {
        "icon": "error",
        "icon_class": "text-rose-500",
        "text_class": "text-rose-600",
    },
}

_bus_registered = False
_api_registered = False
_static_registered = False


def _register_static_files() -> None:
    global _static_registered
    if _static_registered or not STATIC_DIR.exists():
        return

    app.add_static_files("/static", str(STATIC_DIR))
    _static_registered = True


def _build_open_action(ruta: Path) -> dict[str, str]:
    ruta_serializada = json.dumps(str(ruta))
    return {
        "label": "Abrir informe",
        "color": "white",
        ":handler": (
            "async () => {"
            "  await fetch('/api/abrir-informe?ruta=' + encodeURIComponent("
            + ruta_serializada
            + "));"
            "}"
        ),
    }


def _register_api_routes() -> None:
    global _api_registered
    if _api_registered:
        return

    @app.get("/api/abrir-informe")
    async def abrir_informe_api(ruta: str) -> dict[str, str]:
        path = Path(ruta)
        ok = _abrir_archivo(path)
        return {
            "status": "ok" if ok else "error",
            "detail": "" if ok else f"No se pudo abrir el archivo: {path}",
        }

    _api_registered = True


def _shorten(text: str, length: int = 60) -> str:
    clean = str(text)
    return clean if len(clean) <= length else clean[: length - 1] + "…"


def actualizar_estado(kind: str, mensaje: str) -> None:
    config = STATUS_UI.get(kind, STATUS_UI["idle"])
    if state.status_icon:
        for cls in state.status_icon_classes:
            state.status_icon.classes(remove=cls)
        state.status_icon.set_text(config["icon"])
        for cls in config["classes"]:
            state.status_icon.classes(add=cls)
        state.status_icon_classes = config["classes"]
    if state.status:
        state.status.text = mensaje


def agregar_log(msg: str, kind: str = "info") -> None:
    if not state.log or not state.empty:
        return

    state.empty.classes(add="hidden")
    state.log.classes(remove="hidden")

    config = LOG_STYLES.get(kind, LOG_STYLES["info"])
    with state.log:
        with ui.row().classes(
            "items-start gap-2 bg-gray-50 rounded-xl px-3 py-2 border border-gray-200"
        ):
            ui.icon(config["icon"]).classes(f"text-base {config['icon_class']}")
            ui.label(msg).classes(f"text-sm leading-snug {config['text_class']}")


def touch_last_update() -> None:
    if state.last_update:
        state.last_update.text = (
            f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
        )


def limpiar_log() -> None:
    if state.log:
        state.log.clear()
        state.log.classes(add="hidden")
    if state.empty:
        state.empty.classes(remove="hidden")
    if state.last_update:
        state.last_update.text = "Última actualización: —"
    actualizar_estado("idle", "Sistema listo")


def copiar_ruta() -> None:
    ruta = str(settings.ruta_plantilla)
    ui.run_javascript(f"navigator.clipboard.writeText({json.dumps(ruta)})")
    agregar_log("Ruta de la plantilla copiada al portapapeles.")
    ui.notify("Ruta copiada al portapapeles.", type="positive", position="top")
    touch_last_update()


def _abrir_en_sistema(destino: Path, descripcion: str) -> bool:
    try:
        if sys.platform.startswith("win"):
            os.startfile(destino)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(destino)], check=False)
        else:
            subprocess.run(["xdg-open", str(destino)], check=False)
        return True
    except Exception as exc:  # pragma: no cover - defensivo
        mensaje = f"No se pudo abrir {descripcion}: {exc}"
        agregar_log(mensaje, "error")
        ui.notify(mensaje, type="negative", position="top")
        touch_last_update()
        return False


def abrir_carpeta() -> None:
    carpeta = settings.ruta_plantilla.parent
    if not carpeta.exists():
        mensaje = f"No se encontró la carpeta: {carpeta}"
        agregar_log(mensaje, "error")
        actualizar_estado("error", "Ruta de la plantilla no encontrada")
        ui.notify(mensaje, type="negative", position="top")
        touch_last_update()
        return

    if _abrir_en_sistema(carpeta, "la carpeta"):
        agregar_log(f"Carpeta abierta: {carpeta}")
        touch_last_update()


def _abrir_archivo(path: Path) -> bool:
    if not path.exists():
        mensaje = f"El archivo no existe: {path}"
        agregar_log(mensaje, "error")
        ui.notify(mensaje, type="negative", position="top")
        touch_last_update()
        return False

    if _abrir_en_sistema(path, "el archivo"):
        agregar_log(f"Archivo abierto: {path}", "success")
        touch_last_update()
        return True

    return False


def _extraer_ruta_informe(msg: str) -> Path | None:
    prefijo = "Informe generado:"
    if not msg.startswith(prefijo):
        return None

    posible = msg[len(prefijo) :].strip()
    if not posible:
        return None

    ruta = Path(posible)
    return ruta if ruta.exists() else None


def _register_bus_handlers() -> None:
    global _bus_registered
    if _bus_registered:
        return

    def _on_log(msg: str) -> None:
        agregar_log(msg, "info")
        touch_last_update()

    def _on_done(msg: str) -> None:
        agregar_log(msg, "success")
        actualizar_estado("success", "Proceso completado")
        acciones = []
        ruta_informe = _extraer_ruta_informe(msg)
        if ruta_informe:
            acciones.append(_build_open_action(ruta_informe))
        ui.notify(msg, type="positive", position="top", actions=acciones or None)
        touch_last_update()

    def _on_error(msg: str) -> None:
        agregar_log(msg, "error")
        actualizar_estado("error", "Revisa los registros")
        ui.notify(msg, type="negative", position="top")
        touch_last_update()

    bus.subscribe("log", _on_log)
    bus.subscribe("done", _on_done)
    bus.subscribe("error", _on_error)
    _bus_registered = True


def setup_ui() -> None:
    """Construye la interfaz web siguiendo el estilo solicitado."""

    _register_static_files()
    logo_url = f"/static/{LOGO_FILE.name}" if LOGO_FILE.exists() else None

    def ejecutar_auto() -> None:
        actualizar_estado("running", "Generando informe automático…")
        agregar_log("Iniciando generación automática del informe.")
        uc_auto(
            GenerarInformeRequest(ruta_plantilla=str(settings.ruta_plantilla)),
            bus,
        )

    def ejecutar_manual() -> None:
        fecha = (fecha_input.value or "").strip() or None
        objetivo = fecha or "día anterior"
        actualizar_estado("running", f"Generando informe para {objetivo}…")
        agregar_log(f"Iniciando generación manual del informe ({objetivo}).")
        uc_manual(
            GenerarInformeRequest(
                ruta_plantilla=str(settings.ruta_plantilla),
                fecha=fecha,
            ),
            bus,
        )

    def ejecutar_listado() -> None:
        actualizar_estado("running", "Generando listado de productos…")
        agregar_log("Iniciando generación del listado de productos.")
        uc_listado(bus)

    with ui.column().classes("max-w-5xl mx-auto py-10 gap-6"):
        with ui.row().classes("items-center gap-4 w-full"):
            if logo_url:
                ui.image(logo_url).classes("h-12 w-auto object-contain")
            ui.label("Rentabilidad").classes("text-2xl font-semibold text-gray-800")

        with ui.column().classes("gap-2 w-full"):
            with ui.row().classes("items-center gap-2"):
                ui.icon("folder_open").classes("text-gray-600")
                ui.label("Plantilla base").classes("font-medium")

            with ui.row().classes("items-center gap-2 w-full"):
                ruta_input = ui.input(value=str(settings.ruta_plantilla))
                ruta_input.props("readonly")
                ruta_input.classes(
                    "flex-1 bg-gray-50 rounded-xl p-2 h-10 min-h-0 text-sm"
                )
                ui.button("Copiar", on_click=copiar_ruta)
                ui.button("Abrir carpeta", on_click=abrir_carpeta)

        with ui.row().classes("gap-4 flex-wrap w-full"):
            with ui.card().classes(
                "rounded-2xl shadow-sm border border-gray-200 bg-white flex-1 min-w-[260px]"
            ):
                with ui.row().classes("items-center gap-2 px-4 pt-4"):
                    ui.icon("bolt").classes("text-violet-500")
                    ui.label("Informe automático").classes("font-medium")
                ui.label(
                    "Genera el informe del día anterior usando el EXCZ más reciente disponible."
                ).classes("px-4 pb-2 text-sm text-gray-500")
                btn_auto = ui.button(
                    "Generar informe automático", on_click=ejecutar_auto
                )
                btn_auto.classes("mx-4 mb-2 w-full")
                btn_auto.props("color=primary")
                nota_auto = ui.label(
                    f"Prefijo EXCZ: {settings.excz_prefix} · Carpeta: {_shorten(settings.excz_dir)}"
                ).classes("px-4 pb-4 text-xs text-gray-400")
                with nota_auto:
                    ui.tooltip(str(settings.excz_dir))

            with ui.card().classes(
                "rounded-2xl shadow-sm border border-gray-200 bg-white flex-1 min-w-[260px]"
            ):
                with ui.row().classes("items-center gap-2 px-4 pt-4"):
                    ui.icon("calendar_month").classes("text-violet-500")
                    ui.label("Informe manual").classes("font-medium")
                ui.label(
                    "Permite elegir una fecha específica para regenerar el informe."
                ).classes("px-4 pb-2 text-sm text-gray-500")
                fecha_input = ui.input(
                    label="Fecha objetivo",
                    value=ayer_str(),
                )
                fecha_input.props("type=date")
                fecha_input.classes(
                    "mx-4 mb-2 w-full rounded-xl border border-gray-200 px-3 py-2 text-sm"
                )
                btn_manual = ui.button(
                    "Generar informe manual", on_click=ejecutar_manual
                )
                btn_manual.classes("mx-4 mb-2 w-full")
                btn_manual.props("color=primary")
                ui.label(
                    "Deja la fecha vacía para utilizar el día anterior de forma automática."
                ).classes("px-4 pb-4 text-xs text-gray-400")

            with ui.card().classes(
                "rounded-2xl shadow-sm border border-gray-200 bg-white flex-1 min-w-[260px]"
            ):
                with ui.row().classes("items-center gap-2 px-4 pt-4"):
                    ui.icon("inventory_2").classes("text-violet-500")
                    ui.label("Listado de productos").classes("font-medium")
                ui.label(
                    "Descarga y limpia el catálogo directamente desde SIIGO."
                ).classes("px-4 pb-2 text-sm text-gray-500")
                btn_listado = ui.button(
                    "Generar listado de productos", on_click=ejecutar_listado
                )
                btn_listado.classes("mx-4 mb-2 w-full")
                btn_listado.props("color=primary")
                nota_productos = ui.label(
                    f"Destino: {_shorten(settings.context.productos_dir)}"
                ).classes("px-4 pb-4 text-xs text-gray-400")
                with nota_productos:
                    ui.tooltip(str(settings.context.productos_dir))

        with ui.card().classes(
            "rounded-2xl shadow-sm border border-gray-200 bg-white mt-6"
        ):
            with ui.row().classes("items-center gap-2 px-4 pt-4"):
                ui.icon("activity").classes("text-violet-500")
                ui.label("Registro de Actividades").classes("font-medium")
                ui.button("Limpiar", icon="delete", on_click=limpiar_log).props(
                    "flat"
                ).classes("ml-auto")

            with ui.element("div").classes("px-4 pb-4"):
                state.empty = ui.column().classes(
                    "items-center justify-center h-56 w-full text-gray-400 bg-gray-50 rounded-xl gap-2"
                )
                with state.empty:
                    ui.icon("inbox").classes("text-4xl")
                    ui.label(
                        "El registro de actividades aparecerá aquí"
                    ).classes("text-sm")

                state.log = ui.column().classes("hidden w-full gap-2 mt-3")

        with ui.row().classes(
            "items-center justify-between text-xs text-gray-500 mt-2"
        ):
            with ui.row().classes("items-center gap-1"):
                state.status_icon = ui.icon("check_circle").classes("text-emerald-500")
                state.status = ui.label("Sistema listo")
            state.last_update = ui.label("Última actualización: —")

    _register_api_routes()
    _register_bus_handlers()
    actualizar_estado("idle", "Sistema listo")


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
