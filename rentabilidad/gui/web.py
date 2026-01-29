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

from rentabilidad.app.dto import (
    GenerarConsolidadoMalosCobrosRequest,
    GenerarInformeCodigosIncorrectosRequest,
    GenerarInformeRequest,
)
from rentabilidad.app.use_cases.ejecutar_productos_script import (
    run as uc_productos,
)
from rentabilidad.app.use_cases.generar_consolidado_malos_cobros import (
    run as uc_malos_cobros,
)
from rentabilidad.app.use_cases.generar_informe_automatico import run as uc_auto
from rentabilidad.app.use_cases.generar_informe_codigos_incorrectos import (
    run as uc_codigos_incorrectos,
)
from rentabilidad.app.use_cases.listar_meses_informes import run as uc_listar_meses
from rentabilidad.config import bus, settings
from rentabilidad.infra.fs import ayer_str

state = SimpleNamespace(
    empty=None,
    log=None,
    last_update=None,
    status=None,
    status_icon=None,
    status_icon_classes=tuple(),
    status_action=None,
    status_target=None,
    progress=None,
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


def _set_status_action(path: Path | None) -> None:
    state.status_target = path
    if not state.status_action:
        return

    if path:
        state.status_action.enable()
        state.status_action.classes(remove="hidden")
    else:
        state.status_action.disable()
        state.status_action.classes(add="hidden")


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


def _abrir_estado_destino() -> None:
    destino = state.status_target
    if not destino:
        return

    _abrir_archivo(destino)


def actualizar_estado(kind: str, mensaje: str) -> None:
    config = STATUS_UI.get(kind, STATUS_UI["idle"])
    if state.status_icon:
        for cls in state.status_icon_classes:
            state.status_icon.classes(remove=cls)
        state.status_icon.set_name(config["icon"])
        for cls in config["classes"]:
            state.status_icon.classes(add=cls)
        state.status_icon_classes = config["classes"]
    if state.status:
        state.status.text = mensaje
    if kind != "success":
        _set_status_action(None)
    if state.progress:
        if kind == "running":
            state.progress.classes(remove="hidden")
        else:
            state.progress.classes(add="hidden")


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
        touch_last_update()
        return False


def abrir_carpeta() -> None:
    carpeta = settings.ruta_plantilla.parent
    if not carpeta.exists():
        mensaje = f"No se encontró la carpeta: {carpeta}"
        agregar_log(mensaje, "error")
        actualizar_estado("error", "Ruta de la plantilla no encontrada")
        touch_last_update()
        return

    if _abrir_en_sistema(carpeta, "la carpeta"):
        agregar_log(f"Carpeta abierta: {carpeta}")
        touch_last_update()


def _abrir_archivo(path: Path) -> bool:
    if not path.exists():
        mensaje = f"El archivo no existe: {path}"
        agregar_log(mensaje, "error")
        touch_last_update()
        return False

    if _abrir_en_sistema(path, "el archivo"):
        agregar_log(f"Archivo abierto: {path}", "success")
        touch_last_update()
        return True

    return False


def _extraer_ruta_resultado(msg: str) -> Path | None:
    if ":" not in msg:
        return None

    posible = msg.split(":", 1)[1].strip()
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
        ruta_resultado = _extraer_ruta_resultado(msg)
        if ruta_resultado:
            detalle = f"Archivo generado: {ruta_resultado}"
        else:
            detalle = msg

        mensaje = f"Proceso finalizó correctamente. {detalle}".strip()
        agregar_log(mensaje, "success")
        actualizar_estado("success", "Proceso completado")
        _set_status_action(ruta_resultado)
        touch_last_update()

    def _on_error(msg: str) -> None:
        detalle = msg.strip()
        mensaje = (
            f"El proceso no finalizó correctamente. Detalle: {detalle}"
            if detalle
            else "El proceso no finalizó correctamente."
        )
        agregar_log(mensaje, "error")
        actualizar_estado("error", "Revisa los registros")
        _set_status_action(None)
        touch_last_update()

    bus.subscribe("log", _on_log)
    bus.subscribe("done", _on_done)
    bus.subscribe("error", _on_error)
    _bus_registered = True


def setup_ui() -> None:
    """Construye la interfaz web siguiendo el estilo solicitado."""

    _register_static_files()
    logo_url = f"/static/{LOGO_FILE.name}" if LOGO_FILE.exists() else None
    month_options = uc_listar_meses()
    default_month = month_options[-1] if month_options else None
    month_select = None
    manual_date_input = None

    def ejecutar_auto(usar_sql: bool | None) -> None:
        modo = "SQL" if usar_sql else "EXCZ"
        actualizar_estado("running", f"Generando informe automático ({modo})…")
        agregar_log(f"Iniciando generación automática del informe ({modo}).")
        uc_auto(
            GenerarInformeRequest(
                ruta_plantilla=str(settings.ruta_plantilla),
                usar_sql=usar_sql,
            ),
            bus,
        )

    def ejecutar_manual(usar_sql: bool | None) -> None:
        fecha_texto = (
            (manual_date_input.value or "").strip() if manual_date_input else ""
        )
        if not fecha_texto:
            agregar_log("Debes seleccionar una fecha válida.", "error")
            actualizar_estado("error", "Selecciona una fecha")
            return
        try:
            datetime.strptime(fecha_texto, "%Y-%m-%d")
        except ValueError:
            agregar_log("La fecha debe tener el formato AAAA-MM-DD.", "error")
            actualizar_estado("error", "Fecha inválida")
            return
        modo = "SQL" if usar_sql else "EXCZ"
        actualizar_estado(
            "running",
            f"Generando informe manual ({fecha_texto}, {modo})…",
        )
        agregar_log(
            f"Iniciando generación manual del informe para {fecha_texto} ({modo})."
        )
        uc_auto(
            GenerarInformeRequest(
                ruta_plantilla=str(settings.ruta_plantilla),
                fecha=fecha_texto,
                usar_sql=usar_sql,
            ),
            bus,
        )

    def ejecutar_productos() -> None:
        actualizar_estado("running", "Generando listado de productos…")
        agregar_log(
            "Iniciando generación del listado de productos (Productos.bat)."
        )
        uc_productos(bus)

    def ejecutar_codigos() -> None:
        mes = (month_select.value or "").strip() if month_select else ""
        if not mes:
            agregar_log("Debes seleccionar un mes disponible.", "error")
            actualizar_estado("error", "Selecciona un mes válido")
            return
        actualizar_estado(
            "running", f"Generando informe de códigos incorrectos ({mes})…"
        )
        agregar_log(
            f"Iniciando generación del informe de códigos incorrectos para {mes}."
        )
        uc_codigos_incorrectos(
            GenerarInformeCodigosIncorrectosRequest(mes=mes),
            bus,
        )

    def ejecutar_cobros() -> None:
        mes = (month_select.value or "").strip() if month_select else ""
        if not mes:
            agregar_log("Debes seleccionar un mes disponible.", "error")
            actualizar_estado("error", "Selecciona un mes válido")
            return
        actualizar_estado(
            "running", f"Generando consolidado de malos cobros ({mes})…"
        )
        agregar_log(
            f"Iniciando consolidado de malos cobros para {mes}."
        )
        uc_malos_cobros(
            GenerarConsolidadoMalosCobrosRequest(mes=mes),
            bus,
        )

    with ui.column().classes("max-w-5xl mx-auto py-10 space-y-6"):
        with ui.row().classes("items-center gap-3 w-full"):
            if logo_url:
                ui.image(logo_url).classes("h-12 w-auto object-contain")
            ui.label("Rentabilidad").classes("text-2xl font-semibold text-gray-800")

        with ui.card().classes(
            "rounded-2xl shadow-sm border border-gray-200 bg-white w-full"
        ):
            with ui.row().classes("items-center gap-2 px-5 pt-4"):
                ui.icon("folder_open").classes("text-gray-600")
                ui.label("Plantilla base").classes("font-medium")

            with ui.row().classes(
                "items-center gap-3 px-5 pb-4 w-full flex-wrap md:flex-nowrap"
            ):
                ruta_input = ui.input(value=str(settings.ruta_plantilla))
                ruta_input.props("readonly")
                ruta_input.classes(
                    "grow bg-gray-50 rounded-xl px-3 h-10 min-h-0 text-sm"
                )
                ui.button("Copiar", on_click=copiar_ruta).classes("w-full md:w-auto")
                ui.button("Abrir carpeta", on_click=abrir_carpeta).classes(
                    "w-full md:w-auto"
                )

        def _card_container():
            return ui.card().classes(
                "h-full rounded-2xl shadow-sm border border-gray-200 bg-white flex flex-col"
            )

        with ui.element("div").classes(
            "grid grid-cols-1 gap-4 w-full md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-5"
        ):
            with _card_container():
                with ui.row().classes("items-center gap-2 px-5 pt-4"):
                    ui.icon("bolt").classes("text-violet-500")
                    ui.label("Informe automático").classes("font-medium")
                ui.label(
                    "Genera el informe del día anterior desde SQL Server con respaldo EXCZ."
                ).classes("px-5 pb-3 text-sm text-gray-500 leading-snug")
                btn_auto = ui.button(
                    "Generar informe automático (SQL)",
                    on_click=lambda: ejecutar_auto(True),
                )
                btn_auto.classes("mx-5 mb-2 w-full")
                btn_auto.props("color=primary")
                btn_auto_backup = ui.button(
                    "Generar informe automático (EXCZ respaldo)",
                    on_click=lambda: ejecutar_auto(False),
                )
                btn_auto_backup.classes("mx-5 mb-2 w-full")
                nota_auto = ui.label(
                    "Flujo principal: SQL Server. Respaldo EXCZ en: "
                    f"{_shorten(settings.excz_dir)}"
                ).classes("px-5 pb-5 text-xs text-gray-400")
                with nota_auto:
                    ui.tooltip(str(settings.excz_dir))

            with _card_container():
                with ui.row().classes("items-center gap-2 px-5 pt-4"):
                    ui.icon("calendar_month").classes("text-violet-500")
                    ui.label("Informe manual").classes("font-medium")
                ui.label(
                    "Selecciona una fecha para generar el informe desde SQL Server "
                    "o con respaldo EXCZ."
                ).classes("px-5 pb-3 text-sm text-gray-500 leading-snug")
                manual_date_input = ui.input(
                    label="Fecha objetivo",
                    value=ayer_str(),
                )
                manual_date_input.props("type=date outlined")
                manual_date_input.classes("mx-5 mb-2 w-full")
                btn_manual = ui.button(
                    "Generar informe manual (SQL)",
                    on_click=lambda: ejecutar_manual(True),
                )
                btn_manual.classes("mx-5 mb-2 w-full")
                btn_manual.props("color=primary")
                btn_manual_backup = ui.button(
                    "Generar informe manual (EXCZ respaldo)",
                    on_click=lambda: ejecutar_manual(False),
                )
                btn_manual_backup.classes("mx-5 mb-2 w-full")
                ui.label(
                    "Se recomienda SQL como fuente principal. Usa EXCZ si SQL no está disponible."
                ).classes("px-5 pb-5 text-xs text-gray-400")

            with _card_container():
                with ui.row().classes("items-center gap-2 px-5 pt-4"):
                    ui.icon("inventory_2").classes("text-violet-500")
                    ui.label("Generar productos").classes("font-medium")
                ui.label(
                    "Ejecuta el script Productos.bat para actualizar el listado de productos."
                ).classes("px-5 pb-3 text-sm text-gray-500 leading-snug")
                productos_script = settings.productos_batch_script
                script_text = (
                    _shorten(productos_script) if productos_script else "No configurado"
                )
                script_label = ui.label(f"Script: {script_text}")
                script_label.classes("px-5 pb-3 text-xs text-gray-400")
                if productos_script:
                    with script_label:
                        ui.tooltip(str(productos_script))
                btn_productos = ui.button(
                    "Generar productos", on_click=ejecutar_productos
                )
                btn_productos.classes("mx-5 mb-2 w-full")
                btn_productos.props("color=primary")
                if not productos_script:
                    btn_productos.disable()
                    ui.label(
                        "Configura la ruta del script Productos.bat antes de ejecutar esta acción."
                    ).classes("px-5 pb-5 text-xs text-amber-500")
                else:
                    ui.label(
                        "El resultado se guardará en la carpeta de productos configurada."
                    ).classes("px-5 pb-5 text-xs text-gray-400")

            with _card_container():
                with ui.row().classes("items-center gap-2 px-5 pt-4"):
                    ui.icon("insights").classes("text-violet-500")
                    ui.label("Informes mensuales").classes("font-medium")
                ui.label(
                    "Genera consolidaciones mensuales desde los informes existentes."
                ).classes("px-5 pb-3 text-sm text-gray-500 leading-snug")
                month_select = ui.select(
                    options=month_options,
                    value=default_month,
                    label="Mes",
                )
                month_select.props("outlined")
                month_select.classes("mx-5 mb-2 w-full text-sm")
                btn_codigos = ui.button(
                    "Informe códigos incorrectos",
                    on_click=ejecutar_codigos,
                )
                btn_codigos.classes("mx-5 mb-2 w-full")
                btn_codigos.props("color=primary")
                btn_cobros = ui.button(
                    "Consolidado malos cobros",
                    on_click=ejecutar_cobros,
                )
                btn_cobros.classes("mx-5 mb-2 w-full")
                nota_meses = ui.label(
                    "Los resultados se guardarán en las carpetas de consolidados configuradas."
                ).classes("px-5 pb-5 text-xs text-gray-400")
                if not month_options:
                    btn_codigos.disable()
                    btn_cobros.disable()
                    nota_meses.text = (
                        "No se encontraron carpetas de meses en la ruta de informes."
                    )

            with _card_container():
                with ui.column().classes(
                    "px-5 py-5 items-center text-center gap-3 flex-1"
                ):
                    ui.icon("route").classes("text-violet-500 text-3xl")
                    ui.label("Rutas de trabajo").classes("font-medium")
                    ui.label("Pronto…").classes(
                        "text-lg font-semibold text-gray-500"
                    )
                    ui.label(
                        "Estamos preparando nuevas herramientas para administrar rutas de trabajo."
                    ).classes("text-sm text-gray-500 leading-snug")

        def _path_entry(nombre: str, ruta: Path) -> None:
            with ui.row().classes(
                "items-center gap-2 px-5 text-sm text-gray-600 flex-wrap"
            ):
                ui.label(f"{nombre}:").classes("font-semibold text-gray-700")
                valor = ui.label(_shorten(ruta)).classes("text-gray-600")
                with valor:
                    ui.tooltip(str(ruta))

        with ui.card().classes(
            "rounded-2xl shadow-sm border border-gray-200 bg-white mt-4 w-full"
        ):
            with ui.row().classes("items-center gap-2 px-5 pt-4"):
                ui.icon("map").classes("text-violet-500")
                ui.label("Ubicaciones configuradas").classes("font-medium")
            with ui.column().classes("px-5 pb-4 gap-2"):
                _path_entry("Informes", settings.context.informes_dir)
                _path_entry("Productos", settings.context.productos_dir)
                _path_entry("Plantilla", settings.ruta_plantilla)
                ui.label(
                    "Puedes modificar estas rutas mediante variables de entorno."
                ).classes("text-xs text-gray-400")

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

        with ui.column().classes(
            "mt-2 w-full gap-2 text-xs text-gray-500"
        ):
            with ui.row().classes("items-center justify-between w-full"):
                with ui.row().classes("items-center gap-1"):
                    state.status_icon = ui.icon("check_circle").classes(
                        "text-emerald-500"
                    )
                    state.status = ui.label("Sistema listo")
                state.status_action = ui.button(
                    "Abrir informe", on_click=_abrir_estado_destino
                ).props("flat color=primary")
                state.status_action.classes("text-xs hidden")
                state.status_action.disable()
                state.last_update = ui.label("Última actualización: —")
            state.progress = (
                ui.linear_progress()
                .props("color=primary indeterminate")
                .classes("hidden w-full")
            )

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
