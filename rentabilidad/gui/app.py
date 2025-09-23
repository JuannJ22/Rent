from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

from nicegui import ui

from rentabilidad.app.dto import GenerarInformeRequest
from rentabilidad.app.use_cases.generar_informe_automatico import run as uc_auto
from rentabilidad.app.use_cases.generar_informe_manual import run as uc_manual
from rentabilidad.app.use_cases.listar_productos import run as uc_listado
from rentabilidad.config import bus, settings
from rentabilidad.infra.fs import ayer_str

state = SimpleNamespace(empty=None, log=None, last_update=None, status=None)

_subscriptions_registered = False

LOG_ICONS = {
    "info": "info",
    "success": "check_circle",
    "error": "error",
}

LOG_ICON_CLASSES = {
    "info": "text-sky-500",
    "success": "text-emerald-500",
    "error": "text-rose-500",
}

LOG_ENTRY_CLASSES = {
    "info": "log-info",
    "success": "log-success",
    "error": "log-error",
}


def _status_markup(kind: str, text: str) -> str:
    return (
        f'<div class="status-chip status-{kind}">'  # noqa: E501
        "<span class=\"status-dot\"></span>"
        f"<span>{text}</span>"
        "</div>"
    )


def update_status(kind: str, text: str) -> None:
    if state.status is None:
        return
    state.status.content = _status_markup(kind, text)


def _shorten(text: str, length: int = 42) -> str:
    clean = str(text)
    return clean if len(clean) <= length else clean[: length - 1] + "…"


def agregar_log(msg: str, kind: str = "info") -> None:
    if state.empty is None or state.log is None:
        return

    state.empty.classes(add="hidden")
    state.log.classes(remove="hidden")

    css_class = LOG_ENTRY_CLASSES.get(kind, "log-info")
    icon = LOG_ICONS.get(kind, "info")
    icon_class = LOG_ICON_CLASSES.get(kind, "text-slate-500")

    with state.log:
        with ui.row().classes(f"log-entry {css_class}"):
            ui.icon(icon).classes(f"log-entry-icon {icon_class}")
            ui.label(msg).classes("log-entry-text")


def touch_last_update() -> None:
    if state.last_update is None:
        return
    state.last_update.text = (
        f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
    )


def limpiar_log() -> None:
    if state.log is not None:
        state.log.clear()
        state.log.classes(add="hidden")
    if state.empty is not None:
        state.empty.classes(remove="hidden")
    if state.last_update is not None:
        state.last_update.text = "Última actualización: —"
    update_status("idle", "Sistema listo")


def copiar_ruta() -> None:
    ruta = str(settings.ruta_plantilla)
    ui.run_javascript(f"navigator.clipboard.writeText({json.dumps(ruta)})")
    agregar_log("Ruta de la plantilla copiada al portapapeles.")


def abrir_carpeta() -> None:
    carpeta = Path(settings.ruta_plantilla).parent
    if not carpeta.exists():
        mensaje = f"No se encontró la carpeta: {carpeta}"
        agregar_log(mensaje, "error")
        update_status("error", "Ruta de la plantilla no encontrada")
        ui.notify(mensaje, type="negative", position="top")
        return

    try:
        if sys.platform.startswith("win"):
            os.startfile(carpeta)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(carpeta)], check=False)
        else:
            subprocess.run(["xdg-open", str(carpeta)], check=False)
        agregar_log(f"Carpeta abierta: {carpeta}")
    except Exception as exc:  # pragma: no cover - defensivo
        mensaje = f"No se pudo abrir la carpeta: {exc}"
        agregar_log(mensaje, "error")
        update_status("error", "No fue posible abrir la carpeta")
        ui.notify(mensaje, type="negative", position="top")


def _path_line(label: str, value: Path) -> None:
    with ui.row().classes("path-line w-full flex-wrap"):
        ui.icon("chevron_right").classes("path-line-icon")
        ui.label(f"{label}:").classes("path-line-label")
        shortened = _shorten(value)
        component = ui.label(shortened).classes("path-line-value")
        with component:
            ui.tooltip(str(value))


def _register_bus_subscriptions() -> None:
    global _subscriptions_registered
    if _subscriptions_registered:
        return

    def _on_log(msg: str) -> None:
        agregar_log(msg, "info")

    def _on_done(msg: str) -> None:
        agregar_log(msg, "success")
        touch_last_update()
        update_status("success", "Proceso completado")
        ui.notify(msg, type="positive", position="top")

    def _on_error(msg: str) -> None:
        agregar_log(msg, "error")
        update_status("error", "Revisa los registros")
        ui.notify(msg, type="negative", position="top")

    bus.subscribe("log", _on_log)
    bus.subscribe("done", _on_done)
    bus.subscribe("error", _on_error)

    _subscriptions_registered = True


def build_ui() -> None:
    ui.add_head_html(
        """
<style>
  :root {
    color-scheme: light;
    font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif;
  }
  body {
    background: #f5f7fb;
    color: #0f172a;
  }
  .hero-gradient {
    background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 45%, #38bdf8 100%);
    border-bottom-left-radius: 48px;
    border-bottom-right-radius: 48px;
    box-shadow: 0 28px 70px rgba(15, 23, 42, 0.25);
  }
  .hero-title {
    font-size: 2.75rem;
    font-weight: 700;
    line-height: 1.12;
    letter-spacing: -0.02em;
  }
  .hero-subtitle {
    font-size: 1.05rem;
    max-width: 640px;
    color: rgba(255, 255, 255, 0.78);
  }
  .quick-card {
    min-width: 220px;
    background: rgba(255, 255, 255, 0.14);
    border-radius: 1.2rem;
    border: 1px solid rgba(255, 255, 255, 0.28);
    padding: 1.1rem 1.3rem;
    box-shadow: 0 16px 40px rgba(15, 23, 42, 0.18);
    backdrop-filter: blur(14px);
    display: flex;
    flex-direction: column;
    gap: 0.35rem;
  }
  .quick-card-title {
    font-size: 0.8rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: rgba(255, 255, 255, 0.72);
    font-weight: 700;
  }
  .quick-card-value {
    font-size: 1.05rem;
    font-weight: 600;
    color: #fff;
  }
  .quick-card-foot {
    font-size: 0.78rem;
    color: rgba(255, 255, 255, 0.72);
  }
  .panel-card {
    background: #ffffff;
    border-radius: 1.6rem;
    border: 1px solid rgba(148, 163, 184, 0.2);
    box-shadow: 0 22px 55px rgba(15, 23, 42, 0.09);
  }
  .panel-card .content {
    padding: 1.8rem;
    display: flex;
    flex-direction: column;
    gap: 1.3rem;
  }
  .section-title {
    font-size: 1.28rem;
    font-weight: 700;
    color: #0f172a;
  }
  .icon-bubble {
    width: 48px;
    height: 48px;
    border-radius: 16px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: #fff;
    font-size: 22px;
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.18);
  }
  .icon-amber {
    background: linear-gradient(135deg, #f59e0b 0%, #f97316 100%);
  }
  .icon-blue {
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
  }
  .icon-emerald {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%);
  }
  .icon-purple {
    background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%);
  }
  .action-primary {
    background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
    color: #fff !important;
    border-radius: 0.9rem !important;
    font-weight: 600 !important;
    height: 3rem;
  }
  .action-primary:hover {
    filter: brightness(1.05);
  }
  .action-secondary {
    background: rgba(37, 99, 235, 0.08) !important;
    color: #1d4ed8 !important;
    border-radius: 0.85rem !important;
    font-weight: 600 !important;
    height: 2.75rem;
  }
  .action-secondary:hover {
    background: rgba(37, 99, 235, 0.12) !important;
  }
  .action-note {
    font-size: 0.78rem;
    color: #64748b;
    line-height: 1.45;
  }
  .path-line {
    align-items: center;
    gap: 0.5rem;
    color: #475569;
    font-size: 0.85rem;
  }
  .path-line-icon {
    font-size: 1rem;
    color: #94a3b8;
  }
  .path-line-label {
    font-weight: 600;
  }
  .path-line-value {
    font-weight: 600;
    color: #0f172a;
  }
  .log-wrapper {
    border-radius: 1.25rem;
    background: #f8fafc;
    border: 1px dashed rgba(148, 163, 184, 0.35);
    padding: 1.5rem;
  }
  .log-empty {
    align-items: center;
    justify-content: center;
    height: 220px;
    width: 100%;
    gap: 0.5rem;
    color: #94a3b8;
    text-align: center;
  }
  .log-entry {
    width: 100%;
    border-radius: 1.1rem;
    padding: 0.9rem 1.1rem;
    align-items: flex-start;
    gap: 0.75rem;
    border: 1px solid transparent;
    background: rgba(226, 232, 240, 0.6);
  }
  .log-entry-text {
    font-size: 0.9rem;
    color: #0f172a;
    line-height: 1.45;
  }
  .log-entry-icon {
    font-size: 1.2rem;
  }
  .log-entry.log-info {
    background: rgba(219, 234, 254, 0.7);
    border-color: rgba(59, 130, 246, 0.15);
  }
  .log-entry.log-success {
    background: rgba(209, 250, 229, 0.8);
    border-color: rgba(16, 185, 129, 0.18);
  }
  .log-entry.log-error {
    background: rgba(254, 226, 226, 0.85);
    border-color: rgba(239, 68, 68, 0.25);
  }
  .status-chip {
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.45rem 0.95rem;
    border-radius: 9999px;
    font-weight: 600;
    font-size: 0.75rem;
    letter-spacing: 0.02em;
    transition: all 0.2s ease;
  }
  .status-chip .status-dot {
    width: 0.5rem;
    height: 0.5rem;
    border-radius: 9999px;
    background: currentColor;
    box-shadow: 0 0 0 4px rgba(255, 255, 255, 0.4);
  }
  .status-chip.status-idle {
    background: rgba(148, 163, 184, 0.2);
    color: #1e293b;
  }
  .status-chip.status-running {
    background: rgba(191, 219, 254, 0.7);
    color: #1d4ed8;
  }
  .status-chip.status-success {
    background: rgba(187, 247, 208, 0.7);
    color: #047857;
  }
  .status-chip.status-error {
    background: rgba(254, 226, 226, 0.8);
    color: #b91c1c;
  }
</style>
"""
    )

    def hero_card(title: str, value: str, foot: str) -> None:
        with ui.column().classes("quick-card w-full"):
            ui.label(title).classes("quick-card-title")
            display = _shorten(value, 28)
            label = ui.label(display).classes("quick-card-value")
            with label:
                ui.tooltip(value)
            ui.label(foot).classes("quick-card-foot")

    with ui.column().classes("min-h-screen w-full pb-16 text-slate-700 items-stretch"):
        with ui.element("section").classes("hero-gradient text-white pb-24 w-full"):
            with ui.column().classes(
                "max-w-6xl w-full mx-auto px-6 py-14 gap-8"
            ):
                ui.label("Centro de control de rentabilidad").classes("hero-title")
                ui.label(
                    "Administra la generación de informes, el cargue de EXCZ "
                    "y los listados de productos desde un solo lugar."
                ).classes("hero-subtitle")

                with ui.grid().classes(
                    "w-full gap-4 grid-cols-1 sm:grid-cols-2 xl:grid-cols-3"
                ):
                    hero_card(
                        "Plantilla base",
                        settings.ruta_plantilla.name,
                        _shorten(settings.ruta_plantilla.parent, 36),
                    )
                    hero_card(
                        "Prefijo EXCZ",
                        settings.excz_prefix,
                        _shorten(settings.excz_dir, 36),
                    )
                    hero_card(
                        "Informes",
                        settings.context.informes_dir.name,
                        _shorten(settings.context.informes_dir, 36),
                    )

        with ui.column().classes(
            "max-w-6xl w-full mx-auto px-6 -mt-16 gap-6"
        ):
            with ui.card().classes("panel-card w-full"):
                with ui.column().classes("content"):
                    with ui.row().classes(
                        "items-center gap-4 flex-wrap w-full justify-between"
                    ):
                        with ui.element("div").classes("icon-bubble icon-blue"):
                            ui.icon("folder_open").classes("text-white text-lg")
                        with ui.column().classes("gap-1"):
                            ui.label("Plantilla base de informes").classes("section-title")
                            ui.label(
                                "Esta es la plantilla utilizada para cada informe generado."
                            ).classes("action-note")
                    with ui.row().classes("items-center gap-3 flex-wrap w-full"):
                        ui.input(value=str(settings.ruta_plantilla)) \
                            .props("readonly") \
                            .classes(
                                "flex-1 min-w-[260px] bg-slate-50 border border-slate-200 "
                                "rounded-xl px-3 py-2 text-sm"
                            )
                        ui.button(
                            "Copiar ruta",
                            icon="content_copy",
                            on_click=copiar_ruta,
                        ).classes("action-secondary w-full sm:w-auto")
                        ui.button(
                            "Abrir carpeta",
                            icon="folder_open",
                            on_click=abrir_carpeta,
                        ).classes("action-secondary w-full sm:w-auto")

            with ui.row().classes(
                "gap-6 w-full flex-col lg:flex-row items-stretch"
            ):
                with ui.column().classes("flex-1 w-full gap-6"):
                    with ui.card().classes("panel-card"):
                        with ui.column().classes("content"):
                            with ui.row().classes(
                                "items-center gap-3 w-full flex-wrap"
                            ):
                                with ui.element("div").classes("icon-bubble icon-amber"):
                                    ui.icon("bolt").classes("text-white text-xl")
                                with ui.column().classes("gap-1"):
                                    ui.label("Informe automático").classes("section-title")
                                    ui.label(
                                        "Genera el informe del día anterior usando el EXCZ más reciente disponible."
                                    ).classes("action-note")

                            def ejecutar_auto() -> None:
                                update_status("running", "Generando informe automático…")
                                agregar_log(
                                    "Iniciando generación automática del informe.",
                                    "info",
                                )
                                uc_auto(
                                    GenerarInformeRequest(
                                        ruta_plantilla=str(settings.ruta_plantilla)
                                    ),
                                    bus,
                                )

                            ui.button(
                                "Generar informe automático",
                                icon="play_arrow",
                                on_click=ejecutar_auto,
                            ).classes("action-primary w-full sm:w-auto")
                            ui.label(
                                "Se buscará el archivo con prefijo configurado en la carpeta de EXCZ."
                            ).classes("action-note")

                    with ui.card().classes("panel-card"):
                        with ui.column().classes("content"):
                            with ui.row().classes(
                                "items-center gap-3 w-full flex-wrap"
                            ):
                                with ui.element("div").classes("icon-bubble icon-purple"):
                                    ui.icon("calendar_month").classes("text-white text-xl")
                                with ui.column().classes("gap-1"):
                                    ui.label("Informe manual").classes("section-title")
                                    ui.label(
                                        "Selecciona la fecha objetivo para regenerar un informe específico."
                                    ).classes("action-note")

                            fecha_input = (
                                ui.input(
                                    label="Fecha objetivo",
                                    value=ayer_str(),
                                )
                                .props("type=date")
                                .classes("w-full rounded-xl border border-slate-200 px-3 py-2")
                            )

                            def ejecutar_manual() -> None:
                                fecha = (fecha_input.value or "").strip() or None
                                if fecha:
                                    estado = f"Generando informe para {fecha}…"
                                else:
                                    estado = "Generando informe manual…"
                                update_status("running", estado)
                                agregar_log(
                                    f"Iniciando generación manual del informe ({fecha or 'día anterior'}).",
                                    "info",
                                )
                                uc_manual(
                                    GenerarInformeRequest(
                                        ruta_plantilla=str(settings.ruta_plantilla),
                                        fecha=fecha,
                                    ),
                                    bus,
                                )

                            ui.button(
                                "Generar informe manual",
                                icon="event",
                                on_click=ejecutar_manual,
                            ).classes("action-primary w-full sm:w-auto")
                            ui.label(
                                "Si dejas la fecha vacía se utilizará el día anterior de forma automática."
                            ).classes("action-note")

                with ui.column().classes("flex-1 w-full gap-6"):
                    with ui.card().classes("panel-card"):
                        with ui.column().classes("content"):
                            with ui.row().classes(
                                "items-center gap-3 w-full flex-wrap"
                            ):
                                with ui.element("div").classes("icon-bubble icon-emerald"):
                                    ui.icon("inventory_2").classes("text-white text-xl")
                                with ui.column().classes("gap-1"):
                                    ui.label("Listado de productos").classes("section-title")
                                    ui.label(
                                        "Descarga y depura el catálogo directamente desde SIIGO."
                                    ).classes("action-note")

                            def ejecutar_listado() -> None:
                                update_status(
                                    "running", "Generando listado de productos…"
                                )
                                agregar_log(
                                    "Iniciando generación del listado de productos.", "info"
                                )
                                uc_listado(bus)

                            ui.button(
                                "Generar listado de productos",
                                icon="download",
                                on_click=ejecutar_listado,
                            ).classes("action-primary w-full sm:w-auto")
                            ui.label(
                                "Se conservarán las columnas configuradas y solo se incluirán productos activos."
                            ).classes("action-note")
                            _path_line("Destino", settings.context.productos_dir)

                    with ui.card().classes("panel-card"):
                        with ui.column().classes("content"):
                            with ui.row().classes(
                                "items-center gap-3 w-full flex-wrap"
                            ):
                                with ui.element("div").classes("icon-bubble icon-blue"):
                                    ui.icon("map").classes("text-white text-xl")
                                with ui.column().classes("gap-1"):
                                    ui.label("Rutas de trabajo").classes("section-title")
                                    ui.label(
                                        "Ubicaciones donde se guardan los archivos generados."
                                    ).classes("action-note")
                            _path_line("Informes", settings.context.informes_dir)
                            _path_line("Productos", settings.context.productos_dir)
                            _path_line("Plantilla", settings.ruta_plantilla)
                            ui.label(
                                "Puedes modificar estas rutas mediante variables de entorno."
                            ).classes("action-note")

            with ui.card().classes("panel-card w-full"):
                with ui.column().classes("content"):
                    with ui.row().classes(
                        "items-center gap-3 w-full flex-wrap"
                    ):
                        with ui.element("div").classes("icon-bubble icon-blue"):
                            ui.icon("activity").classes("text-white text-xl")
                        with ui.column().classes("gap-1"):
                            ui.label("Registro de actividades").classes("section-title")
                            ui.label(
                                "Consulta el detalle de cada paso ejecutado por los procesos."
                            ).classes("action-note")
                        ui.button(
                            "Limpiar historial",
                            icon="delete",
                            on_click=limpiar_log,
                        ).props("flat color=grey").classes(
                            "ml-auto mt-3 sm:mt-0"
                        )

                    with ui.element("div").classes("log-wrapper w-full"):
                        state.empty = ui.column().classes("log-empty")
                        with state.empty:
                            ui.icon("inbox").classes("text-4xl text-slate-300")
                            ui.label(
                                "Aún no hay registros. Ejecuta una acción para comenzar."
                            ).classes("text-sm text-slate-400")
                        state.log = ui.column().classes("hidden w-full gap-3")

                    with ui.row().classes(
                        "items-center justify-between text-xs text-slate-500 w-full flex-wrap gap-3"
                    ):
                        state.status = ui.html(_status_markup("idle", "Sistema listo"))
                        state.last_update = ui.label("Última actualización: —")

    _register_bus_subscriptions()


def main() -> None:  # pragma: no cover - entrada manual
    build_ui()
    ui.run(
        native=True,
        title="Rentabilidad",
        window_size=(1200, 800),
        fullscreen=False,
        reload=False,
        port=0,
    )


if __name__ in {"__main__", "__mp_main__"}:  # pragma: no cover
    main()
