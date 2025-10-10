from __future__ import annotations

import asyncio
import base64
import html
import json
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from typing import Any

from nicegui import app, ui
from importlib import resources
from contextlib import suppress

from rentabilidad.app.dto import GenerarInformeRequest
from rentabilidad.app.use_cases.generar_informe_automatico import run as uc_auto
from rentabilidad.app.use_cases.generar_informe_manual import run as uc_manual
from rentabilidad.app.use_cases.listar_productos import run as uc_listado
from rentabilidad.config import bus, settings
from rentabilidad.infra.fs import ayer_str, find_latest_informe, find_latest_producto

@dataclass(slots=True)
class LatestResourceComponents:
    name: Any = None
    meta: Any = None
    path_display: Any = None
    action: Any = None
    path: Path | None = None


@dataclass(slots=True)
class UIState:
    empty: Any = None
    log: Any = None
    last_update: Any = None
    status: Any = None
    status_button: Any = None
    status_path: Path | None = None
    status_kind: str = "idle"
    latest_report: LatestResourceComponents = field(default_factory=LatestResourceComponents)
    latest_products: LatestResourceComponents = field(default_factory=LatestResourceComponents)


state = UIState()

BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
STATIC_DIR = BASE_DIR / "static"
LOGO_FILENAME = "logo.svg"

_subscriptions_registered = False
_api_registered = False
_static_registered = False
_static_temp_dir: tempfile.TemporaryDirectory[str] | None = None
_static_dir_cache: Path | None = None

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

LOG_NOTIFY_TYPES = {
    "info": "info",
    "success": "positive",
    "error": "negative",
}


class StatusManager:
    def __init__(self, ui_state: UIState) -> None:
        self._state = ui_state

    @property
    def current_path(self) -> Path | None:
        return self._state.status_path

    def render(self, kind: str, text: str) -> str:
        return (
            f'<div class="status-chip status-{kind}">'  # noqa: E501
            "<span class=\"status-dot\"></span>"
            f"<span>{text}</span>"
            "</div>"
        )

    def _set_action_target(self, path: Path | None) -> None:
        self._state.status_path = path
        button = self._state.status_button
        if button is None:
            return
        if path is None:
            button.disable()
        else:
            button.enable()

    def update(self, kind: str, text: str, open_path: str | Path | None = None) -> None:
        status_component = self._state.status
        self._state.status_kind = kind
        if status_component is None:
            return

        status_component.content = self.render(kind, text)
        target: Path | None = None
        if kind == "success" and open_path:
            target = Path(open_path)
        self._set_action_target(target)


class LogManager:
    def __init__(self, ui_state: UIState) -> None:
        self._state = ui_state

    def add(self, message: str, kind: str = "info") -> None:
        if self._state.empty is None or self._state.log is None:
            return

        self._state.empty.classes(add="hidden")
        self._state.log.classes(remove="hidden")

        css_class = LOG_ENTRY_CLASSES.get(kind, "log-info")
        icon = LOG_ICONS.get(kind, "info")
        icon_class = LOG_ICON_CLASSES.get(kind, "text-slate-500")

        notify_type = LOG_NOTIFY_TYPES.get(kind)
        if notify_type is None:
            notify_type = "info"

        ui.notify(
            message,
            type=notify_type,
            position="top-right",
            close_button="×",
            multi_line=True,
            timeout=6000 if kind != "error" else 0,
        )

        with self._state.log:
            with ui.row().classes(f"log-entry {css_class}"):
                ui.icon(icon).classes(f"log-entry-icon {icon_class}")
                ui.label(message).classes("log-entry-text")

    def touch_last_update(self) -> None:
        if self._state.last_update is None:
            return
        self._state.last_update.text = (
            f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
        )

    def clear(self) -> None:
        if self._state.log is not None:
            self._state.log.clear()
            self._state.log.classes(add="hidden")
        if self._state.empty is not None:
            self._state.empty.classes(remove="hidden")
        if self._state.last_update is not None:
            self._state.last_update.text = "Última actualización: —"


class ResourceManager:
    def __init__(self, status: StatusManager, logs: LogManager) -> None:
        self._status = status
        self._logs = logs

    def copy_template_path(self) -> None:
        ruta = str(settings.ruta_plantilla)
        ui.run_javascript(f"navigator.clipboard.writeText({json.dumps(ruta)})")
        self._logs.add("Ruta de la plantilla copiada al portapapeles.")

    def open_template_folder(self) -> None:
        carpeta = Path(settings.ruta_plantilla).parent
        if not carpeta.exists():
            mensaje = f"No se encontró la carpeta: {carpeta}"
            self._logs.add(mensaje, "error")
            self._status.update("error", "Ruta de la plantilla no encontrada")
            return

        if self._open_destination(carpeta, "la carpeta"):
            self._logs.add(f"Carpeta abierta: {carpeta}")
        else:
            self._status.update("error", "No fue posible abrir la carpeta")

    def _open_destination(self, destino: Path, descripcion: str) -> bool:
        ruta = str(destino)
        try:
            if sys.platform.startswith("win"):
                try:
                    os.startfile(ruta)  # type: ignore[attr-defined]
                except OSError:
                    subprocess.Popen(
                        ["cmd", "/c", "start", "", ruta],
                        shell=True,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
            elif sys.platform == "darwin":
                subprocess.run(["open", ruta], check=False)
            else:
                subprocess.run(["xdg-open", ruta], check=False)
            return True
        except Exception as exc:  # pragma: no cover - defensivo
            mensaje = f"No se pudo abrir {descripcion}: {exc}"
            self._logs.add(mensaje, "error")
            return False

    def open_result(self, destino: Path) -> bool:
        if not destino.exists():
            mensaje = f"No se encontró el recurso generado: {destino}"
            self._logs.add(mensaje, "error")
            return False

        if self._open_destination(destino, "el recurso generado"):
            self._logs.add(f"Recurso abierto: {destino}", "success")
            return True
        return False

    def open_current_result(self) -> None:
        destino = self._status.current_path
        if destino is None:
            return
        self.open_result(destino)

    @staticmethod
    def extract_result_path(msg: str) -> Path | None:
        prefijos = ("Informe generado:", "Listado generado:")
        for prefijo in prefijos:
            if msg.startswith(prefijo):
                posible = msg[len(prefijo) :].strip().strip("'\"")
                if posible:
                    return Path(posible)

        marcadores = (
            " sobre:",
            " Informe:",
            " informe:",
            "INFORME:",
        )
        segmentos = [msg]
        if "|" in msg:
            segmentos = [parte.strip() for parte in msg.split("|")]

        for segmento in segmentos:
            for marcador in marcadores:
                if marcador in segmento:
                    posible = segmento.split(marcador, 1)[1].strip().strip("'\"")
                    if posible:
                        posible = posible.split("|", 1)[0].strip()
                    if posible and any(sep in posible for sep in ("/", "\\")):
                        return Path(posible)

        if ":" in msg:
            partes = msg.split(":")
            for idx in range(len(partes) - 1, 0, -1):
                posible = ":".join(partes[idx:]).strip().strip("'\"")
                if posible and any(sep in posible for sep in ("/", "\\")):
                    return Path(posible.split("|", 1)[0].strip())

        return None


status_manager = StatusManager(state)
log_manager = LogManager(state)
resource_manager = ResourceManager(status_manager, log_manager)


class LatestResourceManager:
    def __init__(self, ui_state: UIState) -> None:
        self._state = ui_state
        self._informes_dir = settings.context.informes_dir
        self._productos_dir = settings.context.productos_dir

    def refresh(self) -> None:
        self._set_components(
            self._state.latest_report, find_latest_informe(self._informes_dir)
        )
        self._set_components(
            self._state.latest_products, find_latest_producto(self._productos_dir)
        )

    def update_from_path(self, path: Path | None) -> None:
        if path is None:
            return
        target = self._resolve(path)
        if target is None:
            return
        if self._is_under(target, self._informes_dir):
            self._set_components(self._state.latest_report, target)
        elif self._is_under(target, self._productos_dir):
            self._set_components(self._state.latest_products, target)

    def _resolve(self, path: Path) -> Path | None:
        try:
            return path.resolve()
        except OSError:
            if path.exists():
                return path
        except RuntimeError:
            pass
        return path if path.exists() else None

    def _set_components(
        self, components: LatestResourceComponents, path: Path | None
    ) -> None:
        actual = path if path and path.exists() else None
        components.path = actual

        if components.name is not None:
            if actual:
                components.name.text = actual.name
            else:
                components.name.text = "Sin registros disponibles."

        if components.meta is not None:
            if actual:
                components.meta.text = (
                    "Actualizado el "
                    + datetime.fromtimestamp(actual.stat().st_mtime).strftime(
                        "%d/%m/%Y %H:%M"
                    )
                )
            else:
                components.meta.text = "Genera este recurso para visualizarlo aquí."

        if components.path_display is not None:
            if actual:
                shortened = shorten(actual)
                escaped_short = html.escape(shortened)
                escaped_full = html.escape(str(actual))
                components.path_display.content = "".join(
                    [
                        '<span class="latest-resource-path-label">Ubicación:</span> ',
                        f'<span class="latest-resource-path-value" title="{escaped_full}">',
                        escaped_short,
                        "</span>",
                    ]
                )
            else:
                components.path_display.content = (
                    '<span class="latest-resource-path-label">Ubicación:</span> '
                    '<span class="latest-resource-path-value">—</span>'
                )

        if components.action is not None:
            if actual:
                components.action.enable()
            else:
                components.action.disable()

    @staticmethod
    def _is_under(path: Path, base: Path) -> bool:
        try:
            path.resolve().relative_to(base.resolve())
            return True
        except (ValueError, OSError):
            try:
                path.relative_to(base)
                return True
            except ValueError:
                return False


latest_resources = LatestResourceManager(state)


def shorten(text: str, length: int = 42) -> str:
    clean = str(text)
    return clean if len(clean) <= length else clean[: length - 1] + "…"


def _is_windows() -> bool:
    return sys.platform.startswith("win")


def _is_remote_session() -> bool:
    if not _is_windows():
        return False

    session_name = os.environ.get("SESSIONNAME", "").strip().lower()
    if session_name and "rdp" in session_name:
        return True

    client_name = os.environ.get("CLIENTNAME", "").strip().lower()
    if client_name and client_name != "console":
        return True

    return False


def _ensure_static_dir() -> Path | None:
    global _static_temp_dir, _static_dir_cache
    if _static_dir_cache and _static_dir_cache.exists():
        return _static_dir_cache

    if STATIC_DIR.exists():
        _static_dir_cache = STATIC_DIR
        return _static_dir_cache

    try:
        package_static = resources.files("rentabilidad.gui").joinpath("static")
    except (FileNotFoundError, ModuleNotFoundError):
        return None

    if not package_static.is_dir():
        return None

    if _static_temp_dir is None:
        with suppress(FileNotFoundError):
            with resources.as_file(package_static) as resolved:
                resolved_path = Path(resolved)
                if resolved_path.exists() and resolved_path.is_dir():
                    temp_dir = tempfile.TemporaryDirectory(prefix="rent_static_")
                    temp_path = Path(temp_dir.name) / "static"
                    try:
                        shutil.copytree(resolved_path, temp_path, dirs_exist_ok=True)
                    except OSError:
                        temp_dir.cleanup()
                    else:
                        _static_temp_dir = temp_dir
                        _static_dir_cache = temp_path

    return _static_dir_cache


def _register_static_files() -> None:
    global _static_registered
    if _static_registered:
        return

    static_dir = _ensure_static_dir()
    if not static_dir:
        return

    app.add_static_files("/static", str(static_dir))
    _static_registered = True


def _get_logo_file() -> Path | None:
    static_dir = _ensure_static_dir()
    if not static_dir:
        return None

    logo_path = static_dir / LOGO_FILENAME
    return logo_path if logo_path.exists() else None


def _logo_source() -> str | None:
    logo_path = _get_logo_file()
    if logo_path and logo_path.exists():
        try:
            encoded = base64.b64encode(logo_path.read_bytes()).decode("ascii")
        except OSError:
            pass
        else:
            return f"data:image/svg+xml;base64,{encoded}"
    if _ensure_static_dir():
        return f"/static/{LOGO_FILENAME}"
    return None


def _inline_logo_markup() -> str | None:
    logo_path = _get_logo_file()
    if not logo_path or not logo_path.exists():
        return None
    try:
        svg = logo_path.read_text(encoding="utf-8")
    except OSError:
        return None

    svg = svg.lstrip()
    if svg.startswith("<?xml"):
        _, _, remainder = svg.partition("?>")
        svg = remainder.lstrip() or svg

    if "<svg" not in svg:
        return None

    return f'<div class="hero-logo-inline">{svg}</div>'

def update_status(
    kind: str, text: str, open_path: str | Path | None = None
) -> None:
    previous_kind = state.status_kind
    status_manager.update(kind, text, open_path)
    if kind == "success" and open_path:
        destino = open_path if isinstance(open_path, Path) else Path(open_path)
        latest_resources.update_from_path(destino)
    if kind in {"success", "error"} and previous_kind != kind:
        if kind == "success":
            agregar_log("El proceso finalizó correctamente.", "success")
        else:
            agregar_log("El proceso finalizó con errores.", "error")


def agregar_log(msg: str, kind: str = "info") -> None:
    log_manager.add(msg, kind)


def touch_last_update() -> None:
    log_manager.touch_last_update()


def limpiar_log() -> None:
    log_manager.clear()
    status_manager.update("idle", "Sistema listo")


def copiar_ruta() -> None:
    resource_manager.copy_template_path()


def abrir_carpeta() -> None:
    resource_manager.open_template_folder()


def abrir_resultado(destino: Path) -> bool:
    return resource_manager.open_result(destino)


def abrir_resultado_actual() -> None:
    resource_manager.open_current_result()


def _extract_result_path(msg: str) -> Path | None:
    return ResourceManager.extract_result_path(msg)


def abrir_ultimo_informe() -> None:
    destino = state.latest_report.path
    if destino is None:
        log_manager.add(
            "No hay un informe de rentabilidad disponible para abrir.", "error"
        )
        return
    resource_manager.open_result(destino)


def abrir_ultimo_listado() -> None:
    destino = state.latest_products.path
    if destino is None:
        log_manager.add(
            "No hay un listado de productos disponible para abrir.", "error"
        )
        return
    resource_manager.open_result(destino)


def _path_line(label: str, value: Path) -> None:
    with ui.row().classes("path-line w-full flex-wrap"):
        ui.icon("chevron_right").classes("path-line-icon")
        ui.label(f"{label}:").classes("path-line-label")
        shortened = shorten(value)
        component = ui.label(shortened).classes("path-line-value")
        with component:
            ui.tooltip(str(value))


def _latest_resource_entry(
    title: str,
    icon_name: str,
    icon_class: str,
    components: LatestResourceComponents,
    on_open,
) -> None:
    with ui.element("div").classes("latest-resource-card"):
        with ui.row().classes("latest-resource-header"):
            ui.icon(icon_name).classes(f"latest-resource-icon {icon_class}")
            ui.label(title).classes("latest-resource-title")
        components.name = ui.label("Sin registros disponibles.").classes(
            "latest-resource-name"
        )
        components.meta = ui.label(
            "Genera este recurso para visualizarlo aquí."
        ).classes("latest-resource-meta")
        components.path_display = ui.html(
            '<span class="latest-resource-path-label">Ubicación:</span> '
            '<span class="latest-resource-path-value">—</span>',
            sanitize=False,
        ).classes("latest-resource-path")
        with ui.row().classes("latest-resource-actions"):
            components.action = (
                ui.button(
                    "Abrir archivo",
                    icon="open_in_new",
                    on_click=on_open,
                )
                .props("flat color=primary")
                .classes("latest-resource-action")
            )
            components.action.disable()


def _register_bus_subscriptions() -> None:
    global _subscriptions_registered
    if _subscriptions_registered:
        return

    def _on_log(msg: str) -> None:
        agregar_log(msg, "info")

    def _on_done(msg: str) -> None:
        agregar_log(msg, "success")
        touch_last_update()
        destino = _extract_result_path(msg)
        if destino is None and state.status_path is not None:
            destino = state.status_path
        if destino is not None:
            latest_resources.update_from_path(destino)
        update_status("success", "Proceso completado", open_path=destino)

    def _on_error(msg: str) -> None:
        agregar_log(msg, "error")
        update_status("error", "Revisa los registros")
        touch_last_update()

    bus.subscribe("log", _on_log)
    bus.subscribe("done", _on_done)
    bus.subscribe("error", _on_error)

    _subscriptions_registered = True


def _register_api_routes() -> None:
    global _api_registered
    if _api_registered:
        return

    @app.post("/api/abrir-recurso")
    async def abrir_recurso_api(payload: dict[str, str]) -> dict[str, str]:
        ruta_texto = payload.get("ruta") if isinstance(payload, dict) else None
        if not ruta_texto:
            mensaje = "No se indicó la ruta del recurso a abrir"
            agregar_log(mensaje, "error")
            return {"status": "error", "detail": mensaje}

        destino = Path(ruta_texto)
        if not destino.exists():
            mensaje = f"No se encontró el recurso generado: {destino}"
            agregar_log(mensaje, "error")
            return {"status": "error", "detail": mensaje}

        if abrir_resultado(destino):
            touch_last_update()
            return {"status": "ok", "detail": ""}

        mensaje = f"No se pudo abrir el recurso generado: {destino}"
        agregar_log(mensaje, "error")
        return {"status": "error", "detail": mensaje}

    _api_registered = True


def build_ui() -> None:
    bus.bind_loop()
    logo_url = _logo_source()
    logo_markup = _inline_logo_markup()

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
  .connection-banner {
    position: fixed;
    top: 1.25rem;
    left: 50%;
    transform: translate(-50%, 0);
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    background: rgba(239, 68, 68, 0.95);
    color: #fff;
    padding: 0.75rem 1.2rem;
    border-radius: 999px;
    box-shadow: 0 18px 40px rgba(248, 113, 113, 0.35);
    z-index: 2000;
    transition: opacity 0.3s ease, transform 0.3s ease;
  }
  .connection-banner.hidden {
    opacity: 0;
    pointer-events: none;
    transform: translate(-50%, -20px);
  }
  .connection-banner-dot {
    width: 0.65rem;
    height: 0.65rem;
    border-radius: 999px;
    background: #fff;
    box-shadow: 0 0 0 0 rgba(255, 255, 255, 0.55);
    animation: pulse-dot 1.8s infinite;
  }
  .connection-banner-text {
    font-weight: 600;
    letter-spacing: 0.02em;
  }
  @keyframes pulse-dot {
    0% {
      box-shadow: 0 0 0 0 rgba(255, 255, 255, 0.55);
    }
    70% {
      box-shadow: 0 0 0 10px rgba(255, 255, 255, 0);
    }
    100% {
      box-shadow: 0 0 0 0 rgba(255, 255, 255, 0);
    }
  }
  .hero-header {
    align-items: center;
    gap: 1.5rem;
  }
  .hero-header > *:first-child {
    flex-shrink: 0;
  }
  .hero-logo {
    height: 3.25rem;
    width: auto;
    filter: drop-shadow(0 10px 24px rgba(15, 23, 42, 0.35));
  }
  .hero-logo-inline {
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .hero-logo-inline svg {
    height: 3.25rem;
    width: auto;
    filter: drop-shadow(0 10px 24px rgba(15, 23, 42, 0.35));
  }
  .hero-logo-fallback {
    width: 3.25rem;
    height: 3.25rem;
    border-radius: 1rem;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: rgba(15, 23, 42, 0.12);
    color: #0f172a;
    box-shadow: 0 14px 28px rgba(15, 23, 42, 0.2);
  }
  .hero-logo-fallback-icon {
    font-size: 1.75rem;
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
    display: flex;
    flex-direction: column;
    gap: 1.5rem;
  }
  .log-stream {
    width: 100%;
    background: #ffffff;
    border-radius: 1rem;
    border: 1px solid rgba(148, 163, 184, 0.2);
    padding: 1.25rem;
    display: flex;
    flex-direction: column;
    gap: 1.25rem;
    min-height: 360px;
    max-height: 360px;
    box-shadow: 0 12px 32px rgba(15, 23, 42, 0.04);
    overflow: hidden;
  }
  .log-empty {
    align-items: center;
    justify-content: center;
    height: 100%;
    min-height: 0;
    width: 100%;
    gap: 0.5rem;
    color: #94a3b8;
    text-align: center;
  }
  .log-list {
    flex: 1 1 auto;
    height: 100%;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 0.75rem;
  }
  .log-list::-webkit-scrollbar {
    width: 6px;
  }
  .log-list::-webkit-scrollbar-thumb {
    background: rgba(148, 163, 184, 0.4);
    border-radius: 9999px;
  }
  .latest-resources {
    width: 100%;
    display: flex;
    flex-direction: column;
    gap: 1rem;
    align-items: stretch;
    align-self: stretch;
  }
  .latest-resources-title {
    font-size: 1.05rem;
    font-weight: 600;
    color: #0f172a;
  }
  .latest-resource-card {
    background: #ffffff;
    border-radius: 1rem;
    border: 1px solid rgba(148, 163, 184, 0.2);
    padding: 1.25rem;
    display: flex;
    flex-direction: column;
    gap: 0.75rem;
    box-shadow: 0 12px 32px rgba(15, 23, 42, 0.04);
    width: 100%;
  }
  .latest-resource-header {
    display: flex;
    align-items: center;
    gap: 0.85rem;
  }
  .latest-resource-icon {
    font-size: 1.2rem;
    color: #1d4ed8;
  }
  .latest-resource-title {
    font-size: 0.95rem;
    font-weight: 700;
    color: #0f172a;
  }
  .latest-resource-name {
    font-size: 0.9rem;
    font-weight: 600;
    color: #1e293b;
  }
  .latest-resource-meta {
    font-size: 0.8rem;
    color: #64748b;
  }
  .latest-resource-path {
    font-size: 0.78rem;
    color: #475569;
  }
  .latest-resource-path-label {
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 600;
    color: #94a3b8;
    margin-right: 0.35rem;
  }
  .latest-resource-path-value {
    font-weight: 600;
    color: #1f2937;
    word-break: break-all;
  }
  .latest-resource-actions {
    margin-top: auto;
    display: flex;
    gap: 0.5rem;
  }
  .latest-resource-action {
    border-radius: 0.75rem !important;
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
  .status-actions {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    flex-wrap: wrap;
  }
  .status-action {
    background: rgba(4, 120, 87, 0.12) !important;
    color: #047857 !important;
    border-radius: 9999px !important;
    font-weight: 600 !important;
    font-size: 0.75rem !important;
    height: 2.25rem;
    padding: 0 1.1rem;
  }
  .status-action:hover {
    background: rgba(4, 120, 87, 0.2) !important;
  }
  .status-action:disabled {
    background: rgba(148, 163, 184, 0.14) !important;
    color: #64748b !important;
    cursor: not-allowed !important;
  }
</style>
"""
    )

    ui.add_body_html(
        """
<div id=\"connection-banner\" class=\"connection-banner hidden\">
  <span class=\"connection-banner-dot\"></span>
  <span class=\"connection-banner-text\">Reconectando…</span>
</div>
<script>
  (function() {
    const banner = document.getElementById('connection-banner');
    if (!banner) return;
    const hide = () => banner.classList.add('hidden');
    const show = () => banner.classList.remove('hidden');
    window.addEventListener('nicegui:connected', hide);
    window.addEventListener('nicegui:disconnected', show);
    hide();
  })();
</script>
        """
    )

    def hero_card(title: str, value: str, foot: str) -> None:
        with ui.column().classes("quick-card w-full"):
            ui.label(title).classes("quick-card-title")
            display = shorten(value, 28)
            label = ui.label(display).classes("quick-card-value")
            with label:
                ui.tooltip(value)
            ui.label(foot).classes("quick-card-foot")

    with ui.column().classes("min-h-screen w-full pb-16 text-slate-700 items-stretch"):
        with ui.element("section").classes("hero-gradient text-white pb-24 w-full"):
            with ui.column().classes(
                "max-w-6xl w-full mx-auto px-6 py-14 gap-8"
            ):
                with ui.row().classes(
                    "items-center gap-5 flex-wrap w-full hero-header"
                ):
                    if logo_markup:
                        ui.html(logo_markup, sanitize=False).classes("hero-logo")
                    elif logo_url:
                        ui.image(logo_url).classes("hero-logo")
                    else:
                        with ui.element("div").classes(
                            "hero-logo hero-logo-fallback"
                        ):
                            ui.icon("apartment").classes("hero-logo-fallback-icon")
                    with ui.column().classes("gap-2"):
                        ui.label(
                            "Centro de control de rentabilidad"
                        ).classes("hero-title")
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
                        shorten(settings.ruta_plantilla.parent, 36),
                    )
                    hero_card(
                        "Prefijo EXCZ",
                        settings.excz_prefix,
                        shorten(settings.excz_dir, 36),
                    )
                    hero_card(
                        "Informes",
                        settings.context.informes_dir.name,
                        shorten(settings.context.informes_dir, 36),
                    )

        with ui.column().classes(
            "max-w-6xl w-full mx-auto px-6 -mt-16 gap-6"
        ):
            with ui.card().classes("panel-card w-full"):
                with ui.column().classes("content w-full items-stretch"):
                    with ui.row().classes(
                        "items-center gap-4 flex-wrap w-full"
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
                        with ui.column().classes("content w-full items-stretch"):
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

                            async def ejecutar_auto() -> None:
                                update_status("running", "Generando informe automático…")
                                agregar_log(
                                    "Iniciando generación automática del informe.",
                                    "info",
                                )
                                try:
                                    resultado = await asyncio.to_thread(
                                        uc_auto,
                                        GenerarInformeRequest(
                                            ruta_plantilla=str(settings.ruta_plantilla)
                                        ),
                                        bus,
                                    )
                                except Exception as exc:  # pragma: no cover - defensivo
                                    bus.publish("error", str(exc))
                                    update_status("error", "Revisa los registros")
                                    return
                                if (
                                    resultado.ok
                                    and resultado.ruta_salida
                                ):
                                    update_status(
                                        "success",
                                        "Proceso completado",
                                        open_path=resultado.ruta_salida,
                                    )
                                else:
                                    mensaje = resultado.mensaje or (
                                        "No se pudo generar el informe automático."
                                    )
                                    agregar_log(
                                        f"Error al generar el informe automático: {mensaje}",
                                        "error",
                                    )
                                    update_status("error", "Revisa los registros")

                            ui.button(
                                "Generar informe automático",
                                icon="play_arrow",
                                on_click=ejecutar_auto,
                            ).classes("action-primary w-full sm:w-auto")
                            ui.label(
                                "Se buscará el archivo con prefijo configurado en la carpeta de EXCZ."
                            ).classes("action-note")

                    with ui.card().classes("panel-card"):
                        with ui.column().classes("content w-full items-stretch"):
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

                            async def ejecutar_manual() -> None:
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
                                try:
                                    resultado = await asyncio.to_thread(
                                        uc_manual,
                                        GenerarInformeRequest(
                                            ruta_plantilla=str(settings.ruta_plantilla),
                                            fecha=fecha,
                                        ),
                                        bus,
                                    )
                                except Exception as exc:  # pragma: no cover - defensivo
                                    bus.publish("error", str(exc))
                                    update_status("error", "Revisa los registros")
                                    return
                                if (
                                    resultado.ok
                                    and resultado.ruta_salida
                                ):
                                    update_status(
                                        "success",
                                        "Proceso completado",
                                        open_path=resultado.ruta_salida,
                                    )
                                else:
                                    mensaje = resultado.mensaje or (
                                        "No se pudo generar el informe manual."
                                    )
                                    agregar_log(
                                        f"Error al generar el informe manual: {mensaje}",
                                        "error",
                                    )
                                    update_status("error", "Revisa los registros")

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
                        with ui.column().classes("content w-full items-stretch"):
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

                            async def ejecutar_listado() -> None:
                                update_status(
                                    "running", "Generando listado de productos…"
                                )
                                agregar_log(
                                    "Iniciando generación del listado de productos.", "info"
                                )
                                try:
                                    ruta = await asyncio.to_thread(uc_listado, bus)
                                except Exception as exc:  # pragma: no cover - defensivo
                                    bus.publish("error", str(exc))
                                    update_status("error", "Revisa los registros")
                                    return
                                if ruta:
                                    update_status(
                                        "success",
                                        "Proceso completado",
                                        open_path=ruta,
                                    )
                                else:
                                    agregar_log(
                                        "No se pudo generar el listado de productos."
                                        " Verifica los registros anteriores para más detalles.",
                                        "error",
                                    )
                                    update_status("error", "Revisa los registros")

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
                        with ui.column().classes("content w-full items-stretch"):
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
                with ui.column().classes("content w-full items-stretch"):
                    with ui.row().classes(
                        "items-center gap-3 w-full flex-wrap"
                    ):
                        with ui.element("div").classes("icon-bubble icon-blue"):
                            ui.icon("history").classes("text-white text-xl")
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
                        with ui.column().classes("latest-resources"):
                            ui.label("Recursos recientes").classes(
                                "latest-resources-title"
                            )
                            _latest_resource_entry(
                                "Informe de rentabilidad",
                                "table_view",
                                "text-blue-500",
                                state.latest_report,
                                abrir_ultimo_informe,
                            )
                            _latest_resource_entry(
                                "Listado de productos",
                                "inventory_2",
                                "text-emerald-500",
                                state.latest_products,
                                abrir_ultimo_listado,
                            )
                        with ui.column().classes("log-stream"):
                            state.empty = ui.column().classes("log-empty")
                            with state.empty:
                                ui.icon("inbox").classes(
                                    "text-4xl text-slate-300"
                                )
                                ui.label(
                                    "Aún no hay registros. Ejecuta una acción para comenzar."
                                ).classes("text-sm text-slate-400")
                            state.log = ui.column().classes(
                                "hidden log-list"
                            )
                    with ui.row().classes(
                        "items-center justify-between text-xs text-slate-500 w-full flex-wrap gap-3"
                    ):
                        with ui.row().classes("status-actions"):
                            state.status = ui.html(
                                status_manager.render("idle", "Sistema listo"),
                                sanitize=False,
                            )
                            state.status_button = (
                                ui.button(
                                    "Abrir",
                                    icon="open_in_new",
                                    on_click=abrir_resultado_actual,
                                )
                                .props("flat")
                                .classes("status-action")
                            )
                            state.status_button.disable()
                        state.last_update = ui.label("Última actualización: —")

    latest_resources.refresh()
    _register_bus_subscriptions()
    _register_api_routes()


def main() -> None:  # pragma: no cover - entrada manual
    build_ui()
    _register_static_files()

    base_kwargs: dict[str, Any] = {
        "title": "Rentabilidad",
        "reload": False,
    }

    if _is_windows() and not _is_remote_session():
        app.native.start_args.setdefault("gui", "edgechromium")

    ui.run(
        native=True,
        window_size=(1200, 800),
        fullscreen=False,
        port=8080,
        **base_kwargs,
    )


if __name__ in {"__main__", "__mp_main__"}:  # pragma: no cover
    main()
