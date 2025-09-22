"""Aplicación gráfica para controlar procesos de rentabilidad."""

from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import threading
import unicodedata
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Optional

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import font as tkfont

from openpyxl.utils import column_index_from_string

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContext, PathContextFactory
from rentabilidad.services.products import (
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
)
from servicios.generar_listado_productos import KEEP_COLUMN_NUMBERS


REPO_ROOT = Path(__file__).resolve().parents[2]
LOADER_SCRIPT = REPO_ROOT / "hojas" / "hoja01_loader.py"


@dataclass
class TaskResult:
    """Valor devuelto por cada tarea en segundo plano."""

    message: str
    output: Optional[Path] = None


class LogStream:
    """Adaptador simple que redirige ``print`` hacia el registro de la GUI."""

    def __init__(self, callback: Callable[[str], None]):
        """Inicializa el stream indicando la función que recibirá cada línea."""

        self._callback = callback

    def write(self, data: str) -> None:  # pragma: no cover - integración I/O
        """Envía ``data`` al registro dividiéndolo por líneas útiles."""

        if not data:
            return
        text = data.strip()
        if not text:
            return
        for line in text.splitlines():
            self._callback(line)

    def flush(self) -> None:  # pragma: no cover - requerido por interface
        """Se expone para cumplir la interfaz de archivo, no realiza acciones."""

        return


def ensure_trailing_backslash(path: str) -> str:
    """Garantiza que ``path`` termine con ``\\`` o ``/``."""

    return path if path.endswith(("\\", "/")) else path + "\\"


def _widget_background(widget: tk.Widget, fallback: str) -> str:
    """Obtiene el color de fondo de ``widget`` con un respaldo seguro."""

    try:
        value = widget.cget("background")
    except tk.TclError:
        return fallback
    if not value or value == "SystemButtonFace":
        return fallback
    return value


class RoundedCard(tk.Frame):
    """Contenedor con esquinas redondeadas dibujadas sobre un lienzo."""

    def __init__(
        self,
        parent: tk.Widget,
        *,
        background: str,
        border: str,
        radius: int = 16,
        padding: int = 20,
    ) -> None:
        container_bg = _widget_background(parent, background)
        super().__init__(parent, bg=container_bg, bd=0, highlightthickness=0)

        self._background_color = background
        self._border_color = border
        self._radius = radius
        self._padding = padding

        self._canvas = tk.Canvas(
            self,
            bg=container_bg,
            bd=0,
            highlightthickness=0,
        )
        self._canvas.pack(fill="both", expand=True)

        self.inner = tk.Frame(self._canvas, bg=background, bd=0, highlightthickness=0)
        self._window = self._canvas.create_window(padding, padding, anchor="nw", window=self.inner)
        self._canvas.bind("<Configure>", self._on_configure)

    def _on_configure(self, event: tk.Event) -> None:
        """Redimensiona el lienzo y redibuja el fondo redondeado."""

        width = max(int(event.width), 2)
        height = max(int(event.height), 2)
        pad = self._padding
        inner_width = max(width - 2 * pad, 1)
        inner_height = max(height - 2 * pad, 1)

        self._canvas.coords(self._window, pad, pad)
        self._canvas.itemconfigure(self._window, width=inner_width, height=inner_height)

        radius = min(self._radius, inner_width // 2, inner_height // 2, width // 2, height // 2)
        self._draw_background(width, height, radius)

    def _draw_background(self, width: int, height: int, radius: int) -> None:
        """Dibuja un polígono suavizado simulando esquinas curvas."""

        if radius <= 0:
            radius = 1

        points = [
            radius,
            0,
            width - radius,
            0,
            width,
            0,
            width,
            radius,
            width,
            height - radius,
            width,
            height,
            width - radius,
            height,
            radius,
            height,
            0,
            height,
            0,
            height - radius,
            0,
            radius,
            0,
            0,
        ]

        self._canvas.delete("card")
        self._canvas.create_polygon(
            points,
            smooth=True,
            splinesteps=24,
            fill=self._background_color,
            outline=self._border_color,
            width=1,
            tags="card",
        )
        self._canvas.tag_lower("card", self._window)


class TerminalPanel(tk.Frame):
    """Panel de aspecto terminal con resaltado de estados y scroll integrado."""

    def __init__(
        self,
        parent: tk.Widget,
        *,
        background: str,
        text_color: str,
        muted_color: str,
        accent_color: str,
        success_color: str,
        warning_color: str,
        error_color: str,
        cursor_color: str,
        font_family: str,
        monospace_family: str,
    ) -> None:
        super().__init__(parent, bg=background, bd=0, highlightthickness=0)

        self._background = background
        self._text_color = text_color
        self._muted_color = muted_color
        self._accent_color = accent_color
        self._success_color = success_color
        self._warning_color = warning_color
        self._error_color = error_color
        self._cursor_color = cursor_color
        self._font_family = font_family
        self._monospace_family = monospace_family

        self._has_content = False

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._text = tk.Text(
            self,
            wrap="word",
            bg=background,
            fg=text_color,
            insertbackground=cursor_color,
            selectbackground=accent_color,
            selectforeground="#f8fafc",
            font=(monospace_family, 11),
            padx=16,
            pady=12,
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self._text.grid(row=0, column=0, sticky="nsew")
        self._text.configure(spacing1=4, spacing3=4, cursor="arrow")
        self._text.bind("<Key>", self._block_key)
        self._text.bind("<Control-a>", self._select_all)
        self._text.bind("<Button-1>", lambda event: self._text.focus_set())
        self._text.bind("<<Paste>>", lambda _event: "break")
        self._text.bind("<<Cut>>", lambda _event: "break")
        self._text.bind("<<Clear>>", lambda _event: "break")

        self._scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            style="Vertical.TScrollbar",
            command=self._text.yview,
        )
        self._scrollbar.grid(row=0, column=1, sticky="ns")
        self._text.configure(yscrollcommand=self._scrollbar.set)

        self._placeholder = tk.Label(
            self,
            text="El registro de actividades aparecerá aquí",
            bg=background,
            fg=muted_color,
            font=(font_family, 11),
            justify="center",
            wraplength=320,
        )
        self._placeholder.place(relx=0.5, rely=0.5, anchor="center")
        self._placeholder.lift()

        self._text.tag_configure(
            "timestamp",
            foreground=accent_color,
            font=(monospace_family, 11, "bold"),
        )
        self._text.tag_configure("info", foreground=text_color)
        self._text.tag_configure("success", foreground=success_color)
        self._text.tag_configure("warning", foreground=warning_color)
        self._text.tag_configure("error", foreground=error_color)
        self._text.tag_configure("muted", foreground=muted_color)

        self.bind("<Configure>", self._on_configure)

    # ----------------------------------------------------------- Internals --
    def _on_configure(self, event: tk.Event) -> None:
        wrap = max(int(event.width) - 80, 200)
        self._placeholder.configure(wraplength=wrap)

    def _block_key(self, event: tk.Event) -> str | None:
        modifiers = event.state & (0x0004 | 0x0008 | 0x0100)
        if modifiers and event.keysym.lower() in {"c", "a"}:
            return None
        return "break"

    def _select_all(self, _event: tk.Event) -> str:
        self._text.tag_add("sel", "1.0", "end-1c")
        return "break"

    def _hide_placeholder(self) -> None:
        if self._placeholder.winfo_manager():
            self._placeholder.place_forget()

    def _show_placeholder(self) -> None:
        self._placeholder.place(relx=0.5, rely=0.5, anchor="center")
        self._placeholder.lift()
        self._has_content = False

    def _write_newline(self) -> None:
        self._hide_placeholder()
        self._text.insert("end", "\n")
        self._text.see("end")

    def _write_segments(self, segments: list[tuple[str, tuple[str, ...]]]) -> None:
        self._hide_placeholder()
        for text, tags in segments:
            self._text.insert("end", text, tags)
        self._text.insert("end", "\n")
        self._text.see("end")
        self._has_content = True

    def _append_line(self, line: str) -> None:
        if not line:
            if self._has_content:
                self._write_newline()
            return

        segments = self._format_segments(line)
        self._write_segments(segments)

    def _format_segments(self, line: str) -> list[tuple[str, tuple[str, ...]]]:
        segments: list[tuple[str, tuple[str, ...]]] = []
        remaining = line

        if line.startswith("[") and "]" in line[:24]:
            closing = line.find("]")
            if 0 < closing < len(line):
                timestamp = line[: closing + 1]
                remaining = line[closing + 1 :].lstrip()
                segments.append((timestamp + " ", ("timestamp",)))

        severity = self._detect_severity(remaining or line)
        if remaining:
            segments.append((remaining, (severity,)))
        elif not segments:
            segments.append((line, (severity,)))
        return segments

    def _detect_severity(self, message: str) -> str:
        normalized = self._normalize_text(message)
        if any(keyword in normalized for keyword in {"error", "traceback", "exception", "fallo", "failed", "critico", "critical"}):
            return "error"
        if any(keyword in normalized for keyword in {"warn", "advertencia", "precaucion", "cuidado", "alerta"}):
            return "warning"
        if any(keyword in normalized for keyword in {"exito", "exitoso", "completado", "generado", "finalizado", "listo", "correctamente", "satisfactorio", "hecho", "ok"}):
            return "success"
        return "info"

    @staticmethod
    def _normalize_text(value: str) -> str:
        decomposed = unicodedata.normalize("NFKD", value)
        return "".join(char for char in decomposed if not unicodedata.combining(char)).casefold()

    # --------------------------------------------------------------- Public --
    def clear(self) -> None:
        self._text.delete("1.0", "end")
        self._text.yview_moveto(0.0)
        self._show_placeholder()

    def append(self, message: str) -> None:
        if message == "":
            self._append_line("")
            return

        lines = message.splitlines()
        if not lines:
            self._append_line(message.rstrip())
            return

        for raw in lines:
            self._append_line(raw.rstrip())

    @property
    def has_content(self) -> bool:
        return self._has_content

class RentApp(tk.Tk):
    """Ventana principal del panel de control."""

    def __init__(self) -> None:
        """Configura la ventana principal y prepara las dependencias comunes."""

        super().__init__()
        load_env()
        self.title("Rentabilidad - Panel de control")
        self.geometry("960x720")
        self.minsize(860, 640)

        self.colors = {
            "background": "#f1f5f9",
            "surface": "#ffffff",
            "surface_alt": "#f8fafc",
            "accent": "#2563eb",
            "accent_hover": "#1d4ed8",
            "accent_soft": "#60a5fa",
            "border": "#e2e8f0",
            "text": "#0f172a",
            "muted": "#64748b",
            "status_bg": "#eef2ff",
            "terminal_bg": "#0f172a",
            "terminal_border": "#1f2937",
            "terminal_fg": "#e2e8f0",
            "terminal_muted": "#475569",
            "terminal_cursor": "#38bdf8",
            "success": "#22c55e",
            "warning": "#f97316",
            "error": "#ef4444",
            "icon_auto_bg": "#dbeafe",
            "icon_auto_fg": "#1d4ed8",
            "icon_manual_bg": "#fef9c3",
            "icon_manual_fg": "#b45309",
            "icon_products_bg": "#dcfce7",
            "icon_products_fg": "#15803d",
            "icon_log_bg": "#bfdbfe",
            "icon_log_fg": "#1d4ed8",
        }
        self.configure(bg=self.colors["background"])

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._current_task: threading.Thread | None = None
        self._action_buttons: list[ttk.Button] = []
        self.context: PathContext = PathContextFactory(os.environ).create()
        self.template_path_var = tk.StringVar(value=str(self.context.template_path()))

        self.status_icon_var = tk.StringVar(value="✅")
        self.status_var = tk.StringVar(value="Sistema listo")
        self.last_update_var = tk.StringVar(value="Última actualización: --:--:--")
        self.manual_date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.products_date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))

        self._header_canvas: tk.Canvas | None = None
        self._header_title_id: int | None = None
        self._header_subtitle_id: int | None = None
        self._header_gradient_image: tk.PhotoImage | None = None
        self._header_gradient_id: int | None = None
        self._header_action_badges: list[int] = []
        self._header_action_icons: list[int] = []
        self._header_action_size = 36

        self._status_icon_label: tk.Label | None = None
        self._status_message_label: ttk.Label | None = None

        self.log_panel: TerminalPanel | None = None

        self._build_styles()
        self._build_layout()
        self._update_clock()
        self.after(150, self._poll_log_queue)

    # ------------------------------------------------------------------ UI --
    def _build_styles(self) -> None:
        """Define la paleta de estilos reutilizada por todos los controles."""

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:  # pragma: no cover - depende del sistema
            pass

        available_fonts = {name.casefold() for name in tkfont.families()}
        if "inter" in available_fonts:
            font_family = "Inter"
        elif sys.platform.startswith("win"):
            font_family = "Segoe UI"
        else:
            font_family = "Helvetica"
        self._font_family = font_family

        monospace_candidates = [
            "Cascadia Mono",
            "Fira Code",
            "JetBrains Mono",
            "Source Code Pro",
            "Consolas",
            "Menlo",
            "DejaVu Sans Mono",
            "Courier New",
        ]
        monospace = next(
            (candidate for candidate in monospace_candidates if candidate.casefold() in available_fonts),
            "Consolas" if sys.platform.startswith("win") else "Menlo",
        )
        self._mono_family = monospace

        emoji_candidates = [
            "Segoe UI Emoji",
            "Noto Color Emoji",
            "Apple Color Emoji",
            "Twemoji Mozilla",
        ]
        emoji_font = next(
            (candidate for candidate in emoji_candidates if candidate.casefold() in available_fonts),
            font_family,
        )
        self._emoji_font = emoji_font

        style.configure("TFrame", background=self.colors["surface"])
        style.configure("Background.TFrame", background=self.colors["background"])
        style.configure("Tab.TFrame", background=self.colors["surface"])
        style.configure("CardInner.TFrame", background=self.colors["surface"])

        style.configure("TLabel", background=self.colors["surface"], foreground=self.colors["text"])
        style.configure(
            "Body.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["text"],
            font=(font_family, 11),
        )
        style.configure(
            "BodyMuted.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            font=(font_family, 10),
        )
        style.configure(
            "FormLabel.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            font=(font_family, 10, "bold"),
        )
        style.configure(
            "SectionHeading.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["text"],
            font=(font_family, 13, "bold"),
        )
        style.configure(
            "Code.TLabel",
            background=self.colors["surface_alt"],
            foreground=self.colors["muted"],
            font=(self._mono_family, 10),
        )
        style.configure(
            "StatusMessage.TLabel",
            background=self.colors["status_bg"],
            foreground=self.colors["text"],
            font=(font_family, 11, "bold"),
        )
        style.configure(
            "StatusTime.TLabel",
            background=self.colors["status_bg"],
            foreground=self.colors["muted"],
            font=(font_family, 10),
        )

        style.configure(
            "Accent.TButton",
            font=(font_family, 11, "bold"),
            padding=(18, 10),
            background=self.colors["accent"],
            foreground="#ffffff",
            borderwidth=0,
            focusthickness=3,
            focuscolor=self.colors["accent"],
        )
        style.map(
            "Accent.TButton",
            background=[
                ("pressed", self.colors["accent_hover"]),
                ("active", self.colors["accent_hover"]),
                ("disabled", self.colors["border"]),
            ],
            foreground=[("disabled", self.colors["muted"])],
        )
        style.configure(
            "Secondary.TButton",
            font=(font_family, 11),
            padding=(16, 8),
            background=self.colors["surface_alt"],
            foreground=self.colors["text"],
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("pressed", self.colors["background"]),
                ("active", self.colors["background"]),
                ("disabled", self.colors["border"]),
            ],
            foreground=[("disabled", self.colors["muted"])],
        )
        style.configure(
            "Link.TButton",
            font=(font_family, 10, "bold"),
            padding=0,
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            borderwidth=0,
        )
        style.map(
            "Link.TButton",
            background=[("active", self.colors["surface"])],
            foreground=[("active", self.colors["accent"])],
        )

        style.configure("Card.TNotebook", background=self.colors["surface"], borderwidth=0, padding=0)
        style.configure(
            "Card.TNotebook.Tab",
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            font=(font_family, 11, "bold"),
            padding=(20, 12),
        )
        style.map(
            "Card.TNotebook.Tab",
            background=[("selected", self.colors["surface_alt"]), ("active", self.colors["surface"])],
            foreground=[("selected", self.colors["accent"])],
        )

        style.configure(
            "Card.TLabelframe",
            background=self.colors["surface"],
            borderwidth=0,
            padding=18,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.colors["surface"],
            foreground=self.colors["text"],
            font=(font_family, 12, "bold"),
        )

        style.configure(
            "Filled.TEntry",
            fieldbackground=self.colors["surface_alt"],
            background=self.colors["surface_alt"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["accent"],
            darkcolor=self.colors["border"],
            padding=8,
        )
        style.map(
            "Filled.TEntry",
            fieldbackground=[("focus", "#ffffff")],
            bordercolor=[("focus", self.colors["accent"])],
        )

        style.configure(
            "Vertical.TScrollbar",
            background=self.colors["surface"],
            troughcolor=self.colors["surface_alt"],
            bordercolor=self.colors["border"],
            arrowcolor=self.colors["muted"],
        )
        style.map("Vertical.TScrollbar", background=[("active", self.colors["accent"])])

    def _create_icon_badge(
        self,
        parent: tk.Widget,
        *,
        icon: str,
        badge_bg: str,
        icon_fg: str,
        background: Optional[str] = None,
    ) -> tk.Canvas:
        """Genera un ícono circular con fondo suave para títulos de secciones."""

        canvas_bg = background or self.colors["surface"]
        canvas = tk.Canvas(
            parent,
            width=48,
            height=48,
            bd=0,
            highlightthickness=0,
            bg=canvas_bg,
        )
        canvas.create_oval(0, 0, 48, 48, fill=badge_bg, outline="")
        canvas.create_text(
            24,
            24,
            text=icon,
            fill=icon_fg,
            font=(self._font_family, 20),
        )
        return canvas

    def _create_card_header(
        self,
        parent: tk.Widget,
        *,
        icon: str,
        badge_bg: str,
        icon_fg: str,
        title: str,
        subtitle: Optional[str] = None,
        wraplength: int = 360,
        background: Optional[str] = None,
    ) -> tk.Frame:
        """Crea una cabecera con icono y descripción para una tarjeta."""

        bg_color = background or self.colors["surface"]
        header = tk.Frame(parent, bg=bg_color, bd=0, highlightthickness=0)
        header.columnconfigure(1, weight=1)

        icon_canvas = self._create_icon_badge(
            header,
            icon=icon,
            badge_bg=badge_bg,
            icon_fg=icon_fg,
            background=bg_color,
        )
        span = 2 if subtitle else 1
        icon_canvas.grid(row=0, column=0, rowspan=span, sticky="nw")

        ttk.Label(header, text=title, style="SectionHeading.TLabel").grid(row=0, column=1, sticky="w")
        if subtitle:
            ttk.Label(
                header,
                text=subtitle,
                style="BodyMuted.TLabel",
                wraplength=wraplength,
            ).grid(row=1, column=1, sticky="w", pady=(6, 0))

        return header

    def _build_layout(self) -> None:
        """Arma la estructura base de pestañas, registro y barra de estado."""

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        container = ttk.Frame(self, padding=32, style="Background.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        self._build_header(container)

        card_outer = tk.Frame(
            container,
            background=self.colors["surface"],
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["border"],
            highlightthickness=1,
            bd=0,
        )
        card_outer.grid(row=1, column=0, sticky="nsew", pady=(24, 0))
        card_outer.columnconfigure(0, weight=1)
        card_outer.rowconfigure(0, weight=1)

        card = ttk.Frame(card_outer, padding=24, style="CardInner.TFrame")
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        notebook = ttk.Notebook(card, style="Card.TNotebook")
        notebook.grid(row=0, column=0, sticky="nsew")

        report_tab = ttk.Frame(notebook, padding=20, style="Tab.TFrame")
        report_tab.columnconfigure(0, weight=1)
        notebook.add(report_tab, text="📄 Informe de Rentabilidad")
        self._build_report_tab(report_tab)

        products_tab = ttk.Frame(notebook, padding=20, style="Tab.TFrame")
        products_tab.columnconfigure(0, weight=1)
        notebook.add(products_tab, text="📦 Listado de Productos")
        self._build_products_tab(products_tab)

        ttk.Separator(card, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=(16, 12))

        log_card = RoundedCard(
            card,
            background=self.colors["surface"],
            border=self.colors["border"],
            radius=20,
            padding=24,
        )
        log_card.grid(row=2, column=0, sticky="nsew")
        log_card.inner.columnconfigure(0, weight=1)
        log_card.inner.rowconfigure(1, weight=1)

        header = tk.Frame(log_card.inner, bg=self.colors["surface"], bd=0, highlightthickness=0)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(1, weight=1)

        icon_canvas = self._create_icon_badge(
            header,
            icon="📋",
            badge_bg=self.colors["icon_log_bg"],
            icon_fg=self.colors["icon_log_fg"],
            background=self.colors["surface"],
        )
        icon_canvas.grid(row=0, column=0, sticky="nw")

        ttk.Label(header, text="Registro de actividades", style="SectionHeading.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
            padx=(12, 0),
        )
        clear_button = ttk.Button(header, text="Limpiar", command=self._clear_log, style="Link.TButton")
        clear_button.grid(row=0, column=2, sticky="e")
        clear_button.configure(cursor="hand2")

        terminal_card = RoundedCard(
            log_card.inner,
            background=self.colors["terminal_bg"],
            border=self.colors["terminal_border"],
            radius=18,
            padding=8,
        )
        terminal_card.grid(row=1, column=0, sticky="nsew")
        terminal_card.inner.columnconfigure(0, weight=1)
        terminal_card.inner.rowconfigure(0, weight=1)

        self.log_panel = TerminalPanel(
            terminal_card.inner,
            background=self.colors["terminal_bg"],
            text_color=self.colors["terminal_fg"],
            muted_color=self.colors["terminal_muted"],
            accent_color=self.colors["accent_soft"],
            success_color=self.colors["success"],
            warning_color=self.colors["warning"],
            error_color=self.colors["error"],
            cursor_color=self.colors["terminal_cursor"],
            font_family=self._font_family,
            monospace_family=self._mono_family,
        )
        self.log_panel.grid(row=0, column=0, sticky="nsew")

        ttk.Separator(card, orient="horizontal").grid(row=3, column=0, sticky="ew", pady=(18, 14))

        status_card = RoundedCard(
            card,
            background=self.colors["status_bg"],
            border=self.colors["border"],
            radius=14,
            padding=16,
        )
        status_card.grid(row=4, column=0, sticky="ew")
        status_card.inner.columnconfigure(1, weight=1)

        self._status_icon_label = tk.Label(
            status_card.inner,
            textvariable=self.status_icon_var,
            font=(self._emoji_font, 18),
            bg=self.colors["status_bg"],
            fg=self.colors["success"],
        )
        self._status_icon_label.grid(row=0, column=0, sticky="w")

        self._status_message_label = ttk.Label(
            status_card.inner,
            textvariable=self.status_var,
            style="StatusMessage.TLabel",
        )
        self._status_message_label.grid(row=0, column=1, sticky="w", padx=(10, 0))

        ttk.Label(
            status_card.inner,
            textvariable=self.last_update_var,
            style="StatusTime.TLabel",
        ).grid(row=0, column=2, sticky="e")

        self._render_empty_log()
        self._set_status("Sistema listo", icon="✅", tone="success")

    def _build_header(self, parent: ttk.Frame) -> None:
        """Crea la cabecera con el degradado principal."""

        header_container = tk.Frame(parent, bg=self.colors["background"], bd=0, highlightthickness=0)
        header_container.grid(row=0, column=0, sticky="ew")
        header_container.columnconfigure(0, weight=1)

        canvas = tk.Canvas(
            header_container,
            height=140,
            highlightthickness=0,
            bd=0,
            bg=self.colors["background"],
        )
        canvas.grid(row=0, column=0, sticky="ew")
        canvas.bind("<Configure>", self._on_header_configure)

        self._header_canvas = canvas
        self._header_title_id = canvas.create_text(
            32,
            60,
            anchor="w",
            text="Panel de rentabilidad",
            font=(self._font_family, 24, "bold"),
            fill="#ffffff",
        )
        self._header_subtitle_id = canvas.create_text(
            32,
            96,
            anchor="w",
            text="Control y automatización de procesos",
            font=(self._font_family, 12),
            fill="#e0e7ff",
        )

        self._header_action_badges.clear()
        self._header_action_icons.clear()
        action_specs = [
            ("⚙️", "#1d4ed8"),
            ("🔔", "#1e40af"),
        ]
        for emoji, fill in action_specs:
            size = self._header_action_size
            badge_id = canvas.create_oval(
                0,
                0,
                size,
                size,
                fill=fill,
                outline="",
                tags=("header_action",),
            )
            icon_id = canvas.create_text(
                0,
                0,
                text=emoji,
                font=(self._emoji_font, 16),
                fill="#f8fafc",
                tags=("header_action",),
            )
            self._header_action_badges.append(badge_id)
            self._header_action_icons.append(icon_id)

        canvas.tag_bind("header_action", "<Enter>", lambda _event: canvas.configure(cursor="hand2"))
        canvas.tag_bind("header_action", "<Leave>", lambda _event: canvas.configure(cursor="arrow"))


    def _on_header_configure(self, event: tk.Event) -> None:
        """Redibuja el degradado y reposiciona los elementos al cambiar de tamaño."""

        if self._header_canvas is None:
            return

        width = max(event.width, 1)
        height = max(event.height, 1)
        self._draw_header_gradient(width, height)

        if self._header_title_id is not None:
            self._header_canvas.coords(self._header_title_id, 32, height / 2 - 18)
        if self._header_subtitle_id is not None:
            self._header_canvas.coords(self._header_subtitle_id, 32, height / 2 + 16)

        if self._header_action_badges and self._header_action_icons:
            size = self._header_action_size
            spacing = 12
            x_right = width - 32
            y_center = height / 2
            for index, (badge_id, icon_id) in enumerate(
                zip(self._header_action_badges, self._header_action_icons)
            ):
                x1 = x_right - index * (size + spacing)
                x0 = x1 - size
                y0 = y_center - size / 2
                y1 = y_center + size / 2
                self._header_canvas.coords(badge_id, x0, y0, x1, y1)
                self._header_canvas.coords(icon_id, x0 + size / 2, y_center)


    def _draw_header_gradient(self, width: int, height: int) -> None:
        """Pinta un degradado diagonal azul en la cabecera."""

        if self._header_canvas is None:
            return

        start = self._hex_to_rgb("#1366f1")
        end = self._hex_to_rgb("#1b5cf6")

        steps_width = max(width - 1, 1)
        steps_height = max(height - 1, 1)
        x_ratios = [x / steps_width for x in range(width)]
        y_ratios = [y / steps_height for y in range(height)]

        gradient_image = tk.PhotoImage(width=width, height=height)
        for y, y_ratio in enumerate(y_ratios):
            row_colors: list[str] = []
            for x_ratio in x_ratios:
                ratio = min(1.0, max(0.0, (x_ratio + y_ratio) / 2))
                r = int(start[0] + (end[0] - start[0]) * ratio)
                g = int(start[1] + (end[1] - start[1]) * ratio)
                b = int(start[2] + (end[2] - start[2]) * ratio)
                row_colors.append(f"#{r:02x}{g:02x}{b:02x}")
            gradient_image.put("{" + " ".join(row_colors) + "}", to=(0, y))

        if self._header_gradient_id is not None:
            self._header_canvas.delete(self._header_gradient_id)

        self._header_gradient_image = gradient_image
        self._header_gradient_id = self._header_canvas.create_image(0, 0, anchor="nw", image=gradient_image)
        self._header_canvas.tag_lower(self._header_gradient_id)

    @staticmethod
    def _hex_to_rgb(value: str) -> tuple[int, int, int]:
        """Convierte ``value`` de formato ``#RRGGBB`` a tupla RGB."""

        value = value.lstrip("#")
        return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))

    def _build_report_tab(self, parent: ttk.Frame) -> None:
        """Construye los controles relacionados con los informes de rentabilidad."""

        info_frame = ttk.Frame(parent, style="Tab.TFrame")
        info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        info_frame.columnconfigure(0, weight=1)

        ttk.Label(info_frame, text="📁 Plantilla base", style="SectionHeading.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        template_card = RoundedCard(
            info_frame,
            background=self.colors["surface_alt"],
            border=self.colors["border"],
            radius=12,
            padding=6,
        )
        template_card.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        template_card.inner.columnconfigure(0, weight=1)
        template_card.inner.columnconfigure(1, weight=0)
        template_card.inner.columnconfigure(2, weight=0)

        self.template_path_var.set(str(self.context.template_path()))
        template_entry = ttk.Entry(
            template_card.inner,
            textvariable=self.template_path_var,
            style="Filled.TEntry",
            state="readonly",
            font=("Consolas", 10),
        )
        template_entry.grid(row=0, column=0, sticky="ew")

        copy_button = ttk.Button(
            template_card.inner,
            text="Copiar",
            style="Secondary.TButton",
            command=lambda: self._copy_to_clipboard(self.template_path_var.get()),
        )
        copy_button.grid(row=0, column=1, padx=(12, 8))
        copy_button.configure(cursor="hand2")

        open_button = ttk.Button(
            template_card.inner,
            text="Abrir carpeta",
            style="Secondary.TButton",
            command=self._open_template_folder,
        )
        open_button.grid(row=0, column=2)
        open_button.configure(cursor="hand2")

        parent.rowconfigure(2, weight=1)

        cards_frame = ttk.Frame(parent, style="Tab.TFrame")
        cards_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        cards_frame.columnconfigure(0, weight=1, uniform="report")
        cards_frame.columnconfigure(1, weight=1, uniform="report")
        cards_frame.rowconfigure(0, weight=1)

        auto_card = RoundedCard(
            cards_frame,
            background=self.colors["surface"],
            border=self.colors["border"],
            radius=18,
            padding=24,
        )
        auto_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        auto_inner = auto_card.inner
        auto_inner.columnconfigure(0, weight=1)

        auto_header = self._create_card_header(
            auto_inner,
            icon="🕒",
            badge_bg=self.colors["icon_auto_bg"],
            icon_fg=self.colors["icon_auto_fg"],
            title="Informe automático",
            subtitle=(
                "Genera automáticamente el informe del día anterior utilizando los archivos "
                "más recientes y actualizando todas las hojas de la plantilla."
            ),
            wraplength=320,
        )
        auto_header.grid(row=0, column=0, sticky="ew")

        auto_button = ttk.Button(
            auto_inner,
            text="Generar informe automático",
            style="Accent.TButton",
            command=self._on_generate_auto,
        )
        auto_button.grid(row=1, column=0, sticky="ew", pady=(20, 0))
        self._register_action(auto_button)

        manual_card = RoundedCard(
            cards_frame,
            background=self.colors["surface"],
            border=self.colors["border"],
            radius=18,
            padding=24,
        )
        manual_card.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        manual_inner = manual_card.inner
        manual_inner.columnconfigure(0, weight=1)

        manual_header = self._create_card_header(
            manual_inner,
            icon="📅",
            badge_bg=self.colors["icon_manual_bg"],
            icon_fg=self.colors["icon_manual_fg"],
            title="Informe manual",
            subtitle=(
                "Genera un informe para la fecha indicada utilizando los archivos cuyo nombre "
                "coincida con ese día."
            ),
            wraplength=320,
        )
        manual_header.grid(row=0, column=0, sticky="ew")

        form_row = ttk.Frame(manual_inner, style="CardInner.TFrame")
        form_row.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        form_row.columnconfigure(1, weight=1)

        ttk.Label(form_row, text="Fecha (YYYY-MM-DD):", style="FormLabel.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        entry = ttk.Entry(form_row, textvariable=self.manual_date_var, width=20, style="Filled.TEntry")
        entry.grid(row=0, column=1, sticky="w", padx=(10, 0))

        today_button = ttk.Button(
            form_row,
            text="Hoy",
            command=lambda: self.manual_date_var.set(date.today().strftime("%Y-%m-%d")),
            style="Secondary.TButton",
        )
        today_button.grid(row=0, column=2, padx=(12, 0))

        manual_button = ttk.Button(
            manual_inner,
            text="Generar informe manual",
            style="Accent.TButton",
            command=self._on_generate_manual,
        )
        manual_button.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        self._register_action(manual_button)

    def _build_products_tab(self, parent: ttk.Frame) -> None:
        """Configura la pestaña para generar listados de productos SIIGO."""

        info = ttk.Label(
            parent,
            text=(
                "Genera el listado de productos ejecutando ExcelSIIGO y limpiando el resultado "
                "con las credenciales configuradas."
            ),
            style="BodyMuted.TLabel",
            wraplength=660,
        )
        info.grid(row=0, column=0, sticky="w", pady=(0, 8))

        form_card = RoundedCard(
            parent,
            background=self.colors["surface"],
            border=self.colors["border"],
            radius=18,
            padding=24,
        )
        form_card.grid(row=1, column=0, sticky="ew", pady=(20, 0))
        form_card.inner.columnconfigure(0, weight=1)

        header = self._create_card_header(
            form_card.inner,
            icon="📋",
            badge_bg=self.colors["icon_products_bg"],
            icon_fg=self.colors["icon_products_fg"],
            title="Generación de listado",
            subtitle=(
                "La aplicación obtendrá la información desde SIIGO, filtrará las columnas "
                "necesarias y guardará el archivo listo para su revisión."
            ),
            wraplength=600,
        )
        header.grid(row=0, column=0, sticky="ew")

        form_row = ttk.Frame(form_card.inner, style="CardInner.TFrame")
        form_row.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        form_row.columnconfigure(1, weight=1)

        ttk.Label(form_row, text="Fecha (YYYY-MM-DD):", style="FormLabel.TLabel").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(form_row, textvariable=self.products_date_var, width=20, style="Filled.TEntry")
        entry.grid(row=0, column=1, sticky="w", padx=(10, 0))

        set_today = ttk.Button(
            form_row,
            text="Hoy",
            command=lambda: self.products_date_var.set(date.today().strftime("%Y-%m-%d")),
            style="Secondary.TButton",
        )
        set_today.grid(row=0, column=2, padx=(12, 0))

        button = ttk.Button(
            form_card.inner,
            text="Generar listado de productos",
            style="Accent.TButton",
            command=self._on_generate_products,
        )
        button.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        self._register_action(button)

    # -------------------------------------------------------------- Helpers --
    def _register_action(self, button: ttk.Button) -> None:
        """Mantiene una referencia a ``button`` para gestionar su estado conjunto."""

        button.configure(cursor="hand2")
        self._action_buttons.append(button)

    def _set_actions_state(self, state: str) -> None:
        """Activa o desactiva todos los botones de acción."""

        for btn in self._action_buttons:
            btn.state([state]) if state == "disabled" else btn.state(["!disabled"])

    def _update_clock(self) -> None:
        """Actualiza la marca de tiempo que se muestra en la barra de estado."""

        now = datetime.now().strftime("%H:%M:%S")
        self.last_update_var.set(f"Última actualización: {now}")
        self.after(1000, self._update_clock)

    def _poll_log_queue(self) -> None:
        """Transfiere los mensajes pendientes de la cola al registro visual."""

        try:
            while True:
                message = self._log_queue.get_nowait()
                self._append_log(message)
        except queue.Empty:
            pass
        finally:
            self.after(150, self._poll_log_queue)

    def _append_log(self, message: str) -> None:
        """Añade ``message`` al panel de registro manteniendo el estado vacío."""

        if self.log_panel is None:
            return
        self.log_panel.append(message)

    def _clear_log(self) -> None:
        """Elimina por completo el contenido del registro de actividades."""

        self._render_empty_log()
        self._set_status("Sistema listo", icon="✅", tone="success")

    def _render_empty_log(self) -> None:
        """Muestra un mensaje neutro cuando no hay actividades registradas."""

        if self.log_panel is None:
            return
        self.log_panel.clear()

    def _set_status(self, message: str, *, icon: str, tone: str) -> None:
        """Actualiza el indicador inferior aplicando el color según ``tone``."""

        color_map = {
            "success": self.colors["success"],
            "warning": self.colors["warning"],
            "error": self.colors["error"],
            "info": self.colors["accent"],
        }
        color = color_map.get(tone, self.colors["muted"])

        self.status_icon_var.set(icon)
        self.status_var.set(message)
        if self._status_icon_label is not None:
            self._status_icon_label.configure(fg=color)
        if self._status_message_label is not None:
            self._status_message_label.configure(foreground=color)

    def _copy_to_clipboard(self, value: str) -> None:
        """Copia ``value`` al portapapeles del sistema y notifica al usuario."""

        try:
            self.clipboard_clear()
            self.clipboard_append(value)
        except tk.TclError as error:  # pragma: no cover - depende del entorno
            messagebox.showerror("No se pudo copiar", str(error))
        else:
            self._set_status("Ruta copiada al portapapeles", icon="✅", tone="success")

    def _open_template_folder(self) -> None:
        """Abre la carpeta que contiene la plantilla base en el explorador."""

        path = self.context.template_path()
        directory = path.parent
        if not directory.exists():
            messagebox.showerror(
                "Carpeta no encontrada",
                f"No existe la carpeta de la plantilla: {directory}",
            )
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(directory)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(directory)])
            else:
                subprocess.Popen(["xdg-open", str(directory)])
        except FileNotFoundError:
            messagebox.showerror(
                "Comando no disponible",
                "No se encontró una forma compatible de abrir la carpeta automáticamente.",
            )
        except Exception as error:  # noqa: BLE001 - mostrar cualquier fallo
            messagebox.showerror("No se pudo abrir la carpeta", str(error))
        else:
            self._set_status("Carpeta de la plantilla abierta", icon="✅", tone="success")

    def _log(self, message: str) -> None:
        """Encola ``message`` con marca temporal para mostrarlo en pantalla."""

        timestamp = datetime.now().strftime("%H:%M:%S")
        self._log_queue.put(f"[{timestamp}] {message}")

    # -------------------------------------------------------------- Actions --
    def _on_generate_auto(self) -> None:
        """Genera el informe del día anterior reutilizando los EXCZ más recientes."""

        target_date = date.today() - timedelta(days=1)
        self._start_task(
            f"Generando informe automático del {target_date:%Y-%m-%d}",
            lambda: self._task_generate_report(target_date, use_latest=True),
        )

    def _on_generate_manual(self) -> None:
        """Solicita al usuario una fecha específica y lanza la actualización."""

        raw = self.manual_date_var.get().strip()
        if not raw:
            messagebox.showerror("Fecha requerida", "Ingresa una fecha en formato YYYY-MM-DD.")
            return
        try:
            target_date = datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Formato inválido", "La fecha debe tener el formato YYYY-MM-DD.")
            return

        self._start_task(
            f"Generando informe para {target_date:%Y-%m-%d}",
            lambda: self._task_generate_report(target_date, use_latest=False),
        )

    def _on_generate_products(self) -> None:
        """Inicia la generación del listado de productos para la fecha elegida."""

        raw = self.products_date_var.get().strip()
        if raw:
            try:
                target_date = datetime.strptime(raw, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Formato inválido", "La fecha debe tener el formato YYYY-MM-DD.")
                return
        else:
            target_date = date.today()

        self._start_task(
            f"Generando listado de productos para {target_date:%Y-%m-%d}",
            lambda: self._task_generate_products(target_date),
        )

    # ------------------------------------------------------------ Task flow --
    def _start_task(self, status_message: str, task: Callable[[], TaskResult]) -> None:
        """Ejecuta ``task`` en un hilo aparte y actualiza los indicadores UI."""

        if self._current_task and self._current_task.is_alive():
            messagebox.showinfo("Tarea en curso", "Espera a que finalice la operación actual.")
            return

        self._set_status(status_message, icon="⏳", tone="info")
        if self.log_panel and self.log_panel.has_content:
            self._log_queue.put("")
        self._log(status_message)
        self._set_actions_state("disabled")

        def runner() -> None:
            try:
                result = task()
            except Exception as exc:  # noqa: BLE001 - mostrar cualquier error
                self._log(f"ERROR: {exc}")
                self.after(0, lambda: self._finish_task(False, str(exc), None))
            else:
                self.after(0, lambda: self._finish_task(True, result.message, result.output))

        thread = threading.Thread(target=runner, daemon=True)
        self._current_task = thread
        thread.start()

    def _finish_task(self, success: bool, message: str, output: Optional[Path]) -> None:
        """Restaura el estado de la interfaz y comunica el resultado al usuario."""

        self._current_task = None
        self._set_actions_state("normal")
        if success:
            self._set_status(message, icon="✅", tone="success")
            if output:
                self._log(f"Archivo generado: {output}")
            messagebox.showinfo("Proceso completado", message)
        else:
            self._set_status(message, icon="⚠️", tone="error")
            messagebox.showerror("Ocurrió un problema", message)

    # ----------------------------------------------------------- Operations --
    def _task_generate_report(self, target_date: date, *, use_latest: bool) -> TaskResult:
        """Clona la plantilla y ejecuta el loader para ``target_date``."""

        template = self.context.template_path()
        if not template.exists():
            raise FileNotFoundError(f"No existe la plantilla base: {template}")

        output_dir = self.context.informe_month_dir(target_date)
        output_path = output_dir / self.context.informe_filename(target_date)
        self._log(f"Clonando plantilla hacia {output_path}")
        shutil.copyfile(template, output_path)

        self._log("Ejecutando loader de rentabilidad...")
        return_code = self._run_loader(output_path, target_date, use_latest=use_latest)
        if return_code != 0:
            raise RuntimeError(f"El loader finalizó con código {return_code}")

        return TaskResult(
            message=f"Informe actualizado correctamente ({output_path.name})",
            output=output_path,
        )

    def _run_loader(self, excel_path: Path, target_date: date, *, use_latest: bool) -> int:
        """Lanza ``hoja01_loader.py`` y reenvía sus mensajes al registro local."""

        if not LOADER_SCRIPT.exists():
            raise FileNotFoundError(f"No se encuentra el script: {LOADER_SCRIPT}")

        cmd = [
            sys.executable,
            str(LOADER_SCRIPT),
            "--excel",
            str(excel_path),
            "--fecha",
            target_date.isoformat(),
        ]
        if use_latest:
            cmd.append("--use-latest-sources")

        env = os.environ.copy()
        process = subprocess.Popen(  # noqa: S603, S607 - ejecución controlada
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env,
            cwd=str(REPO_ROOT),
        )

        assert process.stdout is not None
        for line in process.stdout:  # pragma: no cover - flujo interactivo
            self._log(line.rstrip())
        return process.wait()

    def _collect_product_settings(self) -> tuple[PathContext, ProductGenerationConfig]:
        """Obtiene contexto y configuración para ``ProductListingService``."""

        context = PathContextFactory(os.environ).create()
        defaults = {
            "SIIGO_DIR": os.environ.get("SIIGO_DIR", r"C:\\Siigo"),
            "SIIGO_BASE": os.environ.get("SIIGO_BASE", r"D:\\SIIWI01"),
            "SIIGO_LOG": os.environ.get(
                "SIIGO_LOG",
                str(Path(os.environ.get("SIIGO_BASE", r"D:\\SIIWI01")) / "LOGS" / "log_catalogos.txt"),
            ),
            "SIIGO_REPORTE": os.environ.get("SIIGO_REPORTE", "GETINV"),
            "SIIGO_EMPRESA": os.environ.get("SIIGO_EMPRESA", "L"),
            "SIIGO_USUARIO": os.environ.get("SIIGO_USUARIO", "JUAN"),
            "SIIGO_CLAVE": os.environ.get("SIIGO_CLAVE", "0110"),
            "SIIGO_ESTADO_PARAM": os.environ.get("SIIGO_ESTADO_PARAM", "S"),
            "SIIGO_RANGO_INI": os.environ.get("SIIGO_RANGO_INI", "0010001000001"),
            "SIIGO_RANGO_FIN": os.environ.get("SIIGO_RANGO_FIN", "0400027999999"),
            "SIIGO_ACTIVO_COL": os.environ.get("SIIGO_ACTIVO_COL", "AX"),
            "PRODUCTOS_DIR": os.environ.get("PRODUCTOS_DIR", str(context.productos_dir)),
        }

        productos_dir = Path(defaults["PRODUCTOS_DIR"])
        context = PathContext(
            base_dir=context.base_dir,
            productos_dir=productos_dir,
            informes_dir=context.informes_dir,
        )
        context.ensure_structure()

        siigo_dir = Path(defaults["SIIGO_DIR"])
        if not siigo_dir.exists():
            raise FileNotFoundError(f"No existe la carpeta de SIIGO: {siigo_dir}")

        credentials = SiigoCredentials(
            reporte=defaults["SIIGO_REPORTE"],
            empresa=defaults["SIIGO_EMPRESA"],
            usuario=defaults["SIIGO_USUARIO"],
            clave=defaults["SIIGO_CLAVE"],
            estado_param=defaults["SIIGO_ESTADO_PARAM"],
            rango_ini=defaults["SIIGO_RANGO_INI"],
            rango_fin=defaults["SIIGO_RANGO_FIN"],
        )

        activo_col = defaults["SIIGO_ACTIVO_COL"]
        keep_columns = KEEP_COLUMN_NUMBERS + (column_index_from_string(activo_col),)

        config = ProductGenerationConfig(
            siigo_dir=siigo_dir,
            base_path=ensure_trailing_backslash(defaults["SIIGO_BASE"]),
            log_path=defaults["SIIGO_LOG"],
            credentials=credentials,
            activo_column=activo_col,
            keep_columns=keep_columns,
        )
        return context, config

    def _task_generate_products(self, target_date: date) -> TaskResult:
        """Genera el Excel de productos y devuelve su ubicación final."""

        context, config = self._collect_product_settings()
        service = ProductListingService(context, config)

        log_stream = LogStream(self._log)
        with redirect_stdout(log_stream), redirect_stderr(log_stream):
            output_path = service.generate(target_date)

        return TaskResult(
            message=f"Listado de productos generado ({output_path.name})",
            output=output_path,
        )


def main() -> None:
    """Inicia la aplicación gráfica y entra en el bucle principal de Tk."""

    app = RentApp()
    app.mainloop()


if __name__ == "__main__":  # pragma: no cover - punto de entrada manual
    main()
