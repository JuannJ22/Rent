"""Aplicación gráfica para controlar procesos de rentabilidad."""

from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import threading
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Optional

import tkinter as tk
from tkinter import messagebox, ttk

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
            "background": "#0f172a",
            "surface": "#111f30",
            "surface_alt": "#1c2a44",
            "accent": "#6366f1",
            "accent_hover": "#818cf8",
            "border": "#1f2a3d",
            "text": "#f9fafb",
            "muted": "#c7d2fe",
            "status_bg": "#0b1628",
            "log_bg": "#0b1220",
            "log_fg": "#e2e8f0",
        }
        self.configure(bg=self.colors["background"])

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._current_task: threading.Thread | None = None
        self._action_buttons: list[ttk.Button] = []
        self.context: PathContext = PathContextFactory(os.environ).create()

        self.status_var = tk.StringVar(value="Listo")
        self.manual_date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.products_date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))

        self._build_styles()
        self._build_layout()
        self.after(150, self._poll_log_queue)

    # ------------------------------------------------------------------ UI --
    def _build_styles(self) -> None:
        """Define la paleta de estilos reutilizada por todos los controles."""

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:  # pragma: no cover - depende del sistema
            pass

        font_family = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"
        style.configure("TFrame", background=self.colors["surface"])
        style.configure("Background.TFrame", background=self.colors["background"])
        style.configure("Tab.TFrame", background=self.colors["surface"])
        style.configure("CardInner.TFrame", background=self.colors["surface"])

        style.configure("TLabel", background=self.colors["surface"], foreground=self.colors["text"])
        style.configure(
            "Header.TLabel",
            background=self.colors["background"],
            foreground=self.colors["text"],
            font=(font_family, 22, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.colors["background"],
            foreground=self.colors["muted"],
            font=(font_family, 12),
        )
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
            "Status.TLabel",
            background=self.colors["status_bg"],
            foreground=self.colors["muted"],
            font=(font_family, 10),
        )

        style.configure(
            "Accent.TButton",
            font=(font_family, 11, "bold"),
            padding=(18, 10),
            background=self.colors["accent"],
            foreground=self.colors["text"],
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
            padding=(14, 8),
            background=self.colors["surface_alt"],
            foreground=self.colors["text"],
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("pressed", self.colors["surface"]),
                ("active", self.colors["surface"]),
                ("disabled", self.colors["surface"]),
            ],
            foreground=[("disabled", self.colors["muted"])],
        )

        style.configure("Card.TNotebook", background=self.colors["background"], borderwidth=0, padding=0)
        style.configure(
            "Card.TNotebook.Tab",
            background=self.colors["surface_alt"],
            foreground=self.colors["muted"],
            font=(font_family, 11, "bold"),
            padding=(20, 12),
        )
        style.map(
            "Card.TNotebook.Tab",
            background=[("selected", self.colors["accent"]), ("active", self.colors["surface_alt"])],
            foreground=[("selected", self.colors["text"])],
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
            fieldbackground=[("focus", self.colors["background"])],
            bordercolor=[("focus", self.colors["accent"])],
        )

        style.configure(
            "Vertical.TScrollbar",
            background=self.colors["surface"],
            troughcolor=self.colors["surface"],
            bordercolor=self.colors["border"],
            arrowcolor=self.colors["text"],
        )
        style.map("Vertical.TScrollbar", background=[("active", self.colors["accent"])])

    def _build_layout(self) -> None:
        """Arma la estructura base de pestañas, registro y barra de estado."""

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        container = ttk.Frame(self, padding=32, style="Background.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=3)
        container.rowconfigure(3, weight=2)

        title = ttk.Label(container, text="Panel de automatización", style="Header.TLabel")
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            container,
            text="Gestiona informes de rentabilidad y listados de productos desde una interfaz amigable.",
            style="Subtitle.TLabel",
            wraplength=760,
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(12, 28))

        notebook = ttk.Notebook(container, style="Card.TNotebook")
        notebook.grid(row=2, column=0, sticky="nsew", pady=(0, 24))

        report_tab = ttk.Frame(notebook, padding=24, style="Tab.TFrame")
        report_tab.columnconfigure(0, weight=1)
        notebook.add(report_tab, text="Informe de rentabilidad")
        self._build_report_tab(report_tab)

        products_tab = ttk.Frame(notebook, padding=24, style="Tab.TFrame")
        products_tab.columnconfigure(0, weight=1)
        notebook.add(products_tab, text="Listado de productos")
        self._build_products_tab(products_tab)

        log_frame = ttk.LabelFrame(container, text="Registro de actividades", style="Card.TLabelframe")
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=12,
            state="disabled",
            wrap="word",
            background=self.colors["log_bg"],
            foreground=self.colors["log_fg"],
            insertbackground=self.colors["text"],
            borderwidth=0,
            relief="flat",
            font=("Consolas", 10),
            padx=12,
            pady=12,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            style="Vertical.TScrollbar",
            command=self.log_text.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        actions_frame = ttk.Frame(log_frame, style="CardInner.TFrame")
        actions_frame.grid(row=1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(
            actions_frame,
            text="Limpiar",
            command=self._clear_log,
            style="Secondary.TButton",
        ).grid(row=0, column=0, padx=(0, 8))

        status_bar = ttk.Label(
            container,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor="w",
            padding=(0, 12),
        )
        status_bar.grid(row=4, column=0, sticky="ew", pady=(20, 0))

    def _build_report_tab(self, parent: ttk.Frame) -> None:
        """Construye los controles relacionados con los informes de rentabilidad."""

        info_frame = ttk.Frame(parent, style="Tab.TFrame")
        info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        info_frame.columnconfigure(0, weight=1)

        template_label = ttk.Label(
            info_frame,
            text=f"Plantilla base: {self.context.template_path()}",
            style="BodyMuted.TLabel",
            wraplength=700,
        )
        template_label.grid(row=0, column=0, sticky="w")

        auto_frame = ttk.LabelFrame(parent, text="Informe del día anterior", style="Card.TLabelframe")
        auto_frame.grid(row=1, column=0, sticky="ew")
        auto_frame.columnconfigure(0, weight=1)

        desc = ttk.Label(
            auto_frame,
            text=(
                "Genera automáticamente el informe del día anterior. "
                "La aplicación clonará la plantilla, localizará los EXCZ más recientes "
                "y actualizará todas las hojas correspondientes."
            ),
            style="BodyMuted.TLabel",
            wraplength=660,
        )
        desc.grid(row=0, column=0, sticky="w")

        auto_button = ttk.Button(
            auto_frame,
            text="Generar informe automático",
            style="Accent.TButton",
            command=self._on_generate_auto,
        )
        auto_button.grid(row=1, column=0, sticky="e", pady=(18, 0))
        self._register_action(auto_button)

        manual_frame = ttk.LabelFrame(parent, text="Informe por fecha específica", style="Card.TLabelframe")
        manual_frame.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        manual_frame.columnconfigure(1, weight=1)

        manual_desc = ttk.Label(
            manual_frame,
            text=(
                "Selecciona la fecha del informe y se utilizarán los archivos cuyo nombre "
                "contenga la fecha indicada."
            ),
            style="BodyMuted.TLabel",
            wraplength=660,
        )
        manual_desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(manual_frame, text="Fecha (YYYY-MM-DD):", style="FormLabel.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
        )
        entry = ttk.Entry(manual_frame, textvariable=self.manual_date_var, width=20, style="Filled.TEntry")
        entry.grid(row=1, column=1, sticky="w", padx=(10, 0))

        today_button = ttk.Button(
            manual_frame,
            text="Hoy",
            command=lambda: self.manual_date_var.set(date.today().strftime("%Y-%m-%d")),
            style="Secondary.TButton",
        )
        today_button.grid(row=1, column=2, padx=(12, 0))

        manual_button = ttk.Button(
            manual_frame,
            text="Generar informe",
            style="Accent.TButton",
            command=self._on_generate_manual,
        )
        manual_button.grid(row=2, column=0, columnspan=3, sticky="e", pady=(20, 0))
        self._register_action(manual_button)

    def _build_products_tab(self, parent: ttk.Frame) -> None:
        """Configura la pestaña para generar listados de productos SIIGO."""

        info = ttk.Label(
            parent,
            text=(
                "Crea el listado de productos ejecutando ExcelSIIGO y limpiando el resultado. "
                "Se emplearán las credenciales configuradas en las variables de entorno."
            ),
            style="BodyMuted.TLabel",
            wraplength=660,
        )
        info.grid(row=0, column=0, sticky="w")

        form = ttk.LabelFrame(parent, text="Generación de listado", style="Card.TLabelframe")
        form.grid(row=1, column=0, sticky="ew", pady=(20, 0))
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Fecha (YYYY-MM-DD):", style="FormLabel.TLabel").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(form, textvariable=self.products_date_var, width=20, style="Filled.TEntry")
        entry.grid(row=0, column=1, sticky="w", padx=(10, 0))

        set_today = ttk.Button(
            form,
            text="Hoy",
            command=lambda: self.products_date_var.set(date.today().strftime("%Y-%m-%d")),
            style="Secondary.TButton",
        )
        set_today.grid(row=0, column=2, padx=(12, 0))

        button = ttk.Button(
            form,
            text="Generar listado de productos",
            style="Accent.TButton",
            command=self._on_generate_products,
        )
        button.grid(row=1, column=0, columnspan=3, sticky="e", pady=(20, 0))
        self._register_action(button)

    # -------------------------------------------------------------- Helpers --
    def _register_action(self, button: ttk.Button) -> None:
        """Mantiene una referencia a ``button`` para gestionar su estado conjunto."""

        self._action_buttons.append(button)

    def _set_actions_state(self, state: str) -> None:
        """Activa o desactiva todos los botones de acción."""

        for btn in self._action_buttons:
            btn.state([state]) if state == "disabled" else btn.state(["!disabled"])

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
        """Añade ``message`` al cuadro de texto bloqueado de la interfaz."""

        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.insert("end", "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def _clear_log(self) -> None:
        """Elimina por completo el contenido del registro de actividades."""

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

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

        self.status_var.set(status_message)
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
            self.status_var.set(f"✅ {message}")
            if output:
                self._log(f"Archivo generado: {output}")
            messagebox.showinfo("Proceso completado", message)
        else:
            self.status_var.set(f"❌ {message}")
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
