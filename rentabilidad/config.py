from __future__ import annotations

import os
import re
from pathlib import Path

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContext, PathContextFactory
from rentabilidad.services.products import (
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
    resolve_column_index,
)
from servicios.generar_listado_productos import KEEP_COLUMN_NUMBERS

from .infra.logging_bus import EventBus


def _ensure_trailing_backslash(path: str) -> str:
    if not path:
        return "\\"

    stripped = path.rstrip("\\/")
    if not stripped:
        return path

    return stripped + "\\"


def _parse_required_files_env(value: str | None) -> tuple[str, ...] | None:
    if value is None:
        return None

    parts = [part.strip() for part in re.split(r"[;,]", value) if part.strip()]
    return tuple(parts)


def _read_float_env(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


class Settings:
    def __init__(self) -> None:
        load_env()
        self.context: PathContext = PathContextFactory(os.environ).create()

        self.ruta_plantilla: Path = self.context.template_path()
        self.plantilla_hoja: str | None = os.environ.get("PLANTILLA_HOJA")

        self.excz_dir: Path = Path(os.environ.get("EXCZDIR", r"D:\\SIIWI01\\LISTADOS"))
        self.excz_prefix: str = os.environ.get("EXCZPREFIX", "EXCZ980")
        self.excz_sheet: str = os.environ.get("EXCZ_SHEET", "Hoja1")

        self._product_config = self._build_product_config()

    def _build_product_config(self) -> ProductGenerationConfig:
        siigo_dir = Path(os.environ.get("SIIGO_DIR", r"C:\\Siigo"))
        base_path = _ensure_trailing_backslash(os.environ.get("SIIGO_BASE", r"D:\\SIIWI01"))
        log_path = os.environ.get("SIIGO_LOG", str(Path(base_path.rstrip("\\/")) / "LOGS" / "log_catalogos.txt"))

        credenciales = SiigoCredentials(
            reporte=os.environ.get("SIIGO_REPORTE", "GETINV"),
            empresa=os.environ.get("SIIGO_EMPRESA", "L"),
            usuario=os.environ.get("SIIGO_USUARIO", "JUAN"),
            clave=os.environ.get("SIIGO_CLAVE", "0110"),
            estado_param=os.environ.get("SIIGO_ESTADO_PARAM", "S"),
            rango_ini=os.environ.get("SIIGO_RANGO_INI", "0010001000001"),
            rango_fin=os.environ.get("SIIGO_RANGO_FIN", "0400027999999"),
        )

        activo_column = os.environ.get("SIIGO_ACTIVO_COL", "AX")
        keep_columns = KEEP_COLUMN_NUMBERS + (resolve_column_index(activo_column),)

        batch_script_env = os.environ.get("SIIGO_BATCH", "GenerarListadoProductos.bat")
        batch_script: Path | None = None
        if batch_script_env:
            candidate = Path(batch_script_env)
            if not candidate.is_absolute():
                candidate = self.context.base_dir / candidate
            batch_script = candidate

        required_files_override = _parse_required_files_env(
            os.environ.get("SIIGO_REQUIRED_FILES")
        )

        config_kwargs: dict[str, object] = {}
        if required_files_override is not None:
            config_kwargs["required_files"] = required_files_override

        return ProductGenerationConfig(
            siigo_dir=siigo_dir,
            base_path=base_path,
            log_path=log_path,
            credentials=credenciales,
            activo_column=activo_column,
            keep_columns=keep_columns,
            siigo_command=os.environ.get("SIIGO_COMMAND", "ExcelSIIGO.exe"),
            siigo_output_filename=os.environ.get("SIIGO_OUTPUT_FILENAME", "ProductosMesDia.xlsx"),
            wait_timeout=_read_float_env("SIIGO_WAIT_TIMEOUT", 60.0),
            wait_interval=_read_float_env("SIIGO_WAIT_INTERVAL", 0.2),
            batch_script=batch_script,
            **config_kwargs,
        )

    def build_product_service(self) -> ProductListingService:
        return ProductListingService(self.context, self._product_config)


bus = EventBus()
settings = Settings()
