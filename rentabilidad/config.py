from __future__ import annotations

import os
from pathlib import Path

from openpyxl.utils import column_index_from_string

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContext, PathContextFactory
from rentabilidad.services.products import (
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
)
from servicios.generar_listado_productos import KEEP_COLUMN_NUMBERS

from .infra.logging_bus import EventBus


def _ensure_trailing_backslash(path: str) -> str:
    return path if path.endswith(("\\", "/")) else path + "\\"


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
        keep_columns = KEEP_COLUMN_NUMBERS + (column_index_from_string(activo_column),)

        return ProductGenerationConfig(
            siigo_dir=siigo_dir,
            base_path=base_path,
            log_path=log_path,
            credentials=credenciales,
            activo_column=activo_column,
            keep_columns=keep_columns,
        )

    def build_product_service(self) -> ProductListingService:
        return ProductListingService(self.context, self._product_config)


bus = EventBus()
settings = Settings()
