"""Servicios para la generación y depuración del listado de productos."""

from __future__ import annotations

import subprocess
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from rentabilidad.core.paths import PathContext


@dataclass
class SiigoCredentials:
    """Agrupa los parámetros de autenticación para ExcelSIIGO."""

    reporte: str
    empresa: str
    usuario: str
    clave: str
    estado_param: str
    rango_ini: str
    rango_fin: str


@dataclass
class ProductGenerationConfig:
    """Configuración necesaria para generar el archivo de productos."""

    siigo_dir: Path
    base_path: str
    log_path: str
    credentials: SiigoCredentials
    activo_column: str
    keep_columns: Sequence[int]


class ExcelSiigoFacade:
    """Fachada que ejecuta ExcelSIIGO y gestiona los parámetros requeridos."""

    def __init__(self, config: ProductGenerationConfig):
        self._config = config

    def run(self, output_path: Path, year: str) -> None:
        command = [
            "ExcelSIIGO",
            self._config.base_path,
            year,
            self._config.credentials.reporte,
            self._config.credentials.empresa,
            self._config.credentials.usuario,
            self._config.credentials.clave,
            self._config.log_path,
            self._config.credentials.estado_param,
            self._config.credentials.rango_ini,
            self._config.credentials.rango_fin,
            str(output_path),
        ]

        result = subprocess.run(
            command,
            cwd=str(self._config.siigo_dir),
            check=False,
            capture_output=True,
            text=True,
        )

        if result.stdout:
            print(result.stdout.strip())
        if result.stderr:
            print(result.stderr.strip())

        if result.returncode != 0:
            raise RuntimeError(
                "ExcelSIIGO falló con código "
                f"{result.returncode}: {result.stderr.strip() or result.stdout.strip()}"
            )


class WorkbookCleaner:
    """Responsable de filtrar columnas y filas en el archivo generado."""

    def __init__(self, activo_column: str, keep_columns: Iterable[int]):
        self._activo_idx = column_index_from_string(activo_column)
        self._keep_columns = {int(idx) for idx in keep_columns}
        self._keep_columns.add(self._activo_idx)

    def clean(self, file_path: Path) -> None:
        wb = load_workbook(filename=file_path)
        ws = wb.active

        if ws.max_column < self._activo_idx:
            raise RuntimeError(
                "La hoja activa no tiene la columna requerida para estado del producto."
            )

        for row in range(ws.max_row, 1, -1):
            value = ws.cell(row=row, column=self._activo_idx).value
            if self._normalize(value) != "S":
                ws.delete_rows(row, 1)

        for col in range(ws.max_column, 0, -1):
            if col not in self._keep_columns:
                ws.delete_cols(col, 1)

        wb.save(file_path)

    @staticmethod
    def _normalize(value) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip().upper()
        return str(value).strip().upper()


@contextmanager
def safe_backup(path: Path):
    """Crea un ``.bak`` temporal y lo restaura si ocurre un error."""

    backup = None
    if path.exists():
        backup = path.with_suffix(path.suffix + ".bak")
        if backup.exists():
            backup.unlink()
        path.replace(backup)
    try:
        yield
    except Exception:
        if backup and backup.exists():
            if path.exists():
                path.unlink()
            backup.replace(path)
        raise
    else:
        if backup and backup.exists():
            backup.unlink()


class ProductListingService:
    """Servicio principal para crear y depurar el archivo de productos."""

    def __init__(
        self,
        context: PathContext,
        config: ProductGenerationConfig,
    ):
        self._context = context
        self._config = config
        self._facade = ExcelSiigoFacade(config)
        self._cleaner = WorkbookCleaner(
            activo_column=config.activo_column, keep_columns=config.keep_columns
        )

    def generate(self, target_date: date) -> Path:
        output_path = self._context.productos_path(target_date)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        print(f"INFO: Ejecutando ExcelSIIGO para generar {output_path}")
        with safe_backup(output_path):
            self._facade.run(output_path, target_date.strftime("%Y"))
            print("INFO: Limpiando el archivo generado...")
            self._cleaner.clean(output_path)
        print(f"OK: Archivo final listo en {output_path}")
        return output_path


__all__ = [
    "ProductGenerationConfig",
    "ProductListingService",
    "SiigoCredentials",
    "WorkbookCleaner",
]
