"""Servicios relacionados con la generación del listado de productos.

El módulo orquesta la ejecución de ``ExcelSIIGO`` y la posterior depuración del
Excel resultante. Las piezas se encuentran desacopladas siguiendo los
principios SOLID: ``ExcelSiigoFacade`` encapsula la interacción externa (Single
Responsibility y Facade Pattern), ``WorkbookCleaner`` aplica la lógica de
post-procesamiento (*Strategy* reemplazable mediante extensión) y
``ProductListingService`` coordina ambas piezas actuando como *Service Layer*
abierto a nuevas variaciones.
"""

from __future__ import annotations

import os
import subprocess
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from rentabilidad.core.paths import PathContext


def _quote_windows(arg: str) -> str:
    """Devuelve ``arg`` listo para mostrarse como parte de un comando de Windows."""

    if not arg:
        return '""'
    if any(ch in arg for ch in ' \t"'):
        escaped = arg.replace('"', r'\"')
        return f'"{escaped}"'
    return arg


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
    activo_column: int | str
    keep_columns: Sequence[int | str]


class ExcelSiigoFacade:
    """Fachada que ejecuta ExcelSIIGO y gestiona los parámetros requeridos."""

    def __init__(self, config: ProductGenerationConfig):
        self._config = config

    def run(self, output_path: Path, year: str) -> None:
        """Ejecuta ``ExcelSIIGO`` generando el archivo temporal ``output_path``.

        Parameters
        ----------
        output_path:
            Ruta completa donde se escribirá el reporte exportado por
            ``ExcelSIIGO``.
        year:
            Año fiscal que se pasa como segundo parámetro al ejecutable.

        Raises
        ------
        RuntimeError
            Cuando ``ExcelSIIGO`` termina con un código distinto de cero. El
            mensaje incluye la salida estándar y de error para facilitar la
            depuración.
        """

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

        printable_command = " ".join(_quote_windows(arg) for arg in command)
        print(f"CMD> {printable_command}")
        print(f"CWD> {self._config.siigo_dir}")

        if os.name == "nt":
            cd_command = f"cd /d {self._config.siigo_dir}"
            joined_command = " ".join(_quote_windows(arg) for arg in command)
            cmdline = f"{cd_command} && {joined_command}"
            result = subprocess.run(
                ["cmd.exe", "/d", "/c", cmdline],
                check=False,
                capture_output=True,
                text=True,
            )
        else:
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

    def __init__(self, activo_column: int | str, keep_columns: Iterable[int | str]):
        self._activo_idx = self._resolve_column(activo_column)
        self._keep_columns = {self._resolve_column(idx) for idx in keep_columns}
        self._keep_columns.add(self._activo_idx)

    def clean(self, file_path: Path) -> None:
        """Filtra filas y columnas dejando únicamente productos activos.

        El método carga el libro indicado, conserva solo las columnas
        declaradas en ``keep_columns`` (más la columna ``activo``) y elimina las
        filas cuyo indicador de producto activo no sea ``"S"``.
        """

        wb = load_workbook(filename=file_path)
        ws = wb.active

        if ws.max_column < self._activo_idx:
            raise RuntimeError(
                "La hoja activa no tiene la columna requerida para estado del producto."
            )

        removed_rows = 0
        for row in range(ws.max_row, 1, -1):
            value = ws.cell(row=row, column=self._activo_idx).value
            if self._normalize(value) != "S":
                ws.delete_rows(row, 1)
                removed_rows += 1

        removed_columns = 0
        for col in range(ws.max_column, 0, -1):
            if col not in self._keep_columns:
                ws.delete_cols(col, 1)
                removed_columns += 1

        wb.save(file_path)
        print(
            "INFO: Limpieza completada -"
            f" filas eliminadas: {removed_rows}, columnas eliminadas: {removed_columns}."
        )
        columnas_conservadas = ", ".join(str(idx) for idx in sorted(self._keep_columns))
        print(f"INFO: Columnas conservadas: {columnas_conservadas}")

    @staticmethod
    def _normalize(value) -> str:
        """Normaliza el contenido de la celda a mayúsculas sin espacios."""

        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip().upper()
        return str(value).strip().upper()

    @staticmethod
    def _resolve_column(value: int | str) -> int:
        """Convierte ``value`` en un índice de columna de Excel (1-based)."""

        if isinstance(value, int):
            if value < 1:
                raise ValueError("El índice de columna debe ser mayor o igual a 1")
            return value
        if isinstance(value, str):
            text = value.strip()
            if not text:
                raise ValueError("El identificador de columna no puede estar vacío")
            if text.isdigit():
                idx = int(text)
            else:
                try:
                    idx = column_index_from_string(text)
                except ValueError as exc:  # pragma: no cover - conversión de openpyxl
                    raise ValueError(f"Columna inválida: {value!r}") from exc
            if idx < 1:
                raise ValueError("El índice de columna debe ser mayor o igual a 1")
            return idx
        raise TypeError("El identificador de columna debe ser entero o cadena")


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
        """Genera el listado para ``target_date`` delegando en los componentes.

        Parameters
        ----------
        target_date:
            Fecha objetivo utilizada para construir el nombre final del
            archivo. También se utiliza para determinar el parámetro ``year``
            de ``ExcelSIIGO``.

        Returns
        -------
        Path
            Ruta del archivo Excel procesado y listo para distribución.
        """

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
