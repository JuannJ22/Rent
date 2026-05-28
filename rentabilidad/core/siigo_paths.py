"""Rutas por defecto compartidas para integraciones con SIIGO.

Centralizar estos valores evita que los cambios de infraestructura queden
replicados en varios módulos y mantiene los puntos de configuración alineados.
"""

from __future__ import annotations

DEFAULT_SIIGO_BASE = r"Z:\SIIWI01"
DEFAULT_EXCZ_DIR = rf"{DEFAULT_SIIGO_BASE}\LISTADOS"
DEFAULT_SIIGO_LOG_FILENAME = "log_catalogos.txt"


def build_siigo_log_path(base_path: str = DEFAULT_SIIGO_BASE) -> str:
    """Construye la ruta Windows del log de ExcelSIIGO desde la base indicada."""

    normalized_base = base_path.rstrip("\\/")
    if not normalized_base:
        return rf"LOGS\{DEFAULT_SIIGO_LOG_FILENAME}"

    return rf"{normalized_base}\LOGS\{DEFAULT_SIIGO_LOG_FILENAME}"


__all__ = [
    "DEFAULT_EXCZ_DIR",
    "DEFAULT_SIIGO_BASE",
    "DEFAULT_SIIGO_LOG_FILENAME",
    "build_siigo_log_path",
]
