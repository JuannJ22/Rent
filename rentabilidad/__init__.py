"""Paquete principal para la automatización de informes de rentabilidad."""

from importlib.metadata import PackageNotFoundError, version

try:  # pragma: no cover - mejor esfuerzo
    __version__ = version("rent")
except PackageNotFoundError:  # pragma: no cover - entorno editable
    __version__ = "0.0.0"

__all__ = ["__version__"]
