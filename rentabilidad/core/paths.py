"""Administración de rutas utilizadas por los procesos de rentabilidad."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Mapping

SPANISH_MONTHS: Mapping[int, str] = {
    1: "01 - Enero",
    2: "02 - Febrero",
    3: "03 - Marzo",
    4: "04 - Abril",
    5: "05 - Mayo",
    6: "06 - Junio",
    7: "07 - Julio",
    8: "08 - Agosto",
    9: "09 - Septiembre",
    10: "10 - Octubre",
    11: "11 - Noviembre",
    12: "12 - Diciembre",
}


@dataclass(frozen=True)
class PathContext:
    """Representa las rutas relevantes del proceso."""

    base_dir: Path
    productos_dir: Path
    informes_dir: Path

    def ensure_structure(self) -> None:
        """Garantiza la existencia de las carpetas base."""

        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.productos_dir.mkdir(parents=True, exist_ok=True)
        self.informes_dir.mkdir(parents=True, exist_ok=True)

    def informe_month_dir(self, target_date: date) -> Path:
        """Retorna la carpeta del mes para ``target_date`` dentro de Informes."""

        month_name = SPANISH_MONTHS[target_date.month]
        year_dir = self.informes_dir / str(target_date.year)
        month_dir = year_dir / month_name
        month_dir.mkdir(parents=True, exist_ok=True)
        return month_dir

    def informe_path(self, target_date: date) -> Path:
        """Ruta completa al informe estándar para ``target_date``."""

        month_dir = self.informe_month_dir(target_date)
        return month_dir / f"INFORME_{target_date:%Y%m%d}.xlsx"

    def productos_path(self, target_date: date) -> Path:
        """Ruta completa al archivo de productos para ``target_date``."""

        self.productos_dir.mkdir(parents=True, exist_ok=True)
        return self.productos_dir / f"Productos{target_date:%m%d}.xlsx"

    def template_path(self) -> Path:
        """Ruta esperada de la plantilla base."""

        return self.base_dir / "PLANTILLA.xlsx"


class PathContextFactory:
    """Fábrica que construye :class:`PathContext` a partir del entorno."""

    def __init__(self, environ: Mapping[str, str]):
        self._environ = environ

    def create(self) -> PathContext:
        base_dir = Path(self._environ.get("RENT_DIR", r"C:\\Rentabilidad"))
        productos_dir = Path(
            self._environ.get("PRODUCTOS_DIR", str(base_dir / "Productos"))
        )
        informes_dir = Path(
            self._environ.get("INFORMES_DIR", str(base_dir / "Informes"))
        )
        context = PathContext(base_dir=base_dir, productos_dir=productos_dir, informes_dir=informes_dir)
        context.ensure_structure()
        return context


__all__ = ["PathContext", "PathContextFactory", "SPANISH_MONTHS"]
