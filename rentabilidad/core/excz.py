"""Herramientas para localizar archivos EXCZ por fecha y prefijo."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Protocol


@dataclass(frozen=True)
class ExczMetadata:
    """Información derivada del nombre de un archivo EXCZ."""

    path: Path
    prefix: str
    timestamp: datetime

    @property
    def date_key(self) -> str:
        return self.timestamp.strftime("%Y%m%d")

    @property
    def modified_at(self) -> float:
        """Marca de tiempo de última modificación en segundos."""

        try:
            return self.path.stat().st_mtime
        except FileNotFoundError:
            return 0.0


class ExczPattern(Protocol):
    """Protocolo de coincidencia para nombres de archivos EXCZ."""

    def match(self, name: str, prefix: str) -> datetime | None:
        """Devuelve un ``datetime`` si ``name`` corresponde al ``prefix``."""


class TimestampedExczPattern:
    """Implementación que busca ``prefix`` seguido de un bloque AAAAMMDDHHMMSS."""

    _regex_template = r"^{prefix}(?P<ts>\d{{14}})"

    def match(self, name: str, prefix: str) -> datetime | None:
        """Intenta extraer un ``datetime`` del nombre ``name`` usando ``prefix``."""

        pattern = self._regex_template.format(prefix=re.escape(prefix.lower()))
        match = re.match(pattern, name.lower())
        if not match:
            return None
        try:
            return datetime.strptime(match.group("ts"), "%Y%m%d%H%M%S")
        except ValueError:
            return None


class ExczFileFinder:
    """Localiza archivos EXCZ para una fecha específica."""

    def __init__(self, directory: Path, pattern: ExczPattern | None = None):
        """Inicializa el buscador indicando carpeta y patrón opcional."""

        self.directory = directory
        self.pattern = pattern or TimestampedExczPattern()

    def iter_matches(self, prefix: str) -> Iterable[ExczMetadata]:
        """Itera sobre los archivos que coinciden con ``prefix`` en ``directory``."""

        if not self.directory.exists():
            return []
        results: list[ExczMetadata] = []
        for child in self.directory.iterdir():
            if not child.is_file():
                continue
            ts = self.pattern.match(child.name, prefix)
            if not ts:
                continue
            results.append(ExczMetadata(path=child, prefix=prefix, timestamp=ts))
        results.sort(key=lambda meta: meta.timestamp, reverse=True)
        return results

    def find_for_date(self, prefix: str, target_date: date) -> Path | None:
        """Busca un EXCZ cuyo sello temporal coincide con ``target_date``."""

        for meta in self.iter_matches(prefix):
            if meta.timestamp.date() == target_date:
                return meta.path
        return None

    def find_latest(self, prefix: str) -> Path | None:
        """Devuelve el archivo más reciente según fecha de modificación."""

        matches = list(self.iter_matches(prefix))
        if not matches:
            return None
        latest = max(matches, key=lambda meta: meta.modified_at)
        return latest.path


__all__ = ["ExczFileFinder", "ExczMetadata", "TimestampedExczPattern"]
