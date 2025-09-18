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


class ExczPattern(Protocol):
    """Protocolo de coincidencia para nombres de archivos EXCZ."""

    def match(self, name: str, prefix: str) -> datetime | None:
        """Devuelve un ``datetime`` si ``name`` corresponde al ``prefix``."""


class TimestampedExczPattern:
    """Implementación que busca ``prefix`` seguido de un bloque AAAAMMDDHHMMSS."""

    _regex_template = r"^{prefix}(?P<ts>\d{{14}})"

    def match(self, name: str, prefix: str) -> datetime | None:
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
        self.directory = directory
        self.pattern = pattern or TimestampedExczPattern()

    def iter_matches(self, prefix: str) -> Iterable[ExczMetadata]:
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
        for meta in self.iter_matches(prefix):
            if meta.timestamp.date() == target_date:
                return meta.path
        return None


__all__ = ["ExczFileFinder", "ExczMetadata", "TimestampedExczPattern"]
