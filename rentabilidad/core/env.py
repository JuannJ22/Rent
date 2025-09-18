"""Utilidades para cargar variables de entorno desde un archivo ``.env``."""

from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Iterable


@lru_cache(maxsize=1)
def load_env(extra_paths: Iterable[Path] | None = None) -> None:
    """Carga pares clave-valor de archivos ``.env`` hacia ``os.environ``.

    Se busca un archivo ``.env`` en la raíz del repositorio (dos niveles por
    encima de este módulo) y en las rutas adicionales proporcionadas. Los
    valores existentes en ``os.environ`` prevalecen sobre los del archivo.
    """

    candidate_paths: list[Path] = []

    base = Path(__file__).resolve().parents[2]
    candidate_paths.append(base / ".env")

    if extra_paths:
        candidate_paths.extend(Path(p) for p in extra_paths)

    for env_path in candidate_paths:
        if not env_path.exists():
            continue
        for raw_line in env_path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


__all__ = ["load_env"]
