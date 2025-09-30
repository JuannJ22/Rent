from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable


def ayer_str() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")


def asegurar_carpeta(p: Path) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)


def _iter_files(root: Path, patterns: Iterable[str]) -> list[Path]:
    if not root.exists():
        return []
    files: list[Path] = []
    for pattern in patterns:
        files.extend(
            path
            for path in root.rglob(pattern)
            if path.is_file() and not path.name.startswith("~$")
        )
    return files


def _find_latest(root: Path, patterns: Iterable[str]) -> Path | None:
    candidates = _iter_files(root, patterns)
    if not candidates:
        return None
    try:
        return max(candidates, key=lambda item: item.stat().st_mtime)
    except OSError:
        return None


def find_latest_informe(root: Path) -> Path | None:
    """Devuelve el informe más reciente dentro de ``root``."""

    return _find_latest(root, ("*.xlsx",))


def find_latest_producto(root: Path) -> Path | None:
    """Devuelve el listado de productos más reciente dentro de ``root``."""

    return _find_latest(root, ("*.xlsx",))
