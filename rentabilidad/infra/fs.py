from datetime import datetime, timedelta
from pathlib import Path


def ayer_str() -> str:
    return (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")


def asegurar_carpeta(p: Path) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
