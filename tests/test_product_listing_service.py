from __future__ import annotations

import threading
import time
from datetime import date
from pathlib import Path

import pytest

from rentabilidad.core.paths import PathContext
from rentabilidad.services.products import (
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
)


class _DelayedFacade:
    def __init__(self, delay: float = 0.0, create_file: bool = True) -> None:
        self._delay = delay
        self._create_file = create_file

    def run(self, output_path: Path, year: str) -> None:  # noqa: ARG002 - firma requerida
        if not self._create_file:
            return

        def _writer() -> None:
            time.sleep(self._delay)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(b"contenido")

        thread = threading.Thread(target=_writer, daemon=True)
        thread.start()


class _DummyCleaner:
    def __init__(self) -> None:
        self.cleaned: list[Path] = []

    def clean(self, path: Path) -> None:
        if not path.exists():
            raise AssertionError("El archivo debe existir antes de limpiar")
        self.cleaned.append(path)


def _build_service(tmp_path: Path) -> ProductListingService:
    base_dir = tmp_path / "base"
    productos_dir = base_dir / "Productos"
    informes_dir = base_dir / "Informes"

    context = PathContext(base_dir=base_dir, productos_dir=productos_dir, informes_dir=informes_dir)
    context.ensure_structure()

    credentials = SiigoCredentials(
        reporte="REP",
        empresa="EMP",
        usuario="USR",
        clave="PWD",
        estado_param="S",
        rango_ini="0001",
        rango_fin="9999",
    )

    config = ProductGenerationConfig(
        siigo_dir=tmp_path,
        base_path="D:\\SIIWI01\\",
        log_path="D:\\SIIWI01\\LOGS\\log_catalogos.txt",
        credentials=credentials,
        activo_column=1,
        keep_columns=(1,),
        wait_timeout=1.0,
        wait_interval=0.01,
    )

    service = ProductListingService(context, config)
    return service


def test_generate_waits_for_delayed_file(tmp_path: Path) -> None:
    service = _build_service(tmp_path)
    cleaner = _DummyCleaner()
    service._cleaner = cleaner
    service._facade = _DelayedFacade(delay=0.05)

    result = service.generate(date(2024, 9, 27))

    assert result.exists()
    assert cleaner.cleaned == [result]


def test_generate_fails_when_file_never_appears(tmp_path: Path) -> None:
    service = _build_service(tmp_path)
    service._facade = _DelayedFacade(create_file=False)

    with pytest.raises(FileNotFoundError):
        service.generate(date(2024, 1, 15))
