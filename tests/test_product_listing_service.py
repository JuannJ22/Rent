from __future__ import annotations

import threading
import time
from datetime import date
from pathlib import Path

import pytest

from rentabilidad.core.paths import PathContext
from rentabilidad.services.products import (
    ExcelSiigoFacade,
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
)


class _DelayedFacade:
    def __init__(
        self,
        delay: float = 0.0,
        create_file: bool = True,
        payload_size: int = 2048,
    ) -> None:
        self._delay = delay
        self._create_file = create_file
        self._payload_size = payload_size

    def run(self, output_path: Path, year: str) -> None:  # noqa: ARG002 - firma requerida
        if not self._create_file:
            return

        def _writer() -> None:
            time.sleep(self._delay)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(b"0" * self._payload_size)

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


def test_siigo_output_filename_replaces_placeholders(tmp_path: Path) -> None:
    service = _build_service(tmp_path)
    cleaner = _DummyCleaner()
    service._cleaner = cleaner

    captured_paths: list[Path] = []

    class _CapturingFacade:
        def run(self, output_path: Path, year: str) -> None:  # noqa: ARG002 - compatibilidad de firma
            captured_paths.append(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_bytes(b"0" * 2048)

    service._facade = _CapturingFacade()
    service._config.siigo_output_filename = "ProductosMesDia.xlsx"

    target_date = date(2024, 9, 5)
    result = service.generate(target_date)

    assert captured_paths, "Debe llamarse a la fachada con una ruta de salida"
    assert captured_paths[0].name == "ProductosSeptiembre05.xlsx"
    assert result.exists()
    assert cleaner.cleaned == [result]


def test_generate_fails_when_file_never_appears(tmp_path: Path) -> None:
    service = _build_service(tmp_path)
    service._facade = _DelayedFacade(create_file=False)

    with pytest.raises(FileNotFoundError):
        service.generate(date(2024, 1, 15))


def test_facade_runs_with_cwd_and_executable(monkeypatch, tmp_path: Path) -> None:
    from rentabilidad.services import products as products_module

    siigo_dir = tmp_path / "Siigo Dir"
    siigo_dir.mkdir()
    output_path = tmp_path / "Productos Finales" / "salida.xlsx"

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
        siigo_dir=siigo_dir,
        base_path="D:\\SIIWI01\\",
        log_path="D:\\SIIWI01\\LOGS\\log_catalogos.txt",
        credentials=credentials,
        activo_column=1,
        keep_columns=(1,),
    )

    facade = ExcelSiigoFacade(config)

    captured: dict[str, object] = {}

    class DummyResult:
        returncode = 0
        stdout = ""
        stderr = ""

    def fake_run(command, **kwargs):  # noqa: ANN001 - firma flexible para imitar subprocess.run
        captured["command"] = command
        captured["kwargs"] = kwargs
        return DummyResult()

    monkeypatch.setattr(products_module.subprocess, "run", fake_run)

    facade.run(output_path, "2024")

    assert captured, "El comando de ExcelSIIGO debe ejecutarse"
    cmd = captured["command"]
    kwargs = captured["kwargs"]
    assert cmd[0] == str(siigo_dir / "ExcelSIIGO.exe")
    assert kwargs.get("cwd") == str(siigo_dir)
    assert kwargs.get("capture_output") is True
    assert kwargs.get("text") is True


def test_facade_accepts_custom_executable(monkeypatch, tmp_path: Path) -> None:
    from rentabilidad.services import products as products_module

    siigo_dir = tmp_path / "Siigo"
    siigo_dir.mkdir()
    output_path = tmp_path / "salida.xlsx"

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
        siigo_dir=siigo_dir,
        base_path="D:\\SIIWI01\\",
        log_path="D:\\SIIWI01\\LOGS\\log_catalogos.txt",
        credentials=credentials,
        activo_column=1,
        keep_columns=(1,),
        siigo_command="Excel Custom.exe",
    )

    facade = ExcelSiigoFacade(config)

    captured: dict[str, object] = {}

    class DummyResult:
        returncode = 0
        stdout = ""
        stderr = ""

    def fake_run(command, **kwargs):  # noqa: ANN001 - firma flexible
        captured["command"] = command
        return DummyResult()

    monkeypatch.setattr(products_module.subprocess, "run", fake_run)

    facade.run(output_path, "2024")

    assert captured, "Debe ejecutarse ExcelSIIGO"
    cmd = captured["command"]
    assert cmd[0].endswith("Excel Custom.exe")


def test_generate_fails_when_file_too_small(tmp_path: Path) -> None:
    service = _build_service(tmp_path)
    service._facade = _DelayedFacade(payload_size=10)

    with pytest.raises(RuntimeError) as excinfo:
        service.generate(date(2024, 3, 1))

    assert "No se generó el archivo de productos" in str(excinfo.value)
