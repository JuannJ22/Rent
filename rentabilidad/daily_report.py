"""Ejecución diaria SQL sin GUI: python -m rentabilidad.daily_report."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from datetime import date, datetime, timedelta
import logging
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

from rentabilidad.core.env import load_env
from rentabilidad.core.paths import SPANISH_MONTHS

ROOT = Path(__file__).resolve().parents[1]


@contextmanager
def report_lock(path: Path):
    """El sistema libera el bloqueo incluso cuando termina el proceso."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as handle:
        if path.stat().st_size == 0:
            handle.write(b"0")
            handle.flush()
        handle.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise RuntimeError("Ya hay otra ejecución del informe en curso.") from exc
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == "nt":
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def target_date(value: str | None, timezone: str) -> date:
    if value:
        return date.fromisoformat(value)
    return datetime.now(ZoneInfo(timezone)).date() - timedelta(days=1)


def generate_report(
    template: Path, destination: Path, report_date: date, *,
    sql_config: str | None, timeout: int, replace: bool, check_only: bool = False,
) -> None:
    command = [
        sys.executable, str(ROOT / "hojas" / "hoja01_loader.py"),
        "--sql", "--require-sql-data", "--fecha", report_date.isoformat(),
    ]
    if sql_config:
        command.extend(["--sql-config", sql_config])
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONDONTWRITEBYTECODE="1")

    def run_loader(args: list[str]) -> None:
        result = subprocess.run(
            args, cwd=ROOT, env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout,
        )
        for output in (result.stdout, result.stderr):
            if output.strip():
                logging.info("%s", output.strip())
        if result.returncode:
            raise RuntimeError(f"El motor SQL terminó con código {result.returncode}.")

    if check_only:
        run_loader(command + ["--check-sql"])
        return

    with report_lock(destination.parent.parent.parent / ".daily-report.lock"):
        if destination.exists() and not replace:
            logging.info("Informe existente; se conserva: %s", destination)
            return
        if not template.is_file():
            raise FileNotFoundError(f"No existe la plantilla: {template}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        # Mantener el nombre del informe: el motor lo usa para las fechas y hojas.
        with tempfile.TemporaryDirectory(prefix=".daily-", dir=destination.parent) as stage:
            staged = Path(stage) / destination.name
            shutil.copyfile(template, staged)
            run_loader(command + ["--excel", str(staged)])
            workbook = load_workbook(staged, read_only=True)
            workbook.close()
            os.replace(staged, destination)
        logging.info("Informe SQL generado: %s", destination)


def main(argv: list[str] | None = None) -> int:
    load_env()
    parser = argparse.ArgumentParser(description="Genera el informe diario exclusivamente desde SQL.")
    parser.add_argument("--fecha", help="YYYY-MM-DD; predeterminado: ayer en Colombia")
    parser.add_argument("--timezone", default="America/Bogota")
    parser.add_argument("--base-dir", default=os.environ.get("RENT_DIR", r"C:\Rentabilidad"))
    parser.add_argument("--template", help="Predeterminado: BASE_DIR/PLANTILLA.xlsx")
    parser.add_argument("--sql-config", default=os.environ.get("SQL_CONFIG"))
    parser.add_argument("--timeout", type=int, default=1800, help="Límite del motor en segundos")
    parser.add_argument("--replace", action="store_true", help="Regenera un informe existente de forma atómica")
    parser.add_argument("--check-sql", action="store_true", help="Valida las consultas sin generar Excel")
    args = parser.parse_args(argv)
    if args.timeout <= 0:
        parser.error("--timeout debe ser positivo")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    base = Path(args.base_dir).resolve()
    logs = base / "Logs"
    logs.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler(
            logs / f"daily-{datetime.now(ZoneInfo('America/Bogota')):%Y-%m-%d}.log",
            encoding="utf-8",
        )], force=True,
    )
    try:
        day = target_date(args.fecha, args.timezone)
        month = SPANISH_MONTHS[day.month]
        destination = base / "InformesDiarios" / str(day.year) / month / f"{month} {day:%d}.xlsx"
        os.environ["RENT_DIR"] = str(base)
        logging.info("Fecha objetivo: %s; fuente: SQL", day)
        generate_report(
            Path(args.template).resolve() if args.template else base / "PLANTILLA.xlsx",
            destination, day, sql_config=args.sql_config, timeout=args.timeout,
            replace=args.replace, check_only=args.check_sql,
        )
        return 0
    except (Exception, SystemExit) as exc:
        logging.error("No se completó la ejecución: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
