import argparse
import os
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def _load_env() -> None:
    """Load environment variables from a .env file if present."""
    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


_load_env()

DEFAULT_SIIGO_DIR = os.environ.get("SIIGO_DIR", r"C:\\Siigo")
DEFAULT_SIIGO_BASE = os.environ.get("SIIGO_BASE", r"D:\\SIIWI01")
DEFAULT_LOG_PATH = os.environ.get(
    "SIIGO_LOG", str(Path(DEFAULT_SIIGO_BASE) / "LOGS" / "log_catalogos.txt")
)
DEFAULT_PRODUCTOS_DIR = os.environ.get(
    "PRODUCTOS_DIR", r"C:\\Rentabilidad\\Productos"
)
DEFAULT_REPORTE = os.environ.get("SIIGO_REPORTE", "GETINV")
DEFAULT_EMPRESA = os.environ.get("SIIGO_EMPRESA", "L")
DEFAULT_USUARIO = os.environ.get("SIIGO_USUARIO", "JUAN")
DEFAULT_CLAVE = os.environ.get("SIIGO_CLAVE", "0110")
DEFAULT_ESTADO_PARAM = os.environ.get("SIIGO_ESTADO_PARAM", "S")
DEFAULT_RANGO_INI = os.environ.get("SIIGO_RANGO_INI", "0010001000001")
DEFAULT_RANGO_FIN = os.environ.get("SIIGO_RANGO_FIN", "0400027999999")
DEFAULT_ACTIVO_COL = os.environ.get("SIIGO_ACTIVO_COL", "AX")
KEEP_COLUMN_NUMBERS = {4, *range(7, 19)}
KEEP_COLUMN_NUMBERS.add(column_index_from_string(DEFAULT_ACTIVO_COL))


def _ensure_trailing_backslash(path: str) -> str:
    if path.endswith(("\\", "/")):
        return path
    return path + "\\"


def _build_output_path(productos_dir: Path, fecha: datetime) -> Path:
    productos_dir.mkdir(parents=True, exist_ok=True)
    filename = f"Productos{fecha.strftime('%m')}{fecha.strftime('%d')}.xlsx"
    return productos_dir / filename


def _run_excel_siigo(
    *,
    siigo_dir: Path,
    base_path: str,
    ano: str,
    reporte: str,
    empresa: str,
    usuario: str,
    clave: str,
    log_path: str,
    estado_param: str,
    rango_ini: str,
    rango_fin: str,
    output_path: Path,
) -> None:
    command = [
        "ExcelSIIGO",
        base_path,
        ano,
        reporte,
        empresa,
        usuario,
        clave,
        log_path,
        estado_param,
        rango_ini,
        rango_fin,
        str(output_path),
    ]

    result = subprocess.run(
        command,
        cwd=str(siigo_dir),
        check=False,
        capture_output=True,
        text=True,
    )

    if result.stdout:
        print(result.stdout.strip())
    if result.stderr:
        print(result.stderr.strip())

    if result.returncode != 0:
        raise RuntimeError(
            "ExcelSIIGO fallo con codigo "
            f"{result.returncode}: {result.stderr.strip() or result.stdout.strip()}"
        )


def _normalize_activo(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip().upper()
    return str(value).strip().upper()


def _clean_excel(
    file_path: Path,
    *,
    activo_column: str,
    keep_columns: Iterable[int],
) -> None:
    wb = load_workbook(filename=file_path)
    ws = wb.active

    activo_idx = column_index_from_string(activo_column)
    if ws.max_column < activo_idx:
        raise RuntimeError(
            f"La hoja activa no tiene la columna {activo_column} (indice {activo_idx})."
        )

    for row in range(ws.max_row, 1, -1):
        value = ws.cell(row=row, column=activo_idx).value
        if _normalize_activo(value) != "S":
            ws.delete_rows(row, 1)

    keep_set = set(keep_columns)
    # Ensure the activo column stays even if it was not part of the defaults
    keep_set.add(activo_idx)

    max_col = ws.max_column
    for col in range(max_col, 0, -1):
        if col not in keep_set:
            ws.delete_cols(col, 1)

    wb.save(file_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Ejecuta ExcelSIIGO para generar el listado de productos y "
            "depura el archivo dejando solo columnas relevantes y "
            "productos activos."
        )
    )
    parser.add_argument("--siigo-dir", default=DEFAULT_SIIGO_DIR, help="Carpeta donde se ubica ExcelSIIGO.exe")
    parser.add_argument(
        "--siigo-base",
        default=DEFAULT_SIIGO_BASE,
        help="Ruta base usada como primer argumento para ExcelSIIGO (ej. D:\\SIIWI01)",
    )
    parser.add_argument(
        "--productos-dir",
        default=DEFAULT_PRODUCTOS_DIR,
        help="Carpeta destino donde se guardara el listado de productos",
    )
    parser.add_argument(
        "--log",
        default=DEFAULT_LOG_PATH,
        help="Ruta del archivo de log que usara ExcelSIIGO",
    )
    parser.add_argument("--fecha", help="Fecha a usar en formato YYYY-MM-DD (por defecto hoy)")
    parser.add_argument("--reporte", default=DEFAULT_REPORTE, help="Codigo de reporte a solicitar (GETINV por defecto)")
    parser.add_argument("--empresa", default=DEFAULT_EMPRESA, help="Codigo de empresa para ExcelSIIGO")
    parser.add_argument("--usuario", default=DEFAULT_USUARIO, help="Usuario para ExcelSIIGO")
    parser.add_argument("--clave", default=DEFAULT_CLAVE, help="Clave de ExcelSIIGO")
    parser.add_argument(
        "--estado-param",
        default=DEFAULT_ESTADO_PARAM,
        help="Parametro de estado para ExcelSIIGO (S por defecto)",
    )
    parser.add_argument(
        "--rango-inicial",
        default=DEFAULT_RANGO_INI,
        help="Codigo inicial de rango de productos",
    )
    parser.add_argument(
        "--rango-final",
        default=DEFAULT_RANGO_FIN,
        help="Codigo final de rango de productos",
    )
    parser.add_argument(
        "--activo-column",
        default=DEFAULT_ACTIVO_COL,
        help="Columna (por letra) que indica si el producto esta activo",
    )

    args = parser.parse_args()

    try:
        siigo_dir = Path(args.siigo_dir)
        if not siigo_dir.exists():
            raise RuntimeError(f"No existe la carpeta de SIIGO: {siigo_dir}")

        fecha = (
            datetime.strptime(args.fecha, "%Y-%m-%d")
            if args.fecha
            else datetime.now()
        )
        ano = fecha.strftime("%Y")
        productos_dir = Path(args.productos_dir)
        output_path = _build_output_path(productos_dir, fecha)

        base_path = _ensure_trailing_backslash(args.siigo_base)

        log_path = args.log
        if not log_path:
            log_path = str(Path(base_path) / "LOGS" / "log_catalogos.txt")

        backup_path = None
        if output_path.exists():
            backup_path = output_path.with_suffix(output_path.suffix + ".bak")
            try:
                if backup_path.exists():
                    backup_path.unlink()
                output_path.rename(backup_path)
            except OSError as exc:
                raise RuntimeError(
                    f"No se pudo respaldar el archivo existente {output_path}: {exc}"
                ) from exc

        try:
            print(f"INFO: Ejecutando ExcelSIIGO para generar {output_path}")
            _run_excel_siigo(
                siigo_dir=siigo_dir,
                base_path=base_path,
                ano=ano,
                reporte=args.reporte,
                empresa=args.empresa,
                usuario=args.usuario,
                clave=args.clave,
                log_path=log_path,
                estado_param=args.estado_param,
                rango_ini=args.rango_inicial,
                rango_fin=args.rango_final,
                output_path=output_path,
            )

            print("INFO: Limpiando el archivo generado...")
            _clean_excel(
                output_path,
                activo_column=args.activo_column,
                keep_columns=KEEP_COLUMN_NUMBERS,
            )
            print(f"OK: Archivo final listo en {output_path}")
        except Exception:
            if backup_path and backup_path.exists():
                try:
                    if output_path.exists():
                        output_path.unlink()
                except OSError:
                    pass
                try:
                    backup_path.rename(output_path)
                except OSError:
                    pass
            raise
        else:
            if backup_path and backup_path.exists():
                backup_path.unlink()
    except Exception as exc:  # noqa: BLE001 - queremos mostrar cualquier error amigablemente
        raise SystemExit(str(exc))


if __name__ == "__main__":
    main()
