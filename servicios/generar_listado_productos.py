"""CLI para generar el listado de productos utilizando ExcelSIIGO."""

from __future__ import annotations

import argparse
import os
from pathlib import Path

from openpyxl.utils import column_index_from_string

from rentabilidad.core.env import load_env
from rentabilidad.core.dates import DateResolver, TodayStrategy
from rentabilidad.core.paths import PathContext, PathContextFactory
from rentabilidad.services.products import (
    ProductGenerationConfig,
    ProductListingService,
    SiigoCredentials,
)

KEEP_COLUMN_NUMBERS = (4, *range(7, 19))


def _ensure_trailing_backslash(path: str) -> str:
    """Devuelve ``path`` asegurando un separador final."""

    return path if path.endswith(("\\", "/")) else path + "\\"


def build_parser(defaults: dict[str, str]) -> argparse.ArgumentParser:
    """Crea el ``ArgumentParser`` principal para la herramienta de consola."""

    parser = argparse.ArgumentParser(
        description=(
            "Ejecuta ExcelSIIGO para generar el listado de productos, "
            "depura el archivo dejando solo columnas relevantes y productos activos."
        )
    )
    parser.add_argument("--siigo-dir", default=defaults["SIIGO_DIR"], help="Carpeta donde se ubica ExcelSIIGO.exe")
    parser.add_argument(
        "--siigo-base",
        default=defaults["SIIGO_BASE"],
        help="Ruta base usada como primer argumento para ExcelSIIGO (ej. D:\\SIIWI01)",
    )
    parser.add_argument(
        "--productos-dir",
        default=defaults["PRODUCTOS_DIR"],
        help="Carpeta destino donde se guardará el listado de productos",
    )
    parser.add_argument(
        "--log",
        default=defaults["SIIGO_LOG"],
        help="Ruta del archivo de log que usará ExcelSIIGO",
    )
    parser.add_argument("--fecha", help="Fecha a usar en formato YYYY-MM-DD (por defecto hoy)")
    parser.add_argument("--reporte", default=defaults["SIIGO_REPORTE"], help="Código de reporte a solicitar (GETINV por defecto)")
    parser.add_argument("--empresa", default=defaults["SIIGO_EMPRESA"], help="Código de empresa para ExcelSIIGO")
    parser.add_argument("--usuario", default=defaults["SIIGO_USUARIO"], help="Usuario para ExcelSIIGO")
    parser.add_argument("--clave", default=defaults["SIIGO_CLAVE"], help="Clave de ExcelSIIGO")
    parser.add_argument(
        "--estado-param",
        default=defaults["SIIGO_ESTADO_PARAM"],
        help="Parámetro de estado para ExcelSIIGO (S por defecto)",
    )
    parser.add_argument(
        "--rango-inicial",
        default=defaults["SIIGO_RANGO_INI"],
        help="Código inicial de rango de productos",
    )
    parser.add_argument(
        "--rango-final",
        default=defaults["SIIGO_RANGO_FIN"],
        help="Código final de rango de productos",
    )
    parser.add_argument(
        "--activo-column",
        default=defaults["SIIGO_ACTIVO_COL"],
        help="Columna (por letra) que indica si el producto está activo",
    )
    return parser


def _collect_defaults() -> dict[str, str]:
    """Lee variables de entorno y prepara valores por defecto configurables."""

    context = PathContextFactory(os.environ).create()
    defaults = {
        "SIIGO_DIR": os.environ.get("SIIGO_DIR", r"C:\\Siigo"),
        "SIIGO_BASE": os.environ.get("SIIGO_BASE", r"D:\\SIIWI01"),
        "SIIGO_LOG": os.environ.get("SIIGO_LOG", str(Path(os.environ.get("SIIGO_BASE", r"D:\\SIIWI01")) / "LOGS" / "log_catalogos.txt")),
        "SIIGO_REPORTE": os.environ.get("SIIGO_REPORTE", "GETINV"),
        "SIIGO_EMPRESA": os.environ.get("SIIGO_EMPRESA", "L"),
        "SIIGO_USUARIO": os.environ.get("SIIGO_USUARIO", "JUAN"),
        "SIIGO_CLAVE": os.environ.get("SIIGO_CLAVE", "0110"),
        "SIIGO_ESTADO_PARAM": os.environ.get("SIIGO_ESTADO_PARAM", "S"),
        "SIIGO_RANGO_INI": os.environ.get("SIIGO_RANGO_INI", "0010001000001"),
        "SIIGO_RANGO_FIN": os.environ.get("SIIGO_RANGO_FIN", "0400027999999"),
        "SIIGO_ACTIVO_COL": os.environ.get("SIIGO_ACTIVO_COL", "AX"),
        "PRODUCTOS_DIR": os.environ.get("PRODUCTOS_DIR", str(context.productos_dir)),
    }
    return defaults


def main() -> None:
    """Punto de entrada de la herramienta CLI."""

    load_env()
    defaults = _collect_defaults()
    parser = build_parser(defaults)
    args = parser.parse_args()

    resolver = DateResolver(TodayStrategy())
    fecha = resolver.resolve(args.fecha)

    context = PathContextFactory(os.environ).create()
    if args.productos_dir:
        context = PathContext(
            base_dir=context.base_dir,
            productos_dir=Path(args.productos_dir),
            informes_dir=context.informes_dir,
        )
        context.ensure_structure()

    siigo_dir = Path(args.siigo_dir)
    if not siigo_dir.exists():
        raise SystemExit(f"No existe la carpeta de SIIGO: {siigo_dir}")

    credenciales = SiigoCredentials(
        reporte=args.reporte,
        empresa=args.empresa,
        usuario=args.usuario,
        clave=args.clave,
        estado_param=args.estado_param,
        rango_ini=args.rango_inicial,
        rango_fin=args.rango_final,
    )

    config = ProductGenerationConfig(
        siigo_dir=siigo_dir,
        base_path=_ensure_trailing_backslash(args.siigo_base),
        log_path=args.log,
        credentials=credenciales,
        activo_column=args.activo_column,
        keep_columns=KEEP_COLUMN_NUMBERS + (column_index_from_string(args.activo_column),),
    )

    service = ProductListingService(context, config)
    try:
        service.generate(fecha)
    except Exception as exc:  # noqa: BLE001 - queremos mostrar cualquier error amigablemente
        raise SystemExit(str(exc))


if __name__ == "__main__":
    main()
