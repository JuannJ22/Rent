from __future__ import annotations

import argparse
import os
import shutil
from pathlib import Path

from rentabilidad.core.dates import DateResolver, YesterdayStrategy
from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContextFactory


def main() -> None:
    load_env()

    context = PathContextFactory(os.environ).create()

    parser = argparse.ArgumentParser(
        description=(
            "Clona PLANTILLA.xlsx a '<Mes> DD.xlsx' dentro de la estructura de "
            r"C:\Rentabilidad\Informes."
        )
    )
    parser.add_argument("--template", default=str(context.template_path()), help="Ruta a PLANTILLA.xlsx")
    parser.add_argument(
        "--outdir",
        default=None,
        help=(
            "Carpeta de salida. Por defecto se utiliza la carpeta del mes dentro de"
            r" C:\Rentabilidad\Informes"
        ),
    )
    parser.add_argument("--fecha", default=None, help="YYYY-MM-DD (por defecto el día anterior)")
    args = parser.parse_args()

    resolver = DateResolver(YesterdayStrategy())
    target_date = resolver.resolve(args.fecha)

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"ERROR: No existe {template_path}")
        raise SystemExit(2)

    if args.outdir:
        outdir = Path(args.outdir)
        outdir.mkdir(parents=True, exist_ok=True)
    else:
        outdir = context.informe_month_dir(target_date)

    out_path = outdir / context.informe_filename(target_date)

    shutil.copyfile(template_path, out_path)
    print(out_path)


if __name__ == "__main__":
    main()
