"""CLI para clonar la plantilla base de rentabilidad.

El script se encarga de copiar ``PLANTILLA.xlsx`` hacia la carpeta del mes
correspondiente generando un archivo con el formato ``<Mes> DD.xlsx``. Todas
las rutas pueden personalizarse a través de variables de entorno o mediante
argumentos explícitos, lo que facilita extender el comportamiento sin modificar
el código (principio de *Open/Closed*).
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path


CURRENT_DIR = Path(__file__).resolve().parent
REPO_ROOT = CURRENT_DIR.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


from rentabilidad.core.dates import DateResolver, YesterdayStrategy
from rentabilidad.core.env import load_env
from rentabilidad.core.paths import PathContext, PathContextFactory


@dataclass(frozen=True)
class TemplateCloneService:
    """Encapsula la lógica de copiado de la plantilla a un destino específico."""

    context: PathContext

    def clone(self, template_path: Path, target_date: date, outdir: Path | None = None) -> Path:
        """Copia ``template_path`` al archivo estándar de ``target_date``.

        Parameters
        ----------
        template_path:
            Ruta a la plantilla base que se desea duplicar.
        target_date:
            Fecha objetivo usada para determinar el nombre del archivo
            resultante y la carpeta del mes.
        outdir:
            Carpeta opcional donde crear el archivo. Si es ``None`` se utiliza
            la carpeta calculada por :class:`PathContext` en función de la
            fecha.

        Returns
        -------
        Path
            Ruta final del archivo generado.
        """

        if not template_path.exists():
            raise FileNotFoundError(f"No existe la plantilla indicada: {template_path}")

        destination_dir = outdir or self.context.informe_month_dir(target_date)
        destination_dir.mkdir(parents=True, exist_ok=True)

        destination = destination_dir / self.context.informe_filename(target_date)
        shutil.copyfile(template_path, destination)
        return destination


def _build_parser(context: PathContext) -> argparse.ArgumentParser:
    """Construye el analizador de argumentos para la interfaz de línea de comandos."""

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
    return parser


def main() -> None:
    """Punto de entrada del script de consola.

    1. Carga variables de entorno soportadas por el proyecto.
    2. Construye el contexto de rutas para localizar carpetas relevantes.
    3. Interpreta los argumentos suministrados por el usuario.
    4. Clona la plantilla hacia la ubicación calculada mostrando la ruta final.
    """

    load_env()
    context = PathContextFactory(os.environ).create()

    parser = _build_parser(context)
    args = parser.parse_args()

    resolver = DateResolver(YesterdayStrategy())
    target_date = resolver.resolve(args.fecha)

    service = TemplateCloneService(context)
    template_path = Path(args.template)
    outdir = Path(args.outdir) if args.outdir else None
    result = service.clone(template_path, target_date, outdir)

    print(result)


if __name__ == "__main__":  # pragma: no cover - ejecución directa
    main()
