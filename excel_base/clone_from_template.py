import argparse, shutil
from pathlib import Path
from datetime import datetime

DEFAULT_TEMPLATE = r"C:\\Rentabilidad\\PLANTILLA.xlsx"
DEFAULT_OUTDIR   = r"C:\\Rentabilidad"

def main():
    p = argparse.ArgumentParser(description="Clona PLANTILLA.xlsx a INFORME_YYYYMMDD.xlsx en C:\\Rentabilidad.")
    p.add_argument("--template", default=DEFAULT_TEMPLATE, help="Ruta a PLANTILLA.xlsx")
    p.add_argument("--outdir",   default=DEFAULT_OUTDIR,   help="Carpeta de salida")
    p.add_argument("--fecha",    default=None,             help="YYYY-MM-DD (por defecto hoy)")
    args = p.parse_args()

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"ERROR: No existe {template_path}")
        raise SystemExit(2)

    fecha = args.fecha or datetime.now().strftime("%Y-%m-%d")
    yyyymmdd = fecha.replace("-", "")
    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out_path = outdir / f"INFORME_{yyyymmdd}.xlsx"

    shutil.copyfile(template_path, out_path)
    print(out_path)

if __name__ == "__main__":
    main()
