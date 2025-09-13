import argparse, os, shutil
from pathlib import Path
from datetime import datetime


def _load_env():
    """Load environment variables from a .env file if present."""
    env_file = Path(__file__).resolve().parent.parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


_load_env()
DEFAULT_RENT_DIR = os.environ.get("RENT_DIR", r"C:\\Rentabilidad")
DEFAULT_TEMPLATE = os.environ.get(
    "TEMPLATE", str(Path(DEFAULT_RENT_DIR) / "PLANTILLA.xlsx")
)
DEFAULT_OUTDIR = os.environ.get("RENT_DIR", DEFAULT_RENT_DIR)

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
