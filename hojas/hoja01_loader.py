import argparse, os, re
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def _load_env():
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
DEFAULT_EXCEL = os.environ.get(
    "EXCEL",
    str(Path(DEFAULT_RENT_DIR) / f"INFORME_{datetime.now().strftime('%Y%m%d')}.xlsx"),
)
DEFAULT_EXCZDIR = os.environ.get("EXCZDIR", r"D:\\SIIWI01\\LISTADOS")

def _norm(s: str) -> str:
    return (str(s).strip().lower()
            .replace("%","").replace(".","")
            .replace("_"," ").replace("-"," ").replace("  "," "))

def _find_header_row_and_map(ws):
    header_row = None
    header_map = {}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 100), values_only=True), start=1):
        non_empty = [(j+1, c) for j, c in enumerate(row) if c not in (None, "")]
        if len(non_empty) >= 3:
            header_row = i
            for col_idx, val in non_empty:
                header_map[_norm(val)] = (val, col_idx)
            break
    return header_row, header_map

def _letter_from_header(header_map, *candidates):
    for c in candidates:
        key = _norm(c)
        if key in header_map:
            return header_map[key][1]
    return None

def _pick_latest_excz(path: Path):
    patterns = [r'^EXCZ.*\.(xlsx|xls|csv)$', r'^EXZ.*\.(xlsx|xls|csv)$', r'^EXC.*\.(xlsx|xls|csv)$']
    candidates = []
    for p in path.iterdir():
        if not p.is_file(): 
            continue
        name = p.name.upper()
        if any(re.match(pat, name) for pat in patterns):
            candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

def _read_excz_df(file: Path):
    suffix = file.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        try:
            return pd.read_excel(file, sheet_name="Hoja1")
        except Exception:
            return pd.read_excel(file)
    elif suffix == ".csv":
        try:
            return pd.read_csv(file, sep=";", engine="python")
        except Exception:
            return pd.read_csv(file, engine="python")
    else:
        raise ValueError("Formato no soportado: " + suffix)

def _guess_map(df_cols):
    def norm(s): 
        return (str(s).strip().lower()
                .replace("%","").replace(".","")
                .replace("_"," ").replace("-"," ").replace("  "," "))
    cols = { norm(c): c for c in df_cols }
    def pick(*keys):
        for k in keys:
            if k in cols: return cols[k]
        return None
    return {
        "nit": pick("nit","nit cliente","identificacion","identificación"),
        "cliente_combo": pick("nit - sucursal - cliente","cliente sucursal","cliente","razon social","razón social"),
        "descripcion": pick("descripcion","descripción","producto","nombre producto","item"),
        "cantidad": pick("cantidad","cant"),
        "ventas": pick("ventas","subtotal sin iva","total sin iva","valor venta","base"),
        "costos": pick("costos","costo","costo total","costo sin iva"),
        "renta": pick("% renta","renta","rentabilidad","rentabilidad venta"),
        "utili": pick("% utili","utili","utilidad","utilidad %","utilidad porcentaje"),
    }

def main():
    p = argparse.ArgumentParser(description="Importa último EXCZ a Hoja 1 y aplica fórmulas fijas.")
    p.add_argument("--excel",   default=DEFAULT_EXCEL,   help="Ruta al INFORME_YYYYMMDD.xlsx")
    p.add_argument("--exczdir", default=DEFAULT_EXCZDIR, help="Carpeta de EXCZ")
    p.add_argument("--hoja",    default=None,            help="Nombre de la Hoja 1 (por defecto la primera)")
    p.add_argument("--max-rows", type=int, default=0,    help="Forzar número de filas (0 = según datos)")
    p.add_argument("--skip-import", action="store_true", help="No importar EXCZ, sólo aplicar fórmulas")
    p.add_argument("--safe-fill",  action="store_true", default=True, help="Sólo escribir en filas con datos")
    args = p.parse_args()

    path = Path(args.excel)
    if not path.exists():
        print(f"ERROR: No existe el informe: {path}")
        raise SystemExit(2)

    wb = load_workbook(path)
    ws = wb[args.hoja] if args.hoja else wb.worksheets[0]

    header_row, hmap = _find_header_row_and_map(ws)
    if not header_row:
        print("ERROR: No se detectaron cabeceras en Hoja 1")
        raise SystemExit(3)

    def idx(*names): 
        return _letter_from_header(hmap, *names)

    col_nit = idx("nit")
    col_cliente_combo = idx("nit - sucursal - cliente","cliente")
    col_desc = idx("descripcion","descripción","producto")
    col_cant = idx("cantidad")
    col_ventas = idx("ventas")
    col_costos = idx("costos","costo")
    col_renta = idx("% renta.","% renta","renta","rentabilidad")
    col_utili = idx("% utili.","% utili","utili","utilidad")
    col_vendedor = idx("vendedor","cod vendedor","cod. vendedor","codigo vendedor","cód vendedor")
    col_precio = idx("precio")
    col_descuento = idx("descuento")
    col_excz = idx("excz")

    start_row = header_row + 1

    # Importar EXCZ más reciente
    n_rows = 0
    if not args.skip_import:
        excz_dir = Path(args.exczdir)
        if not excz_dir.exists():
            print(f"ERROR: No existe la carpeta de EXCZ: {excz_dir}")
            raise SystemExit(4)

        latest = _pick_latest_excz(excz_dir)
        if not latest:
            print("ERROR: No se encontró EXCZ en la carpeta.")
            raise SystemExit(5)

        df = _read_excz_df(latest)
        m = _guess_map(df.columns)

        cols_needed = {k:v for k,v in m.items() if v is not None}
        sub = df[list(cols_needed.values())].copy()
        sub.rename(columns={v:k for k,v in cols_needed.items()}, inplace=True)

        # Filtrar % RENTA < 100% si viene
        if "renta" in sub.columns:
            sub = sub[pd.to_numeric(sub["renta"], errors="coerce") < 1.0]

        if args.max_rows and len(sub) > args.max_rows:
            sub = sub.iloc[:args.max_rows].copy()

        # Escribir al Excel
        for i, row in enumerate(sub.itertuples(index=False), start=start_row):
            if col_nit and "nit" in sub.columns: ws.cell(i, col_nit, getattr(row, "nit"))
            if col_cliente_combo and "cliente_combo" in sub.columns: ws.cell(i, col_cliente_combo, getattr(row, "cliente_combo"))
            if col_desc and "descripcion" in sub.columns: ws.cell(i, col_desc, getattr(row, "descripcion"))
            if col_cant and "cantidad" in sub.columns: ws.cell(i, col_cant, getattr(row, "cantidad"))
            if col_ventas and "ventas" in sub.columns: ws.cell(i, col_ventas, getattr(row, "ventas"))
            if col_costos and "costos" in sub.columns: ws.cell(i, col_costos, getattr(row, "costos"))
            if col_renta and "renta" in sub.columns: ws.cell(i, col_renta, getattr(row, "renta"))
            if col_utili and "utili" in sub.columns: ws.cell(i, col_utili, getattr(row, "utili"))
            if col_excz: ws.cell(i, col_excz, latest.stem)

        n_rows = len(sub)

    # Aplicar fórmulas fijas
    vend_range = "G:H"   # VENDEDORES (NIT en G, COD_VENDEDOR en H)
    prec_range = "A:B"   # PRECIOS (DESCRIPCION en A, PRECIO en B)

    L = lambda c: get_column_letter(c) if c else None
    L_vend = L(col_vendedor); L_prec = L(col_precio); L_desc = L(col_descuento)
    L_nit = L(col_nit); L_desc_src = L(col_desc); L_cant = L(col_cant); L_vent = L(col_ventas)

    end_row = ws.max_row if n_rows == 0 else (start_row + n_rows - 1)
    if args.max_rows and end_row < start_row + args.max_rows - 1:
        end_row = start_row + args.max_rows - 1

    for r in range(start_row, end_row + 1):
        if args.safe_fill:
            has_any = False
            for cidx in [col_nit, col_desc, col_ventas, col_cant]:
                if cidx and ws.cell(r, cidx).value not in (None, ""):
                    has_any = True; break
            if not has_any:
                continue
        if L_vend and L_nit:
            ws[f"{L_vend}{r}"] = f"=VLOOKUP({L_nit}{r},VENDEDORES!{vend_range},2,0)"
        if L_prec and L_desc_src:
            ws[f"{L_prec}{r}"] = f"=VLOOKUP({L_desc_src}{r},PRECIOS!{prec_range},2,0)"
        if L_desc and L_vent and L_cant and L_prec:
            ws[f"{L_desc}{r}"] = f"=1-(({L_vent}{r}*1.19)/{L_cant}{r}/{L_prec}{r})"

    wb.save(path)
    print(f"OK. Procesadas {n_rows} filas y fórmulas aplicadas sobre: {path}")

if __name__ == "__main__":
    main()
