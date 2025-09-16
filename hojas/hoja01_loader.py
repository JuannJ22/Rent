import argparse, os, re
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Font


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
DEFAULT_EXCZ_PREFIX = os.environ.get("EXCZPREFIX", "EXCZ980")
DEFAULT_CCOSTO_EXCZ_PREFIX = os.environ.get("CCOSTO_EXCZPREFIX", "EXCZ979")

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

def _pick_latest_excz(path: Path, prefix: str):
    pattern = rf'^{re.escape(prefix.lower())}.*\.(xlsx|xls|csv)$'
    candidates = []
    for p in path.iterdir():
        if not p.is_file():
            continue
        name = p.name.lower()
        if re.match(pattern, name):
            candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

def _read_excz_df(file: Path):
    """
    Lee un archivo EXCZ en distintos formatos intentando detectar la fila de
    cabeceras real.  Los reportes EXCZ suelen traer filas preliminares antes de
    las columnas, por lo que se importa sin cabecera y luego se busca la primera
    fila con suficientes datos para considerarla cabecera.
    """
    suffix = file.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        df_raw = pd.read_excel(file, sheet_name=None, header=None)
        # tomar la primera hoja
        if isinstance(df_raw, dict):
            df_raw = next(iter(df_raw.values()))
    elif suffix == ".csv":
        df_raw = pd.read_csv(file, sep=";", header=None, engine="python")
    else:
        raise ValueError("Formato no soportado: " + suffix)

    # Detectar fila de cabeceras
    header_row = None
    for i in range(min(len(df_raw), 50)):
        row = df_raw.iloc[i].dropna().astype(str).str.strip()
        if len(row) >= 3:
            header_row = i
            break
    if header_row is None:
        return pd.DataFrame()

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = df_raw.iloc[header_row].astype(str).tolist()
    return df

def _guess_map(df_cols):
    cols = { _norm(c): c for c in df_cols }
    def pick(*keys):
        for k in keys:
            nk = _norm(k)
            if nk in cols:
                return cols[nk]
        return None
    return {
        "centro_costo": pick(
            "centro de costo",
            "centro costo",
            "centro de costos",
            "punto de venta",
            "pto de venta",
            "punto",

            "centro",
            "zona"

        ),
        "nit": pick("nit","nit cliente","identificacion","identificación"),
        "cliente_combo": pick("nit - sucursal - cliente","cliente sucursal","cliente","razon social","razón social"),
        "descripcion": pick("descripcion","descripción","producto","nombre producto","item"),
        "cantidad": pick("cantidad","cant"),
        "ventas": pick("ventas","subtotal sin iva","total sin iva","valor venta","base"),
        "costos": pick("costos","costo","costo total","costo sin iva"),
        "renta": pick("% renta","renta","rentabilidad","rentabilidad venta"),
        "utili": pick("% utili","utili","utilidad","utilidad %","utilidad porcentaje"),
    }


def _normalize_spaces(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def _update_ccosto_sheets(wb, excz_dir, prefix, currency_fmt, border):
    config = [
        ("CCOSTO1", "0001   MOST. PRINCIPAL"),
        ("CCOSTO2", "0002   MOST. SUCURSAL"),
        ("CCOSTO3", "0003   MOSTRADOR CALARCA"),
        ("CCOSTO4", "0007   TIENDA PINTUCO"),
    ]

    excz_dir = Path(excz_dir)
    if not excz_dir.exists():
        print(f"ERROR: No existe la carpeta de EXCZ para CCOSTO: {excz_dir}")
        raise SystemExit(8)
    latest = _pick_latest_excz(excz_dir, prefix)
    if not latest:
        print(f"ERROR: No se encontró EXCZ para CCOSTO con prefijo {prefix} en {excz_dir}")
        raise SystemExit(6)

    df = _read_excz_df(latest)
    if df.empty:
        df = pd.DataFrame()

    mapping = _guess_map(df.columns)
    centro_col = mapping.get("centro_costo")
    if not centro_col:

        print("ERROR: El EXCZ para CCOSTO no contiene columna de Centro de Costo o Zona")

        raise SystemExit(7)

    columns = {
        key: mapping[key]
        for key in ["centro_costo", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]
        if mapping.get(key)
    }

    sub = df[list(columns.values())].copy() if columns else pd.DataFrame()
    sub.rename(columns={v: k for k, v in columns.items()}, inplace=True)

    for col in ["centro_costo", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]:
        if col not in sub.columns:
            sub[col] = pd.NA

    sub = sub.dropna(how="all")

    for col in ["cantidad", "ventas", "costos", "renta", "utili"]:
        if col in sub.columns:
            sub[col] = pd.to_numeric(sub[col], errors="coerce")

    sub["ccosto_norm"] = sub["centro_costo"].map(_normalize_spaces)

    order = ["centro_costo", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]
    headers = [
        "CENTRO DE COSTO",
        "DESCRIPCION",
        "CANTIDAD",
        "VENTAS",
        "COSTOS",
        "% RENTA",
        "% UTIL.",
    ]

    summary = {}

    for sheet_name, label in config:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)

        target_norm = _normalize_spaces(label)
        data = sub[sub["ccosto_norm"] == target_norm].copy()

        if data.empty:
            ws["A1"] = "ESTE PUNTO DE VENTA NO ABRIÓ HOY"
            summary[sheet_name] = 0
            continue

        data = data[order]

        mask_valid = data[["descripcion", "cantidad", "ventas", "costos", "renta", "utili"]].notna().any(axis=1)
        data = data[mask_valid]

        if data.empty:
            ws["A1"] = "ESTE PUNTO DE VENTA NO ABRIÓ HOY"
            summary[sheet_name] = 0
            continue

        subtotal_mask = data["descripcion"].astype(str).str.contains("subtotal", case=False, na=False)
        detail = data[~subtotal_mask]
        subtotal_rows = data[subtotal_mask]

        if not detail.empty and detail["renta"].notna().any():
            detail = detail.sort_values(by="renta", ascending=True, na_position="last")

        data = pd.concat([detail, subtotal_rows], ignore_index=True)

        for idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=idx, value=header)

        for i, row in enumerate(data.itertuples(index=False), start=2):
            values = [
                getattr(row, "centro_costo"),
                getattr(row, "descripcion"),
                getattr(row, "cantidad"),
                getattr(row, "ventas"),
                getattr(row, "costos"),
                getattr(row, "renta"),
                getattr(row, "utili"),
            ]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row=i, column=col_idx)
                cell.value = None if pd.isna(value) else value
                if col_idx in (4, 5) and cell.value is not None:
                    cell.number_format = currency_fmt
                cell.border = border

        summary[sheet_name] = len(data)

    return summary, latest

def main():
    p = argparse.ArgumentParser(description="Importa último EXCZ a Hoja 1 y aplica fórmulas fijas.")
    p.add_argument("--excel",   default=DEFAULT_EXCEL,   help="Ruta al INFORME_YYYYMMDD.xlsx")
    p.add_argument("--exczdir", default=DEFAULT_EXCZDIR, help="Carpeta de EXCZ")
    p.add_argument("--hoja",    default=None,            help="Nombre de la Hoja 1 (por defecto la primera)")
    p.add_argument("--excz-prefix", default=DEFAULT_EXCZ_PREFIX,
                   help="Prefijo del archivo EXCZ a buscar")
    p.add_argument("--max-rows", type=int, default=0,    help="Forzar número de filas (0 = según datos)")
    p.add_argument("--skip-import", action="store_true", help="No importar EXCZ, sólo aplicar fórmulas")
    p.add_argument("--safe-fill",  action="store_true", default=True, help="Sólo escribir en filas con datos")
    p.add_argument("--skip-ccosto", action="store_true", help="No actualizar hojas CCOSTO")
    p.add_argument("--ccosto-excz-prefix", default=DEFAULT_CCOSTO_EXCZ_PREFIX,
                   help="Prefijo del archivo EXCZ para hojas CCOSTO")
    args = p.parse_args()

    path = Path(args.excel)
    if not path.exists():
        print(f"ERROR: No existe el informe: {path}")
        raise SystemExit(2)

    wb = load_workbook(path)
    ws = wb[args.hoja] if args.hoja else wb.worksheets[0]

    # Asegurar encabezado fijo para código de vendedor
    ws["D6"] = "COD. VENDEDOR"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True)
    currency_fmt = "$#,##0.00"

    ccosto_summary = {}
    ccosto_file = None

    # --- Actualizar encabezado con fechas dinámicas -----------------------
    now = datetime.now()

    m = re.search(r"(\d{4})(\d{2})(\d{2})", path.stem)
    report_date = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else now

    for row in ws.iter_rows(min_row=1, max_row=6, max_col=ws.max_column):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            val = cell.value
            if "MES/DIA/ANIO" in val:
                cell.value = now.strftime("%m/%d/%Y")
            elif "FECHA DEL INFORME" in val:
                cell.value = val.replace("FECHA DEL INFORME", report_date.strftime("%m/%d/%Y"))
            elif "Procesado en" in val:
                cell.value = f"Procesado en: {now.strftime('%Y/%m/%d %H:%M:%S:%f')[:-3]}"
    # ---------------------------------------------------------------------

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
    col_vendedor = 4  # Columna D reservada para COD. VENDEDOR
    col_precio = idx("precio")
    col_descuento = idx("descuento")
    col_excz = idx("excz")

    start_row = header_row + 1

    # Congelar filas superiores para mantener visible el encabezado
    ws.freeze_panes = ws.cell(row=start_row, column=1)

    # Importar EXCZ más reciente
    n_rows = 0
    if not args.skip_import:
        excz_dir = Path(args.exczdir)
        if not excz_dir.exists():
            print(f"ERROR: No existe la carpeta de EXCZ: {excz_dir}")
            raise SystemExit(4)

        latest = _pick_latest_excz(excz_dir, args.excz_prefix)
        if not latest:
            print("ERROR: No se encontró EXCZ en la carpeta.")
            raise SystemExit(5)

        df = _read_excz_df(latest)
        m = _guess_map(df.columns)

        cols_needed = {k: v for k, v in m.items() if v is not None}
        sub = df[list(cols_needed.values())].copy()
        sub.rename(columns={v: k for k, v in cols_needed.items()}, inplace=True)

        # Derivar NIT desde "cliente_combo" si hace falta
        if "nit" not in sub.columns and "cliente_combo" in sub.columns:
            sub["nit"] = (
                sub["cliente_combo"].astype(str).str.extract(r"^(\d+)")[0]
            )

        # Convertir datos numéricos y ordenar por rentabilidad
        for col in ["ventas", "costos", "renta", "utili"]:
            if col in sub.columns:
                sub[col] = pd.to_numeric(sub[col], errors="coerce")
        if "renta" in sub.columns:
            sub = sub.sort_values(by="renta", ascending=True, na_position="last")

        # Eliminar filas de totales o cabeceras repetidas
        if "descripcion" in sub.columns:
            sub = sub[~sub["descripcion"].astype(str).str.contains("total", case=False, na=False)]
            sub = sub[sub["descripcion"].notna()]

        if args.max_rows and len(sub) > args.max_rows:
            sub = sub.iloc[:args.max_rows].copy()

        # Escribir al Excel
        for i, row in enumerate(sub.itertuples(index=False), start=start_row):
            cells = []
            if col_nit and "nit" in sub.columns:
                cells.append(ws.cell(i, col_nit, getattr(row, "nit")))
            if col_cliente_combo and "cliente_combo" in sub.columns:
                cells.append(ws.cell(i, col_cliente_combo, getattr(row, "cliente_combo")))
            if col_desc and "descripcion" in sub.columns:
                cells.append(ws.cell(i, col_desc, getattr(row, "descripcion")))
            if col_cant and "cantidad" in sub.columns:
                cells.append(ws.cell(i, col_cant, getattr(row, "cantidad")))
            if col_ventas and "ventas" in sub.columns:
                c = ws.cell(i, col_ventas, getattr(row, "ventas"))
                c.number_format = currency_fmt
                cells.append(c)
            if col_costos and "costos" in sub.columns:
                c = ws.cell(i, col_costos, getattr(row, "costos"))
                c.number_format = currency_fmt
                cells.append(c)
            if col_renta and "renta" in sub.columns:
                cells.append(ws.cell(i, col_renta, getattr(row, "renta")))
            if col_utili and "utili" in sub.columns:
                cells.append(ws.cell(i, col_utili, getattr(row, "utili")))
            if col_excz:
                cells.append(ws.cell(i, col_excz, latest.stem))
            for c in cells:
                c.border = border

        n_rows = len(sub)

    if not args.skip_import and not args.skip_ccosto:
        ccosto_summary, ccosto_file = _update_ccosto_sheets(
            wb, args.exczdir, args.ccosto_excz_prefix, currency_fmt, border
        )

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
            c = ws[f"{L_vend}{r}"]
            c.value = f"=VLOOKUP({L_nit}{r},VENDEDORES!{vend_range},2,0)"
            c.border = border
        if L_prec and L_desc_src:
            c = ws[f"{L_prec}{r}"]
            c.value = f"=VLOOKUP({L_desc_src}{r},PRECIOS!{prec_range},2,0)"
            c.border = border
        if L_desc and L_vent and L_cant and L_prec:
            c = ws[f"{L_desc}{r}"]
            c.value = f"=1-(({L_vent}{r}*1.19)/{L_cant}{r}/{L_prec}{r})"
            c.border = border

    # --- Fila de Total General -------------------------------------------
    total_label = "Total General"
    total_label_col = col_desc or col_cliente_combo or col_nit or 1

    # Eliminar totales previos para evitar duplicados
    to_delete = []
    for r in range(start_row, ws.max_row + 1):
        cell_val = ws.cell(r, total_label_col).value
        if isinstance(cell_val, str) and cell_val.strip().lower() == total_label.lower():
            to_delete.append(r)
    for offset, r in enumerate(to_delete):
        ws.delete_rows(r - offset)

    data_check_cols = [c for c in [col_nit, col_cliente_combo, col_desc, col_cant, col_ventas, col_costos] if c]
    last_data_row = start_row - 1
    for r in range(ws.max_row, start_row - 1, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in data_check_cols):
            last_data_row = r
            break

    total_row = last_data_row + 1
    label_cell = ws.cell(total_row, total_label_col, total_label)
    label_cell.font = bold
    label_cell.border = border

    def set_sum(col_idx, number_format=None):
        if not col_idx:
            return None
        cell = ws.cell(total_row, col_idx)
        if last_data_row >= start_row:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last_data_row})"
        else:
            cell.value = 0
        if number_format:
            cell.number_format = number_format
        cell.font = bold
        cell.border = border
        return cell

    total_cant_cell = set_sum(col_cant)
    total_ventas_cell = set_sum(col_ventas, currency_fmt)
    total_costos_cell = set_sum(col_costos, currency_fmt)

    if col_renta and total_ventas_cell and total_costos_cell:
        ventas_ref = f"{get_column_letter(col_ventas)}{total_row}"
        costos_ref = f"{get_column_letter(col_costos)}{total_row}"
        rent_cell = ws.cell(total_row, col_renta)
        rent_cell.value = f"=IF({ventas_ref}=0,0,({ventas_ref}-{costos_ref})/{ventas_ref})"
        rent_cell.number_format = "0.00%"
        rent_cell.font = bold
        rent_cell.border = border

    if col_utili and total_ventas_cell and total_costos_cell:
        util_cell = ws.cell(total_row, col_utili)
        util_cell.value = f"={get_column_letter(col_ventas)}{total_row}-{get_column_letter(col_costos)}{total_row}"
        util_cell.number_format = currency_fmt
        util_cell.font = bold
        util_cell.border = border

    for _, (_, col_idx) in hmap.items():
        cell = ws.cell(total_row, col_idx)
        cell.border = border

    wb.save(path)
    msg = f"OK. Procesadas {n_rows} filas y fórmulas aplicadas sobre: {path}"
    if ccosto_file:
        items = ", ".join(f"{k}={v}" for k, v in sorted(ccosto_summary.items())) or "sin datos"
        msg += f" | CCOSTO ({ccosto_file.name}): {items}"
    print(msg)

if __name__ == "__main__":
    main()
