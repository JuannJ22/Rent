from __future__ import annotations

import argparse
import numbers
import os
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Tuple

import pandas as pd
from pandas.api.types import is_numeric_dtype
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    # Garantiza que las importaciones absolutas funcionen incluso al ejecutar el
    # script directamente (por ejemplo, desde un .bat).
    sys.path.insert(0, str(PROJECT_ROOT))

from rentabilidad.core.dates import DateResolver, YesterdayStrategy
from rentabilidad.core.env import load_env
from rentabilidad.core.excz import ExczFileFinder, ExczMetadata
from rentabilidad.core.paths import PathContext, PathContextFactory, SPANISH_MONTHS


load_env()
PATH_CONTEXT: PathContext = PathContextFactory(os.environ).create()
DEFAULT_RENT_DIR = os.environ.get("RENT_DIR", str(PATH_CONTEXT.base_dir))
DEFAULT_EXCZDIR = os.environ.get("EXCZDIR", r"D:\\SIIWI01\\LISTADOS")
DEFAULT_EXCZ_PREFIX = os.environ.get("EXCZPREFIX", "EXCZ980")
DEFAULT_CCOSTO_EXCZ_PREFIX = os.environ.get("CCOSTO_EXCZPREFIX", "EXCZ979")
DEFAULT_COD_EXCZ_PREFIX = os.environ.get("COD_EXCZPREFIX", "EXCZ978")
DEFAULT_PRECIOS_DIR = os.environ.get("PRECIOS_DIR", str(PATH_CONTEXT.productos_dir))
DEFAULT_PRECIOS_PREFIX = os.environ.get("PRECIOS_PREFIX", "productos")
DEFAULT_VENDEDORES_DIR = os.environ.get(
    "VENDEDORES_DIR", str(PATH_CONTEXT.base_dir / "CodVendedor")
)
DEFAULT_VENDEDORES_PREFIX = os.environ.get(
    "VENDEDORES_PREFIX", "movimientocontable"
)
ACCOUNTING_FORMAT = '_-[$$-409]* #,##0.00_-;_-[$$-409]* (#,##0.00);_-[$$-409]* "-"??_-;_-@_-'


def _normalize_month_string(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value))
    stripped = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    cleaned = re.sub(r"[\-_/]+", " ", stripped)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip().lower()


_MONTH_NAME_LOOKUP = {
    _normalize_month_string(name): month for month, name in SPANISH_MONTHS.items()
}


def _extract_report_datetime(path: Path, fallback: date) -> datetime:
    stem = path.stem

    match = re.search(r"(\d{4})(\d{2})(\d{2})", stem)
    if match:
        year, month, day = map(int, match.groups())
        try:
            return datetime(year, month, day)
        except ValueError:
            pass

    normalized = _normalize_month_string(stem)
    for month_key, month_number in _MONTH_NAME_LOOKUP.items():
        pattern = rf"{re.escape(month_key)}\s*(\d{{1,2}})"
        match = re.search(pattern, normalized)
        if match:
            day = int(match.group(1))
            year = fallback.year
            try:
                return datetime(year, month_number, day)
            except ValueError:
                continue

    return datetime.combine(fallback, datetime.min.time())


def _make_unique_sheet_title(base: str, existing_titles: set[str]) -> str:
    """Genera un título de hoja único basado en ``base``.

    Si ``base`` ya existe en ``existing_titles`` se agregan sufijos numéricos
    hasta encontrar un nombre disponible.  Se respeta el límite de 31
    caracteres impuesto por Excel ajustando el prefijo cuando es necesario.
    """

    max_len = 31
    suffix = 1
    while True:
        suffix_text = f"_{suffix}"
        trimmed_base = base
        if len(trimmed_base + suffix_text) > max_len:
            trimmed_base = trimmed_base[: max_len - len(suffix_text)]
            if not trimmed_base:
                # Como último recurso usar sólo el número de sufijo.
                trimmed_base = str(suffix)
                suffix_text = ""
        candidate = f"{trimmed_base}{suffix_text}"
        if candidate not in existing_titles:
            return candidate
        suffix += 1


def _ensure_primary_sheet_title(wb, desired_title: str) -> None:
    """Renombra la hoja principal del libro a ``desired_title``.

    En caso de que exista otra hoja con el mismo nombre se renombra primero a
    un título único para evitar que openpyxl agregue sufijos automáticos.
    """

    if not wb.worksheets:
        return

    primary_sheet = wb.worksheets[0]
    if primary_sheet.title == desired_title:
        return

    existing_titles = {sheet.title for sheet in wb.worksheets if sheet is not primary_sheet}
    for sheet in wb.worksheets[1:]:
        if sheet.title != desired_title:
            continue
        existing_titles.discard(sheet.title)
        new_title = _make_unique_sheet_title(sheet.title, existing_titles)
        sheet.title = new_title
        existing_titles.add(new_title)

    primary_sheet.title = desired_title


def _clean_cell_value(value, *, strip: bool = True):
    if value is None:
        return None
    if value is pd.NA:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, str):
        if strip:
            cleaned = value.strip()
            return cleaned if cleaned else None
        return value if value != "" else None
    return value


def _normalize_nit_value(value):
    """Normaliza un NIT eliminando espacios y convirtiéndolo a número si es posible."""

    if value is None or value is pd.NA:
        return None

    if isinstance(value, str):
        text = value.strip()
    else:
        if pd.isna(value):
            return None
        if isinstance(value, numbers.Integral):
            text = f"{int(value)}"
        elif isinstance(value, numbers.Real):
            value_float = float(value)
            if value_float.is_integer():
                text = f"{int(value_float)}"
            else:
                text = str(value).strip()
        else:
            text = str(value).strip()

    text = re.sub(r"\s+", "", text)
    if not text:
        return None

    numeric_value = _try_convert_numeric(text)
    return numeric_value if numeric_value is not None else text


def _try_convert_numeric(text: str):
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return None

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

def _pick_excz_for_date(
    path: Path,
    prefix: str,
    report_date: date,
    *,
    use_latest: bool = False,
) -> Tuple[Path | None, list[ExczMetadata]]:
    directory = Path(path)
    finder = ExczFileFinder(directory)
    matches = list(finder.iter_matches(prefix))
    if use_latest:
        matches_sorted = sorted(matches, key=lambda meta: meta.modified_at, reverse=True)
        if matches_sorted:
            return matches_sorted[0].path, matches
        return None, matches
    for meta in matches:
        if meta.timestamp.date() == report_date:
            return meta.path, matches
    return None, matches

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


def _precios_candidate_dirs(base_dir):
    candidates = []
    seen = set()

    def add(path):
        if not path:
            return
        p = Path(path)
        key = str(p).lower()
        if key in seen:
            return
        seen.add(key)
        candidates.append(p)

    if base_dir:
        p = Path(base_dir)
        if p.is_file():
            add(p)
        else:
            add(p)
            if p.name.lower() != "productos":
                add(p / "Productos")

    default_dir = Path(DEFAULT_PRECIOS_DIR)
    if default_dir.is_file():
        add(default_dir)
    else:
        add(default_dir)
        if default_dir.name.lower() != "productos":
            add(default_dir / "Productos")

    rent_dir = Path(DEFAULT_RENT_DIR)
    add(rent_dir / "Productos")
    add(rent_dir)

    return candidates


def _find_latest_file_by_prefix(
    candidate_dirs,
    prefix: str | None,
    extensions: Tuple[str, ...],
):
    best_path: Path | None = None
    best_mtime = float("-inf")
    prefix_lower = prefix.lower() if prefix else ""
    allowed_exts = tuple(ext.lower() for ext in extensions)

    for dir_path in candidate_dirs:
        if not dir_path:
            continue
        path_obj = Path(dir_path)
        if path_obj.is_file():
            if (not prefix_lower or path_obj.name.lower().startswith(prefix_lower)) and (
                not allowed_exts or path_obj.suffix.lower() in allowed_exts
            ):
                try:
                    mtime = path_obj.stat().st_mtime
                except FileNotFoundError:
                    continue
                if mtime > best_mtime:
                    best_mtime = mtime
                    best_path = path_obj
            continue

        if not path_obj.exists():
            continue

        for child in path_obj.iterdir():
            if not child.is_file():
                continue
            if allowed_exts and child.suffix.lower() not in allowed_exts:
                continue
            if prefix_lower and not child.name.lower().startswith(prefix_lower):
                continue
            try:
                mtime = child.stat().st_mtime
            except FileNotFoundError:
                continue
            if mtime > best_mtime:
                best_mtime = mtime
                best_path = child

    return best_path


def _resolve_precios_path(
    report_date,
    *,
    explicit_file=None,
    directory=None,
    prefix=None,
    use_latest=False,
):
    prefix = prefix or DEFAULT_PRECIOS_PREFIX

    if explicit_file:
        path = Path(explicit_file)
        return (path if path.exists() else None), [path.parent], [path.name], True

    candidate_dirs = _precios_candidate_dirs(directory)
    extensions = (".xlsx", ".xlsm", ".xls")

    if use_latest:
        pattern_names = [
            f"{prefix or ''}*{ext}".strip() if prefix else f"*{ext}"
            for ext in extensions
        ]
        latest = _find_latest_file_by_prefix(candidate_dirs, prefix, extensions)
        if latest:
            return latest, candidate_dirs, pattern_names, False
        candidate_names = pattern_names
    else:
        base_name = report_date.strftime("%m%d")
        if prefix:
            base_name = f"{prefix}{base_name}"

        candidate_names = [f"{base_name}{ext}" for ext in extensions]

    for dir_path in candidate_dirs:
        if dir_path.is_file():
            if dir_path.exists():
                return dir_path, candidate_dirs, candidate_names, False
            continue
        if not dir_path.exists():
            continue
        for name in candidate_names:
            candidate = dir_path / name
            if candidate.exists():
                return candidate, candidate_dirs, candidate_names, False

    return None, candidate_dirs, candidate_names, False


def _vendedores_candidate_dirs(base_dir):
    candidates = []
    seen = set()

    def add(path):
        if not path:
            return
        p = Path(path)
        key = str(p).lower()
        if key in seen:
            return
        seen.add(key)
        candidates.append(p)

    if base_dir:
        p = Path(base_dir)
        add(p)

    default_dir = Path(DEFAULT_VENDEDORES_DIR)
    add(default_dir)

    rent_dir = Path(DEFAULT_RENT_DIR)
    add(rent_dir / "CodVendedor")
    add(rent_dir)

    return candidates


def _resolve_vendedores_path(
    report_date,
    *,
    explicit_file=None,
    directory=None,
    prefix=None,
    use_latest=False,
):
    prefix = prefix or DEFAULT_VENDEDORES_PREFIX

    if explicit_file:
        path = Path(explicit_file)
        return (path if path.exists() else None), [path.parent], [path.name], True

    candidate_dirs = _vendedores_candidate_dirs(directory)
    search_dirs = []

    extensions = (".xlsx", ".xlsm", ".xls", ".csv")

    if use_latest:
        pattern_names = [
            f"{prefix or ''}*{ext}".strip() if prefix else f"*{ext}"
            for ext in extensions
        ]
        latest = _find_latest_file_by_prefix(candidate_dirs, prefix, extensions)
        if latest:
            return latest, candidate_dirs, pattern_names, False
        candidate_names = pattern_names
    else:
        base_name = report_date.strftime("%d%m")
        if prefix:
            base_name = f"{prefix}{base_name}"

        candidate_names = [f"{base_name}{ext}" for ext in extensions]

    for dir_path in candidate_dirs:
        p = Path(dir_path)
        if p.is_file():
            if p.exists() and p.name.lower() in (name.lower() for name in candidate_names):
                return p, [p.parent], candidate_names, False
            search_dirs.append(p.parent)
            continue

        if not p.exists():
            search_dirs.append(p)
            continue

        search_dirs.append(p)
        for name in candidate_names:
            candidate = p / name
            if candidate.exists():
                return candidate, search_dirs, candidate_names, False

    return None, search_dirs or candidate_dirs, candidate_names, False


def _update_vendedores_sheet(
    wb,
    *,
    report_date,
    vendedores_file=None,
    vendedores_dir=None,
    vendedores_prefix=None,
    use_latest=False,
):
    sheet_name = "VENDEDORES"
    if sheet_name not in wb.sheetnames:
        return {}, None

    path, search_dirs, candidate_names, explicit = _resolve_vendedores_path(
        report_date,
        explicit_file=vendedores_file,
        directory=vendedores_dir,
        prefix=vendedores_prefix,
        use_latest=use_latest,
    )

    if not path or not path.exists():
        if explicit:
            print(
                "ERROR: No existe el archivo de vendedores especificado "
                f"({vendedores_file})."
            )
        else:
            locations = [str(d) for d in search_dirs if d]
            if not locations:
                base_location = (
                    Path(vendedores_dir)
                    if vendedores_dir
                    else Path(DEFAULT_VENDEDORES_DIR)
                )
                locations = [str(base_location)]
            names = ", ".join(candidate_names) if candidate_names else ""
            if use_latest:
                print(
                    "ERROR: No se encontró un archivo de vendedores más "
                    "reciente "
                    f"con prefijo {vendedores_prefix or DEFAULT_VENDEDORES_PREFIX} "
                    f"en: {', '.join(locations)}"
                )
            else:
                print(
                    "ERROR: No se encontró archivo de vendedores "
                    f"({names}) en: {', '.join(locations)}"
                )
        raise SystemExit(20)

    src_wb = load_workbook(filename=path, data_only=True, read_only=True)
    try:
        src_ws = src_wb.active
        rows = [
            (row[0], row[1])
            for row in src_ws.iter_rows(min_row=1, max_col=2, values_only=True)
        ]
    finally:
        src_wb.close()

    ws = wb[sheet_name]
    ws.delete_rows(1, ws.max_row)

    rows_written = 0
    for col_a, col_b in rows:
        if col_a in (None, "") and col_b in (None, ""):
            continue
        rows_written += 1
        ws.cell(row=rows_written, column=1, value=None if col_b in ("", None) else col_b)
        ws.cell(row=rows_written, column=2, value=None if col_a in ("", None) else col_a)

    summary = {"rows": rows_written}
    return summary, path


def _guess_map(df_cols):
    cols = {_norm(c): c for c in df_cols}

    def pick(*keys, contains=None):
        for k in keys:
            nk = _norm(k)
            if nk in cols:
                return cols[nk]

        if contains:
            candidates = tuple(_norm(c) for c in contains)
            for col_norm, original in cols.items():
                for needle in candidates:
                    if needle and needle in col_norm:
                        return original

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
        "vendedor": pick(
            "cod vendedor",
            "cod. vendedor",
            "codigo vendedor",
            "código vendedor",
            "vendedor",
            "nom vendedor",
            "nombre vendedor",
            "vendedor cod",
        ),
        "nit": pick("nit","nit cliente","identificacion","identificación"),
        "cliente_combo": pick("nit - sucursal - cliente","cliente sucursal","cliente","razon social","razón social"),
        "linea": pick("linea", "línea"),
        "grupo": pick("grupo", "grupo descripción", contains=("grupo",)),
        "descripcion": pick("descripcion","descripción","producto","nombre producto","item"),
        "cantidad": pick("cantidad","cant"),
        "ventas": pick("ventas","subtotal sin iva","total sin iva","valor venta","base"),
        "costos": pick("costos","costo","costo total","costo sin iva"),
        "renta": pick(
            "% renta",
            "renta",
            "rentabilidad",
            "rentabilidad venta",
            contains=("rentab", "rentabilidad", "renta"),
        ),
        "utili": pick(
            "% utili",
            "utili",
            "utilidad",
            "utilidad %",
            "utilidad porcentaje",
            contains=("utili", "utilid", "util"),
        ),
    }


def _normalize_spaces(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def _strip_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def _normalize_lookup_value(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = _strip_accents(text)
    text = re.sub(r"[^0-9a-z]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _normalize_ccosto_value(value) -> str:
    return _normalize_lookup_value(value)


def _parse_numeric_series(series: pd.Series, *, is_percent: bool = False) -> pd.Series:
    """Convert a series with numeric-like strings into floats.

    Handles values that include percentage symbols as well as numbers that use
    comma or dot as decimal/thousand separators.
    """

    if series.empty:
        return series

    if is_numeric_dtype(series):
        result = pd.to_numeric(series, errors="coerce")
    else:
        text = series.astype(str).str.strip()
        has_percent = text.str.contains("%", regex=False, na=False)
        cleaned = text.str.replace(r"[^0-9,\-.]+", "", regex=True)
        both_sep = cleaned.str.contains(",", na=False) & cleaned.str.contains(".", na=False)
        cleaned = cleaned.where(~both_sep, cleaned.str.replace(".", "", regex=False))
        cleaned = cleaned.str.replace(",", ".", regex=False)
        result = pd.to_numeric(cleaned, errors="coerce")

    if is_percent:
        if "has_percent" not in locals():
            text = series.astype(str).str.strip()
            has_percent = text.str.contains("%", regex=False, na=False)
        result = result.where(~has_percent, result / 100)
        adjust_mask = (~has_percent) & result.notna() & result.abs().between(1, 100)
        result.loc[adjust_mask] = result.loc[adjust_mask] / 100

    return result


def _select_rows_by_norm(df: pd.DataFrame, label: str, norm_col: str) -> pd.DataFrame:
    if norm_col not in df.columns:
        return df.iloc[0:0].copy()

    target_norm = _normalize_lookup_value(label)
    if not target_norm:
        return df.iloc[0:0].copy()

    series = df[norm_col].fillna("")
    masks = []

    masks.append(series == target_norm)
    compact_target = target_norm.replace(" ", "")
    masks.append(series.str.replace(" ", "", regex=False) == compact_target)

    target_tokens = tuple(token for token in target_norm.split() if token)
    if target_tokens:
        target_set = set(target_tokens)

        def token_match(val: str) -> bool:
            tokens = tuple(token for token in val.split() if token)
            if not tokens:
                return False
            val_set = set(tokens)
            if target_set.issubset(val_set) or val_set.issubset(target_set):
                return True

            # Verificación adicional considerando coincidencias parciales y números sin ceros a la izquierda
            def check_token(target_token: str) -> bool:
                if target_token.isdigit():
                    target_num = int(target_token)
                    for token in tokens:
                        if token.isdigit() and int(token) == target_num:
                            return True
                    return target_token in val
                return target_token in val

            return all(check_token(t) for t in target_tokens)

        masks.append(series.apply(token_match))

    masks.append(series.str.contains(target_norm, na=False, regex=False))

    for mask in masks:
        if mask.any():
            return df.loc[mask].copy()

    return df.iloc[0:0].copy()


def _select_ccosto_rows(df: pd.DataFrame, label: str) -> pd.DataFrame:
    return _select_rows_by_norm(df, label, "ccosto_norm")


def _select_cod_rows(df: pd.DataFrame, label: str) -> pd.DataFrame:
    return _select_rows_by_norm(df, label, "cod_norm")


def _update_ccosto_sheets(
    wb,
    excz_dir,
    prefix,
    accounting_fmt,
    border,
    report_date: date,
    *,
    use_latest: bool = False,
):
    config = [
        ("CCOSTO 1", "0001   MOST. PRINCIPAL"),
        ("CCOSTO 2", "0002   MOST. SUCURSAL"),
        ("CCOSTO 3", "0003   MOSTRADOR CALARCA"),
        ("CCOSTO 4", "0007   TIENDA PINTUCO"),
    ]

    excz_dir = Path(excz_dir)
    if not excz_dir.exists():
        print(f"ERROR: No existe la carpeta de EXCZ para CCOSTO: {excz_dir}")
        raise SystemExit(8)
    latest, matches = _pick_excz_for_date(
        excz_dir,
        prefix,
        report_date,
        use_latest=use_latest,
    )
    if not latest:
        available = ", ".join(meta.path.name for meta in matches) or "sin archivos"
        if use_latest:
            print(
                "ERROR: No se encontró un EXCZ para CCOSTO más reciente "
                f"con prefijo {prefix} en {excz_dir}. Disponibles: {available}"
            )
        else:
            print(
                "ERROR: No se encontró EXCZ para CCOSTO "
                f"con prefijo {prefix} y fecha {report_date:%Y-%m-%d} en {excz_dir}. "
                f"Disponibles: {available}"
            )
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

    for col in ["cantidad", "ventas", "costos"]:
        if col in sub.columns:
            sub[col] = _parse_numeric_series(sub[col])
    for col in ["renta", "utili"]:
        if col in sub.columns:
            sub[col] = _parse_numeric_series(sub[col])

    sub["ccosto_norm"] = sub["centro_costo"].map(_normalize_ccosto_value)

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
    bold_font = Font(bold=True)

    for sheet_name, label in config:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)

        data = _select_ccosto_rows(sub, label)

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

        start_row = 2
        last_data_row = start_row - 1

        for i, row in enumerate(data.itertuples(index=False), start=start_row):
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
                    cell.number_format = accounting_fmt
                cell.border = border

            last_data_row = i

        summary[sheet_name] = len(data)

        if last_data_row >= start_row:
            total_row = last_data_row + 1
            label_col_idx = order.index("descripcion") + 1
            label_cell = ws.cell(total_row, label_col_idx, "Total General")
            label_cell.font = bold_font
            label_cell.border = border

            def set_sum_for(col_key, number_format=None):
                col_idx = order.index(col_key) + 1
                cell = ws.cell(total_row, col_idx)
                col_letter = get_column_letter(col_idx)
                cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last_data_row})"
                if number_format:
                    cell.number_format = number_format
                cell.font = bold_font
                cell.border = border
                return cell

            set_sum_for("cantidad")
            total_ventas_cell = set_sum_for("ventas", accounting_fmt)
            total_costos_cell = set_sum_for("costos", accounting_fmt)

            util_col_idx = order.index("utili") + 1
            util_cell = ws.cell(total_row, util_col_idx)
            if total_ventas_cell and total_costos_cell:
                ventas_ref = total_ventas_cell.coordinate
                costos_ref = total_costos_cell.coordinate
                util_cell.value = f"=IF({costos_ref}=0,0,({ventas_ref}/{costos_ref})-1)"
            else:
                util_cell.value = 0
            util_cell.number_format = "0.00%"
            util_cell.font = bold_font
            util_cell.border = border

            ventas_ref = f"{get_column_letter(order.index('ventas') + 1)}{total_row}"
            costos_ref = f"{get_column_letter(order.index('costos') + 1)}{total_row}"

            rent_cell = ws.cell(total_row, order.index("renta") + 1)
            rent_cell.value = f"=IF({ventas_ref}=0,0,1-({costos_ref}/{ventas_ref}))"
            rent_cell.number_format = "0.00%"
            rent_cell.font = bold_font
            rent_cell.border = border

            for col_idx in range(1, len(order) + 1):
                cell = ws.cell(total_row, col_idx)
                cell.border = border

            if sheet_name == "CCOSTO 4":
                for row_idx in range(start_row, last_data_row + 1):
                    cell = ws.cell(row_idx, 1)
                    value = cell.value
                    if value in (None, ""):
                        continue
                    text_value = str(value)
                    new_value = text_value.replace("7", "4")
                    if new_value != text_value:
                        cell.value = new_value

    return summary, latest


def _update_lineas_sheet(wb, data: pd.DataFrame, accounting_fmt: str, border):
    sheet_name = "LINEAS"
    if sheet_name not in wb.sheetnames:
        return {}

    ws = wb[sheet_name]
    ws.delete_rows(1, ws.max_row)

    headers = [
        "LÍNEA  DESCRIPCIÓN",
        "GRUPO  DESCRIPCIÓN",
        "CANTIDAD",
        "VENTAS",
        "COSTO",
        "%RENTABILIDAD",
        "%UTILIDAD",
    ]

    bold_font = Font(bold=True)
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        cell.font = bold_font
        cell.border = border

    ws.freeze_panes = ws.cell(row=2, column=1)

    required_cols = ["linea", "grupo", "descripcion", "cantidad", "ventas", "costos"]
    if data is None or data.empty:
        cell = ws.cell(row=2, column=1, value="SIN DATOS PARA MOSTRAR")
        cell.border = border
        return {"lineas": 0, "grupos": 0}

    missing = [col for col in required_cols if col not in data.columns]
    if missing:
        cell = ws.cell(row=2, column=1, value="FALTAN COLUMNAS PARA GENERAR EL REPORTE")
        cell.border = border
        return {}

    detail_mask = (
        data["descripcion"].notna()
        & data["linea"].notna()
        & data["grupo"].notna()
        & ~data["linea"].astype(str).str.contains("total", case=False, na=False)
        & ~data["grupo"].astype(str).str.contains("total", case=False, na=False)
    )
    detail = data.loc[detail_mask, required_cols].copy()

    if detail.empty:
        cell = ws.cell(row=2, column=1, value="SIN DATOS PARA MOSTRAR")
        cell.border = border
        return {"lineas": 0, "grupos": 0}

    def clean_text(value):
        return re.sub(r"\s+", " ", str(value).strip()) if pd.notna(value) else ""

    detail["linea"] = detail["linea"].map(clean_text)
    detail["grupo"] = detail["grupo"].map(clean_text)

    for col in ["cantidad", "ventas", "costos"]:
        detail[col] = _parse_numeric_series(detail[col]).fillna(0)

    aggregated = (
        detail.groupby(["linea", "grupo"], as_index=False)[["cantidad", "ventas", "costos"]]
        .sum()
    )

    if aggregated.empty:
        cell = ws.cell(row=2, column=1, value="SIN DATOS PARA MOSTRAR")
        cell.border = border
        return {"lineas": 0, "grupos": 0}

    def extract_code(text: str) -> int:
        if not text:
            return 10**6
        match = re.search(r"\d+", text)
        if match:
            try:
                return int(match.group())
            except ValueError:
                return 10**6
        return 10**6

    aggregated["line_code"] = aggregated["linea"].map(extract_code)
    aggregated["grupo_code"] = aggregated["grupo"].map(extract_code)
    aggregated.sort_values(["line_code", "linea", "grupo_code", "grupo"], inplace=True)

    line_summary = (
        aggregated.groupby(["linea", "line_code"], as_index=False)[["cantidad", "ventas", "costos"]]
        .sum()
    )
    line_summary.sort_values(["line_code", "linea"], inplace=True)

    groups_by_line = {line: grp for line, grp in aggregated.groupby("linea", sort=False)}

    def format_total_label(text: str) -> str:
        cleaned = re.sub(r"\s+", " ", text.strip()) if text else ""
        cleaned = cleaned.replace("-", " ")
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return f"Total {cleaned}" if cleaned else "Total"

    def compute_metrics(ventas: float, costos: float) -> Tuple[float, float]:
        ventas_val = 0.0 if pd.isna(ventas) else float(ventas)
        costos_val = 0.0 if pd.isna(costos) else float(costos)
        rent = 0.0 if ventas_val == 0 else 1 - (costos_val / ventas_val)
        util = 0.0 if costos_val == 0 else (ventas_val / costos_val) - 1
        return rent, util

    def safe_numeric(value):
        return 0.0 if pd.isna(value) else float(value)

    def write_cell(row, col, value, *, number_format=None, bold=False):
        cell = ws.cell(row=row, column=col)
        cell.value = None if pd.isna(value) else value
        if number_format:
            cell.number_format = number_format
        if bold:
            cell.font = bold_font
        cell.border = border
        return cell

    cantidad_format = "#,##0.00"
    row_idx = 2
    for _, line_row in line_summary.iterrows():
        line_name = line_row["linea"]
        line_label = format_total_label(line_name)
        line_cant = safe_numeric(line_row["cantidad"])
        line_ventas = safe_numeric(line_row["ventas"])
        line_costos = safe_numeric(line_row["costos"])
        line_rent, line_util = compute_metrics(line_ventas, line_costos)

        groups = groups_by_line.get(line_name)
        if groups is not None:
            for _, group_row in groups.iterrows():
                group_label = format_total_label(group_row["grupo"])
                group_cant = safe_numeric(group_row["cantidad"])
                group_ventas = safe_numeric(group_row["ventas"])
                group_costos = safe_numeric(group_row["costos"])
                group_rent, group_util = compute_metrics(group_ventas, group_costos)

                write_cell(row_idx, 1, None)
                write_cell(row_idx, 2, group_label)
                write_cell(row_idx, 3, group_cant, number_format=cantidad_format)
                write_cell(row_idx, 4, group_ventas, number_format=accounting_fmt)
                write_cell(row_idx, 5, group_costos, number_format=accounting_fmt)
                write_cell(row_idx, 6, group_rent, number_format="0.00%")
                write_cell(row_idx, 7, group_util, number_format="0.00%")

                row_idx += 1

        write_cell(row_idx, 1, line_label, bold=True)
        write_cell(row_idx, 2, None, bold=True)
        write_cell(row_idx, 3, line_cant, number_format=cantidad_format, bold=True)
        write_cell(row_idx, 4, line_ventas, number_format=accounting_fmt, bold=True)
        write_cell(row_idx, 5, line_costos, number_format=accounting_fmt, bold=True)
        write_cell(row_idx, 6, line_rent, number_format="0.00%", bold=True)
        write_cell(row_idx, 7, line_util, number_format="0.00%", bold=True)

        row_idx += 1

    totals = line_summary[["cantidad", "ventas", "costos"]].sum()
    total_rent, total_util = compute_metrics(totals["ventas"], totals["costos"])

    write_cell(row_idx, 1, "Total General", bold=True)
    write_cell(row_idx, 2, None, bold=True)
    write_cell(row_idx, 3, safe_numeric(totals["cantidad"]), number_format=cantidad_format, bold=True)
    write_cell(row_idx, 4, safe_numeric(totals["ventas"]), number_format=accounting_fmt, bold=True)
    write_cell(row_idx, 5, safe_numeric(totals["costos"]), number_format=accounting_fmt, bold=True)
    write_cell(row_idx, 6, total_rent, number_format="0.00%", bold=True)
    write_cell(row_idx, 7, total_util, number_format="0.00%", bold=True)

    return {
        "lineas": int(line_summary.shape[0]),
        "grupos": int(aggregated.shape[0]),
    }


def _update_cod_sheets(
    wb,
    excz_dir,
    prefix,
    accounting_fmt,
    border,
    report_date: date,
    *,
    use_latest: bool = False,
):
    config = [
        ("COD24", "0024", "CR CARLOS ALBERTO TOVAR HERRER"),
        ("COD25", "0025", "CO CARLOS ALBERTO TOVAR HERRER"),
        ("COD26", "0026", "CR OMAR SMITH PARRADO BELTRAN"),
        ("COD27", "0027", "CR OMAR SMITH PARRADO BELTRAN"),
        ("COD29", "0029", "CO BEATRIZ LONDOÑO VELASQUEZ"),
        ("COD30", "0030", "CR BEATRIZ LONDOÑO VELASQUEZ"),
        ("COD51", "0051", "CR MARICELY LONDOÑO"),
        ("COD52", "0052", "CO MARICELY LONDOÑO"),
    ]

    excz_dir = Path(excz_dir)
    if not excz_dir.exists():
        print(f"ERROR: No existe la carpeta de EXCZ para COD: {excz_dir}")
        raise SystemExit(18)

    latest, matches = _pick_excz_for_date(
        excz_dir,
        prefix,
        report_date,
        use_latest=use_latest,
    )
    if not latest:
        available = ", ".join(meta.path.name for meta in matches) or "sin archivos"
        if use_latest:
            print(
                "ERROR: No se encontró un EXCZ para COD más reciente "
                f"con prefijo {prefix} en {excz_dir}. Disponibles: {available}"
            )
        else:
            print(
                "ERROR: No se encontró EXCZ para COD "
                f"con prefijo {prefix} y fecha {report_date:%Y-%m-%d} en {excz_dir}. "
                f"Disponibles: {available}"
            )
        raise SystemExit(16)

    df = _read_excz_df(latest)
    if df.empty:
        df = pd.DataFrame()

    mapping = _guess_map(df.columns)
    vendedor_col = mapping.get("vendedor") or mapping.get("centro_costo")
    if not vendedor_col:
        print("ERROR: El EXCZ para COD no contiene columna de Vendedor")
        raise SystemExit(17)

    columns = {}
    for key in ["vendedor", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]:
        if key == "vendedor":
            columns[key] = vendedor_col
        else:
            col = mapping.get(key)
            if col:
                columns[key] = col

    sub = df[list(dict.fromkeys(columns.values()))].copy() if columns else pd.DataFrame()
    sub.rename(columns={v: k for k, v in columns.items()}, inplace=True)

    for col in ["vendedor", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]:
        if col not in sub.columns:
            sub[col] = pd.NA

    sub = sub.dropna(how="all")

    for col in ["cantidad", "ventas", "costos"]:
        if col in sub.columns:
            sub[col] = _parse_numeric_series(sub[col])
    for col in ["renta", "utili"]:
        if col in sub.columns:
            sub[col] = _parse_numeric_series(sub[col])

    sub["cod_norm"] = sub["vendedor"].map(_normalize_lookup_value)

    order = ["vendedor", "descripcion", "cantidad", "ventas", "costos", "renta", "utili"]
    headers = [
        "COD. VENDEDOR",
        "DESCRIPCION",
        "CANTIDAD",
        "VENTAS",
        "COSTOS",
        "% RENTA",
        "% UTIL.",
    ]

    summary = {}
    bold_font = Font(bold=True)

    for sheet_name, code, description in config:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)

        data = _select_cod_rows(sub, code)
        if data.empty and description:
            data = _select_cod_rows(sub, description)
        if data.empty and description:
            combo_label = f"{code} {description}".strip()
            data = _select_cod_rows(sub, combo_label)

        if data.empty:
            ws["A1"] = "ESTE VENDEDOR NO REGISTRA VENTAS"
            summary[sheet_name] = 0
            continue

        data = data[order]

        mask_valid = data[["descripcion", "cantidad", "ventas", "costos", "renta", "utili"]].notna().any(axis=1)
        data = data[mask_valid]

        if data.empty:
            ws["A1"] = "ESTE VENDEDOR NO REGISTRA VENTAS"
            summary[sheet_name] = 0
            continue

        total_000_mask = pd.Series(False, index=data.index)
        for col in ["vendedor", "descripcion"]:
            if col in data.columns:
                total_000_mask |= data[col].astype(str).str.contains(
                    r"total\s+0{2,}", case=False, regex=True, na=False
                )
        if total_000_mask.any():
            data = data[~total_000_mask]

        if data.empty:
            ws["A1"] = "ESTE VENDEDOR NO REGISTRA VENTAS"
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

        start_row = 2
        last_data_row = start_row - 1

        for i, row in enumerate(data.itertuples(index=False), start=start_row):
            values = [
                getattr(row, "vendedor"),
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
                    cell.number_format = accounting_fmt
                cell.border = border

            last_data_row = i

        summary[sheet_name] = len(data)

        if last_data_row >= start_row:
            total_row = last_data_row + 1
            label_col_idx = order.index("descripcion") + 1
            label_cell = ws.cell(total_row, label_col_idx, "Total General")
            label_cell.font = bold_font
            label_cell.border = border

            def set_sum_for(col_key, number_format=None):
                col_idx = order.index(col_key) + 1
                cell = ws.cell(total_row, col_idx)
                col_letter = get_column_letter(col_idx)
                cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last_data_row})"
                if number_format:
                    cell.number_format = number_format
                cell.font = bold_font
                cell.border = border
                return cell

            set_sum_for("cantidad")
            total_ventas_cell = set_sum_for("ventas", accounting_fmt)
            total_costos_cell = set_sum_for("costos", accounting_fmt)

            util_col_idx = order.index("utili") + 1
            util_cell = ws.cell(total_row, util_col_idx)
            if total_ventas_cell and total_costos_cell:
                ventas_ref = total_ventas_cell.coordinate
                costos_ref = total_costos_cell.coordinate
                util_cell.value = f"=IF({costos_ref}=0,0,({ventas_ref}/{costos_ref})-1)"
            else:
                util_cell.value = 0
            util_cell.number_format = "0.00%"
            util_cell.font = bold_font
            util_cell.border = border

            ventas_ref = f"{get_column_letter(order.index('ventas') + 1)}{total_row}"
            costos_ref = f"{get_column_letter(order.index('costos') + 1)}{total_row}"

            rent_cell = ws.cell(total_row, order.index("renta") + 1)
            rent_cell.value = f"=IF({ventas_ref}=0,0,1-({costos_ref}/{ventas_ref}))"
            rent_cell.number_format = "0.00%"
            rent_cell.font = bold_font
            rent_cell.border = border

            for col_idx in range(1, len(order) + 1):
                cell = ws.cell(total_row, col_idx)
                cell.border = border

    return summary, latest


def _update_precios_sheet(
    wb,
    *,
    report_date,
    precios_file=None,
    precios_dir=None,
    precios_prefix=None,
    use_latest=False,
):
    sheet_name = "PRECIOS"
    if sheet_name not in wb.sheetnames:
        return {}, None

    path, search_dirs, candidate_names, explicit = _resolve_precios_path(
        report_date,
        explicit_file=precios_file,
        directory=precios_dir,
        prefix=precios_prefix,
        use_latest=use_latest,
    )

    if not path or not path.exists():
        if explicit:
            print(
                "ERROR: No existe el archivo de precios especificado "
                f"({precios_file})."
            )
        else:
            locations = [str(d) for d in search_dirs if d]
            if not locations:
                base_location = Path(precios_dir) if precios_dir else Path(DEFAULT_PRECIOS_DIR)
                locations = [str(base_location)]
            names = ", ".join(candidate_names) if candidate_names else ""
            if use_latest:
                print(
                    "ERROR: No se encontró un archivo de precios más reciente "
                    f"con prefijo {precios_prefix or DEFAULT_PRECIOS_PREFIX} "
                    f"en: {', '.join(locations)}"
                )
            else:
                print(
                    "ERROR: No se encontró archivo de precios "
                    f"({names}) en: {', '.join(locations)}"
                )
        raise SystemExit(19)

    src_wb = load_workbook(filename=path, data_only=True, read_only=True)
    try:
        src_ws = src_wb.active
        rows = [tuple(row) for row in src_ws.iter_rows(values_only=True)]
    finally:
        src_wb.close()

    ws = wb[sheet_name]
    ws.delete_rows(1, ws.max_row)

    rows_with_data = 0
    last_row_index = 0
    max_used_cols = 0

    for row_idx, values in enumerate(rows, start=1):
        if not values:
            continue
        row_has_data = False
        for col_idx, value in enumerate(values, start=1):
            if value in (None, ""):
                continue
            ws.cell(row=row_idx, column=col_idx, value=value)
            row_has_data = True
            if col_idx > max_used_cols:
                max_used_cols = col_idx
        if row_has_data:
            rows_with_data += 1
        last_row_index = row_idx

    summary = {
        "rows": rows_with_data,
        "total_rows": last_row_index,
        "columns": max_used_cols,
    }

    return summary, path


def main():
    p = argparse.ArgumentParser(description="Importa el EXCZ del día previo y aplica fórmulas fijas.")
    p.add_argument(
        "--excel",
        default=None,
        help="Ruta al informe '<Mes> DD.xlsx' a actualizar",
    )
    p.add_argument("--exczdir", default=DEFAULT_EXCZDIR, help="Carpeta de EXCZ")
    p.add_argument("--hoja",    default=None,            help="Nombre de la Hoja 1 (por defecto la primera)")
    p.add_argument("--excz-prefix", default=DEFAULT_EXCZ_PREFIX,
                   help="Prefijo del archivo EXCZ a buscar")
    p.add_argument("--fecha", default=None, help="Fecha del informe (YYYY-MM-DD, por defecto día anterior)")
    p.add_argument("--max-rows", type=int, default=0,    help="Forzar número de filas (0 = según datos)")
    p.add_argument("--skip-import", action="store_true", help="No importar EXCZ, sólo aplicar fórmulas")
    p.add_argument("--safe-fill",  action="store_true", default=True, help="Sólo escribir en filas con datos")
    p.add_argument(
        "--use-latest-sources",
        action="store_true",
        help=(
            "Usa los archivos más recientes por fecha de modificación "
            "para EXCZ, vendedores y precios, ignorando la búsqueda por fecha."
        ),
    )
    p.add_argument("--skip-ccosto", action="store_true", help="No actualizar hojas CCOSTO")
    p.add_argument("--ccosto-excz-prefix", default=DEFAULT_CCOSTO_EXCZ_PREFIX,
                   help="Prefijo del archivo EXCZ para hojas CCOSTO")
    p.add_argument("--skip-cod", action="store_true", help="No actualizar hojas COD")
    p.add_argument("--cod-excz-prefix", default=DEFAULT_COD_EXCZ_PREFIX,
                   help="Prefijo del archivo EXCZ para hojas COD")
    p.add_argument("--skip-vendedores", action="store_true", help="No actualizar hoja VENDEDORES")
    p.add_argument(
        "--vendedores-file",
        default=None,
        help="Ruta exacta del archivo de vendedores a copiar en la hoja VENDEDORES",
    )
    p.add_argument(
        "--vendedores-dir",
        default=DEFAULT_VENDEDORES_DIR,
        help="Carpeta donde se buscará movimientocontableDDMM.* si no se especifica archivo",
    )
    p.add_argument(
        "--vendedores-prefix",
        default=DEFAULT_VENDEDORES_PREFIX,
        help="Prefijo del archivo de vendedores (movimientocontable por defecto)",
    )
    p.add_argument("--skip-precios", action="store_true", help="No actualizar hoja PRECIOS")
    p.add_argument(
        "--precios-file",
        default=None,
        help="Ruta exacta del archivo de precios a copiar en la hoja PRECIOS",
    )
    p.add_argument(
        "--precios-dir",
        default=DEFAULT_PRECIOS_DIR,
        help="Carpeta donde se buscará productosMMDD.xlsx si no se especifica archivo",
    )
    p.add_argument(
        "--precios-prefix",
        default=DEFAULT_PRECIOS_PREFIX,
        help="Prefijo del archivo de precios (productos por defecto)",
    )
    args = p.parse_args()

    resolver = DateResolver(YesterdayStrategy())
    report_date = resolver.resolve(args.fecha)

    use_latest = args.use_latest_sources

    context = PATH_CONTEXT
    if args.excel:
        path = Path(args.excel)
    else:
        path = context.informe_path(report_date)

    if not path.exists():
        print(f"ERROR: No existe el informe: {path}")
        raise SystemExit(2)

    wb = load_workbook(path)

    desired_main_sheet_name = path.stem.upper()
    original_primary_title = wb.worksheets[0].title if wb.worksheets else None
    if desired_main_sheet_name and wb.worksheets:
        _ensure_primary_sheet_title(wb, desired_main_sheet_name)
        if args.hoja and args.hoja == original_primary_title:
            args.hoja = desired_main_sheet_name

    ws = wb[args.hoja] if args.hoja else wb.worksheets[0]

    # Asegurar encabezado fijo para código de vendedor
    ws["D6"] = "COD. VENDEDOR"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True)
    accounting_fmt = ACCOUNTING_FORMAT

    ccosto_summary = {}
    ccosto_file = None
    cod_summary = {}
    cod_file = None
    lineas_summary = {}
    vendedores_summary = {}
    vendedores_file = None
    precios_summary = {}
    precios_file = None

    # --- Actualizar encabezado con fechas dinámicas -----------------------
    now = datetime.now()

    report_dt = _extract_report_datetime(path, report_date)
    report_date = report_dt.date()

    for row in ws.iter_rows(min_row=1, max_row=6, max_col=ws.max_column):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            val = cell.value
            if "MES/DIA/ANIO" in val:
                cell.value = now.strftime("%m/%d/%Y")
            elif "FECHA DEL INFORME" in val:
                cell.value = val.replace("FECHA DEL INFORME", report_dt.strftime("%m/%d/%Y"))
            elif "Procesado en" in val:
                cell.value = f"Procesado en: {now.strftime('%Y/%m/%d %H:%M:%S:%f')[:-3]}"
    # ---------------------------------------------------------------------

    if not args.skip_vendedores:
        resumen, archivo = _update_vendedores_sheet(
            wb,
            report_date=report_date,
            vendedores_file=args.vendedores_file,
            vendedores_dir=args.vendedores_dir,
            vendedores_prefix=args.vendedores_prefix,
            use_latest=use_latest,
        )
        if archivo:
            vendedores_summary = resumen
            vendedores_file = archivo

    if not args.skip_precios:
        resumen, archivo = _update_precios_sheet(
            wb,
            report_date=report_date,
            precios_file=args.precios_file,
            precios_dir=args.precios_dir,
            precios_prefix=args.precios_prefix,
            use_latest=use_latest,
        )
        if archivo:
            precios_summary = resumen
            precios_file = archivo
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

        latest, matches = _pick_excz_for_date(
            excz_dir,
            args.excz_prefix,
            report_date,
            use_latest=use_latest,
        )
        if not latest:
            available = ", ".join(meta.path.name for meta in matches) or "sin archivos"
            if use_latest:
                print(
                    "ERROR: No se encontró un EXCZ principal más reciente "
                    f"con prefijo {args.excz_prefix} en {excz_dir}. "
                    f"Disponibles: {available}"
                )
            else:
                print(
                    "ERROR: No se encontró EXCZ principal "
                    f"con prefijo {args.excz_prefix} y fecha {report_date:%Y-%m-%d} en {excz_dir}. "
                    f"Disponibles: {available}"
                )
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
        for col in ["cantidad", "ventas", "costos"]:
            if col in sub.columns:
                sub[col] = _parse_numeric_series(sub[col])
        for col in ["renta", "utili"]:
            if col in sub.columns:
                sub[col] = _parse_numeric_series(sub[col])
        if "renta" in sub.columns:
            sub = sub.sort_values(by="renta", ascending=True, na_position="last")

        # Eliminar filas de totales o cabeceras repetidas
        if "descripcion" in sub.columns:
            sub = sub[~sub["descripcion"].astype(str).str.contains("total", case=False, na=False)]
            sub = sub[sub["descripcion"].notna()]

        lineas_summary = _update_lineas_sheet(wb, sub.copy(), accounting_fmt, border)

        if args.max_rows and len(sub) > args.max_rows:
            sub = sub.iloc[:args.max_rows].copy()

        # Escribir al Excel
        for i, row in enumerate(sub.itertuples(index=False), start=start_row):
            cells = []
            if col_nit and "nit" in sub.columns:
                raw_value = getattr(row, "nit")
                value = _normalize_nit_value(raw_value)
                cell = ws.cell(i, col_nit, value)
                if isinstance(value, str):
                    cell.number_format = "@"
                    cell.alignment = Alignment(horizontal="left")
                cells.append(cell)
            if col_cliente_combo and "cliente_combo" in sub.columns:
                value = _clean_cell_value(getattr(row, "cliente_combo"))
                cells.append(ws.cell(i, col_cliente_combo, value))
            if col_desc and "descripcion" in sub.columns:
                value = _clean_cell_value(getattr(row, "descripcion"), strip=False)
                cells.append(ws.cell(i, col_desc, value))
            if col_cant and "cantidad" in sub.columns:
                cells.append(ws.cell(i, col_cant, getattr(row, "cantidad")))
            if col_ventas and "ventas" in sub.columns:
                c = ws.cell(i, col_ventas, getattr(row, "ventas"))
                c.number_format = accounting_fmt
                cells.append(c)
            if col_costos and "costos" in sub.columns:
                c = ws.cell(i, col_costos, getattr(row, "costos"))
                c.number_format = accounting_fmt
                cells.append(c)
            if col_renta and "renta" in sub.columns:
                c = ws.cell(i, col_renta, getattr(row, "renta"))
                cells.append(c)
            if col_utili and "utili" in sub.columns:
                c = ws.cell(i, col_utili, getattr(row, "utili"))
                cells.append(c)
            if col_excz:
                cells.append(ws.cell(i, col_excz, _clean_cell_value(latest.stem)))
            for c in cells:
                c.border = border

        n_rows = len(sub)

    if not args.skip_import and not args.skip_ccosto:
        ccosto_summary, ccosto_file = _update_ccosto_sheets(
            wb,
            args.exczdir,
            args.ccosto_excz_prefix,
            accounting_fmt,
            border,
            report_date,
            use_latest=use_latest,
        )

    if not args.skip_import and not args.skip_cod:
        cod_summary, cod_file = _update_cod_sheets(
            wb,
            args.exczdir,
            args.cod_excz_prefix,
            accounting_fmt,
            border,
            report_date,
            use_latest=use_latest,
        )

    # Aplicar fórmulas fijas
    vend_range = "A:B"   # VENDEDORES (NIT en A, COD_VENDEDOR en B)
    prec_range = "A:M"   # PRECIOS (DESCRIPCION en A, PRECIO en M)

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
            c.value = f"=VLOOKUP({L_desc_src}{r},PRECIOS!{prec_range},13,0)"
            c.border = border
        if L_desc and L_vent and L_cant and L_prec:
            c = ws[f"{L_desc}{r}"]
            c.value = f"=1-(({L_vent}{r}*1.19)/{L_cant}{r}/{L_prec}{r})"
            c.border = border
            c.number_format = "0.00%"

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
    total_ventas_cell = set_sum(col_ventas, accounting_fmt)
    total_costos_cell = set_sum(col_costos, accounting_fmt)

    if col_utili:
        total_util_cell = ws.cell(total_row, col_utili)
        if total_ventas_cell and total_costos_cell:
            ventas_ref = f"{get_column_letter(col_ventas)}{total_row}"
            costos_ref = f"{get_column_letter(col_costos)}{total_row}"
            total_util_cell.value = f"=IF({costos_ref}=0,0,({ventas_ref}/{costos_ref})-1)"
        else:
            total_util_cell.value = 0
        total_util_cell.number_format = "0.00%"
        total_util_cell.font = bold
        total_util_cell.border = border

    if col_renta and total_ventas_cell and total_costos_cell:
        ventas_ref = f"{get_column_letter(col_ventas)}{total_row}"
        costos_ref = f"{get_column_letter(col_costos)}{total_row}"
        rent_cell = ws.cell(total_row, col_renta)
        rent_cell.value = f"=IF({ventas_ref}=0,0,1-({costos_ref}/{ventas_ref}))"
        rent_cell.number_format = "0.00%"
        rent_cell.font = bold
        rent_cell.border = border

    for _, (_, col_idx) in hmap.items():
        cell = ws.cell(total_row, col_idx)
        cell.border = border

    wb.save(path)
    msg = f"OK. Procesadas {n_rows} filas y fórmulas aplicadas sobre: {path}"
    msg += f" | FECHA OBJETIVO: {report_date:%Y-%m-%d}"
    if ccosto_file:
        items = ", ".join(f"{k}={v}" for k, v in sorted(ccosto_summary.items())) or "sin datos"
        msg += f" | CCOSTO ({ccosto_file.name}): {items}"
    if cod_file:
        items = ", ".join(f"{k}={v}" for k, v in sorted(cod_summary.items())) or "sin datos"
        msg += f" | COD ({cod_file.name}): {items}"
    if vendedores_file:
        items = ", ".join(
            f"{k}={v}" for k, v in sorted(vendedores_summary.items())
        ) or "sin datos"
        msg += f" | VENDEDORES ({vendedores_file.name}): {items}"
    if lineas_summary:
        items = ", ".join(f"{k}={v}" for k, v in sorted(lineas_summary.items())) or "sin datos"
        msg += f" | LINEAS: {items}"
    if precios_file:
        items = ", ".join(f"{k}={v}" for k, v in sorted(precios_summary.items())) or "sin datos"
        msg += f" | PRECIOS ({precios_file.name}): {items}"
    print(msg)

if __name__ == "__main__":
    main()
