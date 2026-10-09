"""Copias diarias completas de la vista de productos; no reconstruye precios pasados."""
from __future__ import annotations
from calendar import monthrange
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
import os
import re
import tempfile
import pandas as pd
from openpyxl import Workbook, load_workbook
from rentabilidad.infra.sql_server import SqlServerConfig, fetch_dataframe

COLOMBIA = timezone(timedelta(hours=-5))
VIEW = '[SiigoCat].[dbo].[vw_productos_activos]'
PATTERN = re.compile(r'^productos-(\d{4}-\d{2}-\d{2})\.xlsx$')


def snapshot_path(folder: Path, day: date) -> Path:
    return Path(folder) / f'productos-{day.isoformat()}.xlsx'


def month_cutoff(today: date) -> date:
    year, month = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)
    return date(year, month, min(today.day, monthrange(year, month)[1]))


def prune_snapshots(folder: Path, today: date) -> list[Path]:
    removed = []
    cutoff = month_cutoff(today)
    for path in Path(folder).glob('productos-*.xlsx'):
        match = PATTERN.fullmatch(path.name)
        if not match: continue
        try: day = date.fromisoformat(match.group(1))
        except ValueError: continue
        if day < cutoff:
            try: read_snapshot(path, day)
            except Exception: continue
            path.unlink()
            removed.append(path)
    return removed


def read_snapshot(path: Path, day: date) -> pd.DataFrame:
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if 'CAPTURA' not in wb.sheetnames or 'PRODUCTOS' not in wb.sheetnames:
            raise RuntimeError('El archivo no es una captura SQL diaria valida.')
        metadata = dict(wb['CAPTURA'].iter_rows(max_col=2, values_only=True))
        if metadata.get('Fecha') != day.isoformat() or metadata.get('Vista') != VIEW:
            raise RuntimeError('La captura no corresponde a la fecha o vista requerida.')
        rows = wb['PRODUCTOS'].iter_rows(values_only=True)
        headers = next(rows)
        data = list(rows)
        if len(data) != metadata.get('Filas') or not data:
            raise RuntimeError('La captura esta vacia o incompleta.')
        return pd.DataFrame(data, columns=headers)
    finally:
        wb.close()


def capture_snapshot(config: SqlServerConfig, folder: Path, *, now: datetime | None = None) -> Path:
    now = now or datetime.now(COLOMBIA)
    day = now.astimezone(COLOMBIA).date()
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    destination = snapshot_path(folder, day)
    # Primera copia del dia inmutable: nunca se sustituye por precios posteriores.
    if destination.exists():
        read_snapshot(destination, day)
        prune_snapshots(folder, day)
        return destination
    if (now.astimezone(COLOMBIA).hour, now.astimezone(COLOMBIA).minute) < (20, 45):
        raise RuntimeError('La copia de cierre se captura desde las 20:45 de Colombia, despues de la carga nocturna.')
    data = fetch_dataframe(config, f'SELECT * FROM {VIEW}')
    if data.empty or data.columns.duplicated().any():
        raise RuntimeError('La vista de productos esta vacia o sus columnas no son unicas.')
    book = Workbook()
    sheet = book.active
    sheet.title = 'PRODUCTOS'
    sheet.append(list(data.columns))
    for values in data.itertuples(index=False, name=None):
        sheet.append([None if pd.isna(value) else value for value in values])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str): cell.data_type = 's'
    metadata = book.create_sheet('CAPTURA')
    for row in [('Fecha', day.isoformat()), ('CapturadoUTC', now.astimezone(timezone.utc).isoformat()),
                ('Vista', VIEW), ('Filas', len(data))]:
        metadata.append(row)
    metadata.sheet_state = 'hidden'
    fd, name = tempfile.mkstemp(prefix='.productos-', suffix='.xlsx', dir=folder)
    os.close(fd)
    temporary = Path(name)
    try:
        book.save(temporary)
        read_snapshot(temporary, day)
        os.replace(temporary, destination)
    finally:
        book.close()
        temporary.unlink(missing_ok=True)
    prune_snapshots(folder, day)
    return destination
