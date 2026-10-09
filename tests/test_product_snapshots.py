from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import pandas as pd
import pytest
from rentabilidad.infra import product_snapshots as snapshots
from rentabilidad.infra.sql_server import SqlServerConfig


def capture(folder, day, monkeypatch, data=None):
    if data is None:
        data = pd.DataFrame({'DescripcionInv': ['CU' + chr(169) + 'ETE', '=literal'],
                             'Precio1': [Decimal('12345.67'), Decimal('0.00')], 'Precio2': ['50', None]})
    monkeypatch.setattr(snapshots, 'fetch_dataframe', lambda *_: data)
    return snapshots.capture_snapshot(SqlServerConfig('server', 'SiigoRent'), folder,
        now=datetime(day.year, day.month, day.day, 21, tzinfo=snapshots.COLOMBIA))


def test_snapshot_keeps_all_columns_and_sql_strings_and_is_immutable(tmp_path, monkeypatch):
    day = date(2026, 10, 9)
    path = capture(tmp_path, day, monkeypatch)
    data = snapshots.read_snapshot(path, day)
    assert list(data.columns) == ['DescripcionInv', 'Precio1', 'Precio2']
    assert data['DescripcionInv'].tolist() == ['CU' + chr(169) + 'ETE', '=literal']
    assert data.loc[0, 'Precio1'] == 12345.67
    assert data.loc[0, 'Precio2'] == '50'
    original = path.read_bytes()
    monkeypatch.setattr(snapshots, 'fetch_dataframe', lambda *_: pytest.fail('No reemplazar copia diaria'))
    assert snapshots.capture_snapshot(SqlServerConfig('server', 'db'), tmp_path,
        now=datetime(2026, 10, 9, 22, tzinfo=snapshots.COLOMBIA)) == path
    assert path.read_bytes() == original


def test_snapshot_retention_is_one_calendar_month_and_preserves_other_files(tmp_path, monkeypatch):
    old = capture(tmp_path, date(2026, 9, 8), monkeypatch)
    boundary = capture(tmp_path, date(2026, 9, 9), monkeypatch)
    manual = tmp_path / 'productos0901.xlsx'
    manual.write_bytes(b'manual')
    other = tmp_path / 'productos-2026-09-01.xlsx'
    other.write_bytes(b'no es captura SQL')
    capture(tmp_path, date(2026, 10, 9), monkeypatch)
    assert not old.exists() and boundary.exists() and manual.exists() and other.exists()
    assert snapshots.month_cutoff(date(2026, 3, 31)) == date(2026, 2, 28)
    assert snapshots.month_cutoff(date(2026, 1, 31)) == date(2025, 12, 31)


def test_failed_capture_keeps_previous_history(tmp_path, monkeypatch):
    old = capture(tmp_path, date(2026, 9, 8), monkeypatch)
    monkeypatch.setattr(snapshots, 'fetch_dataframe', lambda *_: pd.DataFrame())
    with pytest.raises(RuntimeError, match='vacia'):
        snapshots.capture_snapshot(SqlServerConfig('server', 'db'), tmp_path,
            now=datetime(2026, 10, 9, 21, tzinfo=snapshots.COLOMBIA))
    assert old.exists()
    assert not snapshots.snapshot_path(tmp_path, date(2026, 10, 9)).exists()


def test_snapshot_cannot_be_used_for_another_day(tmp_path, monkeypatch):
    path = capture(tmp_path, date(2026, 10, 9), monkeypatch)
    with pytest.raises(RuntimeError, match='fecha'):
        snapshots.read_snapshot(path, date(2026, 10, 8))


def test_capture_waits_until_nightly_close(tmp_path, monkeypatch):
    monkeypatch.setattr(snapshots, 'fetch_dataframe', lambda *_: pytest.fail('No leer antes del cierre'))
    with pytest.raises(RuntimeError, match='20:45'):
        snapshots.capture_snapshot(SqlServerConfig('server', 'db'), tmp_path,
            now=datetime(2026, 10, 9, 20, 44, tzinfo=snapshots.COLOMBIA))


def test_sql_loader_uses_requested_date_snapshot_not_current_prices(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from hojas import hoja01_loader as loader
    day = date(2026, 10, 9)
    path = capture(tmp_path, day, monkeypatch)
    monkeypatch.setenv('SQL_PRODUCT_SNAPSHOT', str(path))
    @contextmanager
    def reader(_):
        def read(query, params=None):
            assert 'vw_productos_activos' not in query
            return pd.DataFrame()
        yield read
    monkeypatch.setattr(loader, 'dataframe_reader', reader)
    _, _, _, _, prices, _, _ = loader._load_sql_report_data(SqlServerConfig('server', 'db'), day, {})
    assert prices.loc[0, 'DescripcionInv'] == 'CUÑETE'
    assert prices.loc[0, 'Precio1'] == 12345.67
