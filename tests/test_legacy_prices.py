from openpyxl import Workbook
import pytest
from rentabilidad.infra.product_snapshots import read_legacy_prices


def test_manual_catalog_preserves_first_product_and_price_positions(tmp_path):
    path = tmp_path / 'productos0923.xlsx'
    wb = Workbook()
    wb.active.append(['CUÑETE', 100, 200])
    wb.save(path)
    frame = read_legacy_prices(path)
    assert frame.iloc[0]['DESCRIPCION'] == 'CUÑETE'
    assert frame.iloc[0]['PRECIO2'] == 200
    assert len(frame.columns) == 13


def test_empty_manual_catalog_is_rejected(tmp_path):
    path = tmp_path / 'productos0923.xlsx'
    Workbook().save(path)
    with pytest.raises(ValueError, match='vacio'):
        read_legacy_prices(path)
