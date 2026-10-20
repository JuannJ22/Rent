import sys
import pandas as pd
from openpyxl import Workbook, load_workbook
from hojas import hoja01_loader as loader


def test_branch_lookup_never_uses_another_branch_or_last_duplicate():
    wb = Workbook()
    frame = pd.DataFrame([
        {'NitNit': 900123, 'SucNit': '001', 'PrecioNit': 1, 'VendedorNit': 24},
        {'NitNit': 900123, 'SucNit': '002', 'PrecioNit': 2, 'VendedorNit': 26},
    ])
    loader._update_terceros_sheet_from_df(wb, frame)
    lookup = loader._load_terceros_lookup(wb)
    assert loader._resolve_tercero(lookup, 900123, 1, strict_branch=True) == {'lista': 1, 'vendedor': '24'}
    assert loader._resolve_tercero(lookup, 900123, '002', strict_branch=True) == {'lista': 2, 'vendedor': '26'}
    assert loader._resolve_tercero(lookup, 900123, 3, strict_branch=True) is None
    assert loader._resolve_tercero(lookup, 900123, None, strict_branch=True) is None
    assert loader._resolve_tercero(lookup, 900123, None) is None
    wb['TERCEROS'].append([900123, 9, 51, '002'])
    ambiguous = loader._load_terceros_lookup(wb)[(900123, '2')]
    assert ambiguous['ambiguo'] and ambiguous['vendedor'] is None and ambiguous['lista'] is None


def test_sql_report_uses_branch_assignment_gross_account_sales_and_exact_invoice(tmp_path, monkeypatch):
    path = tmp_path / 'Octubre 19.xlsx'
    wb = Workbook()
    ws = wb.active
    headers = ['NIT', 'NIT - SUCURSAL - CLIENTE', 'DESCRIPCION', 'COD. VENDEDOR',
               'CANTIDAD', 'VENTAS', 'COSTOS', '% RENTA.', '% UTILI.', 'PRECIO', 'DESCUENTO', 'OBSERVACIONES']
    for col, value in enumerate(headers, 1): ws.cell(6, col, value)
    wb.create_sheet('PRECIOS')
    wb.save(path)
    specs = [
        (1, 24, 1, 119, 119, 'F202 101'),  # Exento por cuenta, nombre sin palabra "exento".
        (2, 26, 1, 200, 238, 'F202 102'),
        (1, 30, 1, 119, 119, 'F202 103'),  # Precio correcto, vendedor incorrecto.
        (3, 24, 1, 119, 119, 'F202 104'),  # Sucursal inexistente: no asumir lista 1.
        (2, 26, -1, -200, -238, 'J202 105'),
        (2, 26, 0, 200, 238, 'J202 106'),  # Ajuste sin unidades.
    ]
    data = pd.DataFrame([{
        'NIT': 900123, 'NIT - SUCURSAL - CLIENTE': f'900123 - {branch:03} - CLIENTE',
        'SUCURSAL': branch, 'DESCRIPCION': 'PRODUCTO', 'COD. VENDEDOR': vendor,
        'CANTIDAD': qty, 'Ventas': net, 'Costos': net/2,
        'VENTAS_CON_IVA': gross, 'DOCUMENTO': doc, '% RENTA.': 20+i, '% UTILI.': 50,
    } for i, (branch, vendor, qty, net, gross, doc) in enumerate(specs)])
    clients = pd.DataFrame([
        {'NitNit': 900123, 'SucNit': 1, 'PrecioNit': 1, 'VendedorNit': 24},
        {'NitNit': 900123, 'SucNit': 2, 'PrecioNit': 2, 'VendedorNit': 26},
    ])
    prices = pd.DataFrame([{'NOMBRE':'PRODUCTO', **{f'LISTA_PRECIO{i}':119*i for i in range(1,13)}}])
    empty = pd.DataFrame()
    monkeypatch.setenv('SQL_SERVER', 'server')
    monkeypatch.setenv('SQL_DATABASE', 'db')
    monkeypatch.setenv('SQL_USER', 'reader')
    monkeypatch.setattr(loader, '_load_sql_report_data', lambda *a: (data, empty, {}, {}, prices, clients, empty))
    monkeypatch.setattr(sys, 'argv', ['loader', '--sql', '--require-sql-data', '--excel', str(path), '--fecha', '2026-10-19'])
    loader.main()
    wb = load_workbook(path)
    ws = wb.active
    assert not ws['L7'].value  # No falsa diferencia por aplicar 19% a una cuenta exenta.
    assert not ws['L8'].value  # Sucursal 2 usa su lista 2 y su vendedor 26.
    assert 'Se facturó con código 30.' in ws['L9'].value
    assert 'Factura F202 103.' in ws['L9'].value
    assert 'sucursal 3' in ws['L10'].value and 'no se valido' in ws['L10'].value.lower()
    assert not ws['L11'].value  # Devolucion con signos coherentes.
    assert 'Ajuste sin cantidad' in ws['L12'].value
    assert ws['AB7'].value == 119 and ws['AB8'].value == 238 and ws['AB11'].value == -238
    assert ws.column_dimensions['AB'].hidden
    assert 'ABS(AB7)' in ws['K7'].value and 'ABS(E7)' in ws['K7'].value
    assert '13,0)' in ws['J7'].value  # Lista 12 permanece referencia de descuento.
    assert len(list(wb['TERCEROS'].values)) == 2  # Compactado mantiene ambas sucursales.
    wb.close()


def test_legacy_manual_third_parties_remain_usable_without_branch_column():
    wb = Workbook()
    wb.active.title = 'TERCEROS'
    wb.active.append([900123, 2, 26])
    lookup = loader._load_terceros_lookup(wb)
    assert loader._resolve_tercero(lookup, 900123, '001') == {'lista': 2, 'vendedor': '26'}
    assert loader._resolve_tercero(lookup, 900123, '001', strict_branch=True) is None


def test_numeric_vendor_codes_with_padding_do_not_create_false_mismatch():
    assert loader._normalize_vendor_code('024') == '24'
    assert loader._normalize_vendor_code('24.0') == '24'
    assert loader._normalize_vendor_code(24.0) == '24'
    assert loader._normalize_vendor_code('A1') == 'A1'
