import csv
from datetime import date
import os
from pathlib import Path
import re

import pandas as pd
from openpyxl import Workbook
from rentabilidad.infra import sql_server


def test_installer_audit_retains_real_negative_discounts_and_flags_repeated_lines(tmp_path, monkeypatch):
    report = tmp_path / 'InformesDiarios' / 'Octubre 19.xlsx'
    report.parent.mkdir()
    wb = Workbook()
    ws = wb.create_sheet('PRECIOS')
    ws.append(['PRODUCTO'] + [100]*12)
    wb.save(report)
    specs = [
        ('F202 1', 1, 100, 1, 1),  # Cero descuento: no se incluye en pendientes.
        ('F202 2', 1, 180, 1, 1),  # -80% real contra lista 12=100.
        ('J202 3', -1, -180, 1, -1),  # Devolucion coherente: mismo -80%.
        ('J202 4', 0, 50, 1, 0),  # Sin division por cero ni descuento ficticio.
        ('F202 5', 1, 180, 2, 2),  # Cantidad MAX frente a varias filas: requiere revisar.
    ]
    data = pd.DataFrame([{
        'NIT':900123, 'SUCURSAL':1, 'DOCUMENTO':doc, 'DESCRIPCION':'PRODUCTO',
        'CANTIDAD':qty, 'VENTAS':gross/1.19, 'VENTAS_CON_IVA':gross,
        'FILAS_VENTA':count, 'CANTIDAD_VENTA_SUM':summed,
    } for doc, qty, gross, count, summed in specs])
    monkeypatch.setattr(sql_server, 'fetch_dataframe', lambda *a: data)
    for key in ['SQL_SERVER', 'SQL_USER', 'SQL_PASSWORD', 'SQL_TRUSTED']:
        monkeypatch.setenv(key, '0' if key == 'SQL_TRUSTED' else 'test')
    installer = Path(__file__).resolve().parents[1] / 'servicios/SiigoBI.EtlService/Install-BranchValidation.ps1'
    blocks = re.findall(r"@'\n(.*?)\n'@", installer.read_text(), re.S)
    audit = '# Audita' + blocks[-1].split('# Audita', 1)[1]
    exec(compile(audit, str(installer), 'exec'), {'review':tmp_path, 'day':date(2026,10,19), 'os':os, 'csv':csv})
    with (tmp_path / 'revision-descuentos.csv').open(encoding='utf-8-sig') as handle:
        rows = {r['Documento']:r for r in csv.DictReader(handle, delimiter=';')}
    assert set(rows) == {'F202 2','J202 3','J202 4','F202 5'}
    assert float(rows['F202 2']['Descuento %']) == -80
    assert float(rows['J202 3']['Descuento %']) == -80
    assert rows['J202 4']['Descuento %'] == ''
    assert 'varias filas de venta' in rows['F202 5']['Revision']
