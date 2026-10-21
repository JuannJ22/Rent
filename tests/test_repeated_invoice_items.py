from pathlib import Path
import re
import sqlite3
import pytest


def test_production_quantity_aggregate_sums_repeated_items_not_cost_postings():
    # Ejecuta la agregacion SQL real de la vista con los dos ejemplos de factura.
    sql = (Path(__file__).resolve().parents[1] / 'docs/sql/rentabilidad_cliente_sucursales.sql').read_text()
    from servicios.fix_repeated_quantities import patch_quantity_max
    old = sql.replace("SUM(CASE WHEN CuentasMov LIKE '41%' THEN CantidadMov ELSE 0 END)", "MAX(CantidadMov)", 1)
    sql, changes = patch_quantity_max(old)
    assert changes == 1
    aggregate = re.search(r', agg AS \((.*?)\n\)\nSELECT', sql, re.S).group(1)
    conn = sqlite3.connect(':memory:')
    conn.row_factory = sqlite3.Row
    conn.execute('''CREATE TABLE mov (
        TipMov TEXT,ComMov INTEGER,NroMov INTEGER,FechaDctoMov INTEGER,NitMov INTEGER,SucMov INTEGER,
        VendedorMov INTEGER,DescrMov TEXT,LinMov INTEGER,GrpMov INTEGER,ProMov INTEGER,
        CuentasMov TEXT,CantidadMov REAL,ValorMov REAL)''')
    def item(tip, com, number, product, quantity, sales, cost, branch=0):
        identity = (tip, com, number, 20261007, 901800944, branch, 31, product, 1, 1, 1)
        conn.execute('INSERT INTO mov VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', identity + ('4135010100', quantity, sales))
        conn.execute('INSERT INTO mov VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', identity + ('1435010100', quantity, cost))
    item('F',201,25052,'GRANIACRYL',10,1601546.22,1425370)
    item('F',201,25052,'GRANIACRYL',5,800773.11,712685)
    item('F',241,6391,'BROCHA',3,28247.65,23471.82)
    item('F',241,6391,'BROCHA',3,28247.65,23471.82)
    item('J',241,9999,'BROCHA',3,28247.65,23471.82)
    item('J',241,9999,'BROCHA',3,28247.65,23471.82)
    item('F',241,6391,'BROCHA',2,18831.77,15647.88,branch=1)
    result = {(r['TipMov'],r['ComMov'],r['NroMov'],r['SucMov']):r for r in conn.execute(aggregate)}
    gran = result[('F',201,25052,0)]
    brocha = result[('F',241,6391,0)]
    assert gran['Cantidad'] == 15  # MAX=10 era incorrecto; sumar venta+costo daria 30.
    assert brocha['Cantidad'] == 6  # Dos items iguales de 3 cada uno; no MAX ni DISTINCT.
    assert gran['Ventas'] == pytest.approx(2402319.33)
    assert gran['Costos'] == pytest.approx(2138055)
    assert brocha['Ventas'] == pytest.approx(56495.30)
    assert result[('J',241,9999,0)]['Cantidad'] == -6
    assert result[('F',241,6391,1)]['Cantidad'] == 2  # Otra sucursal permanece separada.
    assert 100*(1-gran['VentasConIVA']/gran['Cantidad']/238230) == pytest.approx(20, abs=.0001)
    assert 100*(1-brocha['VentasConIVA']/brocha['Cantidad']/13105.2) == pytest.approx(14.5, abs=.001)
    conn.close()
