"""Migracion de cantidades repetidas: modifica solo MAX(CantidadMov) en vistas del informe."""
from __future__ import annotations
import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from collections import deque
from datetime import date

QUANTITY_MAX = re.compile(
    r'\bMAX\s*\(\s*(?P<prefix>(?:(?:\[[^\]]+\]|[A-Za-z_]\w*)\s*\.\s*)?)'
    r'(?:\[CantidadMov\]|CantidadMov)\s*\)', re.I)


def mask_comments_and_literals(sql: str) -> str:
    """Conserva indices para no modificar texto dentro de comentarios o literales."""
    result = list(sql)
    i = 0
    while i < len(sql):
        start = i
        if sql.startswith('--', i):
            end = sql.find('\n', i)
            i = len(sql) if end < 0 else end
        elif sql.startswith('/*', i):
            i += 2
            depth = 1
            while i < len(sql) and depth:
                if sql.startswith('/*', i): depth += 1; i += 2
                elif sql.startswith('*/', i): depth -= 1; i += 2
                else: i += 1
        elif sql[i] == "'":
            i += 1
            while i < len(sql):
                if sql.startswith("''", i): i += 2
                elif sql[i] == "'": i += 1; break
                else: i += 1
        else:
            i += 1
            continue
        for pos in range(start, i):
            if result[pos] not in '\r\n': result[pos] = ' '
    return ''.join(result)


def patch_quantity_max(sql: str) -> tuple[str, int]:
    masked = mask_comments_and_literals(sql)
    if not re.search(r'\bMovimientoPorComprobante_Actual\b', masked, re.I):
        return sql, 0
    matches = list(QUANTITY_MAX.finditer(masked))
    result = sql
    for match in reversed(matches):
        prefix = sql[match.start('prefix'):match.end('prefix')]
        replacement = (f"SUM(CASE WHEN {prefix}[CuentasMov] LIKE '41%' "
                       f"THEN {prefix}[CantidadMov] ELSE 0 END)")
        result = result[:match.start()] + replacement + result[match.end():]
    return result, len(matches)


def as_alter_view(sql: str) -> str:
    masked = mask_comments_and_literals(sql)
    match = re.search(r'\b(?:CREATE(?:\s+OR\s+ALTER)?|ALTER)\s+VIEW\b', masked, re.I)
    if not match or masked[:match.start()].strip():
        raise RuntimeError('No se encontro una definicion CREATE/ALTER VIEW valida.')
    return sql[:match.start()] + 'CREATE OR ALTER VIEW' + sql[match.end():]


def option(obj, name, default=None):
    return next((value for key, value in obj.items() if key.lower() == name.lower()), default)


def load_target():
    settings = json.loads(Path(r'C:\SiigoBI_ETL\appsettings.json').read_text(encoding='utf-8-sig'))
    target = option(option(settings, 'SqlTargets', {}), 'RentReport')
    if not target: raise RuntimeError('Falta SqlTargets.RentReport.')
    return target


def connect(target):
    import pyodbc
    from rentabilidad.infra.sql_server import SqlServerConfig
    config = SqlServerConfig(server=option(target,'Server'), database='SiigoRent',
        user=option(target,'User'), password=option(target,'Password'),
        trusted_connection=bool(option(target,'TrustedConnection',False)),
        driver='ODBC Driver 18 for SQL Server', encrypt=True, trust_server_certificate=False)
    conn = pyodbc.connect(config.connection_string(), timeout=30)
    conn.timeout = 300
    return conn


def sample_cases(cursor):
    rows = cursor.execute("""
        SELECT DOCUMENTO,DESCRIPCION,CANTIDAD,CANTIDAD_VENTA_SUM,Ventas,Costos
        FROM dbo.vw_rentabilidad_cliente_informe
        WHERE FECHA=CONVERT(date,'20261007',112)
          AND DOCUMENTO IN ('F201 25052','F241 6391')
          AND (DESCRIPCION LIKE '%GRANI%' OR DESCRIPCION LIKE '%BROCHA%')
    """).fetchall()
    if len(rows) != 2 or {row[0] for row in rows} != {'F201 25052','F241 6391'}:
        raise RuntimeError('No se pudieron identificar los dos casos de control; no se aplican cambios a ciegas.')
    return {row[0]: {'descripcion':row[1], 'cantidad':row[2], 'cantidad_ventas':row[3],
                     'ventas':row[4], 'costos':row[5]} for row in rows}


def apply_sql(conn, backup: Path):
    cursor = conn.cursor()
    backup.mkdir(parents=True, exist_ok=True)
    try:
        cursor.execute('SET XACT_ABORT ON')
        before = sample_cases(cursor)
        expected = {'F201 25052':15, 'F241 6391':6}
        for doc, qty in expected.items():
            if before[doc]['cantidad_ventas'] != qty:
                raise RuntimeError(f'Cantidad de ventas inesperada en {doc}; se cancela sin modificar vistas.')
        rows = cursor.execute("""
            SELECT v.object_id,s.name,v.name,m.definition,m.uses_ansi_nulls,m.uses_quoted_identifier
            FROM sys.views v JOIN sys.schemas s ON s.schema_id=v.schema_id
            JOIN sys.sql_modules m ON m.object_id=v.object_id
            WHERE s.name='dbo' AND
              (v.name LIKE 'vw[_]rentabilidad[_]%' OR v.name='vw_movimientos_vendedores_informe')
        """).fetchall()
        views = {row[0]:dict(id=row[0],schema=row[1],name=row[2],definition=row[3],ansi=row[4],quoted=row[5]) for row in rows}
        names = {v['name'] for v in views.values() if v['definition']}
        required = {'vw_rentabilidad_cliente_informe','vw_rentabilidad_base','vw_rentabilidad_base_linea',
                    'vw_rentabilidad_lineas','vw_movimientos_vendedores_informe'}
        if not required.issubset(names): raise RuntimeError('Faltan definiciones de vistas requeridas; no se modifico nada.')
        updates = {}
        for key, view in views.items():
            if not view['definition']: continue
            changed, count = patch_quantity_max(view['definition'])
            if count: updates[key] = (as_alter_view(changed), count)
        (backup / 'vistas-originales.json').write_text(json.dumps(list(views.values()),ensure_ascii=False,indent=2),encoding='utf-8')
        restore = []
        for key in updates:
            v = views[key]
            restore.append(f"SET ANSI_NULLS {'ON' if v['ansi'] else 'OFF'};\nGO\nSET QUOTED_IDENTIFIER {'ON' if v['quoted'] else 'OFF'};\nGO\n" + as_alter_view(v['definition']) + '\nGO\n')
        (backup / 'restaurar-vistas.sql').write_text('\n'.join(restore),encoding='utf-8')
        edges = cursor.execute('SELECT referencing_id,referenced_id FROM sys.sql_expression_dependencies WHERE referenced_id IS NOT NULL').fetchall()
        dependents = {}
        for referencing, referenced in edges:
            if referencing in views: dependents.setdefault(referenced,set()).add(referencing)
        distances = {key:0 for key in updates}
        queue = deque(updates)
        while queue:
            parent = queue.popleft()
            for child in dependents.get(parent, ()):
                if child not in distances:
                    distances[child] = distances[parent]+1
                    queue.append(child)
        for key, (definition, count) in updates.items():
            v = views[key]
            cursor.execute(f"SET ANSI_NULLS {'ON' if v['ansi'] else 'OFF'}")
            cursor.execute(f"SET QUOTED_IDENTIFIER {'ON' if v['quoted'] else 'OFF'}")
            cursor.execute(definition)
            print(f"VISTA AJUSTADA: dbo.{v['name']} ({count} expresiones)")
        for key in sorted(distances, key=distances.get):
            cursor.execute('EXEC sys.sp_refreshview @viewname=?', 'dbo.'+views[key]['name'])
        after = sample_cases(cursor)
        for doc, qty in expected.items():
            if after[doc]['cantidad'] != qty: raise RuntimeError(f'No se corrigio la cantidad de {doc}; se revierte SQL.')
            if (after[doc]['ventas'], after[doc]['costos']) != (before[doc]['ventas'], before[doc]['costos']):
                raise RuntimeError(f'Cambiaron ventas o costos de {doc}; se revierte SQL.')
        (backup / 'controles-cantidades.json').write_text(json.dumps({'antes':before,'despues':after},default=str,ensure_ascii=False,indent=2),encoding='utf-8')
        conn.commit()
        for doc in expected:
            print(f"CONTROL: {doc}: {before[doc]['cantidad']} -> {after[doc]['cantidad']}; ventas y costos conservados.")
        if not updates: print('Las vistas ya tenian las cantidades corregidas.')
    except BaseException:
        conn.rollback()
        print('SQL: transaccion revertida por error; consulta el detalle anterior.')
        raise
    finally:
        cursor.close()


def prepare_review(target, project: Path, review: Path, day: date, products: Path | None):
    source = products or Path(r'C:\Rentabilidad\Productos') / f'productos-{day.isoformat()}.xlsx'
    if not source.exists(): raise RuntimeError(f'Falta el listado de productos para la revision: {source}')
    from rentabilidad.infra.product_snapshots import read_legacy_prices
    read_legacy_prices(source)
    destination = review / 'Productos' / source.name
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    env = dict(os.environ, SQL_SERVER=option(target,'Server'), SQL_DATABASE='SiigoRent',
        SQL_USER=option(target,'User','') or '', SQL_PASSWORD=option(target,'Password','') or '',
        SQL_TRUSTED='1' if option(target,'TrustedConnection',False) else '0', SQL_ENCRYPT='1', SQL_TRUST_CERT='0',
        SQL_TIMEOUT='300', RENT_DIR=str(review), PYTHONUTF8='1', PYTHONIOENCODING='utf-8',
        SQL_LEGACY_PRODUCT_FILE=str(destination))
    env.pop('SQL_PRODUCT_SNAPSHOT',None)
    result = subprocess.run([sys.executable,str(project / 'servicios' / 'etl_rentabilidad.py'),
        '--fecha',day.isoformat(),'--base-dir',str(review),'--template',r'C:\Rentabilidad\PLANTILLA.xlsx'],cwd=project,env=env)
    if result.returncode: raise RuntimeError('SQL corregido, pero fallo la copia de revision. Revisa el resultado anterior.')
    print(f'REVISION PREPARADA: {review}')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--apply-sql',action='store_true')
    parser.add_argument('--review-dir',type=Path)
    parser.add_argument('--backup',type=Path)
    parser.add_argument('--project-dir',type=Path,default=Path(r'C:\Rentabilidad\Rent'))
    parser.add_argument('--fecha',default='2026-10-07')
    parser.add_argument('--productos',type=Path)
    args=parser.parse_args()
    sys.path.insert(0,str(args.project_dir))
    target=load_target()
    if args.apply_sql:
        if not args.backup: parser.error('Falta --backup')
        conn=connect(target)
        try: apply_sql(conn,args.backup)
        finally: conn.close()
    else:
        if not args.review_dir: parser.error('Falta --review-dir')
        prepare_review(target,args.project_dir,args.review_dir,date.fromisoformat(args.fecha),args.productos)


if __name__ == '__main__':
    main()
