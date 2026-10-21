from servicios.fix_repeated_quantities import patch_quantity_max, as_alter_view


def test_only_quantity_max_changes_not_comments_literals_or_money():
    source = """-- MAX(CantidadMov) se comenta y no se cambia.
CREATE VIEW dbo.vw_rentabilidad_base AS
SELECT MAX([m].[CantidadMov]) AS Cantidad,
 SUM(m.ValorMov) AS Ventas, 'MAX(CantidadMov)' AS literal,
 MAX(CASE WHEN m.CuentasMov LIKE '41%' THEN m.CuentasMov END) AS CuentaVenta
FROM dbo.MovimientoPorComprobante_Actual m;
/* comentario /* MAX(CantidadMov) */ anidado */"""
    fixed, count = patch_quantity_max(source)
    assert count == 1
    assert "SUM(CASE WHEN [m].[CuentasMov] LIKE '41%' THEN [m].[CantidadMov] ELSE 0 END) AS Cantidad" in fixed
    assert 'SUM(m.ValorMov) AS Ventas' in fixed
    assert "'MAX(CantidadMov)' AS literal" in fixed
    assert '-- MAX(CantidadMov) se comenta' in fixed
    assert '/* comentario /* MAX(CantidadMov) */ anidado */' in fixed
    assert 'MAX(CASE WHEN' in fixed
    assert patch_quantity_max(fixed) == (fixed, 0)
    ddl = as_alter_view(fixed)
    assert '-- MAX(CantidadMov) se comenta' in ddl
    assert 'CREATE OR ALTER VIEW dbo.vw_rentabilidad_base' in ddl


def test_credit_note_branches_both_use_sum_and_non_source_views_are_untouched():
    source = """CREATE VIEW dbo.vw_movimientos_vendedores_informe AS
SELECT CASE WHEN m.TipMov='J' THEN -MAX(m.CantidadMov) ELSE MAX(m.CantidadMov) END AS Cantidad
FROM dbo.MovimientoPorComprobante_Actual m GROUP BY m.TipMov;"""
    fixed, count = patch_quantity_max(source)
    assert count == 2
    assert "THEN -SUM(CASE WHEN m.[CuentasMov] LIKE '41%'" in fixed
    assert "ELSE SUM(CASE WHEN m.[CuentasMov] LIKE '41%'" in fixed
    other = 'CREATE VIEW dbo.otro AS SELECT MAX(CantidadMov) FROM dbo.OtraTabla;'
    assert patch_quantity_max(other) == (other, 0)


def test_sql_migration_rolls_back_if_financial_control_changes(tmp_path, monkeypatch):
    import pytest
    from servicios import fix_repeated_quantities as fix
    names = ['vw_rentabilidad_cliente_informe','vw_rentabilidad_base','vw_rentabilidad_base_linea',
             'vw_rentabilidad_lineas','vw_movimientos_vendedores_informe']
    definitions = [(i,'dbo',name,f'CREATE VIEW dbo.{name} AS SELECT MAX(CantidadMov) AS Cantidad FROM dbo.MovimientoPorComprobante_Actual;',True,True)
                   for i,name in enumerate(names,1)]
    class Cursor:
        def __init__(self): self.query = ''; self.statements=[]
        def execute(self, query, *args): self.query=query; self.statements.append(query); return self
        def fetchall(self): return definitions if 'FROM sys.views' in self.query else []
        def close(self): pass
    class Connection:
        def __init__(self): self.c=Cursor(); self.committed=False; self.rolled_back=False
        def cursor(self): return self.c
        def commit(self): self.committed=True
        def rollback(self): self.rolled_back=True
    before = {doc:{'cantidad':old,'cantidad_ventas':new,'ventas':100,'costos':60}
              for doc,old,new in [('F201 25052',10,15),('F241 6391',3,6)]}
    after = {doc:dict(value,cantidad=value['cantidad_ventas']) for doc,value in before.items()}
    after['F201 25052']['costos'] = 61
    cases = iter([before,after])
    monkeypatch.setattr(fix,'sample_cases',lambda cursor:next(cases))
    conn=Connection()
    with pytest.raises(RuntimeError,match='Cambiaron ventas o costos'):
        fix.apply_sql(conn,tmp_path)
    assert conn.rolled_back and not conn.committed
    assert (tmp_path/'restaurar-vistas.sql').exists()
    assert all('dbo.vw_' in statement for statement in conn.c.statements if statement.startswith('CREATE OR ALTER'))
