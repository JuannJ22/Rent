-- Fuente nueva; no modifica las tablas ni las vistas existentes.
USE [SiigoRent];
GO
CREATE OR ALTER VIEW dbo.vw_movimientos_vendedores_informe
AS
SELECT
    m.NitMov,
    m.VendedorMov,
    m.TipMov,
    m.ComMov,
    m.NroMov,
    m.DescrMov,
    CASE WHEN m.TipMov = 'J' THEN -MAX(m.CantidadMov)
         ELSE MAX(m.CantidadMov) END AS CantidadMov,
    m.FechaDctoMov
FROM dbo.MovimientoPorComprobante_Actual AS m
WHERE m.TipMov IN ('F','J')
  AND m.CuentasMov LIKE '41%'
GROUP BY m.NitMov, m.SucMov, m.VendedorMov, m.TipMov, m.ComMov, m.NroMov,
         m.FechaDctoMov, m.LinMov, m.GrpMov, m.ProMov, m.DescrMov;
GO
SELECT COUNT_BIG(*) AS FilasDelDia,
       COUNT(DISTINCT NitMov) AS ClientesDelDia
FROM dbo.vw_movimientos_vendedores_informe
WHERE FechaDctoMov = '20261006';
