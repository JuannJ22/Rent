-- Diagnóstico: solo SELECT. No altera vistas ni tablas.
USE [SiigoRent];
DECLARE @Fecha date = '20261006';
DECLARE @FechaSiigo varchar(8) = CONVERT(char(8), @Fecha, 112);

;WITH documentos AS (
    SELECT m.TipMov, m.ComMov, m.NroMov, m.FechaDctoMov,
           m.NitMov, m.VendedorMov, m.LinMov, m.GrpMov, m.ProMov, m.DescrMov,
           MAX(m.CantidadMov) AS Cantidad,
           SUM(CASE WHEN m.CuentasMov LIKE '41%' THEN m.ValorMov ELSE 0 END) AS Ventas,
           SUM(CASE WHEN m.CuentasMov LIKE '1435%' THEN m.ValorMov ELSE 0 END) AS Costos
    FROM dbo.MovimientoPorComprobante_Actual AS m
    WHERE m.FechaDctoMov = @FechaSiigo
      AND m.TipMov IN ('F', 'J')
      AND (m.CuentasMov LIKE '41%' OR m.CuentasMov LIKE '1435%')
    GROUP BY m.TipMov, m.ComMov, m.NroMov, m.FechaDctoMov,
             m.NitMov, m.VendedorMov, m.LinMov, m.GrpMov, m.ProMov, m.DescrMov
), firmados AS (
    SELECT d.*,
           CASE WHEN TipMov = 'J' THEN -Cantidad ELSE Cantidad END AS CantidadNeta,
           CASE WHEN TipMov = 'J' THEN -Ventas ELSE Ventas END AS VentasNetas,
           CASE WHEN TipMov = 'J' THEN -Costos ELSE Costos END AS CostosNetos
    FROM documentos AS d
), aptos AS (
    SELECT * FROM firmados
    WHERE Ventas <> 0 AND Costos <> 0
      AND CAST(100.0 * (1 - ABS(Costos) / NULLIF(ABS(Ventas), 0)) AS decimal(18,2))
          NOT BETWEEN 99.9 AND 100.1
), clientes AS (
    SELECT NitMov, LinMov, GrpMov, ProMov, DescrMov,
           SUM(CantidadNeta) AS Cantidad,
           SUM(VentasNetas) AS Ventas,
           SUM(CostosNetos) AS Costos
    FROM aptos
    GROUP BY NitMov, LinMov, GrpMov, ProMov, DescrMov
)
SELECT '1. Zona 7 actual' AS Escenario,
       SUM(CANTIDAD) AS Cantidad, SUM(VENTAS) AS Ventas, SUM(COSTOS) AS Costos
FROM dbo.vw_rentabilidad_zona7 WHERE FECHA = @Fecha
UNION ALL
SELECT '2. Zona 7 con signo J', SUM(CantidadNeta), SUM(VentasNetas), SUM(CostosNetos)
FROM firmados AS f
WHERE Ventas <> 0 AND Costos <> 0
  AND EXISTS (
      SELECT 1 FROM SiigoCat.dbo.TABLA_DESCRIPCION_VENDEDORES AS v
      WHERE v.VenVen = f.VendedorMov AND v.ZonaVen = 7
  )
UNION ALL
SELECT '3. Clientes por documento', SUM(CantidadNeta), SUM(VentasNetas), SUM(CostosNetos)
FROM aptos
UNION ALL
SELECT '4. Clientes netos sin cantidad cero', SUM(Cantidad), SUM(Ventas), SUM(Costos)
FROM clientes WHERE ABS(Cantidad) >= 0.0001
UNION ALL
SELECT '5. LINEAS actual', CANTIDAD, VENTAS, COSTO
FROM dbo.vw_rentabilidad_lineas
WHERE FECHA = @Fecha AND [LÍNEA  DESCRIPCIÓN] = 'Total General';

-- Escenario 4 es una hipótesis de conciliación, no una regla autorizada:
-- una cantidad neta cero puede representar un ajuste monetario válido.
-- Debe contrastarse con las operaciones originales y el manual.
