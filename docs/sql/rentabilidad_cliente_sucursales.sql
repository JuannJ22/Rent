-- Vista exclusiva del generador: las vistas existentes no se modifican.
CREATE OR ALTER VIEW dbo.vw_rentabilidad_cliente_informe
AS
WITH mov AS (
    SELECT m.TipMov, m.ComMov, m.NroMov, m.FechaDctoMov,
           m.CuentasMov, m.ValorMov, m.NitMov, m.SucMov,
           m.VendedorMov, m.DescrMov, m.CantidadMov,
           m.LinMov, m.GrpMov, m.ProMov
    FROM dbo.MovimientoPorComprobante_Actual m
    WHERE m.TipMov IN ('F','J')
      AND (m.CuentasMov LIKE '41%' OR m.CuentasMov LIKE '1435%')
), agg AS (
    SELECT TipMov, ComMov, NroMov, FechaDctoMov, NitMov, SucMov,
           VendedorMov, DescrMov, LinMov, GrpMov, ProMov,
           (CASE WHEN TipMov='J' THEN -1.0 ELSE 1.0 END) * MAX(CantidadMov) AS Cantidad,
           (CASE WHEN TipMov='J' THEN -1.0 ELSE 1.0 END) *
               SUM(CASE WHEN CuentasMov LIKE '41%' THEN ValorMov ELSE 0 END) AS Ventas,
           (CASE WHEN TipMov='J' THEN -1.0 ELSE 1.0 END) *
               SUM(CASE WHEN CuentasMov LIKE '1435%' THEN ValorMov ELSE 0 END) AS Costos,
           (CASE WHEN TipMov='J' THEN -1.0 ELSE 1.0 END) *
               SUM(CASE WHEN CuentasMov LIKE '41%' THEN ValorMov *
                   CASE WHEN CuentasMov IN (
                       '4135420200','4135420700','4175050200','4175050500','4175050700'
                   ) THEN 1.00 ELSE 1.19 END ELSE 0 END) AS VentasConIVA
    FROM mov
    GROUP BY TipMov, ComMov, NroMov, FechaDctoMov, NitMov, SucMov,
             VendedorMov, DescrMov, LinMov, GrpMov, ProMov
)
SELECT a.NitMov AS NIT,
       CONCAT(RIGHT('0000000000000' + CAST(a.NitMov AS varchar(13)),13), ' - ',
           COALESCE(RIGHT('000' + CONVERT(varchar(20), TRY_CONVERT(int,a.SucMov)),3),'???'), ' - ',
           REPLACE(REPLACE(COALESCE(cli.NombreNit, ter.NombreNit, ven.NombreVen, 'SIN TERCERO'),CHAR(129),'Ñ'),'©','Ñ')
       ) AS [NIT - SUCURSAL - CLIENTE],
       REPLACE(REPLACE(a.DescrMov,CHAR(129),'Ñ'),'©','Ñ') AS DESCRIPCION,
       a.VendedorMov AS [COD. VENDEDOR], a.Cantidad AS CANTIDAD,
       a.Ventas, a.Costos,
       p.RENTA_PCT AS [% RENTA.], p.UTILI_PCT AS [% UTILI.],
       -- PRECIO y DESCUENTO se calculan en Excel con la captura de precios de la fecha.
       CAST(NULL AS decimal(18,5)) AS PRECIO,
       CAST(NULL AS decimal(18,2)) AS DESCUENTO,
       CONCAT(CAST(a.TipMov AS varchar(10)),CAST(a.ComMov AS varchar(10)), ' ',CAST(a.NroMov AS varchar(20))) AS DOCUMENTO,
       CONVERT(date,CONVERT(char(8),a.FechaDctoMov),112) AS FECHA,
       a.SucMov AS SUCURSAL,
       a.VentasConIVA AS VENTAS_CON_IVA
FROM agg a
CROSS APPLY (
    SELECT CAST(100.0 * (1 - ABS(a.Costos)/NULLIF(ABS(a.Ventas),0.0)) AS decimal(18,2)) AS RENTA_PCT,
           CAST(100.0 * (ABS(a.Ventas)/NULLIF(ABS(a.Costos),0.0) - 1) AS decimal(18,2)) AS UTILI_PCT
) p
OUTER APPLY (
    SELECT MAX(c.NombreNit) AS NombreNit
    FROM SiigoCat.dbo.TABLA_IDENTIFICACION_CLIENTES c
    WHERE c.NitNit=a.NitMov AND TRY_CONVERT(int,c.SucNit)=TRY_CONVERT(int,a.SucMov)
) cli
OUTER APPLY (
    SELECT MAX(t.NombreNit) AS NombreNit
    FROM SiigoCat.dbo.TABLA_IDENTIFICACION_TERCEROS t
    WHERE t.NitNit=a.NitMov AND TRY_CONVERT(int,t.SucNit)=TRY_CONVERT(int,a.SucMov)
) ter
OUTER APPLY (
    SELECT MAX(v.NombreVen) AS NombreVen
    FROM SiigoCat.dbo.TABLA_DESCRIPCION_VENDEDORES v WHERE v.CedulaVen=a.NitMov
) ven
WHERE (ABS(a.Ventas)>=0.01 OR ABS(a.Costos)>=0.01 OR ABS(a.Cantidad)>=0.0001)
  AND p.RENTA_PCT IS NOT NULL AND p.UTILI_PCT IS NOT NULL;
