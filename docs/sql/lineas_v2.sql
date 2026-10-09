-- Vista nueva de validación. No reemplaza las vistas actuales.
USE [SiigoRent];
GO
CREATE OR ALTER VIEW dbo.vw_rentabilidad_lineas_v2
AS
WITH documentos AS (
    SELECT m.TipMov, m.ComMov, m.NroMov, m.FechaDctoMov,
           m.NitMov, m.VendedorMov, m.LinMov, m.GrpMov, m.ProMov, m.DescrMov,
           MAX(m.CantidadMov) AS Cantidad,
           SUM(CASE WHEN m.CuentasMov LIKE '41%' THEN m.ValorMov ELSE 0 END) AS Ventas,
           SUM(CASE WHEN m.CuentasMov LIKE '1435%' THEN m.ValorMov ELSE 0 END) AS Costos
    FROM dbo.MovimientoPorComprobante_Actual AS m
    WHERE m.TipMov IN ('F','J')
      AND (m.CuentasMov LIKE '41%' OR m.CuentasMov LIKE '1435%')
    GROUP BY m.TipMov, m.ComMov, m.NroMov, m.FechaDctoMov,
             m.NitMov, m.VendedorMov, m.LinMov, m.GrpMov, m.ProMov, m.DescrMov
), firmados AS (
    SELECT CONVERT(date, CONVERT(char(8), d.FechaDctoMov), 112) AS FECHA,
           d.LinMov, d.GrpMov, d.ProMov, d.DescrMov,
           CASE WHEN d.TipMov = 'J' THEN -d.Cantidad ELSE d.Cantidad END AS CANTIDAD,
           CASE WHEN d.TipMov = 'J' THEN -d.Ventas ELSE d.Ventas END AS VENTAS,
           CASE WHEN d.TipMov = 'J' THEN -d.Costos ELSE d.Costos END AS COSTO
    FROM documentos AS d
    WHERE d.Ventas <> 0 AND d.Costos <> 0
      AND CAST(100.0 * (1 - ABS(d.Costos) / NULLIF(ABS(d.Ventas),0)) AS decimal(18,2))
          NOT BETWEEN 99.9 AND 100.1
), filtrados AS (
    SELECT f.* FROM firmados AS f
    WHERE NOT EXISTS (
        SELECT 1 FROM SiigoCat.dbo.TABLA_DESCRIPCION_INVENTARIOS AS tg
        WHERE tg.LinTinv = f.LinMov AND tg.GruTinv = f.GrpMov
          AND (UPPER(LTRIM(RTRIM(tg.NomGruTinv))) IN (
              'BOLSA PLASTICA','DOMICILIOS','CONCENTRADOS','CONCENTRADO','ADICIONAL CONCENTRADO'
          ) OR UPPER(tg.NomGruTinv) LIKE '%DOMICILIO%'
            OR UPPER(tg.NomGruTinv) LIKE '%CONCENTRAD%')
    ) AND NOT (
        f.LinMov = 17 AND f.GrpMov = 1
        AND UPPER(LTRIM(RTRIM(f.DescrMov))) LIKE '%TAMBOR%VACIO%'
    )
), resumen AS (
    SELECT FECHA, LinMov, GrpMov, GROUPING(LinMov) AS TotalGeneral,
           GROUPING(GrpMov) AS TotalLinea,
           SUM(CANTIDAD) AS CANTIDAD, SUM(VENTAS) AS VENTAS, SUM(COSTO) AS COSTO
    FROM filtrados
    GROUP BY GROUPING SETS ((FECHA,LinMov,GrpMov),(FECHA,LinMov),(FECHA))
)
SELECT
    CASE WHEN r.TotalGeneral = 1 THEN 'Total General'
         WHEN r.TotalLinea = 1 THEN CONCAT('Total ',RIGHT('0000'+CAST(r.LinMov AS varchar(10)),4),' ',COALESCE(tl.Nombre,'SIN CATALOGO'))
         ELSE '' END AS [LÍNEA  DESCRIPCIÓN],
    CASE WHEN r.TotalLinea = 0 THEN CONCAT('Total ',RIGHT('0000'+CAST(r.GrpMov AS varchar(10)),4),' ',COALESCE(tg.Nombre,'SIN CATALOGO'))
         ELSE '' END AS [GRUPO  DESCRIPCIÓN],
    r.CANTIDAD, r.VENTAS, r.COSTO,
    CAST(100.0*(1-r.COSTO/NULLIF(r.VENTAS,0)) AS decimal(18,2)) AS [%RENTABILIDAD],
    CAST(100.0*(r.VENTAS/NULLIF(r.COSTO,0)-1) AS decimal(18,2)) AS [%UTILIDAD],
    r.FECHA,
    CASE WHEN r.TotalGeneral = 1 THEN 9999 ELSE TRY_CONVERT(int,r.LinMov) END AS _LineaOrden,
    CASE WHEN r.TotalGeneral = 1 THEN 2 WHEN r.TotalLinea = 1 THEN 1 ELSE 0 END AS _TipoFila,
    CASE WHEN r.TotalLinea = 1 THEN 9999 ELSE TRY_CONVERT(int,r.GrpMov) END AS _GrupoOrden
FROM resumen AS r
OUTER APPLY (
    SELECT MAX(LTRIM(RTRIM(NomLinTinv))) AS Nombre
    FROM SiigoCat.dbo.TABLA_DESCRIPCION_INVENTARIOS
    WHERE LinTinv = r.LinMov AND GruTinv = 0
) AS tl
OUTER APPLY (
    SELECT MAX(LTRIM(RTRIM(NomGruTinv))) AS Nombre
    FROM SiigoCat.dbo.TABLA_DESCRIPCION_INVENTARIOS
    WHERE LinTinv = r.LinMov AND GruTinv = r.GrpMov
) AS tg;
GO
SELECT [LÍNEA  DESCRIPCIÓN], CANTIDAD, VENTAS, COSTO, [%RENTABILIDAD]
FROM dbo.vw_rentabilidad_lineas_v2
WHERE FECHA = '20261006' AND [LÍNEA  DESCRIPCIÓN] = 'Total General';
