-- Solo lectura. Ejecutar en SSMS para confirmar las cantidades y las vistas base.
USE SiigoRent;

SELECT DOCUMENTO, DESCRIPCION, CANTIDAD, CANTIDAD_VENTA_SUM,
       FILAS_VENTA, Ventas, Costos
FROM dbo.vw_rentabilidad_cliente_informe
WHERE FECHA = CONVERT(date,'20261007',112)
  AND DOCUMENTO IN ('F201 25052','F241 6391')
  AND (DESCRIPCION LIKE '%GRANI%' OR DESCRIPCION LIKE '%BROCHA%');

SELECT TipMov, ComMov, NroMov, SecMov, NitMov, SucMov, DescrMov,
       CuentasMov, CantidadMov, ValorMov
FROM dbo.MovimientoPorComprobante_Actual
WHERE FechaDctoMov = 20261007
  AND ((TipMov='F' AND ComMov=201 AND NroMov=25052)
    OR (TipMov='F' AND ComMov=241 AND NroMov=6391))
  AND (DescrMov LIKE '%GRANI%' OR DescrMov LIKE '%BROCHA%')
  AND (CuentasMov LIKE '41%' OR CuentasMov LIKE '1435%')
ORDER BY ComMov,NroMov,CuentasMov,SecMov;

SELECT v.name AS Vista,m.definition AS Definicion
FROM sys.views v JOIN sys.sql_modules m ON m.object_id=v.object_id
WHERE v.name IN ('vw_rentabilidad_base','vw_rentabilidad_base_linea','vw_rentabilidad_lineas');
