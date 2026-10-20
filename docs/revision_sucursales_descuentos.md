# Validación de sucursales, vendedores y descuentos

El informe usa `SiigoRent.dbo.vw_rentabilidad_cliente_informe`, creada por
`docs/sql/rentabilidad_cliente_sucursales.sql`. Hay que crear esta vista antes de
instalar el loader nuevo o activar un paquete nuevo en un servidor sin ella.
Las vistas existentes siguen disponibles para sus consumidores.

- La agrupación incluye `SucMov`; nombre, vendedor asignado y lista se resuelven
  por NIT y sucursal. Los campos de asignación provienen exclusivamente de
  `TABLA_IDENTIFICACION_CLIENTES`, que contiene `VendedorNit` y `PrecioNit`.
- Una sucursal desconocida o una asignación duplicada y contradictoria se señala
  como pendiente; no se selecciona la sucursal menor ni una lista por defecto.
- El vendedor facturado y el documento se toman de la misma fila SQL. Continúan
  las equivalencias de códigos y la excepción de consumidor final ya configuradas.
- DESCUENTO conserva lista 12 de la captura de la fecha: 1 − precio unitario con
  IVA / lista 12. Ventas con IVA se calculan por cada cuenta de ventas antes de
  sumar, aplicando la regla de cuentas exentas de la vista que estaba instalada.
  Excel guarda ese importe en una columna auxiliar oculta y referencia su celda.
- Los ajustes sin cantidad no tienen descuento. Los descuentos negativos se
  conservan: por ejemplo, unitario con IVA 180 y lista 12 de 100 equivale a −80 %.
  Un negativo no demuestra un error; revisar importe, cantidad, precio de referencia
  y cuenta de venta de ese documento. No se limita el resultado a cero.

`Install-BranchValidation.ps1` crea la vista, respalda los scripts y prepara una
copia de revisión en `C:\Rentabilidad\_Revision\sucursales-<fecha>`, con la captura
histórica de la fecha solicitada. Los informes publicados se conservan. La misma
corrección se aplicará en las siguientes ejecuciones automáticas. No se reconstruyen
precios históricos con el catálogo actual.
