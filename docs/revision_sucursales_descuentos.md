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

## Productos repetidos en una factura

La cantidad se suma exclusivamente en los renglones de ventas (cuentas 41),
conservando las claves de agrupación de cada vista y el signo de las devoluciones.
No se suman los asientos de costo 1435 y no se usa SUM DISTINCT: dos ítems de
3 unidades del mismo producto son 6 unidades, aunque tengan igual cantidad.

`Install-QuantityFix.ps1` respalda las definiciones SQL, sustituye únicamente
MAX(CantidadMov) de las vistas del informe que leen el movimiento original y
actualiza los metadatos de las vistas dependientes. La operación SQL es transaccional.
Comprueba las facturas F201 25052 (15 unidades) y F241 6391 (6 unidades) y que sus
ventas/costos no cambien. Si un control falla, revierte la transacción. El respaldo
incluye `restaurar-vistas.sql` y `controles-cantidades.json`.

El instalador prepara un informe del 7 de octubre en `_Revision` con el catálogo
existente de esa fecha; los informes ya publicados no se regeneran automáticamente.
También aplica la regla solicitada para terceros sin ficha de cliente: cualquier
vendedor puede facturarles, sin aviso de vendedor o lista sin verificar. Para un
cliente registrado cuya sucursal falte o tenga asignaciones contradictorias, sí se
conserva el aviso; nunca se toma la asignación de otra sucursal.
