# Ejemplo de alerta de precio en Hoja 1

La macro de carga marca en amarillo las filas en las que el precio unitario que
se facturó difiere del precio de lista en más de 0,2 %. Esa tolerancia se
controla con la constante `PRICE_TOLERANCE = 0.002` del cargador de la Hoja 1.
Cuando la diferencia supera ese umbral, todo el renglón se llena de amarillo y
la columna de observaciones muestra el mensaje
«`Precio total menor que la lista en …`».【F:hojas/hoja01_loader.py†L49-L52】【F:hojas/hoja01_loader.py†L2292-L2330】

A modo de ejemplo, la fila del cliente 0000094281893 con el producto «ESTUCO
IMPADOC x 25 KLS BLANCO PLUS» muestra el mensaje:

> Precio total menor que la lista en $78.480,25 (-7,65 %).

Los números del renglón permiten reconstruir el cálculo:

1. **Cantidad facturada (CANT)**: 26 unidades.
2. **Venta total sin IVA (`VENTAS`)**: $947.798,32.
3. **Diferencia total indicada por el mensaje**: $78.480,25.

Con esos datos, el script calcula los valores unitarios sin IVA:

- Precio unitario real = 947.798,32 ÷ 26 = **$36.453,78**.
- La diferencia por unidad = 78.480,25 ÷ 26 = **$3.018,47**.
- Precio unitario de lista = 36.453,78 + 3.018,47 = **$39.472,25**.

Por último, el porcentaje de variación relativo es 3.018,47 ÷ 39.472,25 =
0,0765, es decir **7,65 %**, que coincide con el porcentaje mostrado en el
mensaje. Como 7,65 % es mayor que la tolerancia permitida (0,2 %), la línea se
marca de amarillo.【F:hojas/hoja01_loader.py†L2289-L2330】

Este análisis se realiza completamente en valores sin IVA, por lo que las
cantidades y precios deben interpretarse antes de impuestos.
