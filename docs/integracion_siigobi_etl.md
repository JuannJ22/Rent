# Rentabilidad como trabajo de SiigoBI.EtlService

El servicio .NET 8 añade `Kind: Rentabilidad`. Los trabajos existentes conservan su Python y sus argumentos. El nuevo trabajo usa exclusivamente el entorno Python 3.12 de Rent y las credenciales de `SqlTargets.RentReport`, que el instalador deriva del destino SQL instalado. No utiliza el XML DPAPI del administrador.

`RentabilidadDaily` se ejecuta a las 05:00 Colombia (UTC-5), para el día anterior. Si el servicio inicia después, recupera la ejecución pendiente del día actual. No recupera automáticamente todos los días anteriores. Después de un fallo espera 15 minutos y permite hasta tres intentos diarios. Espera el éxito de `Mov45Daily` y `CatalogosPesados`, incluida su ejecución nocturna del día que se informa. No publica si una dependencia tiene un fallo posterior a su último éxito.

El proceso del informe tiene tiempo máximo y se cancela al detener el servicio. Su bloqueo evita ejecuciones simultáneas. Copia la plantilla a una ubicación temporal, valida el resultado y publica mediante reemplazo atómico. Si existe un informe válido, lo conserva. `--replace` permite regenerarlo explícitamente. En fechas sin movimientos SQL, devuelve error y no publica un libro vacío.

Ruta del resultado: `C:\Rentabilidad\InformesDiarios\2026\Octubre\Octubre 06.xlsx`. Registro: `C:\SiigoBI_ETL\logs\RentabilidadDaily_*.log`. Estado y errores: el archivo `Service.StateFile` instalado.

## Instalación en Windows

1. Ejecutar `zona7_con_total.sql` en SQL2025/SiigoRent si aún no se aplicó. Debe devolver 132 / 4820237.02 / 3726154.17 para el 6 de octubre. El total proviene de SQL; Python solo compacta TERCEROS.
2. Descomprimir el paquete en `C:\Rentabilidad\ActualizarETL`. Cerrar Excel. Desde PowerShell como administrador: `& 'C:\Rentabilidad\ActualizarETL\Install-Rentabilidad.ps1'`. Verifica SHA256, conserva la configuración instalada, añade el destino y el trabajo deshabilitado. Guarda respaldos en `C:\SiigoBI_ETL\backups\rentabilidad-*`. Reinicia únicamente SiigoBI.EtlService; no modifica servicios SQL ni ADATEC. Si hay trabajos activos, termina sin instalar y debe repetirse cuando finalicen.
3. Ejecutar `& 'C:\Rentabilidad\ActualizarETL\Test-Rentabilidad.ps1' -CheckOnly`. Comprueba lectura SQL bajo la cuenta del operador; todavía no demuestra la ejecución real bajo LocalSystem. Para generar por fecha debe existir la captura de productos de ese día.
4. Después de que la prueba termine correctamente, ejecutar `& 'C:\Rentabilidad\ActualizarETL\Enable-Rentabilidad.ps1'`. Activa el trabajo y reinicia el servicio, con respaldo de configuración. El registro y estado `RentabilidadDaily` deben confirmar su primera ejecución real bajo LocalSystem. Solo entonces se puede afirmar que quedó operativo en el servidor.

El paquete no incluye el appsettings privado recibido ni modifica datos SQL salvo la vista que se aplica explícitamente. Las claves existentes se conservan. Los scripts de instalación están diseñados para las rutas concretas del servidor; cualquier otra ruta debe actualizarse también en el objeto Rentabilidad del trabajo.

Para una fecha manual: `Test-Rentabilidad.ps1 -Fecha '2026-10-06' -Replace`. Esto genera desde SQL y no modifica el horario. La pantalla del aplicativo debe validarse por separado antes de afirmar que la selección de fecha funciona allí.

## Desarrollo y verificación

Fuentes en `servicios/SiigoBI.EtlService`; no contienen credenciales. `dotnet publish -c Release -r win-x64 --self-contained false` produce el binario para Windows. Requiere .NET 8 instalado, como el servicio original.

Pruebas: `dotnet run --project servicios/SiigoBI.EtlService.Tests -c Release` y `python -m pytest`. Verifican hora colombiana, éxito diario, recuperación tras hora perdida, espera entre intentos, máximo diario y dependencias; Python verifica publicación y conservación del archivo anterior ante fallos.

La compilación y publicación se verificaron en Linux. No se ha ejecutado ni instalado el servicio en el servidor Windows privado desde este entorno.

## Productos y precios históricos

`ProductosDaily` es un trabajo independiente a las 20:45 Colombia, tras el éxito de `CatalogosMedios` y `CatalogosPesados` de ese día. Guarda todas las columnas y registros de `SiigoCat.dbo.vw_productos_activos` en `C:\Rentabilidad\Productos\productos-YYYY-MM-DD.xlsx`. Una hoja oculta registra fecha, instante UTC, vista y cantidad de filas. La primera captura válida de la fecha es inmutable. No se normalizan caracteres del SQL en la copia; se normalizan al utilizarla en el informe.

`RentabilidadDaily` depende también de `ProductosDaily` y usa el archivo de la fecha informada. No utiliza precios actuales como reemplazo de un archivo pasado faltante. La primera fecha informable (`FirstReportDate`) se establece al instalar para no intentar generar días anteriores al inicio del historial. Un informe atrasado se puede regenerar con `Test-Rentabilidad.ps1 -Fecha YYYY-MM-DD -Replace` mientras exista su captura.

Se conserva un mes calendario contado desde la fecha de captura (incluido el límite); por ejemplo, el 9 de octubre se conserva desde el 9 de septiembre. La limpieza se hace después de guardar o validar una captura correctamente. Solo elimina archivos con nombre y metadata de esta función; conserva listados manuales y archivos ajenos. No borra informes ya generados, que mantienen su PRECIOS dentro del propio Excel.

No se pueden reconstruir los precios de días anteriores al inicio leyendo la vista actual. Si se pierde tanto la captura diaria como el catálogo histórico, se informa el faltante. El CLI `--capture-products` rechaza fechar la vista actual como un día pasado y solo captura desde las 20:45 de Colombia.

## Instalación sin descargar ZIP

La actualización se puede obtener desde una rama de GitHub y construir con `servicios\SiigoBI.EtlService\Build-Package.ps1`. Construye el paquete localmente y no modifica el servicio. Si no existe SDK .NET, instala SDK 8 en `C:\Rentabilidad\Herramientas\dotnet` usando el instalador oficial de Microsoft, con TLS validado. Las configuraciones privadas no se incluyen en el repositorio ni se sustituyen por las de ejemplo.

Después de construir: aplicar la vista SQL si falta, ejecutar `Install-Rentabilidad.ps1`, probar la conexión con `Test-Rentabilidad.ps1 -CheckOnly` y activar ambos trabajos con `Enable-Rentabilidad.ps1`. La primera captura se hará en el próximo cierre de las 20:45 (o recuperará ese cierre si se activa después, dentro del mismo día). El primer informe será a las 05:00 del día siguiente. La prueba histórica completa requiere que la captura de esa fecha ya exista; no se debe probar octubre 6 sin una copia válida de octubre 6.

La ejecución real bajo LocalSystem y el resultado de la captura deben confirmarse en los registros del servidor. La pantalla GUI todavía requiere validación de la selección de fecha; el motor SQL ya usa automáticamente la copia correspondiente cuando está disponible y rechaza precios actuales para una fecha pasada sin copia, por defecto.
