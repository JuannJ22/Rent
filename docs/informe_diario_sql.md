# Informe diario desde SQL en Windows Server 2022

Destino acordado: Windows Server 2022 siempre encendido, con SQL Server 2025
en `192.168.5.10,14331`. Ejecución diaria a las **05:00, hora de Colombia**, para
el **día anterior**. No requiere listados EXCZ, ejecutables SIIGO ni una web
abierta. Windows programa el proceso incluso sin una sesión iniciada.

La carga diaria de SQL está programada a las **20:30**. Se mantiene el informe
a las **05:00 del día siguiente**, con 8 horas y media de margen desde el
inicio previsto de la carga. Ese margen no confirma que una carga fallida o
demorada haya terminado. No se adelanta a las 20:00, antes de la carga.
La conexión usará **usuario de SQL**, según la autenticación del servidor.

## Lo que consulta realmente el motor

La configuración histórica decía SiigoBI, pero el informe actual consulta:

| Base | Objetos |
| --- | --- |
| SiigoRent | `vw_rentabilidad_cliente`, `vw_rentabilidad_lineas_ordenadas`, `vw_rentabilidad_zona1/2/3/7`, `vw_rentabilidad_vendedor24/25/26/27/29/30/51/52` |
| SiigoCat | `vw_productos_activos`, `TABLA_DESCRIPCION_VENDEDORES`, `TABLA_IDENTIFICACION_CLIENTES`, `TABLA_IDENTIFICACION_TERCEROS` |
| Siigo2627 | `TABLA_MOVIMIENTO_POR_COMPROBANTE` |

Todos los objetos usan el esquema `dbo`. La cuenta necesita permiso de lectura
en las tres bases, incluidas las dependencias de las vistas. No necesita ser `sa`.
Las vistas de rentabilidad se consultan con `FECHA = ?`, parametrizada con la
fecha del informe. Se conserva la consulta histórica completa de movimientos:
antes de reducirla hay que confirmar la regla de asignación de vendedores y el
tipo de `FactMov`. Puede ser la consulta más costosa.

Los nombres de las bases pueden ajustarse con `SQL_RENT_DATABASE`,
`SQL_CAT_DATABASE` y `SQL_MOV_DATABASE`. Las opciones históricas de tablas
`SQL_MOVIMIENTOS_TABLE`, `SQL_PRECIOS_TABLE`, etc. no sustituyen estas vistas.
El motor reutiliza una conexión para las 19 consultas y limita cada consulta
con `SQL_TIMEOUT`. Una conexión compartida no garantiza una instantánea:
programe después del fin de la carga de SQL.

## Instalación en el servidor

1. Instale Python 3.12 de 64 bits y Microsoft ODBC Driver 18 for SQL Server
   de 64 bits, desde sus distribuidores oficiales.
2. Ubique el repositorio, por ejemplo en `C:\Rent`, y ejecute:

   ```powershell
   Set-Location C:\Rent
   py -3.12 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

3. Coloque la plantilla real en `C:\Rentabilidad\PLANTILLA.xlsx` y copie
   `docs\sql_daily.example.json` a `C:\Rentabilidad\sql_config.json`.
   El ejemplo usa usuario de SQL (`SQL_TRUSTED=false`). Complete `SQL_USER`
   y `SQL_PASSWORD` solamente en esa copia local, fuera del repositorio,
   mediante un editor en el servidor. Use una cuenta SQL con lectura en las
   tres bases. Restrinja los permisos del archivo a la cuenta Windows de la
   tarea, SYSTEM y los administradores autorizados; no lo guarde en Git ni lo
   comparta por chat. La cuenta Windows de la tarea es distinta del usuario
   SQL: necesita leer la plantilla, la configuración y el repositorio, y
   escribir en `C:\Rentabilidad`.
4. Configure la conexión cifrada y el certificado del servidor. El ejemplo
   verifica el certificado: si este no incluye la IP `192.168.5.10`, use en
   `SQL_SERVER` el nombre DNS que figure en el certificado, con `,14331`, y
   confíe en su autoridad emisora mediante el almacén de Windows. No desactive
   la verificación para resolver un error de certificado.

Como alternativa al archivo local, puede inyectar `SQL_USER` y `SQL_PASSWORD`
en el entorno de la cuenta Windows que ejecuta la tarea, dejando vacíos esos
campos en el JSON. Compruebe que la tarea programada recibe esas variables;
las variables de una consola temporal no persisten para la ejecución diaria.
No coloque contraseñas en comandos, documentación o archivos de Git.
La contraseña requiere un valor real accesible al proceso ODBC; una credencial
de proxy HTTPS de un entorno en la nube no autentica esta conexión TCP privada.

## Validación antes de programar

Ejecute bajo la cuenta Windows que usará la tarea:

```powershell
Set-Location C:\Rent
.\.venv\Scripts\python.exe -m rentabilidad.daily_report --base-dir C:\Rentabilidad --sql-config C:\Rentabilidad\sql_config.json --check-sql
.\.venv\Scripts\python.exe -m rentabilidad.daily_report --base-dir C:\Rentabilidad --sql-config C:\Rentabilidad\sql_config.json
```

La primera orden consulta las 19 fuentes sin generar Excel y requiere detalles
para ayer. Comprueba acceso real, objetos y ejecución de consultas; el número
de filas no prueba que la carga esté completa. La segunda carga la plantilla
con el motor existente y publica el archivo solamente cuando termina con éxito.
Compare el primer informe con el informe validado manualmente: ventas, costo,
cantidades, márgenes y hojas por zona/vendedor. Las fórmulas Excel se conservan;
openpyxl no las calcula, por lo que hay que comprobar los resultados al abrir
el libro en Excel. El diagnóstico SQL por sí solo no valida la plantilla.

Salida: `C:\Rentabilidad\InformesDiarios\2026\Octubre\Octubre 05.xlsx`, por
ejemplo. Se usa una carpeta anual distinta de la salida histórica del panel
para evitar sobrescribir el mismo día de otro año. Logs en
`C:\Rentabilidad\Logs\daily-YYYY-MM-DD.log`.

No se publica un informe sin detalles; un día sin operaciones queda registrado
como fallo y debe revisarse. Si el día tiene datos parciales, este control no
los detecta: falta acordar una señal de cierre de carga (tabla de control,
estado del job de SQL o equivalente). Hasta entonces, la hora de las 05:00
debe quedar después de la actualización de las fuentes.

Una repetición conserva el informe existente. Para corregir un día después de
una nueva carga, use explícitamente:

```powershell
.\.venv\Scripts\python.exe -m rentabilidad.daily_report --base-dir C:\Rentabilidad --sql-config C:\Rentabilidad\sql_config.json --fecha 2026-10-05 --replace
```

El reemplazo se realiza desde un archivo temporal en el mismo volumen. Un fallo
de SQL, plantilla o tiempo límite conserva el archivo anterior. Cierre el Excel
durante un reproceso para evitar que Windows impida reemplazarlo.

## Programar a las 05:00

Compruebe que la zona del servidor sea Colombia (`SA Pacific Standard Time`).
El cálculo de ayer usa `America/Bogota` aunque cambie la zona del proceso, pero
el disparador del Programador de tareas usa la zona del servidor.

Desde PowerShell, con permisos para registrar tareas:

```powershell
C:\Rent\servicios\Register-DailyReport.ps1 -RepositoryPath C:\Rent -PythonPath C:\Rent\.venv\Scripts\python.exe
```

El script solicita la cuenta Windows localmente y registra ejecución diaria,
sin sesión abierta, con hasta tres reintentos a intervalos de 15 minutos,
ejecución cuando se recupera una hora perdida y sin instancias simultáneas.
No reemplaza tareas existentes. Puede revisar la operación con `-WhatIf`.
La cuenta requiere el derecho de inicio de sesión como trabajo por lotes;
verifique las políticas del servidor y sus permisos SQL.

```powershell
Start-ScheduledTask -TaskName Rentabilidad-Diaria-SQL
Get-ScheduledTaskInfo -TaskName Rentabilidad-Diaria-SQL
```

Revise el resultado tras finalizar (`LastTaskResult = 0`), el log y el Excel
generado. Un fallo devuelve código 1 y activa los reintentos. El motor tiene
un límite total de 30 minutos por intento; el bloqueo se libera al terminar.
Los logs permiten diagnóstico local; aún no hay avisos por correo ni vigilancia
externa de días faltantes. La interfaz web puede incorporarse para consultar
historial, estados y reprocesar, manteniendo independiente la tarea diaria.

La tarea y la conexión real deben validarse en Windows Server 2022: este cambio
se desarrolla y prueba en Linux, sin acceso a la red privada ni plantilla real.
