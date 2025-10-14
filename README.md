# Rent

Automatiza la creación de un informe de rentabilidad a partir de una plantilla de Excel y la carga de datos EXCZ.

## Estructura de carpetas

Los scripts generan automáticamente la siguiente estructura base (todas
las rutas son configurables mediante variables de entorno):

```
C:\Rentabilidad\
 ├── Informes\<Mes>\<Mes> DD.xlsx
 └── Productos\productosMMDD.xlsx
```

Cada informe se almacena dentro del mes correspondiente y los listados
de productos se guardan en la carpeta `Productos`.

## Flujo

### 1. Clonado de plantilla

- Script principal: `excel_base/clone_from_template.py`.
- Plantilla origen: `C:\Rentabilidad\PLANTILLA.xlsx`.
- Resultado: genera `<Mes> DD.xlsx` en la carpeta del mes calculada para la fecha objetivo (por defecto el día anterior).

### 2. Carga de EXCZ

- Script principal: `hojas/hoja01_loader.py`.
- Origen de datos: busca en `D:\SIIWI01\LISTADOS` el archivo `EXCZ***YYYYMMDDHHMMSS` cuya fecha coincida con la solicitada (por defecto el día anterior).
- Acciones: importa el EXCZ en la Hoja 1, aplica las fórmulas necesarias y actualiza las hojas `CCOSTO` y `COD` con la misma fecha.

### 3. Scripts `.bat`

- `solo_clonar.bat`: crea el informe a partir de la plantilla.
- `solo_loader.bat`: importa el EXCZ a un informe existente.
- `todo_en_un_click.bat`: ejecuta ambos pasos de forma secuencial.
- `GenerarListadoProductos.bat`: genera un catálogo de productos desde SIIGO y lo depura.

## Requisitos previos

- Windows con Python 3 en el `PATH`.
- Dependencias instaladas:

  ```
  pip install -r requirements.txt
  ```
- Archivo `PLANTILLA.xlsx` ubicado en `C:\\Rentabilidad\\`.
- Carpeta con los archivos EXCZ, por defecto `D:\\SIIWI01\\LISTADOS\\`. Los nombres deben seguir el patrón `EXCZ***YYYYMMDDHHMMSS` para permitir la selección por fecha (prefijo configurable con `EXCZPREFIX` o `--excz-prefix`).

## Instalación

1. Clonar este repositorio.
2. Instalar las dependencias con `pip install -r requirements.txt`.
3. Copiar `PLANTILLA.xlsx` a `C:\\Rentabilidad\\`.
4. Ajustar las rutas en los `.bat` si tus ubicaciones son distintas.

## Ejecución

Ejecuta los scripts desde Windows según la tarea que necesites:

```bat
:: Ejecuta todo el flujo en un paso
todo_en_un_click.bat

:: Crea sólo el informe vacío
solo_clonar.bat

:: Carga el EXCZ en un informe existente (usa la fecha del día anterior si no se especifica --fecha)
solo_loader.bat ruta_a_informe.xlsx
```

Reemplaza `ruta_a_informe.xlsx` por la ubicación del archivo a actualizar. El parámetro `--fecha YYYY-MM-DD` es opcional y, si se omite, se utiliza el día anterior.

Cada script muestra mensajes en consola y pausa al final.

## Interfaz gráfica (NiceGUI)

El proyecto incluye un panel web en `rentabilidad/gui/web.py`. Para
ejecutarlo utiliza `python -m rentabilidad.gui.web`. El servidor expone un
encabezado con el logotipo de la empresa y el título del panel.

### Dónde colocar el logotipo

- La aplicación busca un archivo llamado `logo.svg` dentro de la carpeta
  `rentabilidad/gui/static/` del repositorio.
- Puedes reemplazar el archivo de ejemplo (`rentabilidad/gui/static/logo.svg`)
  por el logotipo real de la empresa conservando el mismo nombre.
- Tras guardar los cambios, reinicia el servidor de NiceGUI para ver el nuevo
  logotipo en la esquina superior izquierda.

## Servicio: listado de productos desde SIIGO

`servicios/generar_listado_productos.py` ejecuta el comando `ExcelSIIGO`
para generar un Excel de productos en `C:\\Rentabilidad\\Productos`
(carpeta configurable) y luego deja únicamente las columnas **D**, **G** a
**R** y **AX**, filtrando además los productos cuyo campo `ACTIVO`
(columna AX) sea `S`. El nombre resultante sigue el formato
`productosMMDD.xlsx` y, por defecto, utiliza la fecha actual.

- Ejecución rápida desde Windows:

  ```
  GenerarListadoProductos.bat
  ```

- Variables de entorno (opcionales) que ajustan las rutas por defecto:
  - `SIIGO_DIR`: carpeta donde está instalado SIIGO (por defecto `C:\\Siigo`).
  - `SIIGO_COMMAND`: nombre del ejecutable de SIIGO (por defecto `ExcelSIIGO`).
  - `SIIGO_BASE`: ruta base pasada como primer parámetro a `ExcelSIIGO`
    (por defecto `D:\\SIIWI01`).
  - `PRODUCTOS_DIR`: carpeta destino de los Excel generados
    (por defecto `C:\\Rentabilidad\\Productos`).
  - `SIIGO_LOG`: ruta del archivo de log usado por `ExcelSIIGO`
    (por defecto `D:\\SIIWI01\\LOGS\\log_catalogos.txt`).

El archivo resultante sigue el formato `productosMMDD.xlsx`, usando la fecha
actual si no se indica otra con la opción `--fecha`.
