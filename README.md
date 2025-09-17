# Rent

Automatiza la creación de un informe de rentabilidad a partir de una plantilla de Excel y la carga de datos EXCZ.

## Flujo

1. **Clonado de plantilla**: `excel_base/clone_from_template.py` copia `C:\\Rentabilidad\\PLANTILLA.xlsx` a `INFORME_YYYYMMDD.xlsx`.
2. **Carga de EXCZ**: `hojas/hoja01_loader.py` busca el archivo EXCZ con el prefijo configurado (por defecto `EXCZ980`) más reciente en `D:\\SIIWI01\\LISTADOS`, lo importa a la Hoja 1 aplicando fórmulas y, además, actualiza las hojas `CCOSTO1` a `CCOSTO4` filtrando el EXCZ `EXCZ979` por centro de costo.
3. **Scripts `.bat`**: automatizan el proceso:
   - `solo_clonar.bat` crea el informe a partir de la plantilla.
   - `solo_loader.bat` importa el EXCZ a un informe existente.
   - `todo_en_un_click.bat` ejecuta ambos pasos de forma secuencial.
   - `GenerarListadoProductos.bat` genera un catálogo de productos desde SIIGO y lo depura.

## Requisitos previos

- Windows con Python 3 en el `PATH`.
- Dependencias instaladas:

  ```
  pip install -r requirements.txt
  ```
- Archivo `PLANTILLA.xlsx` ubicado en `C:\\Rentabilidad\\`.
- Carpeta con los archivos EXCZ, por defecto `D:\\SIIWI01\\LISTADOS\\`. El script busca el más reciente cuyo nombre comience con el prefijo configurado (`EXCZ980` por defecto, ajustable con `EXCZPREFIX` o `--excz-prefix`).

## Instalación

1. Clonar este repositorio.
2. Instalar las dependencias con `pip install -r requirements.txt`.
3. Copiar `PLANTILLA.xlsx` a `C:\\Rentabilidad\\`.
4. Ajustar las rutas en los `.bat` si tus ubicaciones son distintas.

## Ejecución

- Para ejecutar todo el flujo en un paso:

  ```
  todo_en_un_click.bat
  ```

- Para crear sólo el informe vacío:

  ```
  solo_clonar.bat
  ```

- Para cargar el EXCZ a un informe existente:

  ```
  solo_loader.bat [ruta_a_informe.xlsx]
  ```

Cada script muestra mensajes en consola y pausa al final.

## Servicio: listado de productos desde SIIGO

`servicios/generar_listado_productos.py` ejecuta el comando `ExcelSIIGO` para
generar un Excel de productos en `C:\\Rentabilidad\\Productos` (carpeta
configurable) y luego deja únicamente las columnas **D**, **G** a **R** y
**AX**, filtrando además los productos cuyo campo `ACTIVO` (columna AX) sea
`S`.

- Ejecución rápida desde Windows:

  ```
  GenerarListadoProductos.bat
  ```

- Variables de entorno (opcionales) que ajustan las rutas por defecto:
  - `SIIGO_DIR`: carpeta donde está instalado SIIGO (por defecto `C:\\Siigo`).
  - `SIIGO_BASE`: ruta base pasada como primer parámetro a `ExcelSIIGO`
    (por defecto `D:\\SIIWI01`).
  - `PRODUCTOS_DIR`: carpeta destino de los Excel generados
    (por defecto `C:\\Rentabilidad\\Productos`).
  - `SIIGO_LOG`: ruta del archivo de log usado por `ExcelSIIGO`
    (por defecto `D:\\SIIWI01\\LOGS\\log_catalogos.txt`).

El archivo resultante sigue el formato `ProductosMMDD.xlsx`, usando la fecha
actual si no se indica otra con la opción `--fecha`.
