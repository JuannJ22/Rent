from pathlib import Path

from openpyxl import Workbook, load_workbook

from rentabilidad.services.products import WorkbookCleaner


def _crear_libro(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Productos"
    ws.append(["COD", "DESCRIPCIÓN", "ACTIVO", "PRECIO", "OTRA"])
    ws.append(["P-1", "Producto 1", "S", 120.0, "IGNORAR"])
    ws.append(["P-2", "Producto 2", "n", 30.0, "IGNORAR"])
    wb.save(path)
    wb.close()


def test_workbook_cleaner_acepta_columnas_en_letras(tmp_path) -> None:
    destino = tmp_path / "productos.xlsx"
    _crear_libro(destino)

    cleaner = WorkbookCleaner(activo_column="C", keep_columns=["A", "4"])
    cleaner.clean(destino)

    libro = load_workbook(destino)
    hoja = libro.active

    assert hoja.max_row == 2
    assert hoja.max_column == 3
    assert [hoja.cell(1, col).value for col in range(1, 4)] == ["COD", "ACTIVO", "PRECIO"]
    assert [hoja.cell(2, col).value for col in range(1, 4)] == ["P-1", "S", 120.0]

    libro.close()


def test_workbook_cleaner_acepta_columna_activo_numerica(tmp_path) -> None:
    destino = tmp_path / "productos.xlsx"
    _crear_libro(destino)

    cleaner = WorkbookCleaner(activo_column="3", keep_columns=[1, "4", 5])
    cleaner.clean(destino)

    libro = load_workbook(destino)
    hoja = libro.active

    assert hoja.max_row == 2
    assert hoja.max_column == 4
    assert [hoja.cell(1, col).value for col in range(1, 5)] == [
        "COD",
        "ACTIVO",
        "PRECIO",
        "OTRA",
    ]

    libro.close()
