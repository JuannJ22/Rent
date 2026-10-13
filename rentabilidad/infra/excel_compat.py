"""Normaliza el identificador de aplicacion para lectores Excel de escritorio."""
from pathlib import Path
import os
import re
import tempfile
from zipfile import ZipFile


def normalize_excel_application(path: Path) -> bool:
    """Cambia solo Application de app.xml; conserva el resto de partes byte a byte."""
    path = Path(path)
    with ZipFile(path) as source:
        bad = source.testzip()
        if bad:
            raise ValueError(f'Parte XLSX corrupta: {bad}')
        original = source.read('docProps/app.xml')
        updated, count = re.subn(
            rb'(<(?:[\w]+:)?Application\b[^>]*>)[^<]*(</(?:[\w]+:)?Application>)',
            rb'\g<1>Microsoft Excel\g<2>', original,
        )
        if count != 1:
            raise ValueError('No se encontro un identificador Application unico.')
        if updated == original:
            return False
        handle, temporary = tempfile.mkstemp(suffix='.xlsx', prefix='.compat-', dir=path.parent)
        os.close(handle)
        try:
            with ZipFile(temporary, 'w') as destination:
                destination.comment = source.comment
                for entry in source.infolist():
                    destination.writestr(entry, updated if entry.filename == 'docProps/app.xml' else source.read(entry.filename))
            with ZipFile(temporary) as check:
                if check.testzip():
                    raise ValueError('La copia de compatibilidad no paso la comprobacion ZIP.')
        except BaseException:
            Path(temporary).unlink(missing_ok=True)
            raise
    try:
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return True


def format_report_presentation(wb):
    """Conserva origen oculto y muestra codigos numericos alineados a la izquierda."""
    from copy import copy
    import unicodedata
    if 'ORIGEN_PRECIOS' in wb.sheetnames:
        wb['ORIGEN_PRECIOS'].sheet_state = 'hidden'
    for ws in wb:
        columns = []
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
            for cell in row:
                label = ''.join(c for c in unicodedata.normalize('NFD', str(cell.value or '').upper()) if not unicodedata.combining(c))
                label = re.sub(r'[^A-Z0-9]', '', label)
                if label in {'CODVENDEDOR', 'CODIGOVENDEDOR'}:
                    columns.append((cell.row, cell.column))
        for header_row, column in columns:
            for row in ws.iter_rows(min_row=header_row+1, min_col=column, max_col=column):
                cell = row[0]
                if cell.data_type == 'f' or cell.value is None:
                    continue
                text = str(cell.value).strip()
                if re.fullmatch(r'[0-9]+(?:\.0+)?', text):
                    cell.value = int(float(text))
                    cell.number_format = '0'
                    alignment = copy(cell.alignment)
                    alignment.horizontal = 'left'
                    cell.alignment = alignment


if __name__ == '__main__':
    import argparse
    import shutil
    parser = argparse.ArgumentParser(description='Crea una copia de prueba sin modificar datos ni formulas.')
    parser.add_argument('original', type=Path)
    parser.add_argument('copia', type=Path)
    args = parser.parse_args()
    if args.copia.exists() or args.original.resolve() == args.copia.resolve():
        parser.error('La copia debe ser un archivo nuevo.')
    shutil.copy2(args.original, args.copia)
    normalize_excel_application(args.copia)
    print(f'COPIA PREPARADA: {args.copia}')
