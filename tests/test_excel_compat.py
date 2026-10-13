from zipfile import ZipFile, ZIP_DEFLATED
from rentabilidad.infra.excel_compat import normalize_excel_application


def test_application_change_preserves_all_other_parts_and_is_idempotent(tmp_path):
    path = tmp_path / 'informe.xlsx'
    original = {
        'docProps/app.xml': b'<Properties><Application>Microsoft Excel Compatible / Openpyxl 3.1.5</Application><AppVersion>3.1</AppVersion></Properties>',
        'xl/worksheets/sheet1.xml': '<worksheet>Ñ =SUM(A1:A3)</worksheet>'.encode(),
        'xl/styles.xml': b'<styleSheet/>',
    }
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, data in original.items(): z.writestr(name, data)
    assert normalize_excel_application(path)
    with ZipFile(path) as z:
        for name, data in original.items():
            assert z.read(name) == (data.replace(b'Microsoft Excel Compatible / Openpyxl 3.1.5', b'Microsoft Excel') if name == 'docProps/app.xml' else data)
    assert not normalize_excel_application(path)
