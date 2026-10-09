"""Aplica la corrección de columnas SQL con respaldo del cargador existente."""
import ast
from datetime import datetime
import os
from pathlib import Path
import sys
import uuid

path = Path(sys.argv[1])
backup_dir = Path(sys.argv[2])
original = path.read_bytes()
start = original.index(b"def _update_terceros_sheet_from_df(")
end = original.index(b"\ndef _guess_sql_precios_columns(", start)
section = original[start:end]
old = [b"column=2, value=vendedor", b"column=3, value=lista_precio"]
new = [b"column=2, value=lista_precio", b"column=3, value=vendedor"]
if all(section.count(item) == 1 for item in new) and all(item not in section for item in old):
    print("La corrección ya estaba aplicada.")
    raise SystemExit(0)
if not all(section.count(item) == 1 for item in old):
    raise RuntimeError("El código no coincide con la versión esperada. No se modificó.")
for before, after in zip(old, new):
    section = section.replace(before, after, 1)
updated = original[:start] + section + original[end:]
ast.parse(updated.decode("utf-8-sig"))
backup_dir.mkdir(parents=True, exist_ok=True)
backup = backup_dir / f"hoja01_loader-{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}.py"
with backup.open("xb") as handle:
    handle.write(original)
temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
try:
    temporary.write_bytes(updated)
    if path.read_bytes() != original:
        raise RuntimeError("El cargador cambió durante la operación. Se conserva sin reemplazar.")
    os.replace(temporary, path)
finally:
    temporary.unlink(missing_ok=True)
print(f"Corrección aplicada: NIT, lista de precios, vendedor. Respaldo: {backup}")
