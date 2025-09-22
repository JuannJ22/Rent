import subprocess
from pathlib import Path


class SiigoCLI:
    def importar_ordenes(self, archivo: Path) -> None:
        subprocess.run(["cmd", "/c", "echo", f"Importando {archivo}"], check=False)
