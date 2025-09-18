"""Compatibilidad para proyectos que importaban ``Yesterday`` con mayúscula."""

from __future__ import annotations

import importlib
import sys

from yesterday import DateResolver, YesterdayStrategy

__all__ = ["DateResolver", "YesterdayStrategy"]

# Exponer el submódulo legado ``Yesterday.get_date`` reutilizando la versión
# definida en ``yesterday``.
sys.modules[__name__ + ".get_date"] = importlib.import_module("yesterday.get_date")

# Garantizar que ambos nombres remiten al mismo objeto de módulo.
sys.modules.setdefault("yesterday", sys.modules[__name__])
