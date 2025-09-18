"""Compatibilidad con importaciones antiguas de ``yesterday``.

Este paquete mantiene la API original expuesta por ``yesterday`` para que el
código legado continúe funcionando. Los componentes reales ahora viven en
:mod:`rentabilidad.core.dates`.
"""

from __future__ import annotations

import sys
from types import ModuleType

from rentabilidad.core.dates import DateResolver, YesterdayStrategy

__all__ = ["DateResolver", "YesterdayStrategy"]


def _register_submodule() -> None:
    """Crea el submódulo ``yesterday.get_date`` para compatibilidad."""

    module_name = __name__ + ".get_date"
    if module_name in sys.modules:
        return

    shim = ModuleType(module_name)
    shim.DateResolver = DateResolver
    shim.YesterdayStrategy = YesterdayStrategy
    shim.__all__ = ["DateResolver", "YesterdayStrategy"]
    sys.modules[module_name] = shim


_register_submodule()
