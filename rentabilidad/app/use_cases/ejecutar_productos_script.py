"""Generación del listado de productos desde la interfaz gráfica.

Este caso de uso no ejecuta un ``.bat`` directamente. Delega en el servicio de
aplicación para mantener la lógica de generación, espera, validación y limpieza
en un único flujo reutilizable por CLI y GUI.
"""

from __future__ import annotations

from pathlib import Path

from rentabilidad.core.dates import DateResolver, TodayStrategy

from ..dto import GenerarInformeResponse
from ...config import settings


def run(bus) -> GenerarInformeResponse:
    """Genera el listado de productos y publica eventos para la GUI."""

    try:
        resolver = DateResolver(TodayStrategy())
        objetivo = resolver.resolve(None)
        bus.publish("log", f"Generando listado de productos para {objetivo:%Y-%m-%d}")

        service = settings.build_product_service()
        ruta: Path = service.generate(objetivo)
    except Exception as exc:  # pragma: no cover - depende de SIIGO/Windows/archivos externos
        mensaje = f"No se pudo generar el listado de productos: {exc}"
        bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    mensaje = f"Listado de productos generado: {ruta}"
    bus.publish("done", mensaje)
    return GenerarInformeResponse(ok=True, mensaje=mensaje, ruta_salida=str(ruta))
