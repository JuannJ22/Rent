from __future__ import annotations

from rentabilidad.core.dates import DateResolver, TodayStrategy

from ...config import settings


def run(bus):
    try:
        resolver = DateResolver(TodayStrategy())
        objetivo = resolver.resolve(None)
        bus.publish("log", f"Generando listado de productos para {objetivo:%Y-%m-%d}")
        service = settings.build_product_service()
        ruta = service.generate(objetivo)
        bus.publish("done", f"Listado generado: {ruta}")
        return ruta
    except Exception as exc:  # pragma: no cover - depende de entorno externo
        bus.publish("error", str(exc))
        return None
