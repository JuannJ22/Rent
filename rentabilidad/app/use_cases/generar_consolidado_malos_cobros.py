from __future__ import annotations

from rentabilidad.app.dto import (
    GenerarConsolidadoMalosCobrosRequest,
    GenerarInformeResponse,
)
from rentabilidad.config import settings


def run(
    req: GenerarConsolidadoMalosCobrosRequest, bus
) -> GenerarInformeResponse:
    mes = (req.mes or "").strip()
    if not mes:
        mensaje = "Debes seleccionar un mes válido."
        if bus:
            bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)

    service = settings.build_monthly_report_service()
    try:
        if bus:
            bus.publish("log", f"Generando consolidado de malos cobros para {mes}…")
        destino = service.generar_malos_cobros(mes, bus)
        mensaje = f"Informe generado: {destino}"
        if bus:
            bus.publish("done", mensaje)
        return GenerarInformeResponse(
            ok=True,
            mensaje="OK",
            ruta_salida=str(destino),
        )
    except FileNotFoundError as exc:
        mensaje = str(exc)
        if bus:
            bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)
    except ValueError as exc:
        mensaje = str(exc)
        if bus:
            bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)
    except Exception as exc:  # pragma: no cover - defensivo
        mensaje = str(exc)
        if bus:
            bus.publish("error", mensaje)
        return GenerarInformeResponse(ok=False, mensaje=mensaje)
