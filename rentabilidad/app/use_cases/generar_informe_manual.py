from __future__ import annotations

from ..dto import GenerarInformeRequest, GenerarInformeResponse
from .generar_informe_automatico import run as run_auto


def run(req: GenerarInformeRequest, bus) -> GenerarInformeResponse:
    return run_auto(req, bus)
