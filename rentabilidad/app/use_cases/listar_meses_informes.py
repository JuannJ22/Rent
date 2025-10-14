from __future__ import annotations

from rentabilidad.config import settings


def run() -> list[str]:
    service = settings.build_monthly_report_service()
    return service.list_months()
