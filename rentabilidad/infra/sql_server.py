from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence

import pandas as pd


@dataclass(frozen=True)
class SqlServerConfig:
    server: str
    database: str
    user: str | None = None
    password: str | None = None
    driver: str = "ODBC Driver 17 for SQL Server"
    trusted_connection: bool = False
    encrypt: bool = False
    trust_server_certificate: bool = True
    timeout: int = 30

    def connection_string(self) -> str:
        parts = [
            f"DRIVER={{{self.driver}}}",
            f"SERVER={self.server}",
            f"DATABASE={self.database}",
        ]
        if self.trusted_connection:
            parts.append("Trusted_Connection=yes")
        else:
            if self.user is not None:
                parts.append(f"UID={self.user}")
            if self.password is not None:
                parts.append(f"PWD={self.password}")
        if self.encrypt:
            parts.append("Encrypt=yes")
        if self.trust_server_certificate:
            parts.append("TrustServerCertificate=yes")
        return ";".join(parts)


def fetch_dataframe(
    config: SqlServerConfig, query: str, params: Sequence[object] | None = None
) -> pd.DataFrame:
    try:
        import pyodbc
    except ModuleNotFoundError as exc:
        message = (
            "No se encontró el módulo 'pyodbc'. Instálalo con 'pip install pyodbc' "
            "o con 'pip install -r requirements.txt' antes de ejecutar el GUI."
        )
        raise ModuleNotFoundError(message) from exc
    with pyodbc.connect(config.connection_string(), timeout=config.timeout) as conn:
        return pd.read_sql_query(query, conn, params=params)


def normalize_sql_flag(value: str | None) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"1", "true", "yes", "y", "si", "sí"}


def normalize_sql_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]
