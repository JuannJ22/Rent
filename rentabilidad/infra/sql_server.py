from __future__ import annotations

from dataclasses import dataclass
from contextlib import contextmanager
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
    encrypt: bool = True
    trust_server_certificate: bool = False
    timeout: int = 30

    def connection_string(self) -> str:
        def quote(value: str) -> str:
            return "{" + value.replace("}", "}}") + "}"

        parts = [
            f"DRIVER={quote(self.driver)}",
            f"SERVER={quote(self.server)}",
            f"DATABASE={quote(self.database)}",
        ]
        if self.trusted_connection:
            parts.append("Trusted_Connection=yes")
        else:
            if self.user is not None:
                parts.append(f"UID={quote(self.user)}")
            if self.password is not None:
                parts.append(f"PWD={quote(self.password)}")
        parts.append(f"Encrypt={'yes' if self.encrypt else 'no'}")
        parts.append(f"TrustServerCertificate={'yes' if self.trust_server_certificate else 'no'}")
        return ";".join(parts)


@contextmanager
def dataframe_reader(config: SqlServerConfig):
    """Reutiliza una conexión y cierra explícitamente sus recursos."""
    import pyodbc

    conn = pyodbc.connect(config.connection_string(), timeout=config.timeout)
    try:
        conn.timeout = config.timeout

        def read(query: str, params=None) -> pd.DataFrame:
            cursor = conn.cursor()
            try:
                cursor.execute(query, *([] if params is None else params))
                columns = [column[0] for column in cursor.description]
                return pd.DataFrame.from_records(
                    [tuple(row) for row in cursor.fetchall()], columns=columns
                )
            finally:
                cursor.close()

        yield read
    finally:
        conn.close()


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
    with dataframe_reader(config) as read:
        return read(query, params=params)


def normalize_sql_flag(value: str | None) -> bool:
    if value is None:
        return False
    return value.strip().lower() in {"1", "true", "yes", "y", "si", "sí"}


def normalize_sql_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]
