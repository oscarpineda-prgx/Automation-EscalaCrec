from pathlib import Path
from typing import Any, Sequence, Optional

import pandas as pd
import pyodbc

from src.config.db import get_db_connection

QUERIES_DIR = Path(__file__).resolve().parents[1] / "queries"


def _load_query(query_filename: str) -> str:
    query_path = QUERIES_DIR / query_filename
    with open(query_path, "r", encoding="utf-8") as file:
        return file.read()


def fetch_query(
    query_filename: str,
    params: Sequence[Any],
    conn: Optional["pyodbc.Connection"] = None,
) -> pd.DataFrame:
    """Ejecuta un archivo SQL con parámetros y regresa un DataFrame."""
    query = _load_query(query_filename)

    if conn is not None:
        return pd.read_sql(query, conn, params=params)

    with get_db_connection() as new_conn:
        return pd.read_sql(query, new_conn, params=params)
    

def fetch_vendor_range(
    query_filename: str,
    vendor: str,
    fecha_ini: str,
    fecha_fin: str,
    conn: Optional["pyodbc.Connection"] = None,
) -> pd.DataFrame:
    return fetch_query(query_filename, [vendor, fecha_ini, fecha_fin], conn=conn)


def fetch_vendor(
    query_filename: str,
    vendor: str,
    conn: Optional["pyodbc.Connection"] = None,
) -> pd.DataFrame:
    return fetch_query(query_filename, [vendor], conn=conn)
