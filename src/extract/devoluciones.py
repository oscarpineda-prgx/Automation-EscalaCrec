import pandas as pd
from pathlib import Path
from src.config.db import get_db_connection


QUERY_PATH = Path(__file__).resolve().parents[1] / "queries" / "devoluciones.sql"

def get_devoluciones(vendor: str, fecha_ini: str, fecha_fin:str) -> pd.Dataframe:
    """Extrae las devoluciones de un proveedor en un rango de fechas dado."""
    with open(QUERY_PATH, "r", encoding="utf-8") as file:
        query = file.read()

    with get_db_connection() as conn:
        df = pd.read_sql(query, conn, param=[vendor, fecha_ini, fecha_fin])
        return df