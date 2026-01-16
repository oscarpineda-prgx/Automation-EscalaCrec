import pandas as pd

from src.config.db import get_db_connection
from src.extract.repository import fetch_vendor_range, fetch_vendor

QUERY_SPECS = {
    "ap": ("ap.sql", "range"),
    "compras": ("compras.sql", "range"),
    "devoluciones": ("devoluciones.sql", "range"),
    "convenios_sap": ("convenios_sap.sql", "vendor"),
    "convenios_legados_c": ("convenios_legados_c.sql", "vendor"),
    "convenios_legados_s": ("convenios_legados_s.sql", "vendor"),
}

def get_data(vendor: str, fecha_ini: str, fecha_fin: str) -> dict[str, pd.DataFrame]:
    """
    Extrae info para un proveedor. Reutiliza una sola conexión para mejorar performance.
    """
    out: dict[str, pd.DataFrame] = {}

    with get_db_connection() as conn:
        for key, (sql_file, mode) in QUERY_SPECS.items():
            if mode == "range":
                out[key] = fetch_vendor_range(sql_file, vendor, fecha_ini, fecha_fin, conn=conn)
            else:
                out[key] = fetch_vendor(sql_file, vendor, conn=conn)

    return out