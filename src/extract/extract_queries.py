from pathlib import Path
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

OUTPUT_DIR = Path(__file__).resolve().parents[2] / "output"


def get_data(vendor: str, fecha_ini: str, fecha_fin: str) -> dict[str, pd.DataFrame]:
    """Extrae info para un proveedor reutilizando una sola conexion."""
    out: dict[str, pd.DataFrame] = {}

    with get_db_connection() as conn:
        for key, (sql_file, mode) in QUERY_SPECS.items():
            if mode == "range":
                out[key] = fetch_vendor_range(sql_file, vendor, fecha_ini, fecha_fin, conn=conn)
            else:
                out[key] = fetch_vendor(sql_file, vendor, conn=conn)

    return out


def save_data_to_excel(
    data: dict[str, pd.DataFrame],
    filename: str | Path | None = None,
) -> Path:
    """Guarda cada DataFrame del dict en una hoja de Excel en la carpeta output."""
    output_path = Path(filename) if filename else OUTPUT_DIR / "data.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, df in data.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return output_path


def get_data_to_excel(
    vendor: str,
    fecha_ini: str,
    fecha_fin: str,
    filename: str | Path | None = None,
) -> Path:
    """Ejecuta todas las consultas y guarda los resultados en Excel."""
    data = get_data(vendor, fecha_ini, fecha_fin)
    return save_data_to_excel(data, filename)
