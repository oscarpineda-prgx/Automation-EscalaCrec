from pathlib import Path
import pandas as pd
from src.config.db import get_db_connection
from src.extract.repository import fetch_vendor_range, fetch_vendor, fetch_query

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
        # Primero obtenemos convenios_legados_s para usar su FOLIOCONVENIO en ap.sql
        convenios_legados_s_df = fetch_vendor("convenios_legados_s.sql", vendor, conn=conn)
        out["convenios_legados_s"] = convenios_legados_s_df

        folio_series = convenios_legados_s_df.get("FOLIOCONVENIO")
        folio_like = "%"
        if folio_series is not None:
            folio_series = folio_series.dropna().astype(str).str.strip()
            if not folio_series.empty:
                folio_value = folio_series.iloc[0]
                # Si viene como numero con sufijo ".0", quitamos esa parte.
                if folio_value.endswith(".0"):
                    folio_value = folio_value[:-2]
                folio_like = f"%{folio_value}%"

        for key, (sql_file, mode) in QUERY_SPECS.items():
            if key == "convenios_legados_s":
                continue  # ya obtenido arriba
            if key == "ap":
                out[key] = fetch_query(sql_file, [vendor, fecha_ini, fecha_fin, folio_like], conn=conn)
            elif mode == "range":
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
