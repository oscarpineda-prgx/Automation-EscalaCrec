from __future__ import annotations
from pathlib import Path
from typing import Dict, Tuple, Optional
import openpyxl
import pandas as pd
from src.config.db import get_db_connection
from src.extract.repository import fetch_vendor_range, fetch_vendor, fetch_query

# Rutas de plantilla para el archivo de salida y directorio de output.
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "EscalaCrec.xlsx"
OUTPUT_DIR = Path(__file__).resolve().parents[2] / "output"

# Definimos los bloques año actual y año anterior y la fila inicial (columnas fijas).
YEAR_BLOCKS = [
    {"current": 2020, "previous": 2019, "start_row": 8},
    {"current": 2021, "previous": 2020, "start_row": 25},
    {"current": 2022, "previous": 2021, "start_row": 42},
    {"current": 2023, "previous": 2022, "start_row": 59},
    {"current": 2024, "previous": 2023, "start_row": 76},
]
CURRENT_COMPRAS_COL = "E"   # Base compra año actual
PREVIOUS_COMPRAS_COL = "H"  # Base compra año anterior
CURRENT_DEVOL_COL = "F"     # Devoluciones año actual
PREVIOUS_DEVOL_COL = "I"    # Devoluciones año anterior
CURRENT_AP_COL = "N"        # Descuento aplicado (AP) año actual


def _build_lookup(df: pd.DataFrame, value_column: str) -> Dict[Tuple[int, int], float]:
    """Convierte el DF en un dict {(anio, mes): valor} usando la columna indicada."""
    if df is None or df.empty or value_column not in df.columns:
        return {}

    cleaned = df.copy()
    cleaned = cleaned[["anio", "mes", value_column]].dropna()
    cleaned["anio"] = cleaned["anio"].astype(int)
    cleaned["mes"] = cleaned["mes"].astype(int)

    lookup: Dict[Tuple[int, int], float] = {}
    for row in cleaned.itertuples(index=False):
        lookup[(int(row.anio), int(row.mes))] = float(getattr(row, value_column))
    return lookup


def _write_year_values(sheet, lookup: Dict[Tuple[int, int], float], year: int, start_row: int, col: str) -> None:
    """Escribe los 12 meses en las filas contiguas empezando en start_row."""
    for month in range(1, 13):
        value = lookup.get((year, month))
        if value is None:
            continue
        cell = f"{col}{start_row + (month - 1)}"
        sheet[cell].value = value


def fill_template_with_compras(
    compras_df: pd.DataFrame,
    template_path: Path = TEMPLATE_PATH,
    output_path: Path | None = None,
    vendor_number: str | None = None,
    vendor_name: str | None = None,
    devoluciones_df: pd.DataFrame | None = None,
    ap_df: pd.DataFrame | None = None,
) -> Path:
    """
    Carga la plantilla, rellena Base Compra, Devoluciones y AP por año/mes y guarda una copia.
    Solo escribe en columnas E/H (compras), F/I (devoluciones) y N (AP).
    """
    compras_lookup = _build_lookup(compras_df, "compra_neta_mas_impuestos")
    devol_lookup = _build_lookup(devoluciones_df, "devoluciones_mas_impuestos") if devoluciones_df is not None else {}
    ap_lookup = _build_lookup(ap_df, "GrsInvAmt") if ap_df is not None else {}

    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active

    if vendor_number:
        sheet["B8"].value = str(vendor_number)
        title_value = f"{vendor_number} {vendor_name}".strip() if vendor_name else str(vendor_number)
        sheet["B3"].value = title_value

    for block in YEAR_BLOCKS:
        _write_year_values(sheet, compras_lookup, block["current"], block["start_row"], CURRENT_COMPRAS_COL)
        _write_year_values(sheet, compras_lookup, block["previous"], block["start_row"], PREVIOUS_COMPRAS_COL)
        if devol_lookup:
            _write_year_values(sheet, devol_lookup, block["current"], block["start_row"], CURRENT_DEVOL_COL)
            _write_year_values(sheet, devol_lookup, block["previous"], block["start_row"], PREVIOUS_DEVOL_COL)
        if ap_lookup:
            _write_year_values(sheet, ap_lookup, block["current"], block["start_row"], CURRENT_AP_COL)

    final_path = output_path or (OUTPUT_DIR / "EscalaCrec_filled.xlsx")
    final_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(final_path)
    return final_path


def generate_escalacrec_report(
    vendor: str,
    fecha_ini: str,
    fecha_fin: str,
    output_path: Path | None = None,
) -> Path:
    """
    Consulta compras, devoluciones y AP y genera la plantilla poblada.
    Solo toca las columnas Base Compra (E/H), Devoluciones (F/I) y Descuento aplicado (N); deja fórmulas intactas.
    """
    with get_db_connection() as conn:
        compras_df = fetch_vendor_range("compras.sql", vendor, fecha_ini, fecha_fin, conn=conn)
        devoluciones_df = fetch_vendor_range("devoluciones.sql", vendor, fecha_ini, fecha_fin, conn=conn)
        ap_df = _fetch_ap(conn, vendor, fecha_ini, fecha_fin)
        vendor_df = fetch_vendor("vendor_master.sql", vendor, conn=conn)
        vendor_name = None
        if not vendor_df.empty and "Nombre Proveedor" in vendor_df.columns:
            vendor_name = str(vendor_df.iloc[0]["Nombre Proveedor"])

    return fill_template_with_compras(
        compras_df,
        output_path=output_path,
        vendor_number=vendor,
        vendor_name=vendor_name,
        devoluciones_df=devoluciones_df,
        ap_df=ap_df,
    )


def _first_nonempty(series: pd.Series) -> Optional[str]:
    if series is None:
        return None
    cleaned = series.dropna().astype(str).str.strip()
    if cleaned.empty:
        return None
    val = cleaned.iloc[0]
    if val.endswith(".0"):
        val = val[:-2]
    return val or None


def _resolve_ap_key(conn, vendor: str) -> Optional[str]:
    """Obtiene la llave para AP: FOLIOCONVENIO (legado S/C) o acuerdo (SAP)."""
    for sql_file in ("convenios_legados_s.sql", "convenios_legados_c.sql"):
        df = fetch_vendor(sql_file, vendor, conn=conn)
        if not df.empty and "FOLIOCONVENIO" in df.columns:
            folio = _first_nonempty(df["FOLIOCONVENIO"])
            if folio:
                return folio

    df_sap = fetch_vendor("convenios_sap.sql", vendor, conn=conn)
    if not df_sap.empty and "acuerdo" in df_sap.columns:
        acuerdo = _first_nonempty(df_sap["acuerdo"])
        if acuerdo:
            return acuerdo
    return None


def _fetch_ap(conn, vendor: str, fecha_ini: str, fecha_fin: str) -> pd.DataFrame:
    """Consulta AP usando la llave derivada de convenios; retorna DF vacío si no hay llave."""
    key = _resolve_ap_key(conn, vendor)
    if not key:
        return pd.DataFrame()
    like_param = f"%{key}%"
    return fetch_query("ap.sql", [vendor, fecha_ini, fecha_fin, like_param], conn=conn)
