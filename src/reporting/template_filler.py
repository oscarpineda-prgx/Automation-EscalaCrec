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
PORC_INPUT_PATH = Path(__file__).resolve().parents[2] / "input" / "Porcentajes.xlsx"


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
    porcentajes_df: pd.DataFrame | None = None,
) -> Path:
    """
    Carga la plantilla, rellena Base Compra, Devoluciones y AP por año/mes y guarda una copia.
    Solo escribe en columnas E/H (compras), F/I (devoluciones), N (AP) y L (porcentaje de descuento).
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
    if porcentajes_df is not None and vendor_number:
        _apply_porcentajes(sheet, porcentajes_df, vendor_number)

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
    Solo toca las columnas Base Compra (E/H), Devoluciones (F/I), Descuento aplicado (N) y porcentaje de descuento (L); deja fórmulas intactas.
    """
    with get_db_connection() as conn:
        compras_df = fetch_vendor_range("compras.sql", vendor, fecha_ini, fecha_fin, conn=conn)
        devoluciones_df = fetch_vendor_range("devoluciones.sql", vendor, fecha_ini, fecha_fin, conn=conn)
        ap_df = _fetch_ap(conn, vendor, fecha_ini, fecha_fin)
        vendor_df = fetch_vendor("vendor_master.sql", vendor, conn=conn)
        vendor_name = None
        if not vendor_df.empty and "Nombre Proveedor" in vendor_df.columns:
            vendor_name = str(vendor_df.iloc[0]["Nombre Proveedor"])

    porcentajes_df = None
    if PORC_INPUT_PATH.exists():
        porcentajes_df = pd.read_excel(PORC_INPUT_PATH)

    return fill_template_with_compras(
        compras_df,
        output_path=output_path,
        vendor_number=vendor,
        vendor_name=vendor_name,
        devoluciones_df=devoluciones_df,
        ap_df=ap_df,
        porcentajes_df=porcentajes_df,
    )


def _first_nonempty(series: pd.Series) -> Optional[str]:
    """Devuelve el primer valor no vacío de la serie (como texto), limpiando sufijos típicos de Excel."""
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
    """
    Obtiene la llave de búsqueda para AP:
    1) Prioriza FOLIOCONVENIO de convenios_legados_s.
    2) Si no hay, usa convenios_legados_c.
    3) Última opción: acuerdo desde convenios_sap.
    """
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


def _apply_porcentajes(sheet, porcentajes_df: pd.DataFrame, vendor: str) -> None:
    """
    Aplica Porc_Descontar según rangos del archivo Porcentajes.xlsx.
    - Elige la versión más alta del convenio para el proveedor.
    - Si los límites son <=100 se interpreta como porcentaje de crecimiento (columna K).
      Si son mayores, se interpreta como monto (usamos Total Base Compra actual, col G).
    - Escribe el porcentaje en col L (decimal).
    """
    df = porcentajes_df
    if df is None or df.empty:
        return

    df = df[df["Num_Prov"].astype(str) == str(vendor)]
    if df.empty:
        return

    # Tomar la última versión disponible
    max_ver = df["Id_Num_Ver"].max()
    df = df[df["Id_Num_Ver"] == max_ver]
    # Si hay múltiples convenios, priorizamos el de mayor Id_Num_Conv
    if "Id_Num_Conv" in df.columns:
        max_conv = df["Id_Num_Conv"].max()
        df = df[df["Id_Num_Conv"] == max_conv]

    # Normalizamos límites numéricos
    df = df.copy()
    df["Imp_LimInf"] = pd.to_numeric(df["Imp_LimInf"], errors="coerce")
    df["Imp_LimSup"] = pd.to_numeric(df["Imp_LimSup"], errors="coerce")
    df["Porc_Descontar"] = pd.to_numeric(df["Porc_Descontar"], errors="coerce")

    ranges = df[["Imp_LimInf", "Imp_LimSup", "Porc_Descontar"]].dropna()
    if ranges.empty:
        return

    def is_percent(row):
        return (row["Imp_LimSup"] or 0) <= 100 and (row["Imp_LimInf"] or 0) <= 100

    for block in YEAR_BLOCKS:
        start = block["start_row"]
        for offset in range(0, 12):
            row_idx = start + offset
            # Tomamos valores base compra y devoluciones para recomputar total y crecimiento,
            # porque openpyxl no evalúa las fórmulas ya presentes en la hoja.
            e_val = sheet[f"{CURRENT_COMPRAS_COL}{row_idx}"].value
            f_val = sheet[f"{CURRENT_DEVOL_COL}{row_idx}"].value
            h_val = sheet[f"{PREVIOUS_COMPRAS_COL}{row_idx}"].value
            i_val = sheet[f"{PREVIOUS_DEVOL_COL}{row_idx}"].value

            def to_float(v):
                if isinstance(v, str) and v.strip() in {"-", ""}:
                    return 0.0
                try:
                    return float(v)
                except (TypeError, ValueError):
                    return None

            e = to_float(e_val)
            f = to_float(f_val)
            h = to_float(h_val)
            i_prev = to_float(i_val)

            f_num = f if f is not None else 0.0
            i_num = i_prev if i_prev is not None else 0.0

            g_val = e - f_num if e is not None else None
            j_val = h - i_num if h is not None else None

            # Fallback: si no pudimos recomputar, intentar leer G/J directamente
            if g_val is None:
                g_val = to_float(sheet[f"G{row_idx}"].value)
            if j_val is None:
                j_val = to_float(sheet[f"J{row_idx}"].value)

            growth_pct = None
            if g_val is not None and j_val not in (None, 0):
                growth_pct = ((g_val / j_val) * 100) - 100

            for _, r in ranges.iterrows():
                use_percent = is_percent(r)
                target = growth_pct if use_percent else g_val
                if target is None:
                    continue
                low = r["Imp_LimInf"]
                high = r["Imp_LimSup"]
                if low is None or high is None:
                    continue
                if low <= target <= high:
                    porc = r["Porc_Descontar"]
                    if porc is None:
                        continue
                    sheet[f"L{row_idx}"].value = porc / 100  # porcentaje en decimal
                    break
