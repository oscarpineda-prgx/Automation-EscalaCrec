import pyodbc
import pandas as pd
import os


def load_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    print("Excel cargado:", df.shape)
    return df


def unique_ids(df: pd.DataFrame, column: str) -> list[str]:
    values = df[column].dropna().astype(str).unique().tolist()
    print("Valores unicos para consulta SQL:", len(values))
    return values


def connect_sql_server() -> pyodbc.Connection:
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER=ATL20AF2222SQ19;"
        "DATABASE=SORIANA_MX_2024_PROD_F;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)


def normalize_denominacion_externa(value: str):
    """
    Limpia la denominacion externa quitando el sufijo desde el primer guion y los ceros iniciales.
    Ejemplo: '094449-01-S' -> '94449'
    """
    if pd.isna(value):
        return value
    text = str(value).strip()
    base = text.split("-", 1)[0]
    if base.startswith("0") and len(base) > 1:
        base = base.lstrip("0") or "0"
    return base


def apply_denominacion_normalization(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detecta la columna de denominacion externa (cualquiera de sus variantes) y aplica la limpieza.
    """
    candidate_cols = [
        "denominación_externa",
        "denominacion_externa",
        "denominaciA3n_externa",
        "denominaciÃ³n_externa",
    ]
    for col in candidate_cols:
        if col in df.columns:
            df[col] = df[col].apply(normalize_denominacion_externa)
            if col != "denominación_externa" and "denominación_externa" not in df.columns:
                df.rename(columns={col: "denominación_externa"}, inplace=True)
            break
    return df


def invalid_fechacancel_mask(df: pd.DataFrame) -> pd.Series:
    """
    Marca filas con FECHACANCELACION vacia/nula o anterior a FECINI.
    """
    if "FECHACANCELACION" not in df.columns:
        return pd.Series(False, index=df.index)
    fecha_cancel = pd.to_datetime(df["FECHACANCELACION"], errors="coerce")
    fecini = pd.to_datetime(df["FECINI"], errors="coerce") if "FECINI" in df.columns else pd.Series(False, index=df.index)
    earlier_than_start = fecha_cancel < fecini if "FECINI" in df.columns else pd.Series(False, index=df.index)
    return fecha_cancel.isna() | earlier_than_start


def fetch_legado(conn: pyodbc.Connection, ids: list[str], batch_size: int = 500) -> pd.DataFrame:
    resultados = []
    tabla_legado = "dbo.Convenios_Legado_CityClub"
    cols_legado = [
        "ID_NUM_CONV",
        "ID_NUM_VER",
        "ID_NUM_PLAZOPAGO",
        "DESC_CONVEVTO",
        "FECINI",
        "FECFIN",
        "ID_NUM_PROV",
        "DESC_CONVSTAT_NVO",
        "DESC_CONVEVTO_NVO",
        "FECHACANCELACION"
    ]
    cols_legado_select = ", ".join(cols_legado)

    for i in range(0, len(ids), batch_size):
        subset = ids[i: i + batch_size]
        placeholders = ",".join("?" for _ in subset)
        query_legado = f"""
            SELECT 
                CONCAT(CAST(ID_NUM_CONV AS varchar(50)), CAST(ID_NUM_VER AS varchar(50))) AS id,
                {cols_legado_select}
            FROM {tabla_legado}
            WHERE CONCAT(CAST(ID_NUM_CONV AS varchar(50)), CAST(ID_NUM_VER AS varchar(50))) IN ({placeholders})
        """
        print(f"Batch LEGADO {i // batch_size + 1}")
        df_tmp = pd.read_sql(query_legado, conn, params=subset)
        resultados.append(df_tmp)

    df = pd.concat(resultados, ignore_index=True) if resultados else pd.DataFrame()
    print("Filas recuperadas desde LEGADO:", len(df))
    return df


def merge_legado(df_excel: pd.DataFrame, df_legado: pd.DataFrame) -> pd.DataFrame:
    df_excel["Id"] = df_excel["Id"].astype(str)
    df_legado["id"] = df_legado["id"].astype(str)
    df_merged = df_excel.merge(
        df_legado,
        left_on="Id",
        right_on="id",
        how="left",
        suffixes=("", "_legado"),
    )
    df_merged["origen"] = df_merged["ID_NUM_CONV"].notna().map({True: "Legado", False: None})
    print("Filas con origen 'Legado':", df_merged["origen"].eq("Legado").sum())
    return df_merged


def fetch_kona(conn: pyodbc.Connection, faltantes: list[str]) -> pd.DataFrame:
    tabla_kona = "dbo.Catalogo_Convenios_SAP_KONA"
    cols_kona = [
        "acuerdo",
        "clase_acuerdo",
        "descripción_del_acuerdo",
        "no_proveedor",
        "estatus",
        "fecha_modificación_cancelación",
        "denominación_externa"
    ]
    cols_kona_select = ", ".join(cols_kona)
    query_kona = f"""
        SELECT {cols_kona_select}
        FROM {tabla_kona}
        WHERE acuerdo IN ({",".join("?" for _ in faltantes)})
    """
    df = pd.read_sql(query_kona, conn, params=faltantes)
    df = apply_denominacion_normalization(df)
    print("Filas recuperadas desde SAP_KONA:", len(df))
    return df


def fetch_kona_by_denominacion(conn: pyodbc.Connection, convenios: list[str], batch_size: int = 500) -> pd.DataFrame:
    """
    Consulta SAP KONA usando denominacion_externa como clave, con un LIKE para cubrir sufijos/prefijos.
    """
    convenios_limpios = [normalize_denominacion_externa(c) for c in convenios if pd.notna(c)]
    convenios_limpios = [c for c in convenios_limpios if c]
    if not convenios_limpios:
        return pd.DataFrame()

    resultados = []
    tabla_kona = "dbo.Catalogo_Convenios_SAP_KONA"
    cols_kona = [
        "acuerdo",
        "clase_acuerdo",
        "descripción_del_acuerdo",
        "no_proveedor",
        "estatus",
        "fecha_modificación_cancelación",
        "denominación_externa"
    ]
    cols_kona_select = ", ".join(cols_kona)

    for i in range(0, len(convenios_limpios), batch_size):
        subset = convenios_limpios[i: i + batch_size]
        placeholders = " OR ".join("denominación_externa LIKE ?" for _ in subset)
        query_kona = f"""
            SELECT {cols_kona_select}
            FROM {tabla_kona}
            WHERE {placeholders}
        """
        patterns = [f"%{c}%" for c in subset]
        print(f"Batch SAP_KONA (denominacion) {i // batch_size + 1}")
        df_tmp = pd.read_sql(query_kona, conn, params=patterns)
        resultados.append(df_tmp)

    df = pd.concat(resultados, ignore_index=True) if resultados else pd.DataFrame()
    df = apply_denominacion_normalization(df)
    print("Filas recuperadas desde SAP_KONA por denominacion:", len(df))
    return df


def merge_kona(df_sin_legado: pd.DataFrame, df_kona: pd.DataFrame) -> pd.DataFrame:
    df_sin_legado["Id_Num_Conv"] = df_sin_legado["Id_Num_Conv"].astype(str)
    df_kona["acuerdo"] = df_kona["acuerdo"].astype(str)

    df_kona_merge = df_sin_legado.merge(
        df_kona,
        left_on="Id_Num_Conv",
        right_on="acuerdo",
        how="left",
        suffixes=("", "_kona"),
    )
    df_kona_merge["origen"] = df_kona_merge["acuerdo"].notna().map({True: "Sap Kona", False: None})
    print("Filas con origen 'Sap Kona':", df_kona_merge["origen"].eq("Sap Kona").sum())
    return df_kona_merge


def main():
    excel_path = "input/dbo_C_ccConvEscImpPorc_CityClub_20200812.xlsx"
    df_excel = load_excel(excel_path)
    values = unique_ids(df_excel, "Id")

    conn = connect_sql_server()

    df_sql_legado = fetch_legado(conn, values, batch_size=500)
    df_merged = merge_legado(df_excel, df_sql_legado)

    invalid_cancel_mask = invalid_fechacancel_mask(df_merged) & df_merged["origen"].eq("Legado")
    df_legado_valid = df_merged[df_merged["origen"].eq("Legado") & ~invalid_cancel_mask]
    df_legado_invalid = df_merged[invalid_cancel_mask].copy()
    df_sin_legado = df_merged[df_merged["origen"].isna()].copy()
    print("Filas sin cruce en Legado:", df_sin_legado.shape[0])
    print("Filas Legado con FECHACANCELACION invalida:", df_legado_invalid.shape[0])

    faltantes_conv = df_sin_legado["Id_Num_Conv"].dropna().unique().tolist()
    print("Registros para consultar en SAP_KONA:", len(faltantes_conv))

    df_sql_kona = fetch_kona(conn, faltantes_conv)
    df_kona_merge = merge_kona(df_sin_legado, df_sql_kona)

    df_kona_den_merge = pd.DataFrame()
    if not df_legado_invalid.empty:
        convenios_den = df_legado_invalid["Id_Num_Conv"].dropna().tolist()
        df_sql_kona_den = fetch_kona_by_denominacion(conn, convenios_den)
        if not df_sql_kona_den.empty:
            df_sql_kona_den["den_norm"] = df_sql_kona_den["denominación_externa"].astype(str)
            df_sql_kona_den["den_norm"] = df_sql_kona_den["den_norm"].apply(lambda x: normalize_denominacion_externa(x))

            df_legado_invalid["conv_norm"] = df_legado_invalid["Id_Num_Conv"].apply(normalize_denominacion_externa)

            df_kona_den_merge = df_legado_invalid.merge(
                df_sql_kona_den,
                left_on="conv_norm",
                right_on="den_norm",
                how="left",
                suffixes=("", "_kona"),
            )
            df_kona_den_merge["origen"] = df_kona_den_merge["den_norm"].notna().map({True: "Sap Kona", False: None})
            df_kona_den_merge.drop(columns=["conv_norm", "den_norm"], inplace=True)

    df_legado_final = df_legado_valid
    df_kona_final = df_kona_merge[df_kona_merge["origen"] == "Sap Kona"]
    df_kona_den_final = df_kona_den_merge[df_kona_den_merge["origen"] == "Sap Kona"] if not df_kona_den_merge.empty else pd.DataFrame()

    df_total = pd.concat([df_legado_final, df_kona_final, df_kona_den_final], ignore_index=True)

    print("FILAS TOTALES EN EL RESULTADO FINAL:", df_total.shape[0])
    df_total.head()

    # Guardar resultado en Excel
    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", "resultado_final_CityClub08.xlsx")
    df_total.to_excel(output_path, index=False)
    print(f"Archivo guardado en: {output_path}")


if __name__ == "__main__":
    main()
