from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import List

import pandas as pd

from src.config.db import get_db_connection
from src.extract.repository import fetch_vendor
from src.reporting.template_filler import generate_escalacrec_report


def _sanitize_filename(text: str) -> str:
    """Limpia el texto para usarlo en nombres de archivo."""
    cleaned = re.sub(r"[^A-Za-z0-9 _.-]", "", text or "")
    cleaned = cleaned.strip().replace(" ", "_")
    return cleaned or "sin_nombre"


def _obtener_nombre_proveedor(vendor: str, conn) -> str:
    """Devuelve el nombre del proveedor consultando vendor_master; vacío si no existe."""
    df = fetch_vendor("vendor_master.sql", vendor, conn=conn)
    if df.empty or "Nombre Proveedor" not in df.columns:
        return ""
    return str(df.iloc[0]["Nombre Proveedor"])


def run_batch_from_excel(
    excel_path: Path,
    output_dir: Path | None = None,
    vendor_col: str = "vendor",
    fecha_ini_col: str = "fecha_ini",
    fecha_fin_col: str = "fecha_fin",
) -> List[Path]:
    """
    Ejecuta la generación para muchos proveedores listados en un Excel.
    Columnas esperadas (por defecto): vendor, fecha_ini, fecha_fin.
    """
    df = pd.read_excel(excel_path)
    # Intentar alias comunes si faltan columnas estándar
    aliases = {
        vendor_col: [vendor_col, "Num_Prov", "proveedor"],
        fecha_ini_col: [fecha_ini_col, "fecha_inicio", "Fecha_ini"],
        fecha_fin_col: [fecha_fin_col, "fecha_final", "Fecha_fin", "fecha_fin"],
    }

    def pick(col_name: str):
        for cand in aliases[col_name]:
            if cand in df.columns:
                return df[cand]
        raise ValueError(f"No se encontró la columna requerida: {aliases[col_name]}")

    vendors = pick(vendor_col).astype(str).str.strip()
    fechas_ini = pick(fecha_ini_col).astype(str).str.strip()
    fechas_fin = pick(fecha_fin_col).astype(str).str.strip()

    if output_dir is None:
        output_dir = Path(__file__).resolve().parents[2] / "output" / "batch"
    output_dir.mkdir(parents=True, exist_ok=True)

    resultados: List[Path] = []
    with get_db_connection() as conn:
        for vendor, fi, ff in zip(vendors, fechas_ini, fechas_fin):
            nombre = _obtener_nombre_proveedor(vendor, conn) or vendor
            filename = f"{vendor}-{_sanitize_filename(nombre)}.xlsx"
            destino = output_dir / filename
            # Genera el reporte individual
            generate_escalacrec_report(vendor, fi, ff, output_path=destino)
            resultados.append(destino)

    return resultados


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Ejecución por lote de Escala de Crecimiento.")
    parser.add_argument("excel_path", type=Path, help="Ruta del Excel con columnas vendor, fecha_ini, fecha_fin")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Carpeta de salida (por defecto output/batch)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    archivos = run_batch_from_excel(args.excel_path, output_dir=args.output_dir)
    print("Archivos generados:")
    for a in archivos:
        print(a)


if __name__ == "__main__":
    main()
