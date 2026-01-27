# Automation Escala de Crecimiento

Herramienta en Python para automatizar la generación de reportes de escalas de crecimiento para proveedores Soriana/CityClub: cruza compras, devoluciones y AP con convenios (legado/SAP), aplica porcentajes por rangos desde `Porcentajes.xlsx`, y produce un Excel listo para entregar usando la plantilla oficial.

## Contenido rápido
- Consulta SQL Server (compras, devoluciones, AP, convenios legado/SAP, vendor master) con una sola conexión.
- Aplica reglas de descuento por proveedor/convenio/versión según rangos en `Porcentajes.xlsx` (monto o crecimiento %).
- Llena la plantilla `src/templates/EscalaCrec.xlsx` respetando fórmulas y formatos.
- Exporta:
  - `output/EscalaCrec_filled.xlsx` (ejecución individual).
  - `output/batch/<vendor>-<NombreProveedor>.xlsx` (ejecución masiva).
  - `output/data.xlsx` (dump opcional de todas las consultas para un proveedor).

## Requisitos
- Python 3.x.
- Dependencias: `pandas`, `openpyxl`, `pyodbc` (Tkinter viene con Python en Windows).
- ODBC Driver 18 para SQL Server y acceso a `ATL20AF2222SQ19` / base `SORIANA_PROJECTS` (config en `src/config/db.py`).

## Entradas (carpeta `input/`)

| Archivo | Uso principal | Columnas clave |
| --- | --- | --- |
| `Porcentajes.xlsx` | Rangos y % de descuento por proveedor/convenio/versión | `ID, Num_Prov, Negocio, origen, Id_Num_Conv, Id_Num_Ver, Id_Cnsc_Esc, Imp_LimInf, Imp_LimSup, Num_Subtipo, Porc_Descontar` |
| Excel de lote (opcional) | Ejecución masiva | `vendor, fecha_ini, fecha_fin` (alias: `Num_Prov`, `fecha_inicio`, `fecha_final`, etc.) |
| `Soriana-Logo.png`, `Prgx-Logo.png` | Logos para la GUI | Imágenes en `input/` |
| `src/templates/EscalaCrec.xlsx` | Plantilla del reporte | Fórmulas y formato final (no está en `input` pero es requerida) |

## Flujo general (alto nivel)

```
template_filler.py     # arma el Excel final; aplica descuentos y llena datos
extract_queries.py     # orquesta queries SQL (compras, devoluciones, AP, convenios)
repository.py          # ejecución de queries reutilizando la conexión
batch_runner.py        # lee Excel de lote y genera un archivo por proveedor
gui_runner.py          # interfaz gráfica ligera para ejecución individual
src/queries/*.sql      # consultas parametrizadas
```

## Cómo ejecutarlo

Ejecución individual (PowerShell desde la raíz):
```powershell
@'
from src.reporting.template_filler import generate_escalacrec_report
out = generate_escalacrec_report("334177", "2020-01-01", "2025-12-31")
print(out)
'@ | python -
```

GUI:
```powershell
python -m src.reporting.gui_runner
```

Ejecución por lote:
```powershell
python -m src.reporting.batch_runner input/proveedores.xlsx
```
(genera archivos en `output/batch/<vendor>-<NombreProveedor>.xlsx`).

Dump de consultas a Excel:
```powershell
@'
from src.extract.extract_queries import get_data_to_excel
out = get_data_to_excel("334177", "2020-01-01", "2025-12-31")
print(out)
'@ | python -
```

## Componentes principales

| Módulo | Rol |
| --- | --- |
| `src/reporting/template_filler.py` | Llena la plantilla, aplica porcentajes, coloca encabezados dinámicos. |
| `src/extract/extract_queries.py` / `src/extract/repository.py` | Ejecutan consultas SQL parametrizadas con una sola conexión. |
| `src/reporting/batch_runner.py` | Proceso masivo desde Excel de entrada; nombra salidas `<vendor>-<NombreProveedor>.xlsx`. |
| `src/reporting/gui_runner.py` | GUI minimalista con selectores de fecha integrados y logos. |
| `src/config/db.py` | Conexión a SQL Server. |
| `src/queries/*.sql` | Consultas de compras, devoluciones, AP, convenios legado/SAP, vendor master. |

## Reglas de negocio clave
- Llave AP: se prioriza `FOLIOCONVENIO` (legados Soriana/City), luego `acuerdo` SAP; se usa `LIKE %llave%`.
- Crecimiento (%): `(E–F)/(H–I) * 100 – 100`; rangos ≤100 son porcentuales, >100 son montos (comparan contra E–F).
- Rangos: última versión (`Id_Num_Ver` máx) y mayor `Id_Num_Conv`; si cae en un rango se escribe `Porc_Descontar/100` en L (M queda con su fórmula).
- Encabezado dinámico: B3 = `<vendor> <NombreProveedor>`, B8 = `<vendor>`.
- Si faltan datos de AP o porcentajes, se dejan celdas vacías sin fallar la ejecución.

## Salidas
- `output/EscalaCrec_filled.xlsx`: reporte individual.
- `output/batch/<vendor>-<NombreProveedor>.xlsx`: reportes masivos.
- `output/data.xlsx`: dump opcional de consultas (si se usa `get_data_to_excel`).

## Solución rápida de problemas
- Sin descuento aplicado: revisar `Porcentajes.xlsx` (proveedor, versión, límites). Si todos los montos están en el primer tramo (0%), L será 0.
- AP vacío: revisar que exista FOLIOCONVENIO/acuerdo en convenios legados/SAP.
- Conexión fallida: confirmar ODBC 18 y acceso a `ATL20AF2222SQ19`.

## Extensiones sugeridas
- Validar esquema de `Porcentajes.xlsx` y del Excel de lote antes de procesar.
- Añadir pruebas unitarias para cálculo de crecimiento y selección de rangos.
- Externalizar configuración de rutas/DB a un YAML si se requieren ambientes múltiples.
