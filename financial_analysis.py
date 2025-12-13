from pathlib import Path
import pandas as pd

# ==========================================
# CONFIGURATION SECTION
# ==========================================
# Define your Excel file name here. It must be in the same folder as this script.
NOMBRE_ARCHIVO = "Data_Financiera.xlsx"

# Name of the worksheet inside the Excel file
NOMBRE_HOJA = "DataBase Combined"

# Define which year you want to analyze
ANIO_ANALISIS = 2025


COL_FECHA = "Date"
COL_MONTO = "Amount"
COL_CATEGORIA_PRINCIPAL = "Grandparent"
COL_SUB_CATEGORIA = "Parent"
COL_CLASE = "Class"
COL_FUENTE = "Source"

# Specific filters
FILTRO_CATEGORIA_PRINCIPAL = "Income Statement" # Main (top-level) category to filter
FILTRO_ZOOM = "2 COGS"                          # Specific category (zoom-in)



def procesar_datos():
    # 1. Locate the file
    base_dir = Path(__file__).resolve().parent
    ruta_archivo = base_dir / NOMBRE_ARCHIVO

    print(f"--- Starting process ---")
    print(f"Looking for file at: {ruta_archivo}")

    if not ruta_archivo.exists():
        print(f"ERROR: I couldn't find the file '{NOMBRE_ARCHIVO}'. Please check that it's in the folder.")
        return

    # 2. Load data
    try:
        print("Loading Excel... this may take a bit if the file is large.")
        df = pd.read_excel(ruta_archivo, sheet_name=NOMBRE_HOJA)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

 
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], errors="coerce")

   
    print(f"Filtering data for year {ANIO_ANALISIS} and category '{FILTRO_CATEGORIA_PRINCIPAL}'...")
    
    df_filtrado = df[
        (df[COL_CATEGORIA_PRINCIPAL] == FILTRO_CATEGORIA_PRINCIPAL)
        & (df[COL_FECHA].dt.year == ANIO_ANALISIS)
    ].copy()

    if df_filtrado.empty:
        print("Warning: No data found with those filters.")
        return

   
    resumen = (
        df_filtrado
        .groupby([COL_SUB_CATEGORIA, COL_CLASE], dropna=False)[COL_MONTO]
        .sum()
        .reset_index()
        .sort_values([COL_SUB_CATEGORIA, COL_CLASE])
    )

    nombre_csv_resumen = f"reporte_{ANIO_ANALISIS}_resumen_general.csv"
    resumen.to_csv(nombre_csv_resumen, index=False)
    print(f"-> Done! Saved the general summary to: {nombre_csv_resumen}")

 
    print(f"Generating detailed report for: '{FILTRO_ZOOM}'...")
    
    df_zoom = df_filtrado[df_filtrado[COL_SUB_CATEGORIA] == FILTRO_ZOOM].copy()

    if not df_zoom.empty:
        resumen_zoom = (
            df_zoom
            .groupby([COL_FUENTE, COL_CLASE], dropna=False)[COL_MONTO]
            .sum()
            .reset_index()
            .sort_values([COL_FUENTE, COL_CLASE])
        )

        nombre_csv_zoom = f"reporte_{ANIO_ANALISIS}_detalle_{FILTRO_ZOOM.replace(' ', '_')}.csv"
        resumen_zoom.to_csv(nombre_csv_zoom, index=False)
        print(f"-> Done! Saved the specific detail report to: {nombre_csv_zoom}")
    else:
        print(f"No data found for category '{FILTRO_ZOOM}', so that report was not generated.")

if __name__ == "__main__":
    procesar_datos()
