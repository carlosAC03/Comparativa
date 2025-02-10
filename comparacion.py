import pandas as pd

# Cargar los archivos Excel
file_vivess = "VIVESS.xlsx"
file_regular = "Documents_List_Regula.xlsx"

# Leer los archivos en DataFrames
df_vivess = pd.read_excel(file_vivess)
df_regular = pd.read_excel(file_regular)

# Normalizar los datos para facilitar la comparación
df_vivess["supportedDocuments"] = df_vivess["supportedDocuments"].astype(str).str.strip()
df_regular["Document"] = df_regular["Document"].astype(str).str.strip()
df_regular["Year"] = pd.to_numeric(df_regular["Year"], errors='coerce')

# Crear una nueva copia del DataFrame de VIVESS para modificarlo
df_vivess_modificado = df_vivess.copy()

# Fusionar los DataFrames basándonos en el tipo de documento, país y año
df_merged = df_vivess_modificado.merge(
    df_regular,
    left_on=["supportedDocuments", "countryCode"],
    right_on=["Document", "ICAO country code"],
    how="left"
)

# Verificar si el documento tiene soporte por regular basándose en el tipo, país y fecha de emisión
df_merged["Supported by Regular"] = df_merged.apply(
    lambda row: True if pd.notna(row["Year"]) and str(row["Year"]) in str(row["supportedDocuments"]) else False,
    axis=1
)

# Seleccionar las columnas necesarias para el nuevo archivo
df_vivess_modificado = df_merged[df_vivess.columns.tolist() + ["Supported by Regular"]]

# Guardar el nuevo DataFrame en un archivo Excel
output_file = "VIVESS_modificado.xlsx"
df_vivess_modificado.to_excel(output_file, index=False)

# Devolver el nombre del archivo generado
print(f"Archivo generado: {output_file}")
