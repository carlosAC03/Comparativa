import pandas as pd


def load_data():
    """Carga los datos de los archivos Excel."""
    print("Cargando datos de los archivos Excel...")
    xls1 = pd.ExcelFile("Documents_List_Regula.xlsx")
    xls2 = pd.ExcelFile("VIVESS.xlsx")
    df1 = xls1.parse(xls1.sheet_names[0])
    df2 = xls2.parse(xls2.sheet_names[0])
    print("Datos cargados correctamente.\n")
    return df1, df2


def filter_relevant_documents(df1, df2):
    """Filtra los documentos Passport, IdCard y DriverLicense en ambos datasets."""
    print("Filtrando documentos relevantes (Passport, IdCard, DriverLicense)...")
    df1_filtered = df1[df1["Document"].str.contains("Passport|IdCard|Driving License", na=False, case=False)]
    df2_filtered = df2[df2["kind"].isin(["Passport", "IdCard", "DriverLicense"])]
    print("Filtrado completado.\n")
    return df1_filtered, df2_filtered


def compare_countries(df1, df2):
    """Compara los países presentes en ambos archivos."""
    print("Comparando países en ambos archivos...")
    countries_df1 = set(df1["Country / territory"].dropna().unique())
    countries_df2 = set(df2["countryCode"].dropna().unique())
    result = {
        "Países en ambos archivos": list(countries_df1.intersection(countries_df2)),
        "Países solo en Documents_List_Regula": list(countries_df1 - countries_df2),
        "Países solo en VIVESS": list(countries_df2 - countries_df1)
    }
    print("Comparación de países completada.\n")
    return result


def normalize_document_name(name):
    """Normaliza los nombres de documentos en Documents_List_Regula.xlsx."""
    name = name.lower()
    if "passport" in name:
        return "Passport"
    elif "driving license" in name:
        return "DriverLicense"
    elif "id card" in name or "identity card" in name:
        return "IdCard"
    return name


def compare_document_counts(df1, df2):
    """Compara la cantidad de registros por tipo de documento."""
    print("Comparando número de documentos por tipo...")
    df1.loc[:, "Normalized Document"] = df1["Document"].apply(normalize_document_name)
    count_df1 = df1["Normalized Document"].value_counts()
    count_df2 = df2["kind"].value_counts()
   
    result = {
        doc: {
            "Total en Documents_List_Regula": count_df1.get(doc, 0),
            "Total en VIVESS": count_df2.get(doc, 0)
        }
        for doc in ["Passport", "IdCard", "DriverLicense"]
    }
    print("Comparación de documentos completada.\n")
    return result


if __name__ == "__main__":
    print("Iniciando proceso de comparación...\n")
    df1, df2 = load_data()
    df1_filtered, df2_filtered = filter_relevant_documents(df1, df2)
   
    country_comparison = compare_countries(df1_filtered, df2_filtered)
    document_comparison = compare_document_counts(df1_filtered, df2_filtered)
   
    print("\n--- RESULTADOS DE LA COMPARACIÓN ---")
    print("\nComparación de países:")
    for key, value in country_comparison.items():
        print(f"{key}: {value}")
   
    print("\nComparación de documentos:")
    for doc, counts in document_comparison.items():
        print(f"{doc}: {counts}")
   
    print("\nProceso de comparación finalizado.")
