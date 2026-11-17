import pandas as pd
import glob

# List of your Excel file names
excel_files = glob.glob('Otros/Nuevas Bases/Archivos/*.xlsx')

# Read and merge all Excel files
dfs = [pd.read_excel(file) for file in excel_files]
merged_df = pd.concat(dfs, ignore_index=True)
merged_df.to_excel('CarteraUnificadaEnero-Diciembre.xlsx', index=False)
print("Excel files merged successfully into 'CarteraUnificadaEnero-Diciembre.xlsx'")

# Load merged file
mainData = "CarteraUnificadaEnero-Diciembre.xlsx"
df_mainData = pd.read_excel(mainData, sheet_name='Sheet1')

# Ensure DUPLICADO column exists
if "DUPLICADO" not in df_mainData.columns:
    df_mainData["DUPLICADO"] = "NO"

# Mark duplicates
for idx, row in df_mainData.iterrows():
    poliza = row.get("#PÓLIZA", "")
    estado = row.get("ESTADO DE POLIZA", "")
    inicio = row.get("INICIO DE VIGENCIA", "")
    vigencia = row.get("VIGENCIA ACTUAL", "")
    asegurado = row.get("ASEGURADO", "")
    plan = row.get("PLAN", "")
    producto = row.get("PRODUCTO", "")
    fecharenovacion = row.get("FECHA RENOVACIÓN", "")

    matches = df_mainData[df_mainData["#PÓLIZA"] == poliza]
    if len(matches) > 1:
        other_matches = matches.drop(index=idx, errors='ignore')
        duplicate_found = any(
            (other_matches["ESTADO DE POLIZA"] == estado) &
            (other_matches["INICIO DE VIGENCIA"] == inicio) &
            # (other_matches["VIGENCIA ACTUAL"] == vigencia) &
            (other_matches["ASEGURADO"] == asegurado) &
            (other_matches["PLAN"] == plan) &
            (other_matches["PRODUCTO"] == producto) &
            (other_matches["FECHA RENOVACIÓN"] == fecharenovacion)
        )
        df_mainData.at[idx, "DUPLICADO"] = "SI" if duplicate_found else "NO"
    else:
        df_mainData.at[idx, "DUPLICADO"] = "NO"

# Save with DUPLICADO column
df_mainData.to_excel("CarteraDuplicada.xlsx", index=False)
print("Archivo con columna DUPLICADO generado.")

# ✅ Remove duplicates (keep first occurrence)
df_no_duplicates = df_mainData.drop_duplicates(
    subset=["#PÓLIZA", "ESTADO DE POLIZA", "INICIO DE VIGENCIA",
            # "VIGENCIA ACTUAL", 
            "ASEGURADO", "PLAN", "PRODUCTO", "FECHA RENOVACIÓN"],
    keep="first"
)

# Save cleaned file
df_no_duplicates.to_excel("CarteraSinDuplicados.xlsx", index=False)
print("Archivo final sin duplicados generado: 'CarteraSinDuplicados.xlsx'")
