import pandas as pd

# Load the Excel file and sheet
df = pd.read_excel("unpaid_policies_list.xlsx", sheet_name="Missing Cuotas")

# Keep only the first occurrence of each Poliza
df_clean = df.drop_duplicates(subset=["Poliza"], keep="first")

# Save the cleaned data back to Excel (overwrite or new file)
df_clean.to_excel("unpaid_policies_list_clean.xlsx", sheet_name="Missing Cuotas", index=False)

print("Deduplication complete. Cleaned file saved as unpaid_policies_list_clean.xlsx")
