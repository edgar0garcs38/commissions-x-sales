import pandas as pd
import re

# -----------------------------
# Helpers
# -----------------------------

def parse_value(val):
    if pd.isna(val) or str(val).strip() == "":
        return 0.0
    val = str(val).replace("$", "").replace(" ", "").replace("´", "").strip()
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)
    val = val.replace(",", ".")
    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")
    try:
        return float(val)
    except ValueError:
        return 0.0

def classify(val):
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"

def extract_cuota_numbers(cuota_str):
    """
    Extract all cuota numbers from a string like '3 y 4 /12' or '´9/12'.
    Returns a list of normalized cuota numbers as strings.
    """
    if pd.isna(cuota_str):
        return ["1"]
    cuota_str = str(cuota_str).lower()
    cuota_str = cuota_str.replace("´", "").replace("y", " ").replace("/", " ").replace(",", " ")
    numbers = re.findall(r"\d+", cuota_str)
    return [str(int(n)) for n in numbers] if numbers else ["1"]

# -----------------------------
# Load and normalize
# -----------------------------

df = pd.read_excel("Preliquidaciones Unificadas.xlsx")

df["Cliente_norm"] = (
    df["Cliente"].str.lower()
    .str.normalize("NFKD")
    .str.encode("ascii", errors="ignore")
    .str.decode("utf-8")
    .str.strip()
)

df["PrimaCom_val"] = df["Prima Com."].apply(parse_value)
df["ValorFactura_val"] = df["Valor Factura"].apply(parse_value)

# -----------------------------
# Normalize Cuota and explode
# -----------------------------
df["Cuota_list"] = df["Cuota"].apply(extract_cuota_numbers)
df_exploded = df.explode("Cuota_list").copy()
df_exploded["Cuota_norm"] = df_exploded["Cuota_list"]

# -----------------------------
# Cuota-level aggregation
# -----------------------------
agg = (
    df_exploded.groupby(["Poliza", "Cuota_norm"], as_index=False)
    .agg({"PrimaCom_val": "sum", "ValorFactura_val": "sum"})
)

agg["Status"] = agg["PrimaCom_val"].apply(classify)

# -----------------------------
# Detect missing cuotas
# -----------------------------
missing_records = []
for poliza, group in agg.groupby("Poliza"):
    cuotas = group["Cuota_norm"].unique()
    expected = []

    # Use max cuota number as expected total
    numbers = [int(c) for c in cuotas if c.isdigit()]
    if numbers:
        max_cuota = max(numbers)
        expected = [str(i) for i in range(1, max_cuota + 1)]
    else:
        expected = ["1"]

    missing = [c for c in expected if c not in cuotas]
    for m in missing:
        missing_records.append({"Poliza": poliza, "Missing Cuota": m})

missing_df = pd.DataFrame(missing_records)

# -----------------------------
# Save detailed list
# -----------------------------
with pd.ExcelWriter("unpaid_policies_list.xlsx") as writer:
    agg.to_excel(writer, sheet_name="Cuota Status", index=False)
    if not missing_df.empty:
        missing_df.to_excel(writer, sheet_name="Missing Cuotas", index=False)

# -----------------------------
# Policy-level classification
# -----------------------------
policy_status = agg.groupby("Poliza")["Status"].apply(
    lambda x: "Unpaid" if ("Neutral" in x.values or "Anomaly" in x.values) else "Paid"
).reset_index()

# Attach Cliente and Agente once per policy
policy_info = df.groupby("Poliza")[["Cliente_norm", "Agente"]].first().reset_index()
policy_status = policy_status.merge(policy_info, on="Poliza", how="left")

# -----------------------------
# Summaries
# -----------------------------
summary_global = policy_status["Status"].value_counts()

summary_cliente = (
    policy_status.groupby("Cliente_norm")["Status"].value_counts().unstack(fill_value=0)
)

summary_agente = (
    policy_status.groupby("Agente")["Status"].value_counts().unstack(fill_value=0)
)

# -----------------------------
# Save summaries
# -----------------------------
with pd.ExcelWriter("unpaid_policies_summary.xlsx") as writer:
    summary_cliente.to_excel(writer, sheet_name="By Cliente")
    summary_agente.to_excel(writer, sheet_name="By Agente")
    summary_global.to_excel(writer, sheet_name="Global Totals")
