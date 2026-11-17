import pandas as pd
import re

def extract_number(cuota):
    if not cuota or pd.isna(cuota):
        return None
    match = re.search(r"\d+", str(cuota))
    return int(match.group()) if match else None

def parse_value(val):
    if pd.isna(val) or str(val).strip() == "":
        return 0.0

    val = str(val).replace("$", "").replace(" ", "").replace("´", "").strip()

    # Handle European-style numbers: remove thousand separators, fix decimal comma
    # Example: "1.280,22" → "1280.22"
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)  # remove dots only if they are thousand separators
    val = val.replace(",", ".")

    # Handle negatives in parentheses
    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")

    try:
        return float(val)
    except ValueError:
        return 0.0

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

# Load Excel
df = pd.read_excel("Preliquidaciones Unificadas.xlsx")

# Normalize Cliente
df["Cliente_norm"] = (
    df["Cliente"].str.lower()
    .str.normalize("NFKD")
    .str.encode("ascii", errors="ignore")
    .str.decode("utf-8")
    .str.strip()
)

# Normalize Cuota (blank → 1)
df["Cuota_norm"] = df["Cuota"].fillna("1").astype(str)

# Parse values
df["PrimaCom_val"] = df["Prima Com."].apply(parse_value)
df["ValorFactura_val"] = df["Valor Factura"].apply(parse_value)

# Aggregate by Poliza + Cuota
agg = (
    df.groupby(["Poliza", "Cuota_norm", "Cliente_norm", "Agente"], as_index=False)
    .agg({"PrimaCom_val": "sum", "ValorFactura_val": "sum"})
)

# Classify status
def classify(val):
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"

agg["Status"] = agg["PrimaCom_val"].apply(classify)

# Detect missing cuotas per Poliza
missing_records = []
for poliza, group in agg.groupby("Poliza"):
    cuotas = group["Cuota_norm"].unique()
    expected = []

    for c in cuotas:
        if "/" in c:
            try:
                _, total = c.split("/")
                expected = [str(i) for i in range(1, int(total) + 1)]
                break
            except ValueError:
                continue

    if not expected:
        numbers = [extract_number(c) for c in cuotas]
        numbers = [n for n in numbers if n is not None]
        if numbers:
            max_cuota = max(numbers)
            expected = [str(i) for i in range(1, max_cuota + 1)]
        else:
            expected = ["1"]

    missing = [c for c in expected if c not in cuotas]
    for m in missing:
        missing_records.append({"Poliza": poliza, "Missing Cuota": m})

missing_df = pd.DataFrame(missing_records)

# Save detailed list with two sheets
with pd.ExcelWriter("unpaid_policies_list.xlsx") as writer:
    agg.to_excel(writer, sheet_name="Cuota Status", index=False)
    if not missing_df.empty:
        missing_df.to_excel(writer, sheet_name="Missing Cuotas", index=False)

# Build summary at policy level
policy_status = agg.groupby("Poliza")["Status"].apply(
    lambda x: "Unpaid" if ("Neutral" in x.values or "Anomaly" in x.values) else "Paid"
).reset_index()

summary_cliente = df.merge(policy_status, on="Poliza").groupby("Cliente_norm")["Status"].value_counts().unstack(fill_value=0)
summary_agente = df.merge(policy_status, on="Poliza").groupby("Agente")["Status"].value_counts().unstack(fill_value=0)
summary_global = policy_status["Status"].value_counts()

# Save summary
with pd.ExcelWriter("unpaid_policies_summary.xlsx") as writer:
    summary_cliente.to_excel(writer, sheet_name="By Cliente")
    summary_agente.to_excel(writer, sheet_name="By Agente")
    summary_global.to_excel(writer, sheet_name="Global Totals")
