import pandas as pd
import re

# -----------------------------
# Helpers
# -----------------------------

def parse_value(val):
    """Convert Excel string with currency/parenthesis into float, robust to EU/US formats."""
    if pd.isna(val) or str(val).strip() == "":
        return 0.0
    val = str(val).replace("$", "").replace(" ", "").replace("´", "").strip()
    # Remove thousand separators (dots before 3 digits) and fix decimal comma
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)
    val = val.replace(",", ".")
    # Handle negatives in parentheses
    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")
    try:
        return float(val)
    except ValueError:
        return 0.0

def classify(val):
    """Classify cuota status based on PrimaCom value."""
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"

def extract_cuota_numbers(cuota_str):
    """
    Extract all cuota numbers from a string like '3 y 4 /12' or '´9/12'.
    Returns a list of normalized cuota numbers as strings (numerators only).
    """
    if pd.isna(cuota_str):
        return ["1"]
    s = str(cuota_str).lower()
    s = s.replace("´", "").replace("y", " ").replace("/", " ").replace(",", " ")
    numbers = re.findall(r"\d+", s)
    return [str(int(n)) for n in numbers] if numbers else ["1"]

def extract_denominator(cuota_str):
    """
    Extract denominator from cuota strings like '9/12', '3 y 4 /12'.
    Returns int denominator if found, else None.
    """
    if pd.isna(cuota_str):
        return None
    s = str(cuota_str).lower().replace("´", "")
    # Find patterns with slash and a trailing number as denominator
    m = re.search(r"/\s*(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None

def classify_policy_strict(statuses):
    """
    Strict policy classification:
    - If any cuota is Neutral or Anomaly → Unpaid
    - Otherwise → Paid
    """
    if ("Neutral" in statuses.values) or ("Anomaly" in statuses.values):
        return "Unpaid"
    return "Paid"

def classify_policy_all_paid(statuses):
    """
    All-paid policy classification:
    - Policy is Paid only if all cuotas are Paid
    - Otherwise → Unpaid
    """
    total = len(statuses)
    paid_count = (statuses == "Paid").sum()
    return "Paid" if paid_count == total else "Unpaid"

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
# Extract numerators list and a denominator per row (if present)
df["Cuota_list"] = df["Cuota"].apply(extract_cuota_numbers)
df["Cuota_den"] = df["Cuota"].apply(extract_denominator)

df_exploded = df.explode("Cuota_list").copy()
df_exploded["Cuota_norm"] = df_exploded["Cuota_list"]  # normalized numerator as string

# -----------------------------
# Cuota-level aggregation (include Cliente/Agente for consistent payment summaries)
# -----------------------------
agg = (
    df_exploded.groupby(["Poliza", "Cuota_norm", "Cliente_norm", "Agente"], as_index=False)
    .agg({"PrimaCom_val": "sum", "ValorFactura_val": "sum"})
)

agg["Status"] = agg["PrimaCom_val"].apply(classify)

# -----------------------------
# Detect missing cuotas (use denominator if present, else max numerator)
# -----------------------------
missing_records = []
for poliza, group in agg.groupby("Poliza"):
    cuotas = group["Cuota_norm"].unique()

    # Determine expected total:
    # 1) If any denominator exists for this policy, use the max denominator found
    # 2) Else, use max observed numerator
    denoms = df.loc[df["Poliza"] == poliza, "Cuota_den"].dropna().astype(int)
    if len(denoms) > 0:
        total_expected = int(denoms.max())
    else:
        nums = [int(c) for c in cuotas if c.isdigit()]
        total_expected = max(nums) if nums else 1

    expected = [str(i) for i in range(1, total_expected + 1)]
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
# Policy-level classification (two perspectives)
# -----------------------------
# Strict: any Neutral/Anomaly → Unpaid
policy_status_strict = agg.groupby("Poliza")["Status"].apply(classify_policy_strict).reset_index(name="Status")

# All-paid: only Paid if all cuotas are Paid
policy_status_allpaid = agg.groupby("Poliza")["Status"].apply(classify_policy_all_paid).reset_index(name="Status")

# Attach Cliente and Agente once per policy (representative row)
policy_info = df.groupby("Poliza")[["Cliente_norm", "Agente"]].first().reset_index()

policy_status_strict = policy_status_strict.merge(policy_info, on="Poliza", how="left")
policy_status_allpaid = policy_status_allpaid.merge(policy_info, on="Poliza", how="left")

# -----------------------------
# Summaries
# -----------------------------
# 1) Policy-level (strict)
summary_global_policies_strict = policy_status_strict["Status"].value_counts()
summary_cliente_policies_strict = (
    policy_status_strict.groupby("Cliente_norm")["Status"].value_counts().unstack(fill_value=0)
)
summary_agente_policies_strict = (
    policy_status_strict.groupby("Agente")["Status"].value_counts().unstack(fill_value=0)
)

# 2) Policy-level (all-paid)
summary_global_policies_allpaid = policy_status_allpaid["Status"].value_counts()
summary_cliente_policies_allpaid = (
    policy_status_allpaid.groupby("Cliente_norm")["Status"].value_counts().unstack(fill_value=0)
)
summary_agente_policies_allpaid = (
    policy_status_allpaid.groupby("Agente")["Status"].value_counts().unstack(fill_value=0)
)

# 3) Payment-level (transactional, from agg which already has Status)
summary_global_payments = agg["Status"].value_counts()
summary_cliente_payments = (
    agg.groupby("Cliente_norm")["Status"].value_counts().unstack(fill_value=0)
)
summary_agente_payments = (
    agg.groupby("Agente")["Status"].value_counts().unstack(fill_value=0)
)

# -----------------------------
# Save summaries
# -----------------------------
with pd.ExcelWriter("unpaid_policies_summary.xlsx") as writer:
    # Policy-level (strict)
    summary_cliente_policies_strict.to_excel(writer, sheet_name="Policies By Cliente (Strict)")
    summary_agente_policies_strict.to_excel(writer, sheet_name="Policies By Agente (Strict)")
    summary_global_policies_strict.to_excel(writer, sheet_name="Policies Global (Strict)")

    # Policy-level (all-paid)
    summary_cliente_policies_allpaid.to_excel(writer, sheet_name="Policies By Cliente (AllPaid)")
    summary_agente_policies_allpaid.to_excel(writer, sheet_name="Policies By Agente (AllPaid)")
    summary_global_policies_allpaid.to_excel(writer, sheet_name="Policies Global (AllPaid)")

    # Payment-level (transactional)
    summary_cliente_payments.to_excel(writer, sheet_name="Payments By Cliente")
    summary_agente_payments.to_excel(writer, sheet_name="Payments By Agente")
    summary_global_payments.to_excel(writer, sheet_name="Payments Global")
