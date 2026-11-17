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
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)  # remove thousand separators
    val = val.replace(",", ".")
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
    """Extract cuota numerators from strings like '3 y 4 /12' or '´9/12'."""
    if pd.isna(cuota_str):
        return ["1"]
    s = str(cuota_str).lower()
    s = s.replace("´", "").replace("y", " ").replace("/", " ").replace(",", " ")
    # Tighten: only capture numbers that are standalone tokens to avoid stray digits
    numbers = re.findall(r"\b\d+\b", s)
    return [str(int(n)) for n in numbers] if numbers else ["1"]

def extract_denominator(cuota_str):
    """Extract denominator from cuota strings like '9/12' or '3 y 4 /12'."""
    if pd.isna(cuota_str):
        return None
    s = str(cuota_str).lower().replace("´", "")
    m = re.search(r"/\s*(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None

def classify_policy_strict(statuses):
    """Strict policy classification: any Neutral/Anomaly → Unpaid, else Paid."""
    if ("Neutral" in statuses.values) or ("Anomaly" in statuses.values):
        return "Unpaid"
    return "Paid"

def classify_policy_all_paid(statuses):
    """All-paid classification: only Paid if all cuotas are Paid."""
    total = len(statuses)
    paid_count = (statuses == "Paid").sum()
    return "Paid" if paid_count == total else "Unpaid"

def is_subsequence(needle: str, haystack: str) -> bool:
    """
    Linear subsequence match: returns True if all characters of `needle`
    appear in order within `haystack` (contiguous or not).
    Example: needle='GM21Q16958', haystack='TRANSGM21Q16958' → True.
    """
    needle = str(needle)
    haystack = str(haystack)
    if not needle or not haystack:
        return False
    i = 0
    for ch in haystack:
        if i < len(needle) and ch == needle[i]:
            i += 1
            if i == len(needle):
                return True
    return False


# -----------------------------
# Load data
# -----------------------------

# Main transactional data
df = pd.read_excel("Preliquidaciones Unificadas.xlsx", sheet_name="BaseUnificada")

# Admin sheet for company lookup (decision flag only)
df_admin = pd.read_excel("Preliquidaciones Unificadas.xlsx", sheet_name="BaseAdminSeg")

# New base for match logic
df_base = pd.read_excel("BaseCompletaEmiRenoCance.xlsx")  # columns: POLIZA, ASEGURADO, AGENTE, COMPANY, PRODUCTO, PLAN

# -----------------------------
# Company decision rule (BMI/Bupa keep multi-cuota; others single-pay; not found → single-pay)
# -----------------------------

# Build lookup dict: policy → company
company_lookup = dict(zip(df_admin["Poliza"].astype(str).str.strip(), df_admin["COMPANY"].astype(str)))

def decide_cuota(row):
    pol = str(row["Poliza"]).strip()
    company = company_lookup.get(pol, None)
    if company is None:
        # Not found → single pay
        return ["1"], 1
    company_lower = company.lower()
    if "bmi" in company_lower or "bupa" in company_lower:
        # BMI/Bupa → keep original cuota logic
        return extract_cuota_numbers(row["Cuota"]), extract_denominator(row["Cuota"])
    else:
        # Other company → single pay
        return ["1"], 1

df[["Cuota_list", "Cuota_den"]] = df.apply(lambda r: pd.Series(decide_cuota(r)), axis=1)

# -----------------------------
# Normalize fields
# -----------------------------

df["Cliente_norm"] = (
    df["Cliente"].astype(str).str.lower()
    .str.normalize("NFKD")
    .str.encode("ascii", errors="ignore")
    .str.decode("utf-8")
    .str.strip()
)

df["PrimaCom_val"] = df["Prima Com."].apply(parse_value)
df["ValorFactura_val"] = df["Valor Factura"].apply(parse_value)

# -----------------------------
# Explode to cuota-level rows
# -----------------------------

df_exploded = df.explode("Cuota_list").copy()
df_exploded["Cuota_norm"] = df_exploded["Cuota_list"]

# -----------------------------
# Cuota-level aggregation
# -----------------------------

agg = (
    df_exploded.groupby(["Poliza", "Cuota_norm", "Cliente_norm", "Agente"], as_index=False)
    .agg({"PrimaCom_val": "sum", "ValorFactura_val": "sum"})
)
agg["Status"] = agg["PrimaCom_val"].apply(classify)

# -----------------------------
# Missing cuotas detection
# -----------------------------

missing_records = []
for poliza, group in agg.groupby("Poliza"):
    cuotas = group["Cuota_norm"].unique()
    denoms = df.loc[df["Poliza"] == poliza, "Cuota_den"].dropna().astype(int)
    if len(denoms) > 0:
        total_expected = int(denoms.max())
    else:
        nums = [int(c) for c in cuotas if str(c).isdigit()]
        total_expected = max(nums) if nums else 1
    expected = [str(i) for i in range(1, total_expected + 1)]
    missing = [c for c in expected if c not in cuotas]

    # Get agent for this policy (first non-null value)
    agent = group["Agente"].iloc[0] if "Agente" in group.columns else None

    for m in missing:
        missing_records.append({
            "Poliza": poliza,
            "Missing Cuota": m,
            "Agente": agent
        })

missing_df = pd.DataFrame(missing_records)

# -----------------------------
# Policy-level classification (for totals appended to Payments Global)
# -----------------------------

policy_status_strict = agg.groupby("Poliza")["Status"].apply(classify_policy_strict).reset_index(name="Status")
policy_status_allpaid = agg.groupby("Poliza")["Status"].apply(classify_policy_all_paid).reset_index(name="Status")

# Totals (counts of policies by status perspective)
totals_strict = policy_status_strict["Status"].value_counts().rename_axis("Status").reset_index(name="Count")
totals_allpaid = policy_status_allpaid["Status"].value_counts().rename_axis("Status").reset_index(name="Count")

# -----------------------------
# Payments Global summary (transactional)
# -----------------------------

payments_global = agg["Status"].value_counts().rename_axis("Status").reset_index(name="Count")

# Append totals at bottom as extra rows (with a label column)
payments_global_appended = payments_global.copy()
payments_global_appended["Section"] = "Payments Global"

# Add a separator row
sep_row = pd.DataFrame([{"Status": "", "Count": "", "Section": ""}])

# Append strict totals
strict_rows = totals_strict.copy()
strict_rows["Section"] = "Policy Totals (Strict)"
payments_global_appended = pd.concat([payments_global_appended, sep_row, strict_rows], ignore_index=True)

# Append all-paid totals
allpaid_rows = totals_allpaid.copy()
allpaid_rows["Section"] = "Policy Totals (AllPaid)"
payments_global_appended = pd.concat([payments_global_appended, allpaid_rows], ignore_index=True)

# -----------------------------
# Match logic: BaseCompletaEmiRenoCance vs BaseUnificada
# -----------------------------

# Normalize policy strings
df_base["POLIZA_norm"] = df_base["POLIZA"].astype(str).str.strip()
df_unif = df[["Poliza"]].drop_duplicates().copy()
df_unif["Poliza_norm"] = df_unif["Poliza"].astype(str).str.strip()

# Build fast lookup list
unif_policies = df_unif["Poliza_norm"].tolist()

def match_policy(row):
    needle = str(row["POLIZA_norm"]).strip()
    matched_policy = None
    for candidate in unif_policies:
        if is_subsequence(needle, candidate):
            matched_policy = candidate
            break

    if matched_policy is None:
        return pd.Series({
            "MatchedPolicyInUnificada": "",
            "FinalStatus": "Unpaid"
        })
    else:
        return pd.Series({
            "MatchedPolicyInUnificada": matched_policy,
            "FinalStatus": "Paid"
        })

# Apply row-by-row
match_results = df_base.copy()
match_results[["MatchedPolicyInUnificada", "FinalStatus"]] = match_results.apply(match_policy, axis=1)


# -----------------------------
# Single Excel output with minimal sheets
# -----------------------------

with pd.ExcelWriter("consolidated_unpaid_report.xlsx") as writer:
    # 1) Cuota Status
    agg.to_excel(writer, sheet_name="Cuota Status", index=False)

    # 2) Missing Cuotas
    if not missing_df.empty:
        missing_df.to_excel(writer, sheet_name="Missing Cuotas", index=False)
    else:
        pd.DataFrame([{"Info": "No missing cuotas detected"}]).to_excel(writer, sheet_name="Missing Cuotas", index=False)

    # 3) Payments Global (with totals appended at bottom)
    payments_global_appended.to_excel(writer, sheet_name="Payments Global", index=False)

    # 4) Match Results
    match_results.to_excel(writer, sheet_name="Match Results", index=False)
