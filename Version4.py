import pandas as pd
import re

# -----------------------------
# Helpers – Monetary parsing
# -----------------------------

def parse_value(val):
    """Convert Excel string with currency/parenthesis into float, robust to EU/US formats."""
    if pd.isna(val) or str(val).strip() == "":
        return 0.0

    val = str(val)
    val = val.replace("$", "").replace(" ", "").replace("´", "").strip()

    # Remove thousand separators like 1.234,56 → 1234,56
    val = re.sub(r"\.(?=\d{3}(?:,|$))", "", val)
    val = val.replace(",", ".")

    if val.startswith("(") and val.endswith(")"):
        val = "-" + val.strip("()")

    try:
        return float(val)
    except ValueError:
        return 0.0


# -----------------------------
# Helpers – Cuota parsing
# -----------------------------

def normalize_cuota_str(cuota_str: str) -> str:
    """Normalize cuota strings to simplify parsing."""
    s = str(cuota_str).lower()
    s = s.replace("´", "")
    s = s.replace(",", " ")
    s = s.replace("y", " ")
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def extract_denominator(cuota_str):
    """Extract denominator from cuota strings like '3 y 4 /12'."""
    if pd.isna(cuota_str):
        return None

    s = normalize_cuota_str(cuota_str)
    m = re.search(r"/\s*(\d+)", s)
    return int(m.group(1)) if m else None


def extract_cuota_numbers(cuota_str):
    """
    Extract ONLY numerator cuotas.
    Examples:
    '3 y 4 /12' -> ['3', '4']
    '12/12'     -> ['12']
    '1,2,3/12'  -> ['1','2','3']
    """
    if pd.isna(cuota_str):
        return ["1"]

    s = normalize_cuota_str(cuota_str)

    # Remove denominator completely so digits don't leak
    s = re.sub(r"/\s*\d+", "", s)

    nums = re.findall(r"\b\d+\b", s)
    return [str(int(n)) for n in nums] if nums else ["1"]


def extract_cuotas_with_validation(cuota_str):
    """Extract cuota numerators and denominator, validating ranges."""
    nums = extract_cuota_numbers(cuota_str)
    den = extract_denominator(cuota_str)

    if den:
        nums = [n for n in nums if 1 <= int(n) <= den]

    return nums, den


# -----------------------------
# Helpers – Classification
# -----------------------------

def classify(val):
    if val > 0:
        return "Paid"
    elif val == 0:
        return "Neutral"
    else:
        return "Anomaly"


def classify_policy_strict(statuses):
    if ("Neutral" in statuses.values) or ("Anomaly" in statuses.values):
        return "Unpaid"
    return "Paid"


def classify_policy_all_paid(statuses):
    return "Paid" if (statuses == "Paid").all() else "Unpaid"


# -----------------------------
# Helpers – Policy matching
# -----------------------------

def is_subsequence(needle: str, haystack: str) -> bool:
    """Linear subsequence match."""
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

def normalize_policy(pol):
    if pd.isna(pol):
        return ""

    return (
        str(pol)
        .lower()
        .strip()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

# -----------------------------
# Load data
# -----------------------------

df = pd.read_excel("Preliquidaciones Unificadas.xlsx", sheet_name="BaseUnificada")
df_admin = pd.read_excel("Preliquidaciones Unificadas.xlsx", sheet_name="BaseAdminSeg")
df_base = pd.read_excel("BaseCompletaEmiRenoCance.xlsx")
df_cancel = pd.read_excel("cancelreport.xlsx")

# -----------------------------
# Company decision rule
# -----------------------------

company_lookup = dict(
    zip(
        df_admin["Poliza"].astype(str).str.strip(),
        df_admin["COMPANY"].astype(str).str.lower()
    )
)


def decide_cuota(row):
    pol = str(row["Poliza"]).strip()
    company = company_lookup.get(pol)

    if company and ("bmi" in company or "bupa" in company):
        return extract_cuotas_with_validation(row["Cuota"])

    # Default: single payment
    return ["1"], 1


df[["Cuota_list", "Cuota_den"]] = df.apply(
    lambda r: pd.Series(decide_cuota(r)), axis=1
)

# -----------------------------
# Normalize fields
# -----------------------------

df["Cliente_norm"] = (
    df["Cliente"]
    .astype(str)
    .str.lower()
    .str.normalize("NFKD")
    .str.encode("ascii", errors="ignore")
    .str.decode("utf-8")
    .str.strip()
)

df["PrimaCom_val"] = df["Prima Com."].apply(parse_value)
df["ValorFactura_val"] = df["Valor Factura"].apply(parse_value)
# Example: parse Prima Neta like other monetary fields
df["PrimaNeta_val"] = df[" Prima Neta"].apply(parse_value)

df_exploded = df.explode("Cuota_list").copy()
df_exploded["Cuota_norm"] = df_exploded["Cuota_list"]


# -----------------------------
# Explode to cuota-level rows
# -----------------------------

df_exploded = df.explode("Cuota_list").copy()
df_exploded["Cuota_norm"] = df_exploded["Cuota_list"]

# -----------------------------
# Cuota-level aggregation
# -----------------------------
agg = (
    df_exploded
    .groupby(["Poliza", "Cuota_norm", "Cliente_norm", "Agente", "PrimaNeta_val"], as_index=False)
    .agg({
        "PrimaCom_val": "sum",
        "ValorFactura_val": "sum",
        "PrimaNeta_val": "sum"
    })
)

agg["Status"] = agg["PrimaCom_val"].apply(classify)

# -----------------------------
# Step 2: Cancel Report Matching
# -----------------------------
cancel_policies = df_cancel["POLIZA"].apply(normalize_policy).tolist()

# Normalize Poliza in agg and df
agg["Poliza_norm"] = agg["Poliza"].apply(normalize_policy)
df["Poliza_norm"] = df["Poliza"].apply(normalize_policy)

# Filter out canceled policies
agg = agg[~agg["Poliza_norm"].isin(cancel_policies)]
df = df[~df["Poliza_norm"].isin(cancel_policies)]

# -----------------------------
# Missing cuotas detection (new logic)
# -----------------------------
missing_records = []

for poliza, group in agg.groupby("Poliza"):
    # Count positive payments
    positive_count = (group["PrimaCom_val"] > 0).sum()
    positive_count2 = (group["PrimaNeta_val"] > 0).sum()

    # Identify company
    company = company_lookup.get(str(poliza).strip(), "").lower()

    if "bmi" in company or "bupa" in company:
        if positive_count == 1 or positive_count2 == 1:
            # Single payment policy, no missing
            continue
        elif positive_count == 2 or positive_count2 == 2:
            # Semestral policy, no missing
            continue
        elif positive_count == 3 or positive_count2 == 3:
            # Threemesterly policy, no missing
            continue
        elif positive_count == 4 or positive_count2 == 4:
            # Quatrymesterly policy, no missing
            continue        
        elif positive_count <= 12 or positive_count2 <= 12:
            missing_records.append({
                "Poliza": poliza,
                "MissingCuotasCount": 12 - positive_count,
                "Agente": group["Agente"].iloc[0] if "Agente" in group.columns else None
            })
        # If positive_count == 12 → fully paid, no missing
    else:
        # Other companies: always single payment, no missing
        continue

missing_df = pd.DataFrame(missing_records)


# -----------------------------
# Policy-level classifications
# -----------------------------

policy_status_strict = (
    agg.groupby("Poliza")["Status"]
    .apply(classify_policy_strict)
    .reset_index(name="Status")
)

policy_status_allpaid = (
    agg.groupby("Poliza")["Status"]
    .apply(classify_policy_all_paid)
    .reset_index(name="Status")
)

totals_strict = (
    policy_status_strict["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

totals_allpaid = (
    policy_status_allpaid["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

# -----------------------------
# Payments Global summary
# -----------------------------

payments_global = (
    agg["Status"]
    .value_counts()
    .rename_axis("Status")
    .reset_index(name="Count")
)

payments_global["Section"] = "Payments Global"

sep_row = pd.DataFrame([{"Status": "", "Count": "", "Section": ""}])

strict_rows = totals_strict.assign(Section="Policy Totals (Strict)")
allpaid_rows = totals_allpaid.assign(Section="Policy Totals (AllPaid)")

payments_global_appended = pd.concat(
    [payments_global, sep_row, strict_rows, allpaid_rows],
    ignore_index=True
)

# -----------------------------
# Match logic
# -----------------------------
df_base["POLIZA_norm"] = df_base["POLIZA"].apply(normalize_policy)
df_base = df_base[~df_base["POLIZA_norm"].isin(cancel_policies)]  # exclude canceled

df_unif = df[["Poliza"]].drop_duplicates()
df_unif["Poliza_norm"] = df_unif["Poliza"].apply(normalize_policy)

unif_policies = df_unif["Poliza_norm"].tolist()


def match_policy(row):
    needle = row["POLIZA_norm"]

    for candidate in unif_policies:
        if needle in candidate:
            return pd.Series({
                "MatchedPolicyInUnificada": candidate,
                "FinalStatus": "Paid"
            })

    return pd.Series({
        "MatchedPolicyInUnificada": "",
        "FinalStatus": "Unpaid"
    })


match_results = df_base.copy()
match_results[["MatchedPolicyInUnificada", "FinalStatus"]] = (
    match_results.apply(match_policy, axis=1)
)

# -----------------------------
# Excel output
# -----------------------------

with pd.ExcelWriter("consolidated_unpaid_report.xlsx") as writer:
    agg.to_excel(writer, sheet_name="Cuota Status", index=False)

    if missing_df.empty:
        pd.DataFrame(
            [{"Info": "No missing cuotas detected"}]
        ).to_excel(writer, sheet_name="Missing Cuotas", index=False)
    else:
        missing_df.to_excel(writer, sheet_name="Missing Cuotas", index=False)

    payments_global_appended.to_excel(
        writer, sheet_name="Payments Global", index=False
    )

    match_results.to_excel(
        writer, sheet_name="Match Results", index=False
    )