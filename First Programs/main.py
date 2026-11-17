import pandas as pd
import json
import re
import unicodedata
from rapidfuzz import process, fuzz
import numpy as np

# ---------------------
# File paths
# ---------------------
preSettlementsUnified = "Preliquidaciones Unificadas.xlsx"
agentLevels = "REPORTE AGENTES final.xlsx"
specialTables = "niveles.json"   # Special agents tables (by agent name)
normalTables = "niveles1.json"   # Level-based tables (krediseguros nivel 1/2/3)
percentage_of_transactions = "percentage_per_type_of_transaction.json"

# ---------------------
# Text Cleaning Functions
# ---------------------
def remove_accents(text: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(c)
    )

def clean_text(text: str) -> str:
    text = re.sub(r"[()]", "", str(text))
    text = remove_accents(text.strip().upper())
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

# ---------------------
# Load Excel Data
# ---------------------
df_pre_settlements = pd.read_excel(preSettlementsUnified, sheet_name='BaseUnificada')
df_AdminSeg_DB = pd.read_excel(preSettlementsUnified, sheet_name='BaseAdminSeg')
level1 = pd.read_excel(agentLevels, sheet_name='NIVEL 1')
level2 = pd.read_excel(agentLevels, sheet_name='NIVEL 2')
level3 = pd.read_excel(agentLevels, sheet_name='NIVEL 3')
structures = pd.read_excel(agentLevels, sheet_name='ESTRUCTURAS')
specialLeves = pd.read_excel(agentLevels, sheet_name='TABLAS ESPECIALES-NIVEL BR')

# Ensure policy column normalized
df_AdminSeg_DB["#PÓLIZA"] = df_AdminSeg_DB["#PÓLIZA"].astype(str).str.upper()

def normalize_policy(val):
    if pd.isna(val):
        return ""
    return str(val).strip().upper()

df_AdminSeg_DB["#PÓLIZA"] = df_AdminSeg_DB["#PÓLIZA"].apply(normalize_policy)
df_pre_settlements["Poliza"] = df_pre_settlements["Poliza"].apply(normalize_policy)

# ---------------------
# Clean Relevant Columns in pre-settlements
# ---------------------
if "Agente" in df_pre_settlements.columns:
    # Keep original cleaned agent in dataframe for traceability
    df_pre_settlements["Agente"] = df_pre_settlements["Agente"].fillna("").astype(str).apply(clean_text)
if "Prima Com." in df_pre_settlements.columns:
    df_pre_settlements["Prima Com."] = df_pre_settlements["Prima Com."].apply(clean_text)

# ---------------------
# Create Mapping Dictionaries (cleaned plan mapping)
# ---------------------
# Clean PRODUCTO values when building mapping so plan used downstream is normalized
df_policy_company = dict(zip(df_AdminSeg_DB["#PÓLIZA"], df_AdminSeg_DB["COMPANY"].fillna("").astype(str)))
df_policy_plan = dict(zip(df_AdminSeg_DB["#PÓLIZA"], df_AdminSeg_DB["PRODUCTO"].fillna("").astype(str).apply(clean_text)))

# Build lists of (original_name, level) from agent level sheets to avoid cleaned-key collisions
def build_name_level_list(df, name_col="NOMBRES Y APELLIDOS", level_col="NIVEL - TABLA"):
    rows = []
    for _, r in df[[name_col, level_col]].fillna("").iterrows():
        name = str(r[name_col]).strip()
        lvl = r[level_col]
        if name:
            rows.append((name, lvl))
    return rows

level1_list = build_name_level_list(level1)
level2_list = build_name_level_list(level2)
level3_list = build_name_level_list(level3)
structures_list = build_name_level_list(structures)
specialLeves_list = build_name_level_list(specialLeves)

# ---------------------
# Load JSON rules
# ---------------------
with open(percentage_of_transactions, "r", encoding="utf-8") as f:
    commission_rules = json.load(f)
    bmi_rules_clean_map = {clean_text(k): k for k in commission_rules.get("BMI", {})}

with open(specialTables, "r", encoding="utf-8") as f:
    special_tables = json.load(f)  # keys: agent names; nested: companies -> products -> Emision/Renovacion/Cambio BR

with open(normalTables, "r", encoding="utf-8") as f:
    normal_tables = json.load(f)   # keys: "krediseguros nivel 1/2/3"; nested: companies -> products -> Emision/Renovacion/Cambio BR

# ---------------------
# Helper: normalize percentage string from df into xx,xx (and numeric later)
# ---------------------
def normalize_percentage(pct_raw: str) -> str:
    if pct_raw is None or str(pct_raw).strip() == "":
        return ""
    s = str(pct_raw).strip().replace("%", "")
    s = s.replace(".", ",")
    try:
        val = float(s.replace(",", "."))
        if val < 1:  # assume it's a fraction
            val *= 100
        return f"{val:.2f}".replace(".", ",")
    except ValueError:
        return s

# ---------------------
# Helper: Match Agents Names (strict)
# ---------------------
def match_agent_names_strict(agent_input: str, candidate_name: str, fuzzy_threshold: int = 85) -> bool:
    """
    Strict matcher to avoid false positives like 'ALEJANDRA ALAJO' → 'ALEJANDRA CALLE'.

    Rules:
    - If agent_input has 1–2 tokens, require ALL those tokens to be present in candidate.
    - If agent_input has 3+ tokens, require at least 2 overlapping tokens.
    - Fuzzy fallback only allowed if the longest token from agent_input is in candidate.
    """
    a = clean_text(agent_input)
    b = clean_text(candidate_name)

    if not a or not b:
        return False

    tokens_a = a.split()
    tokens_b = b.split()
    set_a = set(tokens_a)
    set_b = set(tokens_b)

    # Short names: strict subset
    if len(tokens_a) <= 2:
        return set_a.issubset(set_b)

    # Longer names: require >=2 overlapping tokens
    overlap = set_a & set_b
    if len(overlap) >= 2:
        return True

    # Guarded fuzzy fallback: longest token must be present
    longest_token_a = max(tokens_a, key=len)
    if longest_token_a not in set_b:
        return False

    score = fuzz.token_set_ratio(a, b)
    return score >= fuzzy_threshold

# ---------------------
# Token-based product/plan matcher
# ---------------------
def match_tokens_strict(source: str, target: str, fuzzy_threshold: int = 85) -> bool:
    """
    Token-based strict matcher for product/plan names.
    - Normalizes both strings with clean_text before tokenizing.
    - If source has 1 token, require that token to appear in target tokens.
    - If source has 2 tokens, require both tokens to appear in target tokens.
    - If source has 3+ tokens, require at least 2 overlapping tokens.
    - Fuzzy fallback allowed only if the longest token from source is present in target tokens.
    """
    s = clean_text(source)
    t = clean_text(target)
    if not s or not t:
        return False

    tokens_s = s.split()
    tokens_t = t.split()
    set_s = set(tokens_s)
    set_t = set(tokens_t)

    # 1 token -> require presence
    if len(tokens_s) == 1:
        return tokens_s[0] in set_t

    # 2 tokens -> require both present
    if len(tokens_s) == 2:
        return set_s.issubset(set_t)

    # 3+ tokens -> require at least 2 overlaps
    overlap = set_s & set_t
    if len(overlap) >= 2:
        return True

    # Guarded fuzzy fallback: longest token must be present
    longest_token = max(tokens_s, key=len)
    if longest_token not in set_t:
        return False

    score = fuzz.token_set_ratio(s, t)
    return score >= fuzzy_threshold

def fuzzy_find_company_and_product(block_dict_or_list, company_candidate: str, plan_candidate: str):
    """
    Handles two block shapes:
    - company-based: { company: { product: txn_map } }
      returns (company_key, product_key)
    - product-based: { product: txn_map } or [product1, product2]
      returns (None, product_key)
    """
    # 1) If block is a list of product names
    if isinstance(block_dict_or_list, list):
        # build a temporary dict mapping product->dummy or empty txn_map
        temp = {p: {} for p in block_dict_or_list}
        prod = find_best_product_in_dict(temp, company_candidate, plan_candidate)
        return None, prod

    # 2) If block is a dict, determine whether it's company->products or product->txn_map
    if isinstance(block_dict_or_list, dict):
        # Inspect first value to guess shape
        sample_vals = list(block_dict_or_list.values())[:3]
        # If values are dicts whose keys look like product names (strings) -> company-based
        company_like = False
        product_like = False
        for v in sample_vals:
            if isinstance(v, dict):
                # if inner dict values are dicts or strings, treat as company->products
                inner_vals = list(v.values())[:1]
                if inner_vals and isinstance(inner_vals[0], dict):
                    company_like = True
                else:
                    # could be product->txn_map (product-> { "Emision": "x", ... })
                    product_like = True
            else:
                product_like = True

        # Heuristic: if keys of block look like company names (contain spaces, words like S.A., etc.)
        # and values are dicts of products, treat as company-based
        if company_like and not product_like:
            # company-based path
            company_keys = list(block_dict_or_list.keys())
            company_clean_keys = [clean_text(k) for k in company_keys]
            company_candidate_clean = clean_text(company_candidate)
            plan_candidate_clean = clean_text(plan_candidate)

            # Try token-based company match first
            for i, ck_clean in enumerate(company_clean_keys):
                if match_tokens_strict(company_candidate_clean, ck_clean, fuzzy_threshold=85):
                    true_company_key = company_keys[i]
                    product_dict = block_dict_or_list.get(true_company_key, {})
                    prod = find_best_product_in_dict(product_dict, company_candidate, plan_candidate)
                    if prod:
                        return true_company_key, prod
                    # continue searching other companies if product not found

            # Fuzzy company fallback
            company_match, company_score, _ = process.extractOne(company_candidate_clean, company_clean_keys, scorer=fuzz.token_set_ratio)
            if company_score < 50:
                alt_match, alt_score, _ = process.extractOne(plan_candidate_clean, company_clean_keys, scorer=fuzz.token_set_ratio)
                if alt_score >= 50:
                    company_match = alt_match
                    company_score = alt_score

            if company_score < 50:
                return None, None

            true_company_key = company_keys[company_clean_keys.index(company_match)]
            product_dict = block_dict_or_list.get(true_company_key, {})
            prod = find_best_product_in_dict(product_dict, company_candidate, plan_candidate)
            return true_company_key, prod

        else:
            # product-based path: keys are products
            prod = find_best_product_in_dict(block_dict_or_list, company_candidate, plan_candidate)
            return None, prod

    # Unknown shape
    return None, None
# ---------------------
# Company + product finder (token-aware)
# ---------------------
def find_best_product_in_dict(product_dict: dict, company_candidate: str, plan_candidate: str):
    """
    Try token-based product matching first, then fuzzy fallback.
    product_dict: { product_name: txn_map } (no company layer)
    Returns product_key or None.
    """
    if not product_dict:
        return None

    product_keys = list(product_dict.keys())
    product_clean_keys = [clean_text(k) for k in product_keys]

    # Prefer plan_candidate, then combined, then company_candidate
    candidates = [plan_candidate, f"{company_candidate} {plan_candidate}", company_candidate]
    best_idx = None
    best_score = -1

    for target in candidates:
        if not target:
            continue
        target_clean = clean_text(target)

        # Token strict pass
        for i, pk_clean in enumerate(product_clean_keys):
            if match_tokens_strict(target_clean, pk_clean, fuzzy_threshold=85):
                return product_keys[i]

        # Keep best fuzzy score
        for i, pk_clean in enumerate(product_clean_keys):
            score = fuzz.token_set_ratio(target_clean, pk_clean)
            if score > best_score:
                best_score = score
                best_idx = i

    if best_idx is not None and best_score >= 80:
        return product_keys[best_idx]

    return None

# ---------------------
# Build cleaned-key maps for JSONs (cleaned -> list(originals))
# ---------------------
def build_clean_key_map_multi(d: dict):
    m = {}
    for k in d.keys():
        ck = clean_text(k)
        m.setdefault(ck, []).append(k)
    return m

special_agent_keys = build_clean_key_map_multi(special_tables)
normal_level_keys = build_clean_key_map_multi(normal_tables)

# ---------------------
# Map agent name to level label (krediseguros nivel 1/2/3)
# ---------------------
def agent_to_level_label(agent_name_raw: str) -> str:
    agent_name_clean = clean_text(agent_name_raw)

    # Search in specialLeves_list, structures_list, level1_list, level2_list, level3_list
    for source_list in [specialLeves_list, structures_list, level1_list, level2_list, level3_list]:
        for candidate_original, lvl in source_list:
            if match_agent_names_strict(agent_name_clean, candidate_original, fuzzy_threshold=85):
                lvl_str = str(lvl).strip()
                if "1" in lvl_str:
                    return "KREDISEGUROS NIVEL 1"
                if "2" in lvl_str:
                    return "KREDISEGUROS NIVEL 2"
                if "3" in lvl_str:
                    return "KREDISEGUROS NIVEL 3"
                return ""
    return ""

# ---------------------
# Classify type of transaction (I/R/CB -> Emision/Renovacion/Cambio BR)
# ---------------------

def normalize_transaction_label(label: str) -> str:
    """
    Normalize various transaction label variants to canonical forms.
    """
    if not label or str(label).startswith("NO ENCONTRADO"):
        return label
    l = clean_text(str(label))
    if "EMISION" in l or l == "I":
        return "Emision"
    if "RENOVACION" in l or "RENOVACIÓN" in l or l == "R":
        return "Renovacion"
    # Any variant that mentions cambio + broker or CB -> Cambio BR
    if "CAMBIO" in l and "BROKER" in l:
        return "Cambio BR"
    if l in ("CB", "CAMBIOBR", "CAMIOBR", "CAMBIOBR.") or "CAMBIO BR" in l:
        return "Cambio BR"
    # If already canonical words
    if l == "CAMBIO BR":
        return "Cambio BR"
    if l == "EMISION":
        return "Emision"
    if l == "RENOVACION":
        return "Renovacion"
    # Default: return original (preserve capitalization)
    return str(label)

def match_plan_tokenwise(plan_candidate: str, plan_key: str, fuzzy_threshold: int = 85) -> bool:
    """
    Token-based plan matcher used for BMI plan keys.
    Reuses the same token rules as match_tokens_strict but exposed for clarity.
    """
    return match_tokens_strict(plan_candidate, plan_key, fuzzy_threshold=fuzzy_threshold)

def find_best_bmi_plan_key(plan_candidate: str, bmi_rules_keys: list):
    plan_candidate_clean = clean_text(plan_candidate)
    keys = list(bmi_rules_keys)
    keys_clean = [clean_text(k) for k in keys]

    # token strict pass
    for i, k_clean in enumerate(keys_clean):
        if match_tokens_strict(plan_candidate_clean, k_clean, fuzzy_threshold=85):
            return keys[i]

    # If plan_candidate is single token, allow lower fuzzy threshold
    if len(plan_candidate_clean.split()) == 1:
        best_idx, best_score = None, -1
        for i, k_clean in enumerate(keys_clean):
            score = fuzz.token_set_ratio(plan_candidate_clean, k_clean)
            if score > best_score:
                best_score, best_idx = score, i
        if best_idx is not None and best_score >= 75:
            return keys[best_idx]

    # default fuzzy threshold for multiword
    best_idx, best_score = None, -1
    for i, k_clean in enumerate(keys_clean):
        score = fuzz.token_set_ratio(plan_candidate_clean, k_clean)
        if score > best_score:
            best_score, best_idx = score, i
    if best_idx is not None and best_score >= 80:
        return keys[best_idx]

    return None

def type_transaction(row) -> str:
    policy_number = row.get("Poliza")
    raw_percentage = row.get("%")
    normalized_percentage = normalize_percentage(raw_percentage)
    company = df_policy_company.get(policy_number, None)

    if not company:
        return "NO ENCONTRADO - Company not found for policy"
    if not normalized_percentage:
        return "NO ENCONTRADO - Percentage missing/invalid"

    # Non-BMI companies (unchanged behavior, but normalize returned label)
    if "BMI" not in str(company).upper():
        rules_for_company = commission_rules.get(str(company).upper())
        if not rules_for_company:
            return f"NO ENCONTRADO - No rules defined for company {company}"
        if normalized_percentage not in rules_for_company:
            return f"NO ENCONTRADO - Percentage {normalized_percentage} not in rules for {company}"
        return normalize_transaction_label(rules_for_company[normalized_percentage])

    # BMI path: token-aware plan matching
    plan_raw = df_policy_plan.get(policy_number, "")
    if not plan_raw:
        return "NO ENCONTRADO - Plan missing for BMI policy"

    bmi_rules = commission_rules.get("BMI", {})
    if not bmi_rules:
        return "NO ENCONTRADO - No BMI rules defined"

    # Find best plan key in bmi_rules using token rules
    plan_match = find_best_bmi_plan_key(plan_raw, list(bmi_rules.keys()))
    if not plan_match:
        # As a last resort, try a looser fuzzy search on cleaned strings (lower threshold)
        plan_candidate_clean = clean_text(plan_raw)
        keys_clean = [clean_text(k) for k in bmi_rules.keys()]
        match_key, score, _ = process.extractOne(plan_candidate_clean, keys_clean, scorer=fuzz.token_set_ratio)
        if score >= 70:
            plan_match = list(bmi_rules.keys())[keys_clean.index(match_key)]

    if not plan_match:
        plan_norm = clean_text(plan_raw).upper()
        return f"NO ENCONTRADO - No matching plan for BMI policy (plan={plan_norm})"

    plan_rules = bmi_rules[plan_match]
    for txn_code, pct in plan_rules.items():
        if str(pct) == normalized_percentage:
            if txn_code == "I":
                return "Emision"
            elif txn_code == "R":
                return "Renovacion"
            elif txn_code == "CB":
                return "Cambio BR"
            else:
                # If JSON already contains a label instead of code, normalize it
                return normalize_transaction_label(txn_code)

    # Fuzzy fallback by policy (existing logic, but normalize returned labels)
    for other_policy in df_policy_company.keys():
        if fuzz.ratio(str(policy_number), str(other_policy)) >= 100:
            other_company = df_policy_company[other_policy]
            if "BMI" not in str(other_company).upper():
                rules_for_company = commission_rules.get(str(other_company).upper())
                if not rules_for_company:
                    return f"NO ENCONTRADO - Fuzzy match found ({other_policy}) but no rules for {other_company}"
                if normalized_percentage not in rules_for_company:
                    return f"NO ENCONTRADO - Fuzzy match found ({other_policy}) but percentage {normalized_percentage} not in rules"
                return normalize_transaction_label(rules_for_company[normalized_percentage])
            else:
                other_plan = df_policy_plan.get(other_policy, "")
                other_plan_match = find_best_bmi_plan_key(other_plan, list(bmi_rules.keys()))
                if not other_plan_match:
                    other_plan_norm = clean_text(other_plan)
                    keys_clean = [clean_text(k) for k in bmi_rules.keys()]
                    match_key, score, _ = process.extractOne(other_plan_norm, keys_clean, scorer=fuzz.token_set_ratio)
                    if score >= 70:
                        other_plan_match = list(bmi_rules.keys())[keys_clean.index(match_key)]

                if other_plan_match:
                    plan_rules = bmi_rules[other_plan_match]
                    for txn_code, pct in plan_rules.items():
                        if str(pct) == normalized_percentage:
                            if txn_code == "I":
                                return "Emision"
                            elif txn_code == "R":
                                return "Renovacion"
                            elif txn_code == "CB":
                                return "Cambio BR"
                            else:
                                return normalize_transaction_label(txn_code)
                return f"NO ENCONTRADO - Fuzzy match found ({other_policy}) but no matching BMI plan"
    return "NO ENCONTRADO - No fuzzy match found"

def classify_with_reason(row):
    result = type_transaction(row)
    if result.startswith("NO ENCONTRADO"):
        return pd.Series({"Tipo_Transaccion": "NO ENCONTRADO", "Error_Cause": result})
    else:
        return pd.Series({"Tipo_Transaccion": result, "Error_Cause": ""})

# ---------------------
# Assign transaction percentage (float) based on agent + level/special + product + Tipo_Transaccion
# ---------------------
def assign_transaction_percentage(row):
    agent_raw = row.get("Agente", "")
    agent = clean_text(agent_raw)
    policy = row.get("Poliza", "")
    company = df_policy_company.get(policy, "")
    plan = df_policy_plan.get(policy, "")  # already cleaned at load

    if not agent or not company:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": "NO ENCONTRADO - Missing agent or company"})

    tipo_txn = row.get("Tipo_Transaccion", "")
    if not tipo_txn or tipo_txn.startswith("NO ENCONTRADO"):
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": f"NO ENCONTRADO - Invalid transaction type ({tipo_txn})"})

    # ---------------------
    # 1) Try special agent table (niveles.json) with strict matcher
    # ---------------------
    special_block_key = None
    for cleaned_key, originals in special_agent_keys.items():
        for original_key in originals:
            if match_agent_names_strict(agent, original_key, fuzzy_threshold=85):
                special_block_key = original_key
                break
        if special_block_key:
            break

    if special_block_key:
        block = special_tables.get(special_block_key, {})
        comp_key, prod_key = fuzzy_find_company_and_product(block, str(company), str(plan))

        # Debugging helpers (temporary)
        print("DEBUG special block match:", {"policy": policy, "company_candidate": company, "plan_candidate": plan, "comp_key": comp_key, "prod_key": prod_key})
        print("DEBUG token checks:", match_tokens_strict("Sigma", "SIGMA"), match_tokens_strict("SIGMA PLUS", "SIGMA-PLUS"))

        # Determine txn_map depending on block shape and returned keys
        if prod_key:
            if comp_key is None:
                # product-based block: product key maps directly to txn_map
                txn_map = block.get(prod_key, {})
            else:
                # company-based block: need both company and product
                company_block = block.get(comp_key, {})
                txn_map = company_block.get(prod_key, {}) if isinstance(company_block, dict) else {}
            if txn_map:
                if tipo_txn in txn_map:
                    try:
                        pct_float = float(str(txn_map[tipo_txn]).replace(",", "."))
                        return pd.Series({"Transaction_Percentage": pct_float,
                                          "Error_Cause_Percentage": ""})
                    except ValueError:
                        return pd.Series({"Transaction_Percentage": None,
                                          "Error_Cause_Percentage": f"NO ENCONTRADO - Percentage not numeric for {tipo_txn}"})
                else:
                    return pd.Series({"Transaction_Percentage": None,
                                      "Error_Cause_Percentage": f"NO ENCONTRADO - Transaction type not found ({tipo_txn}) in special table"})
            else:
                return pd.Series({"Transaction_Percentage": None,
                                  "Error_Cause_Percentage": "NO ENCONTRADO - Product not found in special table"})
        else:
            return pd.Series({"Transaction_Percentage": None,
                              "Error_Cause_Percentage": "NO ENCONTRADO - Product not found in special table"})

    # ---------------------
    # 2) Fallback to level-based table (niveles1.json)
    # ---------------------
    level_label = agent_to_level_label(agent)
    if not level_label:
        # Retry strictly against special tables for edge cases
        retried = False
        for cleaned_key, originals in special_agent_keys.items():
            for original_key in originals:
                if match_agent_names_strict(agent, original_key, fuzzy_threshold=85):
                    special_block_key = original_key
                    block = special_tables.get(special_block_key, {})
                    comp_key, prod_key = fuzzy_find_company_and_product(block, str(company), str(plan))

                    # Debugging helpers (temporary)
                    print("DEBUG special retry match:", {"policy": policy, "company_candidate": company, "plan_candidate": plan, "comp_key": comp_key, "prod_key": prod_key})
                    print("DEBUG token checks:", match_tokens_strict("Sigma", "SIGMA"), match_tokens_strict("SIGMA PLUS", "SIGMA-PLUS"))

                    if prod_key:
                        if comp_key is None:
                            txn_map = block.get(prod_key, {})
                        else:
                            company_block = block.get(comp_key, {})
                            txn_map = company_block.get(prod_key, {}) if isinstance(company_block, dict) else {}
                        if txn_map:
                            if tipo_txn in txn_map:
                                try:
                                    pct_float = float(str(txn_map[tipo_txn]).replace(",", "."))
                                    return pd.Series({"Transaction_Percentage": pct_float,
                                                      "Error_Cause_Percentage": ""})
                                except ValueError:
                                    return pd.Series({"Transaction_Percentage": None,
                                                      "Error_Cause_Percentage": f"NO ENCONTRADO - Percentage not numeric for {tipo_txn}"})
                            else:
                                return pd.Series({"Transaction_Percentage": None,
                                                  "Error_Cause_Percentage": f"NO ENCONTRADO - Transaction type not found ({tipo_txn}) in special table"})
                        else:
                            return pd.Series({"Transaction_Percentage": None,
                                              "Error_Cause_Percentage": "NO ENCONTRADO - Product not found in special table (retry)"})
                    else:
                        retried = True
                        break
            if retried:
                break
        if retried:
            return pd.Series({"Transaction_Percentage": None,
                              "Error_Cause_Percentage": "NO ENCONTRADO - Agent not found in special tables"})
        else:
            return pd.Series({"Transaction_Percentage": None,
                              "Error_Cause_Percentage": "NO ENCONTRADO - Agent level not found"})

    # Level label (e.g., KREDISEGUROS NIVEL 2) → normal_tables block
    normal_block = (normal_tables.get(level_label.title()) or
                    normal_tables.get(level_label.upper()) or
                    normal_tables.get(level_label))
    if not normal_block:
        # Fuzzy match on level label within normal_tables keys
        normal_keys = list(normal_tables.keys())
        if normal_keys:
            mk, sc, _ = process.extractOne(level_label, [clean_text(k) for k in normal_keys], scorer=fuzz.token_set_ratio)
            if sc >= 60:
                true_key = normal_keys[[clean_text(k) for k in normal_keys].index(mk)]
                normal_block = normal_tables.get(true_key)

    if not normal_block:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": f"NO ENCONTRADO - Level block not found ({level_label})"})

    # Find product inside the level block (handles both product-based and company-based shapes)
    comp_key, prod_key = fuzzy_find_company_and_product(normal_block, str(company), str(plan))

    # Debugging helpers (temporary)
    print("DEBUG level product match:", {"policy": policy, "company_candidate": company, "plan_candidate": plan, "comp_key": comp_key, "prod_key": prod_key})
    print("DEBUG token checks:", match_tokens_strict("Sigma", "SIGMA"), match_tokens_strict("SIGMA PLUS", "SIGMA-PLUS"))

    if not prod_key:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": "NO ENCONTRADO - Product not found in level table"})

    # Resolve txn_map depending on block shape
    if comp_key is None:
        # product-based level block
        txn_map = normal_block.get(prod_key, {})
    else:
        # company-based level block
        company_block = normal_block.get(comp_key, {})
        txn_map = company_block.get(prod_key, {}) if isinstance(company_block, dict) else {}

    if not txn_map:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": "NO ENCONTRADO - Product not found in level table"})

    if tipo_txn not in txn_map:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": f"NO ENCONTRADO - Transaction type not found ({tipo_txn}) in level table"})

    try:
        pct_float = float(str(txn_map[tipo_txn]).replace(",", "."))
        return pd.Series({"Transaction_Percentage": pct_float,
                          "Error_Cause_Percentage": ""})
    except ValueError:
        return pd.Series({"Transaction_Percentage": None,
                          "Error_Cause_Percentage": f"NO ENCONTRADO - Percentage not numeric for {tipo_txn} in level table"})

# ---------------------
# Function to know how much has to be paid to the agent based on the %
# ---------------------

def parse_number(val):
    """Parse numbers like '1.234,56', '1234.56', '39,93' into float."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s == "":
        return np.nan
    s = s.replace(" ", "")
    # If both '.' and ',' present, assume '.' thousands and ',' decimal
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

def compute_payment_value(comision_val, pct_val):
    """
    Compute payment using fraction-or-percent logic:
    - If pct <= 1 (e.g., 0.09) treat as fraction: payment = com * pct
    - If pct > 1 (e.g., 9) treat as percent: payment = com * (pct / 100)
    Returns rounded float (2 decimals).
    """
    com = parse_number(comision_val) if not isinstance(comision_val, (int, float)) else comision_val
    pct = parse_number(pct_val) if not isinstance(pct_val, (int, float)) else pct_val

    if pd.isna(com) or pd.isna(pct):
        return 0.0

    if abs(pct) <= 1:
        payment = com * pct
    else:
        payment = com * (pct / 100.0)

    return round(payment, 2)

def format_comma_decimal(x):
    """Optional: format float to string with comma as decimal separator (e.g., 3.59 -> '3,59')."""
    try:
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""

# ---------------------
# 1) Normalize percentage column used for classification
# ---------------------
df_pre_settlements["%_normalized"] = df_pre_settlements["%"].apply(normalize_percentage)

# ---------------------
# 2) Skip rows with MATCH == "NO" (if you use skip logic elsewhere, keep it)
# ---------------------
# (If you already built skip_mask earlier, this will reuse it; otherwise uncomment the next two lines)
# match_col = df_pre_settlements.get("MATCH", pd.Series([""] * len(df_pre_settlements)))
# skip_mask = match_col.fillna("").astype(str).str.upper() == "NO"

# ---------------------
# 3) Classify transaction type (fills Tipo_Transaccion and Error_Cause)
# ---------------------
df_pre_settlements[["Tipo_Transaccion", "Error_Cause"]] = df_pre_settlements.apply(classify_with_reason, axis=1)

# ---------------------
# 4) Assign Transaction_Percentage and Error_Cause_Percentage
# ---------------------
df_pre_settlements[["Transaction_Percentage", "Error_Cause_Percentage"]] = df_pre_settlements.apply(assign_transaction_percentage, axis=1)

# ---------------------
# 5) Compute Payment_to_Agent (after Transaction_Percentage is available)
# ---------------------
# Parse numeric inputs
df_pre_settlements["_com_num"] = df_pre_settlements["Comisión"].apply(parse_number)
# Transaction_Percentage may be numeric or string; parse if needed
df_pre_settlements["_pct_num"] = df_pre_settlements["Transaction_Percentage"].apply(parse_number)

# Vectorized compute using apply on rows (keeps logic explicit)
def _compute_row_payment(r):
    com = r["_com_num"]
    pct = r["_pct_num"]
    if pd.isna(com) or pd.isna(pct):
        return 0.0
    if abs(pct) <= 1:
        return round(com * pct, 2)
    return round(com * (pct / 100.0), 2)

df_pre_settlements["Payment_to_Agent"] = df_pre_settlements.apply(_compute_row_payment, axis=1)

# Optional formatted string column for display with comma decimal
df_pre_settlements["Payment_to_Agent_str"] = df_pre_settlements["Payment_to_Agent"].apply(format_comma_decimal)

# Cleanup helper numeric columns
df_pre_settlements.drop(columns=["_com_num", "_pct_num"], inplace=True)

# ---------------------
#  How much to pay each agent in total (aggregated)
# ---------------------

def create_agent_payment_summary(df: pd.DataFrame, output_path: str = "Preliquidaciones_Con_Tipo_Y_Porcentaje.xlsx",
                                 preliquidaciones_sheet_name: str = "Preliquidaciones",
                                 summary_sheet_name: str = "Agent_Payments",
                                 format_comma: bool = True):
    """
    Create an Excel file with two sheets:
      - preliquidaciones_sheet_name: the full df (df will be written as-is)
      - summary_sheet_name: aggregated payments per agent (Agente, Total_Payment, Total_Payment_str optional)

    Parameters:
      df: DataFrame containing at least columns "Agente" and "Payment_to_Agent".
      output_path: path to write the Excel file (overwrites if exists).
      format_comma: if True, add a formatted string column with comma decimal (e.g., "3,59").
    Returns:
      summary_df: DataFrame with aggregated totals per agent.
    """
    # Validate input
    if "Agente" not in df.columns:
        raise ValueError("DataFrame must contain an 'Agente' column.")
    if "Payment_to_Agent" not in df.columns:
        raise ValueError("DataFrame must contain a 'Payment_to_Agent' column.")

    # Ensure Payment_to_Agent is numeric
    payments = pd.to_numeric(df["Payment_to_Agent"], errors="coerce").fillna(0.0)

    # Build a working copy with cleaned agent names (preserve original)
    working = df.copy()
    working["_Payment_to_Agent_num"] = payments
    working["_Agente_clean"] = working["Agente"].fillna("").astype(str).str.strip()

    # Group by agent and sum payments, ignoring empty agent names
    grouped = (working[working["_Agente_clean"] != ""]
               .groupby("_Agente_clean", dropna=True)["_Payment_to_Agent_num"]
               .sum()
               .reset_index()
               .rename(columns={"_Agente_clean": "Agente", "_Payment_to_Agent_num": "Total_Payment"}))

    # Round totals to 2 decimals
    grouped["Total_Payment"] = grouped["Total_Payment"].round(2)

    # Optional formatted string with comma decimal separator
    if format_comma:
        def _format_comma(x):
            try:
                return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return ""
        grouped["Total_Payment_str"] = grouped["Total_Payment"].apply(_format_comma)

    # Sort by total descending (optional)
    grouped = grouped.sort_values("Total_Payment", ascending=False).reset_index(drop=True)

    # Write both sheets to Excel (overwrites file)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Write the full preliquidaciones dataframe
        df.to_excel(writer, sheet_name=preliquidaciones_sheet_name, index=False)
        # Write the summary
        grouped.to_excel(writer, sheet_name=summary_sheet_name, index=False)

    # Cleanup temporary columns in working copy (not necessary for original df)
    return grouped

# Example usage (after you computed Payment_to_Agent):

# ---------------------
# Apply classification and percentage assignment
# ---------------------
df_pre_settlements["%_normalized"] = df_pre_settlements["%"].apply(normalize_percentage)
df_pre_settlements[["Tipo_Transaccion", "Error_Cause"]] = df_pre_settlements.apply(classify_with_reason, axis=1)

# Assign percentage based on agent-level/special and product
df_pre_settlements[["Transaction_Percentage", "Error_Cause_Percentage"]] = df_pre_settlements.apply(assign_transaction_percentage, axis=1)

# ---------------------
# Save Result
# ---------------------
df_pre_settlements.to_excel("Preliquidaciones_Con_Tipo_Y_Porcentaje.xlsx", index=False)
summary_df = create_agent_payment_summary(df_pre_settlements, output_path="Preliquidaciones_Con_Tipo_Y_Porcentaje.xlsx")
