import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import re
import io

# ──────────────────────────────────────────────────────────────────────────────
# Helper: robust conversion to float (handles EU & US formats)
# ──────────────────────────────────────────────────────────────────────────────
def to_float(val):
    s = str(val).strip().replace(" ", "")
    if s == "" or s.lower() in ("nan", "none"):
        return 0.0
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", ".") if len(parts) == 2 and len(parts[1]) in (1, 2) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ───────── utilities (unchanged) ──────────────────────────────────────────────
def make_columns_unique(df):
    seen, new_cols = {}, []
    for c in df.columns:
        seen[c] = seen.get(c, 0) + 1
        new_cols.append(c if seen[c] == 1 else f"{c}_{seen[c]-1}")
    df.columns = new_cols
    return df


def fuzzy_rename_columns(df, rename_map):
    used = set()
    for final_name, pats in rename_map.items():
        for col in df.columns:
            if col in used:
                continue
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(p.lower().strip() in col.lower() for p in pats):
                df.rename(columns={col: final_name}, inplace=True)
                used.add(col)
                break
    return df


def normalize_key(s):
    return re.sub(r"[^\w\s]", "", re.sub(r"\s+", " ", s.lower())).strip()


# ───────── loaders (unchanged) ────────────────────────────────────────────────
def load_spglobal(file, sheet_name="Sheet1"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if len(full_df) < 6:
            raise ValueError("SPGlobal file does not have enough rows.")

        row5 = full_df.iloc[4].fillna("")
        row6 = full_df.iloc[5].fillna("")
        final_cols = []
        for a, b in zip(row5, row6):
            top = str(a).strip(); bot = str(b).strip()
            col = top if top else ""
            if bot and bot.lower() not in col.lower():
                col = f"{col} {bot}".strip()
            final_cols.append(col)

        sp_df = full_df.iloc[6:].reset_index(drop=True)
        sp_df.columns = final_cols
        sp_df = make_columns_unique(sp_df)

        rename_map_sp = {
            "SP_ENTITY_NAME": ["sp entity name", "entity name"],
            "SP_ENTITY_ID": ["sp entity id", "entity id"],
            "SP_COMPANY_ID": ["sp company id", "company id"],
            "SP_ISIN": ["sp isin"],
            "SP_LEI": ["sp lei"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
            "Coal Share of Revenue": ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector": ["industry sector"],
            ">10MT / >5GW": [">10mt", ">5gw"],
            "expansion": ["expansion"],
        }
        sp_df = fuzzy_rename_columns(sp_df, rename_map_sp).astype(object)

        for col in [
            "Thermal Coal Mining", "Generation (Thermal Coal)",
            "Coal Share of Revenue", "Coal Share of Power Production",
            "Installed Coal Power Capacity (MW)"
        ]:
            if col in sp_df:
                sp_df[col] = sp_df[col].apply(to_float)

        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()


def load_urgewald(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if full_df.empty:
            raise ValueError("Urgewald file is empty.")

        header = full_df.iloc[0].fillna("")
        keep = header.str.strip().str.lower() != "parent company"
        ur_df = full_df.iloc[1:].reset_index(drop=True).loc[:, keep]
        ur_df.columns = [c for c in header if str(c).strip().lower() != "parent company"]
        ur_df = make_columns_unique(ur_df)

        rename_map_ur = {
            "Company": ["company", "issuer name"],
            "ISIN equity": ["isin equity", "isin(eq)", "isin eq"],
            "LEI": ["lei"],
            "BB Ticker": ["bb ticker"],
            "Coal Industry Sector": ["industry sector"],
            ">10MT / >5GW": [">10mt", ">5gw"],
            "expansion": ["expansion"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Coal Share of Revenue": ["coal share of revenue"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
        }
        ur_df = fuzzy_rename_columns(ur_df, rename_map_ur).astype(object)

        for col in [
            "Thermal Coal Mining", "Generation (Thermal Coal)",
            "Coal Share of Revenue", "Coal Share of Power Production",
            "Installed Coal Power Capacity (MW)"
        ]:
            if col in ur_df:
                ur_df[col] = ur_df[col].apply(to_float)

        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()


# ───────── merge (unchanged) ──────────────────────────────────────────────────
def merge_ur_into_sp_opt(sp_df, ur_df):
    sp = sp_df.copy().astype(object)
    ur = ur_df.copy().astype(object)

    sp["norm_isin"] = sp.get("SP_ISIN", "").astype(str).apply(normalize_key)
    sp["norm_lei"] = sp.get("SP_LEI", "").astype(str).apply(normalize_key)
    sp["norm_name"] = sp.get("SP_ENTITY_NAME", "").astype(str).apply(normalize_key)

    for col in ["ISIN equity", "LEI", "Company"]:
        if col not in ur:
            ur[col] = ""
    ur["norm_isin"] = ur["ISIN equity"].astype(str).apply(normalize_key)
    ur["norm_lei"] = ur["LEI"].astype(str).apply(normalize_key)
    ur["norm_company"] = ur["Company"].astype(str).apply(normalize_key)

    isin_map = {k: i for i, k in enumerate(sp["norm_isin"]) if k}
    lei_map = {k: i for i, k in enumerate(sp["norm_lei"]) if k}
    name_map = {k: i for i, k in enumerate(sp["norm_name"]) if k}

    ur_not = []
    for _, r in ur.iterrows():
        target = None
        if r["norm_isin"] in isin_map:
            target = isin_map[r["norm_isin"]]
        elif r["norm_lei"] in lei_map:
            target = lei_map[r["norm_lei"]]
        elif r["norm_company"] in name_map:
            target = name_map[r["norm_company"]]

        if target is not None:
            for c, v in r.items():
                if c.startswith("norm_"):
                    continue
                if c not in sp or pd.isna(sp.at[target, c]) or str(sp.at[target, c]).strip() == "":
                    sp.at[target, c] = v
            sp.at[target, "Merged"] = True
        else:
            ur_not.append(r)

    sp["Merged"] = sp.get("Merged", False).fillna(False)
    ur_only = pd.DataFrame(ur_not)
    ur_only["Merged"] = False

    for c in [c for c in sp.columns if c.startswith("norm_")]:
        sp.drop(columns=c, inplace=True, errors="ignore")
    for c in [c for c in ur_only.columns if c.startswith("norm_")]:
        ur_only.drop(columns=c, inplace=True, errors="ignore")

    return sp, ur_only


# ───────── NEW: helper for > vs ≥ ─────────────────────────────────────────────
def test(val, thr, ge):
    """Return True if value triggers rule ( > or ≥ )."""
    return val >= thr if ge else val > thr


# ───────── compute_exclusion (only comparisons touched) ───────────────────────
def compute_exclusion(row, **params):
    reasons = []

    # numeric fields
    sp_min = row.get("Thermal Coal Mining", 0.0)
    sp_pow = row.get("Generation (Thermal Coal)", 0.0)

    ur_rev_pct = row.get("Coal Share of Revenue", 0.0)
    ur_rev_pct = ur_rev_pct if ur_rev_pct > 1 else ur_rev_pct * 100
    ur_pp_pct = row.get("Coal Share of Power Production", 0.0)
    ur_pp_pct = ur_pp_pct if ur_pp_pct > 1 else ur_pp_pct * 100

    # misc flags
    prod_str = str(row.get(">10MT / >5GW", "")).lower()
    cap = row.get("Installed Coal Power Capacity (MW)", 0.0)
    expansion = str(row.get("expansion", "")).lower()

    has_sp = bool(str(row.get("SP_ENTITY_NAME", "")).strip())
    has_ur = bool(str(row.get("Company", "")).strip())

    # sector parsing
    sector_raw = str(row.get("Coal Industry Sector", "")).lower()
    mining_kw = ("mining", "extraction", "producer")
    power_kw = ("power", "generation", "utility", "electric")
    parts = [p.strip() for p in re.split(r"[;,/]|(?:\s*\n\s*)", sector_raw) if p.strip()]
    mining_parts = [p for p in parts if any(k in p for k in mining_kw)]
    power_parts = [p for p in parts if any(k in p for k in power_kw)]
    other_parts = [p for p in parts if p not in mining_parts + power_parts]

    is_mining_only = bool(mining_parts) and not power_parts and not other_parts
    is_power_only = bool(power_parts) and not mining_parts and not other_parts
    is_mixed = bool(mining_parts) and bool(power_parts) and not other_parts

    # global screens
    if params["exclude_mt"] and "10mt" in prod_str:
        reasons.append(">10 MT indicator")

    if params["exclude_capacity"] and test(cap, params["capacity_threshold"], params["capacity_ge"]):
        reasons.append(f"Installed capacity {cap:.0f} MW")

    if params["exclude_power_prod"] and test(ur_pp_pct, params["power_prod_threshold"], params["power_prod_ge"]):
        reasons.append(f"Coal power production {ur_pp_pct:.2f}%")

    # S&P rules
    if has_sp:
        if params["sp_mining_checkbox"] and test(sp_min, params["sp_mining_threshold"], params["sp_mining_ge"]):
            reasons.append(f"SP mining revenue {sp_min:.2f}%")
        if params["sp_power_checkbox"] and test(sp_pow, params["sp_power_threshold"], params["sp_power_ge"]):
            reasons.append(f"SP power revenue {sp_pow:.2f}%")
        if params["sp_level2_checkbox"]:
            combo = sp_min + sp_pow
            if test(combo, params["sp_level2_threshold"], params["sp_level2_ge"]):
                reasons.append(f"SP level-2 combined {combo:.2f}%")

    # UR rules
    if has_ur:
        if is_mining_only and params["ur_mining_checkbox"] and test(ur_rev_pct, params["ur_mining_threshold"], params["ur_mining_ge"]):
            reasons.append(f"UR mining revenue {ur_rev_pct:.2f}%")
        if is_power_only and params["ur_power_checkbox"] and test(ur_rev_pct, params["ur_power_threshold"], params["ur_power_ge"]):
            reasons.append(f"UR power revenue {ur_rev_pct:.2f}%")
        if is_mixed and params["ur_mixed_checkbox"] and test(ur_rev_pct, params["ur_mixed_threshold"], params["ur_mixed_ge"]):
            reasons.append(f"UR mixed revenue {ur_rev_pct:.2f}%")
        if params["ur_level2_checkbox"] and test(ur_rev_pct, params["ur_level2_threshold"], params["ur_level2_ge"]):
            reasons.append(f"UR level-2 revenue {ur_rev_pct:.2f}%")

    # expansion keywords
    if any(k.lower() in expansion for k in params["expansion_exclude"]):
        reasons.append("Expansion flag")

    return pd.Series([bool(reasons), "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])


# ───────── Streamlit UI (added one small ≥ checkbox beside each threshold) ────
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # file inputs
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", "Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", "GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # helper: numeric input + ≥ toggle side-by-side
    def num_ge(label, default, key):
        col1, col2 = st.columns([3, 1])
        with col1:
            val = st.number_input(label, value=default, key=f"{key}_val")
        with col2:
            ge_flag = st.checkbox("≥", value=False, key=f"{key}_ge")
        return val, ge_flag

    # Mining
    with st.sidebar.expander("Mining", True):
        ur_mining_checkbox = st.checkbox("UR: Exclude mining-only", False)
        ur_mining_threshold, ur_mining_ge = num_ge("UR Mining threshold (%)", 5.0, "urmin")
        sp_mining_checkbox = st.checkbox("SP: Exclude mining-only", True)
        sp_mining_threshold, sp_mining_ge = num_ge("SP Mining threshold (%)", 5.0, "spmin")

        exclude_mt = st.checkbox("Exclude >10MT indicator", True)
        mt_threshold = st.number_input("MT threshold (not used numerically)", value=10.0)

    # Power
    with st.sidebar.expander("Power", True):
        ur_power_checkbox = st.checkbox("UR: Exclude power-only", False)
        ur_power_threshold, ur_power_ge = num_ge("UR Power threshold (%)", 20.0, "urpow")

        sp_power_checkbox = st.checkbox("SP: Exclude power-only", True)
        sp_power_threshold, sp_power_ge = num_ge("SP Power threshold (%)", 20.0, "sppow")

        exclude_power_prod = st.checkbox("Exclude power-production %", True)
        power_prod_threshold, power_prod_ge = num_ge("Power-production threshold (%)", 20.0, "ppp")

        exclude_capacity = st.checkbox("Exclude installed capacity", True)
        capacity_threshold, capacity_ge = num_ge("Capacity threshold (MW)", 10000.0, "cap")

    # Mixed & Level-2
    with st.sidebar.expander("Mixed & Level-2", False):
        ur_mixed_checkbox = st.checkbox("UR: Exclude mining & power (mixed)", False)
        ur_mixed_threshold, ur_mixed_ge = num_ge("UR Mixed threshold (%)", 25.0, "urmix")

        ur_level2_checkbox = st.checkbox("UR: Apply Level-2", False)
        ur_level2_threshold, ur_level2_ge = num_ge("UR Level-2 threshold (%)", 10.0, "url2")

        sp_level2_checkbox = st.checkbox("SP: Apply Level-2", False)
        sp_level2_threshold, sp_level2_ge = num_ge("SP Level-2 threshold (%)", 10.0, "spl2")

    # Expansion
    with st.sidebar.expander("Exclude expansions", False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansion_exclude = st.multiselect("Exclude if expansion text contains", expansions_possible, [])

    st.sidebar.markdown("---")
    if not st.sidebar.button("Run"):
        st.stop()

    # load
    if not sp_file or not ur_file:
        st.warning("Please upload both files")
        st.stop()
    sp_df = load_spglobal(sp_file, sp_sheet)
    ur_df = load_urgewald(ur_file, ur_sheet)
    if sp_df.empty or ur_df.empty:
        st.warning("Error loading data")
        st.stop()

    merged_sp, ur_only = merge_ur_into_sp_opt(sp_df, ur_df)
    for df in (merged_sp, ur_only):
        df["Merged"] = df.get("Merged", False).fillna(False)

    sp_merged = merged_sp[merged_sp.Merged]
    sp_only = merged_sp[~merged_sp.Merged & (
        (merged_sp["Thermal Coal Mining"] > 0) | (merged_sp["Generation (Thermal Coal)"] > 0)
    )]
    ur_unmerged = ur_only[~ur_only.Merged]

    # params dict
    params = {
        # UR Level-1
        "ur_mining_checkbox": ur_mining_checkbox, "ur_mining_threshold": ur_mining_threshold, "ur_mining_ge": ur_mining_ge,
        "ur_power_checkbox": ur_power_checkbox, "ur_power_threshold": ur_power_threshold, "ur_power_ge": ur_power_ge,
        "ur_mixed_checkbox": ur_mixed_checkbox, "ur_mixed_threshold": ur_mixed_threshold, "ur_mixed_ge": ur_mixed_ge,
        # SP Level-1
        "sp_mining_checkbox": sp_mining_checkbox, "sp_mining_threshold": sp_mining_threshold, "sp_mining_ge": sp_mining_ge,
        "sp_power_checkbox": sp_power_checkbox, "sp_power_threshold": sp_power_threshold, "sp_power_ge": sp_power_ge,
        # Global
        "exclude_mt": exclude_mt,
        "exclude_capacity": exclude_capacity, "capacity_threshold": capacity_threshold, "capacity_ge": capacity_ge,
        "exclude_power_prod": exclude_power_prod, "power_prod_threshold": power_prod_threshold, "power_prod_ge": power_prod_ge,
        # Level-2
        "ur_level2_checkbox": ur_level2_checkbox, "ur_level2_threshold": ur_level2_threshold, "ur_level2_ge": ur_level2_ge,
        "sp_level2_checkbox": sp_level2_checkbox, "sp_level2_threshold": sp_level2_threshold, "sp_level2_ge": sp_level2_ge,
        # mixed
        "expansion_exclude": [k.strip() for k in expansion_exclude if k.strip()],
    }

    # apply filter
    def apply(df):
        if df.empty:
            return df.assign(Excluded=False, **{"Exclusion Reasons": ""})
        res = df.apply(lambda r: compute_exclusion(r, **params), axis=1, result_type="expand")
        df["Excluded"], df["Exclusion Reasons"] = res["Excluded"], res["Exclusion Reasons"]
        return df

    sp_merged = apply(sp_merged)
    sp_only = apply(sp_only)
    ur_unmerged = apply(ur_unmerged)

    excluded_final = pd.concat([sp_merged[sp_merged.Excluded], sp_only[sp_only.Excluded], ur_unmerged[ur_unmerged.Excluded]])
    retained_merged = sp_merged[~sp_merged.Excluded]
    sp_retained = sp_only[~sp_only.Excluded]
    ur_retained = ur_unmerged[~ur_unmerged.Excluded]

    final_cols = [
        "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
        "Coal Industry Sector", "Company", ">10MT / >5GW",
        "Installed Coal Power Capacity (MW)",
        "Coal Share of Power Production", "Coal Share of Revenue", "expansion",
        "Generation (Thermal Coal)", "Thermal Coal Mining",
        "BB Ticker", "ISIN equity", "LEI", "Excluded", "Exclusion Reasons"
    ]
    def finalize(d):
        for c in final_cols:
            if c not in d:
                d[c] = ""
        return d[final_cols]

    excluded_final = finalize(excluded_final)
    retained_merged = finalize(retained_merged)
    sp_retained = finalize(sp_retained)
    ur_retained = finalize(ur_retained)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        excluded_final.to_excel(w, "Excluded Companies", index=False)
        retained_merged.to_excel(w, "Retained Companies", index=False)
        sp_retained.to_excel(w, "S&P Only", index=False)
        ur_retained.to_excel(w, "Urgewald Only", index=False)

    st.subheader("Results Summary")
    st.write(f"Excluded Companies: {len(excluded_final)}")
    st.write(f"Retained Companies (Merged & Retained): {len(retained_merged)}")
    st.write(f"S&P Only (Unmatched, Retained): {len(sp_retained)}")
    st.write(f"Urgewald Only (Unmatched, Retained): {len(ur_retained)}")

    st.download_button(
        label="Download Filtered Results",
        data=buf.getvalue(),
        file_name="Coal_Companies_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
