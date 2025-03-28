import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import time
import re
import io

##############################################
# Helper: robust conversion to float
##############################################
def to_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0

##############################################
# 1. MAKE COLUMNS UNIQUE
##############################################
def make_columns_unique(df):
    """Append _1, _2, etc. to duplicate column names."""
    seen = {}
    new_cols = []
    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    df.columns = new_cols
    return df

##############################################
# 2. FUZZY RENAME COLUMNS
##############################################
def fuzzy_rename_columns(df, rename_map):
    used_cols = set()
    for final_name, patterns in rename_map.items():
        for col in df.columns:
            if col in used_cols:
                continue
            # skip "Parent Company" => "Company"
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(p.lower().strip() in col.lower() for p in patterns):
                df.rename(columns={col: final_name}, inplace=True)
                used_cols.add(col)
                break
    return df

##############################################
# 3. NORMALIZE KEY
##############################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

##############################################
# 4. LOAD SPGLOBAL (MULTI-HEADER)
##############################################
def load_spglobal(file, sheet_name="Sheet1"):
    """
    Reads an SPGlobal file that has a multi-header row,
    then normalizes column names with fuzzy_rename_columns.
    """
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if len(full_df) < 6:
            raise ValueError("SPGlobal file does not have enough rows.")

        # Row 5 => index 4, row 6 => index 5
        row5 = full_df.iloc[4].fillna("")
        row6 = full_df.iloc[5].fillna("")
        final_cols = []
        for i in range(full_df.shape[1]):
            top = str(row5[i]).strip()
            bot = str(row6[i]).strip()
            combined = top if top else ""
            if bot and bot.lower() not in combined.lower():
                combined = (combined + " " + bot).strip()
            final_cols.append(combined)

        sp_df = full_df.iloc[6:].reset_index(drop=True)
        sp_df.columns = final_cols
        sp_df = make_columns_unique(sp_df)

        rename_map_sp = {
            "SP_ENTITY_NAME":  ["sp entity name", "s&p entity name", "entity name"],
            "SP_ENTITY_ID":    ["sp entity id", "entity id"],
            "SP_COMPANY_ID":   ["sp company id", "company id"],
            "SP_ISIN":         ["sp isin", "isin code"],
            "SP_LEI":          ["sp lei", "lei code"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining":       ["thermal coal mining"],
            "Coal Share of Revenue":     ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector":      ["coal industry sector", "industry sector"],
            ">10MT / >5GW":              [">10mt", ">5gw"],
            "expansion":                 ["expansion"],
        }
        sp_df = fuzzy_rename_columns(sp_df, rename_map_sp)
        sp_df = sp_df.astype(object)
        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()

##############################################
# 5. LOAD URGEWALD (SINGLE HEADER)
##############################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    """
    Reads a single-header Urgewald file, ignoring any 'Parent Company' column,
    then normalizes column names with fuzzy_rename_columns.
    """
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")

        full_df = pd.DataFrame(data)
        header = full_df.iloc[0].fillna("")
        # Filter out "parent company"
        filtered_header = [col for col in header if str(col).strip().lower() != "parent company"]
        ur_df = full_df.iloc[1:].reset_index(drop=True)
        ur_df = ur_df.loc[:, header.str.strip().str.lower() != "parent company"]
        ur_df.columns = filtered_header
        ur_df = make_columns_unique(ur_df)

        rename_map_ur = {
            "Company":        ["company", "issuer name"],
            "ISIN equity":    ["isin equity", "isin(eq)", "isin eq"],
            "LEI":            ["lei", "lei code"],
            "BB Ticker":      ["bb ticker", "bloomberg ticker"],
            "Coal Industry Sector": ["coal industry sector", "industry sector"],
            ">10MT / >5GW":   [">10mt", ">5gw"],
            "expansion":      ["expansion", "expansion text"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Coal Share of Revenue":          ["coal share of revenue"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Generation (Thermal Coal)":     ["generation (thermal coal)"],
            "Thermal Coal Mining":           ["thermal coal mining"],
        }
        ur_df = fuzzy_rename_columns(ur_df, rename_map_ur)
        ur_df = ur_df.astype(object)
        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

##############################################
# 6. MERGE URGEWALD INTO SP (OPTIMIZED)
##############################################
def merge_ur_into_sp_opt(sp_df, ur_df):
    sp_df = sp_df.copy()
    ur_df = ur_df.copy()
    sp_df = sp_df.astype(object)
    ur_df = ur_df.astype(object)

    sp_df["norm_isin"] = sp_df.get("SP_ISIN", "").astype(str).apply(normalize_key)
    sp_df["norm_lei"]  = sp_df.get("SP_LEI", "").astype(str).apply(normalize_key)
    sp_df["norm_name"] = sp_df.get("SP_ENTITY_NAME", "").astype(str).apply(normalize_key)

    for col in ["ISIN equity", "LEI", "Company"]:
        if col not in ur_df.columns:
            ur_df[col] = ""
    ur_df["norm_isin"]    = ur_df["ISIN equity"].astype(str).apply(normalize_key)
    ur_df["norm_lei"]     = ur_df["LEI"].astype(str).apply(normalize_key)
    ur_df["norm_company"] = ur_df["Company"].astype(str).apply(normalize_key)

    dict_isin = {}
    dict_lei  = {}
    dict_name = {}
    for idx, row in sp_df.iterrows():
        if row["norm_isin"]:
            dict_isin.setdefault(row["norm_isin"], idx)
        if row["norm_lei"]:
            dict_lei.setdefault(row["norm_lei"], idx)
        if row["norm_name"]:
            dict_name.setdefault(row["norm_name"], idx)

    ur_not_merged = []
    for _, ur_row in ur_df.iterrows():
        found_index = None
        if ur_row["norm_isin"] and ur_row["norm_isin"] in dict_isin:
            found_index = dict_isin[ur_row["norm_isin"]]
        elif ur_row["norm_lei"] and ur_row["norm_lei"] in dict_lei:
            found_index = dict_lei[ur_row["norm_lei"]]
        elif ur_row["norm_company"] and ur_row["norm_company"] in dict_name:
            found_index = dict_name[ur_row["norm_company"]]
        
        if found_index is not None:
            for col, val in ur_row.items():
                if col.startswith("norm_"):
                    continue
                if (col not in sp_df.columns) or pd.isna(sp_df.loc[found_index, col]) or str(sp_df.loc[found_index, col]).strip() == "":
                    sp_df.loc[found_index, col] = str(val)
            sp_df.loc[found_index, "Merged"] = True
        else:
            ur_not_merged.append(ur_row)
    
    if "Merged" not in sp_df.columns:
        sp_df["Merged"] = False
    merged_df = sp_df.copy()
    ur_only_df = pd.DataFrame(ur_not_merged)
    for c in ["norm_isin","norm_lei","norm_name"]:
        if c in merged_df.columns:
            merged_df.drop(columns=[c], inplace=True)
    for c in ["norm_isin","norm_lei","norm_company"]:
        if c in ur_only_df.columns:
            ur_only_df.drop(columns=[c], inplace=True)
    if "Merged" not in ur_only_df.columns:
        ur_only_df["Merged"] = False
    return merged_df, ur_only_df

##############################################
# 7. THRESHOLD FILTERING
##############################################
def compute_exclusion(row, **params):
    """
    1) Always check if exclude_mt => if "10mt" in >10MT / >5GW => exclude
    2) Always check if exclude_capacity => if Installed Coal Power Capacity (MW) > threshold => exclude
    3) Always check if exclude_power_prod => if (Coal Share of Power Production * 100) > threshold => exclude
       (We multiply by 100 here because it's stored in decimals in the input file)
    4) S&P (mining => Thermal Coal Mining), (power => Generation (Thermal Coal)) => revenue thresholds (no multiplication)
    5) UR (Coal Share of Revenue) => revenue thresholds apply only when the company’s sector is clearly “mining” or clearly “power”.
       For mining: if the sector contains "mining" but NOT "power" or "generation", then apply UR mining threshold.
       For power: if the sector contains "power" or "generation" but NOT "mining", then apply UR power threshold.
       (In both cases, the UR revenue is multiplied by 100 before comparing with the threshold.)
    6) UR Level 2 => all UR => Coal Share of Revenue (multiplied by 100) vs. threshold.
    7) expansions
    """
    reasons = []

    sp_mining_val = to_float(row.get("Thermal Coal Mining", 0))
    sp_power_val  = to_float(row.get("Generation (Thermal Coal)", 0))
    ur_coal_rev   = to_float(row.get("Coal Share of Revenue", 0))
    raw_coal_power = to_float(row.get("Coal Share of Power Production", 0))
    ur_installed_cap = to_float(row.get("Installed Coal Power Capacity (MW)", 0))

    prod_str = str(row.get(">10MT / >5GW", "")).strip().lower()
    expansion_str = str(row.get("expansion", "")).strip().lower()
    is_sp = bool(str(row.get("SP_ENTITY_NAME", "")).strip())
    sector = str(row.get("Coal Industry Sector", "")).strip().lower()

    # 1) >10MT check (applies regardless)
    if params["exclude_mt"] and "10mt" in prod_str:
        reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")

    # 2) Installed capacity check (applies regardless)
    if params["exclude_capacity"] and ur_installed_cap > params["capacity_threshold"]:
        reasons.append(f"Installed capacity {ur_installed_cap:.2f}MW > {params['capacity_threshold']}MW")

    # 3) Coal Share of Power Production check (stored as decimal; multiply by 100)
    if params["exclude_power_prod"]:
        if (raw_coal_power * 100) > params["power_prod_threshold"]:
            reasons.append(f"Coal power production {raw_coal_power*100:.2f}% > {params['power_prod_threshold']}%")

    # 4) S&P revenue thresholds
    if is_sp:
        if "mining" in sector:
            if params["sp_mining_checkbox"] and sp_mining_val > params["sp_mining_threshold"]:
                reasons.append(f"SP Mining revenue {sp_mining_val:.2f}% > {params['sp_mining_threshold']}%")
        if ("power" in sector or "generation" in sector):
            if params["sp_power_checkbox"] and sp_power_val > params["sp_power_threshold"]:
                reasons.append(f"SP Power revenue {sp_power_val:.2f}% > {params['sp_power_threshold']}%")
    else:
        # 5) UR revenue thresholds: apply only when sector is clearly mining or clearly power
        # For mining-only: sector contains "mining" but NOT "power" or "generation"
        if ("mining" in sector) and not (("power" in sector) or ("generation" in sector)):
            if params["ur_mining_checkbox"] and (ur_coal_rev * 100) > params["ur_mining_threshold"]:
                reasons.append(f"UR Mining revenue {ur_coal_rev*100:.2f}% > {params['ur_mining_threshold']}%")
        # For power-only: sector contains "power" or "generation" but NOT "mining"
        elif (("power" in sector or "generation" in sector) and ("mining" not in sector)):
            if params["ur_power_checkbox"] and (ur_coal_rev * 100) > params["ur_power_threshold"]:
                reasons.append(f"UR Power revenue {ur_coal_rev*100:.2f}% > {params['ur_power_threshold']}%")
        # 6) UR Level 2 threshold (applied to all UR records regardless of sector)
        if params["ur_level2_checkbox"] and (ur_coal_rev * 100) > params["ur_level2_threshold"]:
            reasons.append(f"UR Level 2 revenue {ur_coal_rev*100:.2f}% > {params['ur_level2_threshold']}%")

    # 7) Expansion check
    if params["expansion_exclude"]:
        for kw in params["expansion_exclude"]:
            if kw.lower() in expansion_str:
                reasons.append(f"Expansion matched '{kw}'")
                break

    return pd.Series([bool(reasons), "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])

##############################################
# 8. STREAMLIT MAIN
##############################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Merged & Excluded", layout="wide")
    st.title("Coal Exclusion Filter")

    # Sidebar: File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # ---------------------------
    # MINING
    # ---------------------------
    with st.sidebar.expander("Mining", expanded=True):
        # URGEWALD threshold first
        ur_mining_checkbox = st.checkbox("Urgewald: Exclude if thermal coal mining revenue > threshold", value=False)
        ur_mining_threshold = st.number_input("UR Mining: Level 1 threshold (%)", value=5.0)
        # THEN S&P threshold
        sp_mining_checkbox = st.checkbox("S&P: Exclude if thermal coal mining revenue > threshold", value=True)
        sp_mining_threshold = st.number_input("S&P Mining Threshold (%)", value=5.0)
        exclude_mt = st.checkbox("Exclude if > MT threshold", value=True)
        mt_threshold = st.number_input("Max production (MT) threshold", value=10.0)

    # ---------------------------
    # POWER
    # ---------------------------
    with st.sidebar.expander("Power", expanded=True):
        # URGEWALD threshold first
        ur_power_checkbox = st.checkbox("Urgewald: Exclude if thermal coal generation revenue > threshold", value=False)
        ur_power_threshold = st.number_input("UR Power: Level 1 threshold (%)", value=20.0)
        # THEN S&P threshold
        sp_power_checkbox = st.checkbox("S&P: Exclude if thermal coal generation revenue > threshold", value=True)
        sp_power_threshold = st.number_input("S&P Power Threshold (%)", value=20.0)
        exclude_power_prod = st.checkbox("Exclude if > % production threshold", value=True)
        power_prod_threshold = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity = st.checkbox("Exclude if > capacity (MW) threshold", value=True)
        capacity_threshold = st.number_input("Max installed capacity (MW)", value=10000.0)

    # UR Exclusion Level 2
    with st.sidebar.expander("Urgewald: Exclude if Thermal Coal Mining, Power, and Services revenue > threshold", expanded=False):
        ur_level2_checkbox = st.checkbox("Apply UR Level 2 exclusion", value=False)
        ur_level2_threshold = st.number_input("UR Level 2 revenue threshold (%)", value=10.0)

    # Expansion
    with st.sidebar.expander("Exclude if expansion plans on business", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansion_exclude = st.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[])

    st.sidebar.markdown("---")
    start_time = time.time()

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SP
        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or not loaded.")
            return
        sp_df = make_columns_unique(sp_df)

        # 2) Load UR
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return
        ur_df = make_columns_unique(ur_df)

        # 3) Merge
        merged_sp, ur_only_df = merge_ur_into_sp_opt(sp_df, ur_df)
        if "Merged" not in merged_sp.columns:
            merged_sp["Merged"] = False
        if "Merged" not in ur_only_df.columns:
            ur_only_df["Merged"] = False

        # 4) Split
        sp_merged   = merged_sp[merged_sp["Merged"] == True].copy()
        sp_unmerged = merged_sp[merged_sp["Merged"] == False].copy()
        ur_unmerged = ur_only_df[ur_only_df["Merged"] == False].copy()
        for g in [sp_merged, sp_unmerged, ur_unmerged]:
            if "Merged" in g.columns:
                g.drop(columns=["Merged"], inplace=True)

        # S&P Only => from sp_unmerged with nonzero in "Thermal Coal Mining" or "Generation (Thermal Coal)"
        sp_only = sp_unmerged[
            (pd.to_numeric(sp_unmerged.get("Thermal Coal Mining", "0"), errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(sp_unmerged.get("Generation (Thermal Coal)", "0"), errors='coerce').fillna(0) > 0)
        ].copy()

        # 5) Threshold params
        params = {
            "sp_mining_checkbox": sp_mining_checkbox,
            "sp_mining_threshold": sp_mining_threshold,
            "sp_power_checkbox": sp_power_checkbox,
            "sp_power_threshold": sp_power_threshold,
            "ur_mining_checkbox": ur_mining_checkbox,
            "ur_mining_threshold": ur_mining_threshold,
            "ur_power_checkbox": ur_power_checkbox,
            "ur_power_threshold": ur_power_threshold,
            "exclude_mt": exclude_mt,
            "mt_threshold": mt_threshold,
            "exclude_power_prod": exclude_power_prod,
            "power_prod_threshold": power_prod_threshold,
            "exclude_capacity": exclude_capacity,
            "capacity_threshold": capacity_threshold,
            "ur_level2_checkbox": ur_level2_checkbox,
            "ur_level2_threshold": ur_level2_threshold,
            "expansion_exclude": expansion_exclude
        }

        # 6) Filter
        def apply_filter(df):
            if df.empty:
                df["Excluded"] = False
                df["Exclusion Reasons"] = ""
                return df
            filtered = df.apply(lambda row: compute_exclusion(row, **params), axis=1, result_type="expand")
            if not filtered.empty and "Excluded" in filtered.columns:
                df["Excluded"] = filtered["Excluded"]
                df["Exclusion Reasons"] = filtered["Exclusion Reasons"]
            else:
                df["Excluded"] = False
                df["Exclusion Reasons"] = ""
            return df

        sp_merged = apply_filter(sp_merged)
        sp_only = apply_filter(sp_only)
        ur_unmerged = apply_filter(ur_unmerged)

        # 7) Build final sets
        excluded_final = pd.concat([
            sp_merged[sp_merged["Excluded"] == True],
            sp_only[sp_only["Excluded"] == True],
            ur_unmerged[ur_unmerged["Excluded"] == True]
        ], ignore_index=True)
        retained_merged = sp_merged[sp_merged["Excluded"] == False].copy()
        sp_retained = sp_only[sp_only["Excluded"] == False].copy()
        ur_retained = ur_unmerged[ur_unmerged["Excluded"] == False].copy()

        # 8) Final columns
        final_cols = [
            "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
            "Coal Industry Sector", "Company", ">10MT / >5GW", "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production", "Coal Share of Revenue", "expansion",
            "Generation (Thermal Coal)", "Thermal Coal Mining", "BB Ticker", "ISIN equity", "LEI",
            "Excluded", "Exclusion Reasons"
        ]
        def finalize_cols(df):
            df = df.copy()
            for c in final_cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[final_cols]
            if "BB Ticker" in df.columns:
                # Remove any occurrence of whitespace and the word "Equity"
                df["BB Ticker"] = df["BB Ticker"].astype(str).str.replace(r'\s*Equity', '', regex=True)
            return df


        excluded_final = finalize_cols(excluded_final)
        retained_merged = finalize_cols(retained_merged)
        sp_retained = finalize_cols(sp_retained)
        ur_retained = finalize_cols(ur_retained)

        # 9) Write Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_final.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_merged.to_excel(writer, sheet_name="Retained Companies", index=False)
            sp_retained.to_excel(writer, sheet_name="S&P Only", index=False)
            ur_retained.to_excel(writer, sheet_name="Urgewald Only", index=False)

        elapsed = time.time() - start_time
        st.subheader("Results Summary")
        st.write(f"Excluded Companies: {len(excluded_final)}")
        st.write(f"Retained Companies (Merged & Retained): {len(retained_merged)}")
        st.write(f"S&P Only (Unmatched, Retained): {len(sp_retained)}")
        st.write(f"Urgewald Only (Unmatched, Retained): {len(ur_retained)}")
        st.write(f"Run Time: {elapsed:.2f} seconds")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="Coal_Companies_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
