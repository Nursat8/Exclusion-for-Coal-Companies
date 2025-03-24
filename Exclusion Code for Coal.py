import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
import time
import re

################################################
# 1. MAKE COLUMNS UNIQUE
################################################
def make_columns_unique(df):
    """
    Append _1, _2, etc. to duplicate column names.
    """
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

################################################
# 2. FUZZY RENAME COLUMNS
################################################
def fuzzy_rename_columns(df, rename_map):
    """
    Rename columns based on patterns.
    rename_map: { final_name: [pattern1, pattern2, ...], ... }
    """
    used = set()
    for final, patterns in rename_map.items():
        for col in df.columns:
            if col in used:
                continue
            # When renaming to "Company" in UR, skip if col is "Parent Company"
            if final == "Company" and col.strip().lower() == "parent company":
                continue
            if any(pat.lower().strip() in col.lower() for pat in patterns):
                df.rename(columns={col: final}, inplace=True)
                used.add(col)
    return df

################################################
# 3. REORDER COLUMNS FOR FINAL EXCEL
################################################
def reorder_for_excel(df):
    """
    Force specific columns into fixed positions:
      - "Company" in column G (7th)
      - "BB Ticker" in column AP (42nd)
      - "ISIN equity" in column AQ (43rd)
      - "LEI" in column AT (46th)
    Then move "Excluded" and "Exclusion Reasons" to the very end.
    """
    desired_length = 46
    placeholders = ["(placeholder)"] * desired_length

    # Fixed positions (0-indexed)
    placeholders[6]  = "Company"
    placeholders[41] = "BB Ticker"
    placeholders[42] = "ISIN equity"
    placeholders[45] = "LEI"

    forced_positions = {6, 41, 42, 45}
    forced_cols = {"Company", "BB Ticker", "ISIN equity", "LEI"}
    all_cols = list(df.columns)
    remaining = [c for c in all_cols if c not in forced_cols]

    idx = 0
    for i in range(desired_length):
        if i not in forced_positions and idx < len(remaining):
            placeholders[i] = remaining[idx]
            idx += 1

    leftover = remaining[idx:]
    final_order = placeholders + leftover

    df = df[[c for c in final_order if c in df.columns]]
    cols = list(df.columns)
    for c in ["Excluded", "Exclusion Reasons"]:
        if c in cols:
            cols.remove(c)
            cols.append(c)
    return df[cols]

################################################
# 4. LOAD SPGLOBAL (AUTO-DETECT MULTI-HEADER)
################################################
def load_spglobal(file, sheet_name="Sheet1"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        df = pd.DataFrame(data)
        if len(df) < 6:
            raise ValueError("SPGlobal file does not have enough rows.")
        row5 = df.iloc[4].fillna("")
        row6 = df.iloc[5].fillna("")
        cols = []
        for i in range(df.shape[1]):
            top = str(row5[i]).strip()
            bot = str(row6[i]).strip()
            combined = top if top else ""
            if bot and bot.lower() not in combined.lower():
                combined = (combined + " " + bot).strip()
            cols.append(combined)
        sp_df = df.iloc[6:].reset_index(drop=True)
        sp_df.columns = cols
        sp_df = make_columns_unique(sp_df)
        rename_map = {
            "SP_ENTITY_NAME":  ["sp entity name", "s&p entity name", "entity name"],
            "SP_ENTITY_ID":    ["sp entity id", "entity id"],
            "SP_COMPANY_ID":   ["sp company id", "company id"],
            "SP_ISIN":         ["sp isin", "isin code"],
            "SP_LEI":          ["sp lei", "lei code"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
            "Coal Share of Revenue": ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector": ["coal industry sector", "industry sector"],
            ">10MT / >5GW": [">10mt", ">5gw"],
            "expansion": ["expansion"],
        }
        sp_df = fuzzy_rename_columns(sp_df, rename_map)
        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()

################################################
# 5. LOAD URGEWALD (SINGLE HEADER) – DROP "Parent Company"
################################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")
        df = pd.DataFrame(data)
        # Filter header: remove any column where header equals "Parent Company"
        header = df.iloc[0].fillna("")
        filtered_header = [col for col in header if str(col).strip().lower() != "parent company"]
        ur_df = df.iloc[1:].reset_index(drop=True)
        # Keep only columns corresponding to filtered_header (assume same order)
        # (We assume that the "Parent Company" columns have been removed in the header row.)
        ur_df = ur_df.loc[:, df.iloc[0].str.strip().str.lower() != "parent company"]
        ur_df.columns = filtered_header
        ur_df = make_columns_unique(ur_df)
        rename_map = {
            "Company": ["company", "issuer name"],
            "ISIN equity": ["isin equity", "isin(eq)", "isin eq"],
            "LEI": ["lei", "lei code"],
            "BB Ticker": ["bb ticker", "bloomberg ticker"],
            "Coal Industry Sector": ["coal industry sector", "industry sector"],
            ">10MT / >5GW": [">10mt", ">5gw"],
            "expansion": ["expansion", "expansion text"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Coal Share of Revenue": ["coal share of revenue"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
            "Metallurgical Coal Mining": ["metallurgical coal mining"],
        }
        ur_df = fuzzy_rename_columns(ur_df, rename_map)
        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

################################################
# 6. NORMALIZE KEYS FOR MATCHING
################################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

def unify_name(r):
    sp_val = str(r.get("SP_ENTITY_NAME", ""))
    ur_val = str(r.get("Company", ""))
    return normalize_key(sp_val) if sp_val.strip() else normalize_key(ur_val)

def unify_isin(r):
    sp_val = str(r.get("SP_ISIN", ""))
    ur_val = str(r.get("ISIN equity", ""))
    return normalize_key(sp_val) if sp_val.strip() else normalize_key(ur_val)

def unify_lei(r):
    sp_val = str(r.get("SP_LEI", ""))
    ur_val = str(r.get("LEI", ""))
    return normalize_key(sp_val) if sp_val.strip() else normalize_key(ur_val)

def unify_bbticker(r):
    return normalize_key(str(r.get("BB Ticker", "")))

################################################
# 7. VECTORIZED MATCHING
################################################
def vectorized_match(sp_df, ur_df):
    # Add normalized key columns
    sp_df["norm_name"] = sp_df["SP_ENTITY_NAME"].astype(str).apply(normalize_key)
    sp_df["norm_isin"] = sp_df["SP_ISIN"].astype(str).apply(normalize_key)
    sp_df["norm_lei"]  = sp_df["SP_LEI"].astype(str).apply(normalize_key)
    sp_df["norm_bbticker"] = sp_df["BB Ticker"].astype(str).apply(normalize_key)
    ur_df["norm_name"] = ur_df["Company"].astype(str).apply(normalize_key)
    ur_df["norm_isin"] = ur_df["ISIN equity"].astype(str).apply(normalize_key)
    ur_df["norm_lei"]  = ur_df["LEI"].astype(str).apply(normalize_key)
    ur_df["norm_bbticker"] = ur_df["BB Ticker"].astype(str).apply(normalize_key)
    # For SP, mark as merged if any normalized key is in UR dataset
    sp_df["Merged"] = sp_df["norm_name"].isin(ur_df["norm_name"]) | \
                      sp_df["norm_isin"].isin(ur_df["norm_isin"]) | \
                      sp_df["norm_lei"].isin(ur_df["norm_lei"]) | \
                      sp_df["norm_bbticker"].isin(ur_df["norm_bbticker"])
    # For UR, mark as merged similarly
    ur_df["Merged"] = ur_df["norm_name"].isin(sp_df["norm_name"]) | \
                      ur_df["norm_isin"].isin(sp_df["norm_isin"]) | \
                      ur_df["norm_lei"].isin(sp_df["norm_lei"]) | \
                      ur_df["norm_bbticker"].isin(sp_df["norm_bbticker"])
    # Drop temporary normalized columns
    for col in ["norm_name", "norm_isin", "norm_lei", "norm_bbticker"]:
        sp_df.drop(columns=[col], inplace=True)
        ur_df.drop(columns=[col], inplace=True)
    return sp_df, ur_df

################################################
# 8. THRESHOLD FILTERING VIA APPLY
################################################
def compute_exclusion(row, mining_coal_rev_threshold, exclude_mining_revenue,
                      exclude_mining_prod_mt, mining_prod_mt_threshold,
                      exclude_thermal_coal_mining, thermal_coal_mining_threshold,
                      power_coal_rev_threshold, exclude_power_revenue,
                      exclude_power_prod_percent, power_prod_threshold_percent,
                      capacity_threshold_mw, exclude_capacity_mw,
                      generation_thermal_threshold, exclude_generation_thermal,
                      services_rev_threshold, exclude_services_rev,
                      expansions_global):
    reasons = []
    sector = str(row.get("Coal Industry Sector", "")).lower()
    try:
        coal_rev = float(row.get("Coal Share of Revenue", 0))
    except:
        coal_rev = 0.0
    try:
        coal_power = float(row.get("Coal Share of Power Production", 0))
    except:
        coal_power = 0.0
    try:
        capacity = float(row.get("Installed Coal Power Capacity (MW)", 0))
    except:
        capacity = 0.0
    try:
        gen_val = float(row.get("Generation (Thermal Coal)", 0))
    except:
        gen_val = 0.0
    try:
        thermal_val = float(row.get("Thermal Coal Mining", 0))
    except:
        thermal_val = 0.0
    prod_str = str(row.get(">10MT / >5GW", "")).lower()
    expansion = str(row.get("expansion", "")).lower()
    # Mining revenue check only for sectors containing "mining"
    if "mining" in sector:
        if exclude_mining_revenue and (coal_rev * 100) > mining_coal_rev_threshold:
            reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
    if exclude_mining_prod_mt and (">10mt" in prod_str) and (mining_prod_mt_threshold <= 10):
        reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
    if exclude_thermal_coal_mining and (thermal_val > thermal_coal_mining_threshold):
        reasons.append(f"Thermal Coal Mining {thermal_val:.2f}% > {thermal_coal_mining_threshold}%")
    # Power revenue check only for sectors containing "power" or "generation"
    if ("power" in sector or "generation" in sector):
        if exclude_power_revenue and (coal_rev * 100) > power_coal_rev_threshold:
            reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
        if exclude_power_prod_percent and (coal_power * 100) > power_prod_threshold_percent:
            reasons.append(f"Coal power production {coal_power*100:.2f}% > {power_prod_threshold_percent}%")
        if exclude_capacity_mw and (capacity > capacity_threshold_mw):
            reasons.append(f"Installed capacity {capacity:.2f}MW > {capacity_threshold_mw}MW")
        if exclude_generation_thermal and (gen_val > generation_thermal_threshold):
            reasons.append(f"Generation (Thermal Coal) {gen_val:.2f}% > {generation_thermal_threshold}%")
    if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
        reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")
    if expansions_global:
        for kw in expansions_global:
            if kw.lower() in expansion:
                reasons.append(f"Expansion matched '{kw}'")
                break
    return pd.Series([len(reasons) > 0, "; ".join(reasons)])

def apply_thresholds(df, params):
    res = df.apply(lambda row: compute_exclusion(row, **params), axis=1)
    df["Excluded"] = res[0]
    df["Exclusion Reasons"] = res[1]
    return df

################################################
# 9. MAIN STREAMLIT APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Optimized", layout="wide")
    st.title("Coal Exclusion Filter")

    # Sidebar: File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # Sidebar: Mining Thresholds (Metallurgical removed)
    with st.sidebar.expander("Mining Thresholds", expanded=True):
        exclude_mining_revenue = st.checkbox("Exclude if coal revenue > threshold? (Mining)", value=True)
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0)
        exclude_mining_prod_mt = st.checkbox("Exclude if >10MT indicated?", value=True)
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0)
        exclude_thermal_coal_mining = st.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False)
        thermal_coal_mining_threshold = st.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)

    # Sidebar: Power Thresholds
    with st.sidebar.expander("Power Thresholds", expanded=True):
        exclude_power_revenue = st.checkbox("Exclude if coal revenue > threshold? (Power)", value=True)
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True)
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True)
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0)
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    # Sidebar: Services Thresholds
    with st.sidebar.expander("Services Thresholds", expanded=False):
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)

    # Sidebar: Global Expansion
    with st.sidebar.expander("Global Expansion Exclusion", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansions_global = st.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[])

    st.sidebar.markdown("---")

    # Start runtime timer
    start_time = time.time()

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # Load SPGlobal and Urgewald data
        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or not loaded.")
            return
        sp_df = make_columns_unique(sp_df)
        
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return
        ur_df = make_columns_unique(ur_df)

        # Vectorized matching: add normalized key columns and mark merged records
        sp_df, ur_df = vectorized_match(sp_df, ur_df)

        # Unmatched records become "Only" sheets
        sp_only_df = sp_df[sp_df["Merged"] == False].copy()
        ur_only_df = ur_df[ur_df["Merged"] == False].copy()
        # For S&P Only, further restrict to those with nonzero values in Thermal Coal Mining or Generation (Thermal Coal)
        sp_only_df = sp_only_df[
            (pd.to_numeric(sp_only_df["Thermal Coal Mining"], errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(sp_only_df["Generation (Thermal Coal)"], errors='coerce').fillna(0) > 0)
        ].copy()

        # Prepare threshold parameters
        params = {
            "mining_coal_rev_threshold": mining_coal_rev_threshold,
            "exclude_mining_revenue": exclude_mining_revenue,
            "exclude_mining_prod_mt": exclude_mining_prod_mt,
            "mining_prod_mt_threshold": mining_prod_mt_threshold,
            "exclude_thermal_coal_mining": exclude_thermal_coal_mining,
            "thermal_coal_mining_threshold": thermal_coal_mining_threshold,
            "power_coal_rev_threshold": power_coal_rev_threshold,
            "exclude_power_revenue": exclude_power_revenue,
            "exclude_power_prod_percent": exclude_power_prod_percent,
            "power_prod_threshold_percent": power_prod_threshold_percent,
            "capacity_threshold_mw": capacity_threshold_mw,
            "exclude_capacity_mw": exclude_capacity_mw,
            "generation_thermal_threshold": generation_thermal_threshold,
            "exclude_generation_thermal": exclude_generation_thermal,
            "services_rev_threshold": services_rev_threshold,
            "exclude_services_rev": exclude_services_rev,
            "expansions_global": expansions_global
        }
        # Apply threshold filtering to unmatched sets
        filtered_sp_only = apply_thresholds(sp_only_df.copy(), params)
        filtered_ur_only = apply_thresholds(ur_only_df.copy(), params)

        # For output file, only include retained companies (not excluded) in the "Only" sheets
        sp_retained = filtered_sp_only[filtered_sp_only["Excluded"] == False].copy()
        ur_retained = filtered_ur_only[filtered_ur_only["Excluded"] == False].copy()

        # Also, produce an Excluded Companies sheet from the full datasets:
        full_filtered_sp = apply_thresholds(sp_df.copy(), params)
        full_filtered_ur = apply_thresholds(ur_df.copy(), params)
        excluded_sp = full_filtered_sp[full_filtered_sp["Excluded"] == True].copy()
        excluded_ur = full_filtered_ur[full_filtered_ur["Excluded"] == True].copy()
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)

        # Reorder columns
        sp_retained = reorder_for_excel(sp_retained)
        ur_retained = reorder_for_excel(ur_retained)
        excluded_final = reorder_for_excel(excluded_final)

        # Write output to Excel with three sheets:
        # "S&P Only": retained unmatched SPGlobal records,
        # "Urgewald Only": retained unmatched Urgewald records,
        # "Excluded Companies": all excluded companies from full datasets.
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sp_retained.to_excel(writer, sheet_name="S&P Only", index=False)
            ur_retained.to_excel(writer, sheet_name="Urgewald Only", index=False)
            excluded_final.to_excel(writer, sheet_name="Excluded Companies", index=False)

        elapsed = time.time() - start_time
        st.subheader("Results Summary")
        st.write(f"S&P Only (Retained, Unmatched): {len(sp_retained)}")
        st.write(f"Urgewald Only (Retained, Unmatched): {len(ur_retained)}")
        st.write(f"Excluded Companies (All): {len(excluded_final)}")
        st.write(f"Run Time: {elapsed:.2f} seconds")
        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="Coal_Companies_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
