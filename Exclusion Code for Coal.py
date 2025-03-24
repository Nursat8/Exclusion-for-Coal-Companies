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

################################################
# 2. FUZZY RENAME COLUMNS
################################################
def fuzzy_rename_columns(df, rename_map):
    """Rename columns based on patterns provided in rename_map."""
    used = set()
    for final, patterns in rename_map.items():
        for col in df.columns:
            if col in used:
                continue
            if any(p.lower().strip() in col.lower() for p in patterns):
                df.rename(columns={col: final}, inplace=True)
                used.add(col)
    return df

################################################
# 3. REORDER COLUMNS FOR FINAL EXCEL
################################################
def reorder_for_excel(df):
    """
    Reorder columns so that:
      - "Company" appears in column G,
      - "BB Ticker" in column AP,
      - "ISIN equity" in column AQ,
      - "LEI" in column AT,
    and then "Excluded" and "Exclusion Reasons" are moved to the end.
    """
    desired = 46
    placeholders = ["(placeholder)"] * desired
    placeholders[6] = "Company"
    placeholders[41] = "BB Ticker"
    placeholders[42] = "ISIN equity"
    placeholders[45] = "LEI"
    forced = {6, 41, 42, 45}
    all_cols = list(df.columns)
    remaining = [c for c in all_cols if c not in {"Company", "BB Ticker", "ISIN equity", "LEI"}]
    idx = 0
    for i in range(desired):
        if i not in forced and idx < len(remaining):
            placeholders[i] = remaining[idx]
            idx += 1
    final_order = placeholders + remaining[idx:]
    df = df[[c for c in final_order if c in df.columns]]
    # Move "Excluded" and "Exclusion Reasons" to end:
    cols = list(df.columns)
    for c in ["Excluded", "Exclusion Reasons"]:
        if c in cols:
            cols.remove(c)
            cols.append(c)
    return df[cols]

################################################
# 4. LOAD SPGLOBAL (MULTI-HEADER DETECTION)
################################################
def load_spglobal(file, sheet_name="Sheet1"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        df = pd.DataFrame(data)
        if len(df) < 6:
            raise ValueError("Not enough rows in SPGlobal file.")
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
        sp = df.iloc[6:].reset_index(drop=True)
        sp.columns = cols
        sp = make_columns_unique(sp)
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
        sp = fuzzy_rename_columns(sp, rename_map)
        return sp
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()

################################################
# 5. LOAD URGEWALD (SINGLE HEADER)
################################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")
        df = pd.DataFrame(data)
        header = df.iloc[0].fillna("")
        ur = df.iloc[1:].reset_index(drop=True)
        ur.columns = header
        ur = make_columns_unique(ur)
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
        ur = fuzzy_rename_columns(ur, rename_map)
        return ur
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

################################################
# 6. NORMALIZE KEYS FOR MATCHING (Vectorized)
################################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

def add_normalized_keys(df, source):
    # For SPGlobal, source keys: SP_ENTITY_NAME, SP_ISIN, SP_LEI; for Urgewald: Company, ISIN equity, LEI; plus BB Ticker for both.
    if source == "sp":
        df["norm_name"] = df["SP_ENTITY_NAME"].astype(str).apply(normalize_key)
        df["norm_isin"] = df["SP_ISIN"].astype(str).apply(normalize_key)
        df["norm_lei"]  = df["SP_LEI"].astype(str).apply(normalize_key)
    else:
        df["norm_name"] = df["Company"].astype(str).apply(normalize_key)
        df["norm_isin"] = df["ISIN equity"].astype(str).apply(normalize_key)
        df["norm_lei"]  = df["LEI"].astype(str).apply(normalize_key)
    # For both, add normalized BB Ticker
    df["norm_bbticker"] = df["BB Ticker"].astype(str).apply(normalize_key)
    return df

def vectorized_match(sp_df, ur_df):
    # Add normalized key columns
    sp_df = add_normalized_keys(sp_df.copy(), "sp")
    ur_df = add_normalized_keys(ur_df.copy(), "ur")
    # For each key field, create boolean series indicating if value is in the other dataset:
    sp_df["match_name"] = sp_df["norm_name"].isin(ur_df["norm_name"])
    sp_df["match_isin"] = sp_df["norm_isin"].isin(ur_df["norm_isin"])
    sp_df["match_lei"] = sp_df["norm_lei"].isin(ur_df["norm_lei"])
    sp_df["match_bbticker"] = sp_df["norm_bbticker"].isin(ur_df["norm_bbticker"])
    sp_df["Merged"] = sp_df[["match_name", "match_isin", "match_lei", "match_bbticker"]].any(axis=1)
    # For Urgewald:
    ur_df["match_name"] = ur_df["norm_name"].isin(sp_df["norm_name"])
    ur_df["match_isin"] = ur_df["norm_isin"].isin(sp_df["norm_isin"])
    ur_df["match_lei"] = ur_df["norm_lei"].isin(sp_df["norm_lei"])
    ur_df["match_bbticker"] = ur_df["norm_bbticker"].isin(sp_df["norm_bbticker"])
    ur_df["Merged"] = ur_df[["match_name", "match_isin", "match_lei", "match_bbticker"]].any(axis=1)
    # Drop temporary normalized columns
    for col in ["norm_name", "norm_isin", "norm_lei", "norm_bbticker",
                "match_name", "match_isin", "match_lei", "match_bbticker"]:
        sp_df.drop(columns=[col], inplace=True)
        ur_df.drop(columns=[col], inplace=True)
    return sp_df, ur_df

################################################
# 7. VECTORIZE THRESHOLD FILTERING VIA APPLY
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
    # Convert numeric values
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

    # For Mining: apply revenue check only if sector contains "mining"
    if "mining" in sector:
        if exclude_mining_revenue and (coal_rev * 100) > mining_coal_rev_threshold:
            reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
    if exclude_mining_prod_mt and (">10mt" in prod_str) and (mining_prod_mt_threshold <= 10):
        reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
    if exclude_thermal_coal_mining and (thermal_val > thermal_coal_mining_threshold):
        reasons.append(f"Thermal Coal Mining {thermal_val:.2f}% > {thermal_coal_mining_threshold}%")
    # For Power: apply revenue check only if sector contains "power" or "generation"
    if ("power" in sector or "generation" in sector):
        if exclude_power_revenue and (coal_rev * 100) > power_coal_rev_threshold:
            reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
        if exclude_power_prod_percent and (coal_power * 100) > power_prod_threshold_percent:
            reasons.append(f"Coal power production {coal_power*100:.2f}% > {power_prod_threshold_percent}%")
        if exclude_capacity_mw and (capacity > capacity_threshold_mw):
            reasons.append(f"Installed capacity {capacity:.2f}MW > {capacity_threshold_mw}MW")
        if exclude_generation_thermal and (gen_val > generation_thermal_threshold):
            reasons.append(f"Generation (Thermal Coal) {gen_val:.2f}% > {generation_thermal_threshold}%")
    # Services:
    if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
        reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")
    # Global expansion:
    if expansions_global:
        for kw in expansions_global:
            if kw.lower() in expansion:
                reasons.append(f"Expansion matched '{kw}'")
                break
    return pd.Series([len(reasons) > 0, "; ".join(reasons)])

def apply_thresholds(df, **params):
    res = df.apply(lambda row: compute_exclusion(row, **params), axis=1)
    df["Excluded"] = res[0]
    df["Exclusion Reasons"] = res[1]
    return df

################################################
# 8. MAIN STREAMLIT APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Optimized Unmatched", layout="wide")
    st.title("Coal Exclusion Filter – S&P Only and Urgewald Only (Retained & Excluded)")

    # Sidebar: File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # Sidebar: Mining Thresholds (SPGlobal – Metallurgical removed)
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

    # Start timer
    start_time = time.time()

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # Load data
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

        # --- Vectorized matching: add normalized keys and mark merged records
        sp_df = add_normalized_keys(sp_df.copy(), "sp")
        ur_df = add_normalized_keys(ur_df.copy(), "ur")
        sp_df["Merged"] = sp_df["norm_name"].isin(ur_df["norm_name"]) | \
                          sp_df["norm_isin"].isin(ur_df["norm_isin"]) | \
                          sp_df["norm_lei"].isin(ur_df["norm_lei"]) | \
                          sp_df["norm_bbticker"].isin(ur_df["norm_bbticker"])
        ur_df["Merged"] = ur_df["norm_name"].isin(sp_df["norm_name"]) | \
                          ur_df["norm_isin"].isin(sp_df["norm_isin"]) | \
                          ur_df["norm_lei"].isin(sp_df["norm_lei"]) | \
                          ur_df["norm_bbticker"].isin(sp_df["norm_bbticker"])
        sp_df.drop(columns=["norm_name", "norm_isin", "norm_lei", "norm_bbticker"], inplace=True)
        ur_df.drop(columns=["norm_name", "norm_isin", "norm_lei", "norm_bbticker"], inplace=True)

        # Unmatched records become "Only" sheets:
        sp_only_df = sp_df[~sp_df["Merged"]].copy()
        ur_only_df = ur_df[~ur_df["Merged"]].copy()
        # For S&P Only, keep only records with nonzero values in either Thermal Coal Mining or Generation (Thermal Coal)
        sp_only_df = sp_only_df[
            (pd.to_numeric(sp_only_df["Thermal Coal Mining"], errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(sp_only_df["Generation (Thermal Coal)"], errors='coerce').fillna(0) > 0)
        ].copy()

        # Apply threshold filtering on unmatched sets using vectorized .apply:
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
        filtered_sp_only = sp_only_df.apply(lambda row: compute_exclusion(row, **params), axis=1)
        sp_only_df["Excluded"] = filtered_sp_only[0]
        sp_only_df["Exclusion Reasons"] = filtered_sp_only[1]
        filtered_ur_only = ur_only_df.apply(lambda row: compute_exclusion(row, **params), axis=1)
        ur_only_df["Excluded"] = filtered_ur_only[0]
        ur_only_df["Exclusion Reasons"] = filtered_ur_only[1]

        # Keep only retained companies in the "Only" sheets (for output file)
        sp_retained = sp_only_df[sp_only_df["Excluded"] == False].copy()
        ur_retained = ur_only_df[ur_only_df["Excluded"] == False].copy()

        # Also, produce the Excluded Companies sheet from full datasets:
        filtered_all_sp = sp_df.apply(lambda row: compute_exclusion(row, **params), axis=1)
        sp_df["Excluded"] = filtered_all_sp[0]
        sp_df["Exclusion Reasons"] = filtered_all_sp[1]
        filtered_all_ur = ur_df.apply(lambda row: compute_exclusion(row, **params), axis=1)
        ur_df["Excluded"] = filtered_all_ur[0]
        ur_df["Exclusion Reasons"] = filtered_all_ur[1]
        excluded_sp = sp_df[sp_df["Excluded"] == True].copy()
        excluded_ur = ur_df[ur_df["Excluded"] == True].copy()
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)

        # Reorder columns for final output
        sp_retained = reorder_for_excel(sp_retained)
        ur_retained = reorder_for_excel(ur_retained)
        excluded_final = reorder_for_excel(excluded_final)

        # Write output to Excel with three sheets:
        # - "S&P Only": unmatched & retained SPGlobal companies
        # - "Urgewald Only": unmatched & retained Urgewald companies
        # - "Excluded Companies": all excluded companies from both datasets
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
