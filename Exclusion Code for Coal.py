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
    """Append _1, _2, etc. to duplicate column names to avoid pyarrow errors."""
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
    used_cols = set()
    columns_before = list(df.columns)
    for final_name, patterns in rename_map.items():
        for col in columns_before:
            if col in used_cols:
                continue
            col_lower = col.lower().strip()
            for pat in patterns:
                if pat.lower().strip() in col_lower:
                    df.rename(columns={col: final_name}, inplace=True)
                    used_cols.add(col)
                    break
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
    desired_length = 46  # Force positions for columns A..AT (1..46)
    placeholders = ["(placeholder)"] * desired_length

    # Fixed positions (0-indexed)
    placeholders[6] = "Company"       # 7th column
    placeholders[41] = "BB Ticker"     # 42nd column
    placeholders[42] = "ISIN equity"   # 43rd column
    placeholders[45] = "LEI"           # 46th column

    forced_positions = {6, 41, 42, 45}
    forced_cols = {"Company", "BB Ticker", "ISIN equity", "LEI"}

    all_cols = list(df.columns)
    remaining_cols = [c for c in all_cols if c not in forced_cols]

    idx_remain = 0
    for i in range(desired_length):
        if i not in forced_positions:
            if idx_remain < len(remaining_cols):
                placeholders[i] = remaining_cols[idx_remain]
                idx_remain += 1

    leftover = remaining_cols[idx_remain:]
    final_col_order = placeholders + leftover

    for c in final_col_order:
        if c not in df.columns and c == "(placeholder)":
            df[c] = np.nan

    df = df[final_col_order]
    df = df.loc[:, ~((df.columns == "(placeholder)") & (df.isna().all()))]

    # Move "Excluded" and "Exclusion Reasons" to the end
    cols = list(df.columns)
    for c in ["Excluded", "Exclusion Reasons"]:
        if c in cols:
            cols.remove(c)
            cols.append(c)
    df = df[cols]
    return df

################################################
# 4. LOAD SPGLOBAL (AUTO-DETECT MULTI-HEADER)
################################################
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
        for col_idx in range(full_df.shape[1]):
            top_val = str(row5[col_idx]).strip()
            bot_val = str(row6[col_idx]).strip()
            combined = top_val if top_val else ""
            if bot_val and bot_val.lower() not in combined.lower():
                combined = (combined + " " + bot_val).strip() if combined else bot_val
            final_cols.append(combined.strip())
        sp_data_df = full_df.iloc[6:].reset_index(drop=True)
        sp_data_df.columns = final_cols
        sp_data_df = make_columns_unique(sp_data_df)

        rename_map_sp = {
            "SP_ENTITY_NAME":  ["sp entity name", "s&p entity name", "entity name"],
            "SP_ENTITY_ID":    ["sp entity id", "entity id"],
            "SP_COMPANY_ID":   ["sp company id", "company id"],
            "SP_ISIN":         ["sp isin", "isin code"],
            "SP_LEI":          ["sp lei", "lei code"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
            # Removed Metallurgical Coal Mining options
            "Coal Share of Revenue": ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector": ["coal industry sector", "industry sector"],
            ">10MT / >5GW": [">10mt", ">5gw"],
            "expansion": ["expansion"],
        }
        sp_data_df = fuzzy_rename_columns(sp_data_df, rename_map_sp)
        return sp_data_df
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
        full_df = pd.DataFrame(data)
        header = full_df.iloc[0].fillna("")
        ur_data_df = full_df.iloc[1:].reset_index(drop=True)
        ur_data_df.columns = header
        ur_data_df = make_columns_unique(ur_data_df)

        rename_map_ur = {
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
        ur_data_df = fuzzy_rename_columns(ur_data_df, rename_map_ur)
        return ur_data_df
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

################################################
# 6. NORMALIZE KEYS FOR MERGING
################################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)  # collapse whitespace
    s = re.sub(r'[^\w\s]', '', s)  # remove punctuation
    return s.strip()

def unify_name(r):
    sp_name = normalize_key(str(r.get("SP_ENTITY_NAME", "")))
    ur_name = normalize_key(str(r.get("Company", "")))
    return sp_name if sp_name else (ur_name if ur_name else None)

def unify_isin(r):
    sp_isin = normalize_key(str(r.get("SP_ISIN", "")))
    ur_isin = normalize_key(str(r.get("ISIN equity", "")))
    return sp_isin if sp_isin else (ur_isin if ur_isin else None)

def unify_lei(r):
    sp_lei = normalize_key(str(r.get("SP_LEI", "")))
    ur_lei = normalize_key(str(r.get("LEI", "")))
    return sp_lei if sp_lei else (ur_lei if ur_lei else None)

def unify_bbticker(r):
    return normalize_key(str(r.get("BB Ticker", "")))

################################################
# 7. PERFECT MATCHING: FIND UNMATCHED RECORDS VIA NESTED LOOPS
################################################
def find_unmatched(sp_df, ur_df):
    sp_df = sp_df.copy()
    ur_df = ur_df.copy()
    sp_df["Matched"] = False
    ur_df["Matched"] = False
    for i in range(len(sp_df)):
        sp_row = sp_df.iloc[i].to_dict()
        for j in range(len(ur_df)):
            ur_row = ur_df.iloc[j].to_dict()
            sp_vals = [
                normalize_key(str(sp_row.get("SP_LEI", ""))),
                normalize_key(str(sp_row.get("SP_ISIN", ""))),
                normalize_key(str(sp_row.get("BB Ticker", ""))),
                normalize_key(str(sp_row.get("SP_ENTITY_NAME", "")))
            ]
            ur_vals = [
                normalize_key(str(ur_row.get("LEI", ""))),
                normalize_key(str(ur_row.get("ISIN equity", ""))),
                normalize_key(str(ur_row.get("BB Ticker", ""))),
                normalize_key(str(ur_row.get("Company", "")))
            ]
            for a, b in zip(sp_vals, ur_vals):
                if a and b and a == b:
                    sp_df.at[sp_df.index[i], "Matched"] = True
                    ur_df.at[ur_df.index[j], "Matched"] = True
                    break
    sp_only = sp_df[sp_df["Matched"] == False].copy()
    ur_only = ur_df[ur_df["Matched"] == False].copy()
    sp_only.drop(columns=["Matched"], inplace=True)
    ur_only.drop(columns=["Matched"], inplace=True)
    return sp_only, ur_only

################################################
# 8. FILTER COMPANIES (Thresholds & Exclusion Logic)
################################################
def filter_companies(df,
                     # Mining thresholds:
                     exclude_mining,
                     mining_coal_rev_threshold,       # in %
                     exclude_mining_prod_mt,          # for >10MT string
                     mining_prod_mt_threshold,        # allowed max (MT)
                     exclude_thermal_coal_mining,
                     thermal_coal_mining_threshold,   # in %
                     # Power thresholds:
                     exclude_power,
                     power_coal_rev_threshold,        # in %
                     exclude_power_prod_percent,
                     power_prod_threshold_percent,    # in %
                     exclude_capacity_mw,
                     capacity_threshold_mw,           # in MW
                     exclude_generation_thermal,
                     generation_thermal_threshold,    # in %
                     # Services thresholds:
                     exclude_services,
                     services_rev_threshold,          # in %
                     exclude_services_rev,
                     # Global expansions:
                     expansions_global,
                     # Revenue threshold toggles:
                     exclude_mining_revenue,
                     exclude_power_revenue):
    exclusion_flags = []
    exclusion_reasons = []
    for idx, row in df.iterrows():
        reasons = []
        # Get sector information for revenue checks
        sector = str(row.get("Coal Industry Sector", "")).lower()
        # Numeric columns (coal revenue stored as a decimal, so multiply by 100)
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0

        # S&P identifier values (as percentages; no multiplication)
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0

        expansion_text = str(row.get("expansion", "")).lower()
        prod_str = str(row.get(">10MT / >5GW", "")).lower()

        #### MINING: apply revenue check only if sector contains "mining"
        if exclude_mining:
            if "mining" in sector:
                if exclude_mining_revenue:
                    if (coal_rev * 100) > mining_coal_rev_threshold:
                        reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
            if exclude_mining_prod_mt and (">10mt" in prod_str):
                if mining_prod_mt_threshold <= 10:
                    reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
            if exclude_thermal_coal_mining and (therm_mining_val > thermal_coal_mining_threshold):
                reasons.append(f"Thermal Coal Mining {therm_mining_val:.2f}% > {thermal_coal_mining_threshold}%")
        #### POWER: apply revenue check only if sector contains "power" or "generation"
        if exclude_power:
            if ("power" in sector or "generation" in sector):
                if exclude_power_revenue:
                    if (coal_rev * 100) > power_coal_rev_threshold:
                        reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            if exclude_generation_thermal and (gen_thermal_val > generation_thermal_threshold):
                reasons.append(f"Generation (Thermal Coal) {gen_thermal_val:.2f}% > {generation_thermal_threshold}%")
        #### SERVICES
        if exclude_services:
            if exclude_services_rev:
                if (coal_rev * 100) > services_rev_threshold:
                    reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")
        #### EXPANSIONS (applied universally)
        if expansions_global:
            for kw in expansions_global:
                if kw.lower() in expansion_text:
                    reasons.append(f"Expansion matched '{kw}'")
                    break

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")
    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

################################################
# 9. MAIN STREAMLIT APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Unmatched Retained & Excluded", layout="wide")
    st.title("Coal Exclusion Filter")

    # Sidebar: File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # Sidebar: Mining Thresholds (for SPGlobal; Metallurgical options removed)
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

    # Run Button
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

        # Perfect matching: use nested loops to mark matches
        sp_only_df, ur_only_df = find_unmatched(sp_df, ur_df)

        # Apply threshold filtering separately on:
        # - Unmatched SPGlobal records ("S&P Only")
        # - Unmatched Urgewald records ("Urgewald Only")
        filtered_sp_only = filter_companies(
            df=sp_only_df,
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global,
            exclude_mining_revenue=exclude_mining_revenue,
            exclude_power_revenue=exclude_power_revenue
        )
        filtered_ur_only = filter_companies(
            df=ur_only_df,
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global,
            exclude_mining_revenue=exclude_mining_revenue,
            exclude_power_revenue=exclude_power_revenue
        )

        # In the output file, the "S&P Only" and "Urgewald Only" sheets should include only retained companies.
        sp_retained = filtered_sp_only[filtered_sp_only["Excluded"] == False].copy()
        ur_retained = filtered_ur_only[filtered_ur_only["Excluded"] == False].copy()

        # For the exclusion file, put all excluded companies from both SPGlobal and Urgewald (the entire datasets).
        filtered_all_sp = filter_companies(
            df=sp_df,
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global,
            exclude_mining_revenue=exclude_mining_revenue,
            exclude_power_revenue=exclude_power_revenue
        )
        filtered_all_ur = filter_companies(
            df=ur_df,
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global,
            exclude_mining_revenue=exclude_mining_revenue,
            exclude_power_revenue=exclude_power_revenue
        )
        excluded_sp = filtered_all_sp[filtered_all_sp["Excluded"] == True].copy()
        excluded_ur = filtered_all_ur[filtered_all_ur["Excluded"] == True].copy()
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)

        # Reorder columns for output
        sp_retained = reorder_for_excel(sp_retained)
        ur_retained = reorder_for_excel(ur_retained)
        excluded_final = reorder_for_excel(excluded_final)

        # Write final output to Excel with three sheets:
        # - "S&P Only": retained unmatched SPGlobal companies
        # - "Urgewald Only": retained unmatched Urgewald companies
        # - "Excluded Companies": all excluded companies from SPGlobal and Urgewald
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
