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
    Append _1, _2, etc. to duplicate column names to avoid pyarrow errors.
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
# Force "Company" in column G, "BB Ticker" in AP,
# "ISIN equity" in AQ, "LEI" in AT,
# then move "Excluded" and "Exclusion Reasons" to the very end.
################################################
def reorder_for_excel(df):
    desired_length = 46  # Force positions for columns A..AT (1..46)
    placeholders = ["(placeholder)"] * desired_length

    # Fixed positions (0-indexed):
    placeholders[6]   = "Company"      # Column G (7th)
    placeholders[41]  = "BB Ticker"    # Column AP (42nd)
    placeholders[42]  = "ISIN equity"  # Column AQ (43rd)
    placeholders[45]  = "LEI"          # Column AT (46th)

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
            "Generation (Thermal Coal)":       ["generation (thermal coal)"],
            "Thermal Coal Mining":             ["thermal coal mining"],
            "Metallurgical Coal Mining":       ["metallurgical coal mining"],
            "Coal Share of Revenue":           ["coal share of revenue"],
            "Coal Share of Power Production":  ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector":            ["coal industry sector", "industry sector"],
            ">10MT / >5GW":                    [">10mt", ">5gw"],
            "expansion":                       ["expansion"],
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
            "Metallurgical Coal Mining":     ["metallurgical coal mining"],
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

################################################
# 7. MERGE URGEWALD INTO SPGLOBAL (Optimized with Dictionaries)
################################################
def merge_ur_into_sp(sp_df, ur_df):
    sp_records = sp_df.to_dict("records")
    # Add a "Merged" flag to each SP record; default False.
    for rec in sp_records:
        rec["Merged"] = False
    # Build dictionaries for quick lookup
    name_dict = {}
    isin_dict = {}
    lei_dict = {}
    for i, rec in enumerate(sp_records):
        n = unify_name(rec)
        if n:
            name_dict.setdefault(n, []).append(i)
        iis = unify_isin(rec)
        if iis:
            isin_dict.setdefault(iis, []).append(i)
        l = unify_lei(rec)
        if l:
            lei_dict.setdefault(l, []).append(i)
    merged_records = sp_records.copy()
    ur_only_records = []
    for _, ur_row in ur_df.iterrows():
        n = unify_name(ur_row)
        iis = unify_isin(ur_row)
        l = unify_lei(ur_row)
        indices = set()
        if n and n in name_dict:
            indices.update(name_dict[n])
        if iis and iis in isin_dict:
            indices.update(isin_dict[iis])
        if l and l in lei_dict:
            indices.update(lei_dict[l])
        if indices:
            index = list(indices)[0]
            for k, v in ur_row.items():
                if (k not in merged_records[index]) or (merged_records[index][k] is None) or (str(merged_records[index][k]).strip() == ""):
                    merged_records[index][k] = v
            merged_records[index]["Merged"] = True
        else:
            ur_only_records.append(ur_row.to_dict())
    merged_df = pd.DataFrame(merged_records)
    ur_only_df = pd.DataFrame(ur_only_records)
    return merged_df, ur_only_df

################################################
# 8. FILTER COMPANIES (Thresholds & Exclusion Logic)
################################################
def filter_companies(
    df,
    # Mining thresholds:
    exclude_mining,
    mining_coal_rev_threshold,        # in %
    exclude_mining_prod_mt,           # for >10MT string
    mining_prod_mt_threshold,         # allowed max (MT)
    exclude_thermal_coal_mining,
    thermal_coal_mining_threshold,    # in %
    exclude_metallurgical_coal_mining,
    metallurgical_coal_mining_threshold,  # in %
    # Power thresholds:
    exclude_power,
    power_coal_rev_threshold,         # in %
    exclude_power_prod_percent,
    power_prod_threshold_percent,     # in %
    exclude_capacity_mw,
    capacity_threshold_mw,            # in MW
    exclude_generation_thermal,
    generation_thermal_threshold,     # in %
    # Services thresholds:
    exclude_services,
    services_rev_threshold,           # in %
    exclude_services_rev,
    # Global expansions:
    expansions_global,
    # Revenue threshold toggles:
    exclude_mining_revenue,
    exclude_power_revenue
):
    exclusion_flags = []
    exclusion_reasons = []
    for idx, row in df.iterrows():
        reasons = []
        # For revenue and capacity checks, these are applied based on values.
        sector_val = str(row.get("Coal Industry Sector", "")).lower()
        # For the three S&P identifiers, we always check regardless of sector.
        
        # Numeric columns (coal revenue stored as decimal, so multiply by 100)
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0
        
        # Business involvement columns (percentages; no multiplication)
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0
        
        expansion_text = str(row.get("expansion", "")).lower()
        prod_str = str(row.get(">10MT / >5GW", "")).lower()
        
        #### MINING (S&P identifier checks applied universally)
        if exclude_mining:
            if exclude_mining_revenue:
                if (coal_rev * 100) > mining_coal_rev_threshold:
                    reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
            if exclude_mining_prod_mt and (">10mt" in prod_str):
                if mining_prod_mt_threshold <= 10:
                    reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
            if exclude_thermal_coal_mining and (therm_mining_val > thermal_coal_mining_threshold):
                reasons.append(f"Thermal Coal Mining {therm_mining_val:.2f}% > {thermal_coal_mining_threshold}%")
            if exclude_metallurgical_coal_mining and (met_coal_val > metallurgical_coal_mining_threshold):
                reasons.append(f"Metallurgical Coal Mining {met_coal_val:.2f}% > {metallurgical_coal_mining_threshold}%")
        #### POWER
        if exclude_power:
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
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
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
# 8. MAIN STREAMLIT APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter (Fuzzy Columns)", layout="wide")
    st.title("Coal Exclusion Filter")

    # File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # Mining Thresholds
    with st.sidebar.expander("Mining Thresholds", expanded=True):
        # Revenue check for mining is controlled by this checkbox:
        exclude_mining_revenue = st.checkbox("Exclude if coal revenue > threshold? (Mining)", value=True)
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0)
        exclude_mining_prod_mt = st.checkbox("Exclude if >10MT indicated?", value=True)
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0)
        # Note: GW check removed as per request.
        exclude_thermal_coal_mining = st.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False)
        thermal_coal_mining_threshold = st.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)
        exclude_metallurgical_coal_mining = st.checkbox("Exclude if Metallurgical Coal Mining > threshold?", value=False)
        metallurgical_coal_mining_threshold = st.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    # Power Thresholds
    with st.sidebar.expander("Power Thresholds", expanded=True):
        exclude_power_revenue = st.checkbox("Exclude if coal revenue > threshold? (Power)", value=True)
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True)
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True)
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0)
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    # Services Thresholds
    with st.sidebar.expander("Services Thresholds", expanded=False):
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)

    # Global Expansion
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

        # Load SPGlobal
        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or not loaded.")
            return
        sp_df = make_columns_unique(sp_df)

        # Load Urgewald
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return
        ur_df = make_columns_unique(ur_df)

        # Merge UR into SP (using optimized dictionary lookup with normalization)
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)

        # Apply filtering on merged and UR-only sets
        filtered_merged = filter_companies(
            df=merged_df,
            # Mining thresholds:
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            # Power thresholds:
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services thresholds:
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # Global Expansions:
            expansions_global=expansions_global,
            # Revenue threshold toggles:
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
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
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

        # Separate merged dataset into Excluded and Retained
        excluded_df = filtered_merged[filtered_merged["Excluded"] == True].copy()
        retained_df = filtered_merged[filtered_merged["Excluded"] == False].copy()

        # --- NEW: Create additional sheet for S&P companies that did not merge (SP-only)
        # These are the SPGlobal records (in the merged_df) with Merged==False and
        # with any non-empty value in one or more of the three fields:
        # "Thermal Coal Mining", "Metallurgical Coal Mining", "Generation (Thermal Coal)"
        sp_only_df = merged_df[(merged_df["Merged"] == False) & (
            merged_df["Thermal Coal Mining"].notna() & (merged_df["Thermal Coal Mining"] != "") |
            merged_df["Metallurgical Coal Mining"].notna() & (merged_df["Metallurgical Coal Mining"] != "") |
            merged_df["Generation (Thermal Coal)"].notna() & (merged_df["Generation (Thermal Coal)"] != "")
        )].copy()
        # Drop the "Merged" column for final output
        if "Merged" in sp_only_df.columns:
            sp_only_df.drop(columns=["Merged"], inplace=True)

        # Define final columns (for all sheets)
        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Company","ISIN equity","LEI","BB Ticker",
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "Coal Share of Revenue",
            "expansion",
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            "Excluded","Exclusion Reasons"
        ]
        def ensure_cols_exist(df_):
            for c in final_cols:
                if c not in df_.columns:
                    df_[c] = np.nan
            return df_
        excluded_df = ensure_cols_exist(excluded_df)[final_cols]
        retained_df = ensure_cols_exist(retained_df)[final_cols]
        filtered_ur_only = ensure_cols_exist(filtered_ur_only)[final_cols]
        sp_only_df = ensure_cols_exist(sp_only_df)[final_cols]

        # Reorder columns as required
        excluded_df = reorder_for_excel(excluded_df)
        retained_df = reorder_for_excel(retained_df)
        filtered_ur_only = reorder_for_excel(filtered_ur_only)
        sp_only_df = reorder_for_excel(sp_only_df)

        # Write output to Excel with four sheets:
        # - "Excluded Companies" (merged)
        # - "Retained Companies" (merged)
        # - "Urgewald Only" (UR-only)
        # - "S&P Only" (SPGlobal companies that did not merge and have values in at least one S&P field)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            filtered_ur_only.to_excel(writer, sheet_name="Urgewald Only", index=False)
            sp_only_df.to_excel(writer, sheet_name="S&P Only", index=False)

        elapsed = time.time() - start_time

        st.subheader("Results Summary")
        st.write(f"Merged Total: {len(filtered_merged)}")
        st.write(f"Excluded (Merged): {len(excluded_df)}")
        st.write(f"Retained (Merged): {len(retained_df)}")
        st.write(f"Urgewald Only: {len(filtered_ur_only)}")
        st.write(f"S&P Only (Unmerged with S&P values): {len(sp_only_df)}")
        st.write(f"Run Time: {elapsed:.2f} seconds")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="Coal Companies Exclusion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
