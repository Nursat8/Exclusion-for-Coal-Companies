import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

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
    Rename columns based on a mapping.
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
# Force "Company" in column G, "BB Ticker" in AP, "ISIN equity" in AQ, "LEI" in AT.
# Then move "Excluded" and "Exclusion Reasons" to the very end.
################################################
def reorder_for_excel(df):
    desired_length = 46  # Force positions for columns A..AT (1..46)
    placeholders = ["(placeholder)"] * desired_length

    # Fixed positions (0-indexed):
    placeholders[6]   = "Company"      # G
    placeholders[41]  = "BB Ticker"    # AP
    placeholders[42]  = "ISIN equity"  # AQ
    placeholders[45]  = "LEI"          # AT

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

    # Create any missing columns for placeholders if needed
    for c in final_col_order:
        if c not in df.columns and c == "(placeholder)":
            df[c] = np.nan

    df = df[final_col_order]
    # Drop any placeholder columns that are completely empty
    df = df.loc[:, ~((df.columns == "(placeholder)") & (df.isna().all()))]

    # Move "Excluded" and "Exclusion Reasons" to the very end
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
# 6. MERGE URGEWALD INTO SPGLOBAL (Optimized)
################################################
def unify_name(r):
    sp_name = str(r.get("SP_ENTITY_NAME", "")).strip().lower()
    ur_name = str(r.get("Company", "")).strip().lower()
    return sp_name if sp_name else (ur_name if ur_name else None)

def unify_isin(r):
    sp_isin = str(r.get("SP_ISIN", "")).strip().lower()
    ur_isin = str(r.get("ISIN equity", "")).strip().lower()
    return sp_isin if sp_isin else (ur_isin if ur_isin else None)

def unify_lei(r):
    sp_lei = str(r.get("SP_LEI", "")).strip().lower()
    ur_lei = str(r.get("LEI", "")).strip().lower()
    return sp_lei if sp_lei else (ur_lei if ur_lei else None)

def merge_ur_into_sp(sp_df, ur_df):
    sp_records = sp_df.to_dict("records")
    # Build dictionaries for SP keys
    name_dict = {}
    isin_dict = {}
    lei_dict = {}
    for i, rec in enumerate(sp_records):
        n = unify_name(rec)
        if n:
            if n not in name_dict:
                name_dict[n] = i
        iis = unify_isin(rec)
        if iis:
            if iis not in isin_dict:
                isin_dict[iis] = i
        l = unify_lei(rec)
        if l:
            if l not in lei_dict:
                lei_dict[l] = i

    merged_indices = set()
    ur_only_records = []
    for _, ur_row in ur_df.iterrows():
        merged_flag = False
        n = unify_name(ur_row)
        iis = unify_isin(ur_row)
        l = unify_lei(ur_row)
        index = None
        if n and n in name_dict:
            index = name_dict[n]
        elif iis and iis in isin_dict:
            index = isin_dict[iis]
        elif l and l in lei_dict:
            index = lei_dict[l]
        if index is not None:
            # Merge non-empty values from UR row into sp_records[index]
            for k, v in ur_row.items():
                if (k not in sp_records[index]) or (sp_records[index][k] is None) or (str(sp_records[index][k]).strip() == ""):
                    sp_records[index][k] = v
            merged_flag = True
        if not merged_flag:
            ur_only_records.append(ur_row.to_dict())
    merged_df = pd.DataFrame(sp_records)
    ur_only_df = pd.DataFrame(ur_only_records)
    return merged_df, ur_only_df

################################################
# 7. FILTER COMPANIES (Exclusion Logic)
################################################
def filter_companies(
    df,
    # Mining thresholds:
    exclude_mining,  # Always True (sector filtering is on)
    mining_coal_rev_threshold,        # in %
    exclude_mining_prod_mt,           # for >10MT string check
    mining_prod_mt_threshold,         # allowed max (MT)
    exclude_mining_prod_gw,           # for >5GW string check
    mining_prod_threshold_gw,         # allowed max (GW)
    exclude_thermal_coal_mining,
    thermal_coal_mining_threshold,    # in %
    exclude_metallurgical_coal_mining,
    metallurgical_coal_mining_threshold,  # in %
    # Power thresholds:
    exclude_power,   # Always True (sector filtering is on)
    power_coal_rev_threshold,         # in %
    exclude_power_prod_percent,
    power_prod_threshold_percent,     # in %
    exclude_capacity_mw,
    capacity_threshold_mw,            # in MW
    exclude_generation_thermal,
    generation_thermal_threshold,     # in %
    # Services thresholds:
    exclude_services,   # Always True (sector filtering is on)
    services_rev_threshold,           # in %
    exclude_services_rev,
    # Global expansions:
    expansions_global,
    # New options for revenue thresholds:
    apply_mining_coal_rev,  # boolean to turn off/on mining coal revenue check
    apply_power_coal_rev     # boolean to turn off/on power coal revenue check
):
    exclusion_flags = []
    exclusion_reasons = []
    for idx, row in df.iterrows():
        reasons = []
        sector_val = str(row.get("Coal Industry Sector", "")).lower()
        is_mining = ("mining" in sector_val)
        is_power = ("power" in sector_val) or ("generation" in sector_val)
        is_services = ("service" in sector_val)
        expansion_text = str(row.get("expansion", "")).lower()

        # Numeric columns
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0

        # Business involvement columns (percentages; not multiplied)
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # Production string from ">10MT / >5GW"
        prod_str = str(row.get(">10MT / >5GW", "")).lower()

        #### MINING ####
        if is_mining:
            if apply_mining_coal_rev:
                if (coal_rev * 100) > mining_coal_rev_threshold:
                    reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
            if exclude_mining_prod_mt and (">10mt" in prod_str):
                if mining_prod_mt_threshold <= 10:
                    reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
            if exclude_mining_prod_gw and (">5gw" in prod_str):
                if mining_prod_threshold_gw <= 5:
                    reasons.append(f">5GW indicated (threshold {mining_prod_threshold_gw}GW)")
            if exclude_thermal_coal_mining and (therm_mining_val > thermal_coal_mining_threshold):
                reasons.append(f"Thermal Coal Mining {therm_mining_val:.2f}% > {thermal_coal_mining_threshold}%")
            if exclude_metallurgical_coal_mining and (met_coal_val > metallurgical_coal_mining_threshold):
                reasons.append(f"Metallurgical Coal Mining {met_coal_val:.2f}% > {metallurgical_coal_mining_threshold}%")
        #### POWER ####
        if is_power:
            if apply_power_coal_rev:
                if (coal_rev * 100) > power_coal_rev_threshold:
                    reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            if exclude_generation_thermal and (gen_thermal_val > generation_thermal_threshold):
                reasons.append(f"Generation (Thermal Coal) {gen_thermal_val:.2f}% > {generation_thermal_threshold}% (Power)")
        #### SERVICES ####
        if is_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")
        #### EXPANSIONS ####
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
    st.set_page_config(page_title="Coal Exclusion Filter (Optimized)", layout="wide")
    st.title("Coal Exclusion Filter with Dynamic Column Detection & Optimized Merge")

    # 8.1 File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    # 8.2 Mining Thresholds
    with st.sidebar.expander("Mining Thresholds", expanded=True):
        # Removed "Exclude Mining Sector?" checkbox; always apply mining filters.
        apply_mining_coal_rev = st.checkbox("Apply Mining: Max coal revenue threshold?", value=True)
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0)
        exclude_mining_prod_mt = st.checkbox("Exclude if >10MT indicated?", value=True)
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0)
        exclude_mining_prod_gw = st.checkbox("Exclude if >5GW indicated?", value=True)
        mining_prod_threshold_gw = st.number_input("Mining: Max production (GW)", value=5.0)
        exclude_thermal_coal_mining = st.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False)
        thermal_coal_mining_threshold = st.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)
        exclude_metallurgical_coal_mining = st.checkbox("Exclude if Metallurgical Coal Mining > threshold?", value=False)
        metallurgical_coal_mining_threshold = st.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    # 8.3 Power Thresholds
    with st.sidebar.expander("Power Thresholds", expanded=True):
        # Removed "Exclude Power Sector?" checkbox; always apply power filters.
        apply_power_coal_rev = st.checkbox("Apply Power: Max coal revenue threshold?", value=True)
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True)
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True)
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0)
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    # 8.4 Services Thresholds
    with st.sidebar.expander("Services Thresholds", expanded=False):
        # Removed "Exclude Services Sector?" checkbox; always apply services filters.
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)

    # 8.5 Global Expansion
    with st.sidebar.expander("Global Expansion Exclusion", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansions_global = st.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[])

    st.sidebar.markdown("---")

    # 8.6 Run Button
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
        st.subheader("SPGlobal Data (first 5 rows)")
        st.dataframe(sp_df.head(5))

        # Load Urgewald
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return
        ur_df = make_columns_unique(ur_df)
        st.subheader("Urgewald Data (first 5 rows)")
        st.dataframe(ur_df.head(5))

        # Merge UR into SP (optimized using dictionaries)
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)
        st.write(f"Merged dataset shape: {merged_df.shape}")
        st.write(f"Urgewald-only dataset shape: {ur_only_df.shape}")

        # Apply filtering
        filtered_merged = filter_companies(
            df=merged_df,
            # Mining
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_gw=mining_prod_threshold_gw,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            # Power
            exclude_power=True,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services
            exclude_services=True,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # Global Expansions
            expansions_global=expansions_global,
            # Revenue thresholds toggles:
            apply_mining_coal_rev=apply_mining_coal_rev,
            apply_power_coal_rev=apply_power_coal_rev
        )

        filtered_ur_only = filter_companies(
            df=ur_only_df,
            exclude_mining=True,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_gw=mining_prod_threshold_gw,
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
            apply_mining_coal_rev=apply_mining_coal_rev,
            apply_power_coal_rev=apply_power_coal_rev
        )

        # Separate merged dataset into Excluded and Retained
        excluded_df = filtered_merged[filtered_merged["Excluded"] == True].copy()
        retained_df = filtered_merged[filtered_merged["Excluded"] == False].copy()

        # Define final columns
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

        # Reorder columns as required
        excluded_df = reorder_for_excel(excluded_df)
        retained_df = reorder_for_excel(retained_df)
        filtered_ur_only = reorder_for_excel(filtered_ur_only)

        # Write to Excel with three sheets:
        #   "Excluded Companies" (merged), "Retained Companies" (merged), and "Urgewald Only"
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            filtered_ur_only.to_excel(writer, sheet_name="Urgewald Only", index=False)

        st.subheader("Results Summary")
        st.write(f"Merged Total: {len(filtered_merged)}")
        st.write(f"Excluded (Merged): {len(excluded_df)}")
        st.write(f"Retained (Merged): {len(retained_df)}")
        st.write(f"Urgewald Only: {len(filtered_ur_only)}")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
