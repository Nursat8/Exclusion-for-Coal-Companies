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
    If duplicate column names exist, append _1, _2, etc.
    This avoids errors from pyarrow in st.dataframe.
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
    Given a dictionary mapping final names to a list of patterns,
    rename any column that contains one of the patterns (case-insensitive).
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
#    Force "Company" in column G, "BB Ticker" in AP, "ISIN equity" in AQ, "LEI" in AT.
#    Then move "Excluded" and "Exclusion Reasons" to the very end.
################################################
def reorder_for_excel(df):
    desired_length = 46  # We force positions for the first 46 columns
    placeholders = ["(placeholder)"] * desired_length

    # Force required columns at fixed positions (0-indexed)
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

    # Create empty columns for any remaining placeholder if needed
    for c in final_col_order:
        if c not in df.columns and c == "(placeholder)":
            df[c] = np.nan

    df = df[final_col_order]
    # Drop placeholder columns that are completely empty
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
    """
    Load SPGlobal file assuming:
      - Row 5 (index 4): contains ID columns (e.g. SP_ENTITY_NAME)
      - Row 6 (index 5): contains additional metrics (e.g. Generation (Thermal Coal))
      - Data starts at row 7 (index 6) onward.
    Then perform fuzzy renaming.
    """
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
            "Thermal Coal Mining":       ["thermal coal mining"],
            "Metallurgical Coal Mining": ["metallurgical coal mining"],
            "Coal Share of Revenue":           ["coal share of revenue"],
            "Coal Share of Power Production":  ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector":            ["coal industry sector", "industry sector"],
            ">10MT / >5GW":                    [">10mt", ">5gw"],
            "expansion":                       ["expansion"]
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
    """
    Load Urgewald file assuming:
      - Row 1 (index 0) is the header.
      - Data starts at row 2 (index 1) onward.
    Then perform fuzzy renaming.
    """
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
# 6. MERGE URGEWALD INTO SPGLOBAL
#    If a UR row matches an SP row by (Name OR ISIN OR LEI),
#    merge non-empty values; otherwise, keep UR row separately.
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
    merged_records = []
    ur_only_records = []

    for rec in sp_records:
        rec["Source"] = "SP"
        merged_records.append(rec)

    for _, ur_row in ur_df.iterrows():
        merged_flag = False
        for rec in merged_records:
            if ((unify_name(rec) and unify_name(ur_row) and unify_name(rec) == unify_name(ur_row)) or
                (unify_isin(rec) and unify_isin(ur_row) and unify_isin(rec) == unify_isin(ur_row)) or
                (unify_lei(rec) and unify_lei(ur_row) and unify_lei(rec) == unify_lei(ur_row))):
                for k, v in ur_row.items():
                    if (k not in rec) or (rec[k] is None) or (str(rec[k]).strip() == ""):
                        rec[k] = v
                rec["Source"] = "SP+UR"
                merged_flag = True
                break
        if not merged_flag:
            new_rec = ur_row.to_dict()
            new_rec["Source"] = "UR"
            ur_only_records.append(new_rec)

    merged_df = pd.DataFrame(merged_records)
    ur_only_df = pd.DataFrame(ur_only_records)
    merged_df.drop(columns=["Source"], inplace=True, errors="ignore")
    ur_only_df.drop(columns=["Source"], inplace=True, errors="ignore")
    return merged_df, ur_only_df

################################################
# 7. FILTER COMPANIES (Thresholds & Exclusion Logic)
################################################
def filter_companies(
    df,
    # Mining thresholds:
    exclude_mining,
    mining_coal_rev_threshold,        # in %
    exclude_mining_prod_mt,           # for >10MT string
    mining_prod_mt_threshold,         # allowed max (MT)
    exclude_mining_prod_gw,           # for >5GW string
    mining_prod_threshold_gw,         # allowed max (GW)
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
    # Global expansion keywords:
    expansions_global
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

        # Business involvement columns (percentages, not multiplied)
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # Get production string from ">10MT / >5GW" column
        prod_str = str(row.get(">10MT / >5GW", "")).lower()

        #### MINING ####
        if is_mining and exclude_mining:
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
        if is_power and exclude_power:
            if (coal_rev * 100) > power_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            if exclude_generation_thermal and (gen_thermal_val > generation_thermal_threshold):
                reasons.append(f"Generation (Thermal Coal) {gen_thermal_val:.2f}% > {generation_thermal_threshold}%")
        #### SERVICES ####
        if is_services and exclude_services:
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
        exclude_mining = st.checkbox("Urgewald Coal revenue (%)", value=True)
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0)
        exclude_mining_prod_mt = st.checkbox("Exclude if MT indicated?", value=True)
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0)
        exclude_thermal_coal_mining = st.checkbox("S&P Thermal Coal Mining produc(%)?", value=False)
        thermal_coal_mining_threshold = st.number_input("S&P Thermal Coal Mining revenue (%)", value=20.0)
        exclude_metallurgical_coal_mining = st.checkbox("S&P Metallurgical Coal Mining > threshold?", value=False)
        metallurgical_coal_mining_threshold = st.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    # Power Thresholds
    with st.sidebar.expander("Power Thresholds", expanded=True):
        exclude_power = st.checkbox("Coal revenue (%)", value=True)
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True)
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True)
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0)
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)
       

    # Services Thresholds
    with st.sidebar.expander("Services Thresholds", expanded=False):
        exclude_services = st.checkbox("Coal revenue (%)", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False)

    # Global Expansion
    with st.sidebar.expander("Global Expansion Exclusion", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansions_global = st.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[])

    st.sidebar.markdown("---")

    # Run Button
    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # Load SPGlobal
        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or could not be loaded.")
            return
        sp_df = make_columns_unique(sp_df)
        st.subheader("SPGlobal Data (first 5 rows)")
        st.dataframe(sp_df.head(5))

        # Load Urgewald
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or could not be loaded.")
            return
        ur_df = make_columns_unique(ur_df)
        st.subheader("Urgewald Data (first 5 rows)")
        st.dataframe(ur_df.head(5))

        # Merge UR into SP
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)
        st.write(f"Merged dataset shape: {merged_df.shape}")
        st.write(f"Urgewald-only dataset shape: {ur_only_df.shape}")

        # Apply filtering
        filtered_merged = filter_companies(
            df=merged_df,
            # Mining
            exclude_mining=exclude_mining,
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
            exclude_power=exclude_power,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services
            exclude_services=exclude_services,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # Expansions
            expansions_global=expansions_global
        )

        filtered_ur_only = filter_companies(
            df=ur_only_df,
            # Mining
            exclude_mining=exclude_mining,
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
            exclude_power=exclude_power,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services
            exclude_services=exclude_services,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # Expansions
            expansions_global=expansions_global
        )

        # Separate merged dataset into Excluded and Retained
        excluded_df = filtered_merged[filtered_merged["Excluded"]==True].copy()
        retained_df = filtered_merged[filtered_merged["Excluded"]==False].copy()

        # Define final columns (you can adjust as needed)
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
        # - Excluded Companies (merged)
        # - Retained Companies (merged)
        # - Urgewald Only (UR-only)
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
            file_name="Coal Companies Exclusion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
