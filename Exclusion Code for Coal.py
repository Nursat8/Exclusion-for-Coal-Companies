import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
import time
import re

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
    """
    Rename columns based on provided patterns.
    rename_map: { final_name: [pattern1, pattern2, ...] }
    """
    used_cols = set()
    for final_name, patterns in rename_map.items():
        for col in df.columns:
            if col in used_cols:
                continue
            # For Urgewald, skip renaming if the column is "Parent Company"
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(pat.lower().strip() in col.lower() for pat in patterns):
                df.rename(columns={col: final_name}, inplace=True)
                used_cols.add(col)
                break
    return df

##############################################
# 3. REORDER/SELECT OUTPUT COLUMNS
##############################################
def select_output_columns(df, desired_order):
    """
    Ensure the DataFrame has exactly the desired columns.
    For any missing column, create it as empty.
    """
    df = df.copy()
    for col in desired_order:
        if col not in df.columns:
            df[col] = ""
    return df[desired_order]

##############################################
# 4. LOAD SPGLOBAL (MULTI-HEADER)
##############################################
def load_spglobal(file, sheet_name="Sheet1"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if len(full_df) < 6:
            raise ValueError("SPGlobal file does not have enough rows.")
        # Use row 5 and row 6 for header detection
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
        rename_map = {
            "SP_ENTITY_NAME":  ["sp entity name", "s&p entity name", "entity name"],
            "SP_ENTITY_ID":    ["sp entity id", "entity id"],
            "SP_COMPANY_ID":   ["sp company id", "company id"],
            "SP_ISIN":         ["sp isin", "isin code"],
            "SP_LEI":          ["sp lei", "lei code"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining": ["thermal coal mining"],
            # Metallurgical Coal Mining is kept here for merging but will be excluded from output if not desired.
            "Metallurgical Coal Mining": ["metallurgical coal mining"],
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

##############################################
# 5. LOAD URGEWALD (SINGLE HEADER, EXCLUDING 'Parent Company')
##############################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")
        full_df = pd.DataFrame(data)
        # Filter out any header that is "Parent Company"
        header = full_df.iloc[0].fillna("")
        filtered_header = [col for col in header if str(col).strip().lower() != "parent company"]
        ur_df = full_df.iloc[1:].reset_index(drop=True)
        # Also select only columns whose header (in lowercase) is not "parent company"
        ur_df = ur_df.loc[:, header.str.strip().str.lower() != "parent company"]
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
            # We do not include Metallurgical Coal Mining in output columns if not needed.
        }
        ur_df = fuzzy_rename_columns(ur_df, rename_map)
        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

##############################################
# 6. NORMALIZE KEYS FOR MERGING
##############################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

def unify_name(r):
    sp_name = normalize_key(str(r.get("SP_ENTITY_NAME", "")))
    ur_name = normalize_key(str(r.get("Company", "")))
    return sp_name if sp_name else ur_name

def unify_isin(r):
    sp_isin = normalize_key(str(r.get("SP_ISIN", "")))
    ur_isin = normalize_key(str(r.get("ISIN equity", "")))
    return sp_isin if sp_isin else ur_isin

def unify_lei(r):
    sp_lei = normalize_key(str(r.get("SP_LEI", "")))
    ur_lei = normalize_key(str(r.get("LEI", "")))
    return sp_lei if sp_lei else ur_lei

def unify_bbticker(r):
    return normalize_key(str(r.get("BB Ticker", "")))

##############################################
# 7. MERGE URGEWALD INTO SPGLOBAL
##############################################
def merge_ur_into_sp(sp_df, ur_df):
    sp_records = sp_df.to_dict("records")
    merged_records = []
    ur_only_records = []
    # Add SP records first.
    for rec in sp_records:
        rec["Source"] = "SP"
        merged_records.append(rec)
    # Now, for each UR record, if any key matches, merge into the SP record.
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

##############################################
# 8. FILTER COMPANIES (Thresholds & Exclusion Logic)
##############################################
def filter_companies(df,
                     exclude_mining,
                     mining_coal_rev_threshold,
                     exclude_mining_prod_mt,
                     mining_prod_mt_threshold,
                     exclude_thermal_coal_mining,
                     thermal_coal_mining_threshold,
                     exclude_power,
                     power_coal_rev_threshold,
                     exclude_power_prod_percent,
                     power_prod_threshold_percent,
                     exclude_capacity_mw,
                     capacity_threshold_mw,
                     exclude_generation_thermal,
                     generation_thermal_threshold,
                     exclude_services,
                     services_rev_threshold,
                     exclude_services_rev,
                     expansions_global,
                     apply_mining_coal_rev,
                     apply_power_coal_rev):
    exclusion_flags = []
    exclusion_reasons = []
    for idx, row in df.iterrows():
        reasons = []
        # Numeric values
        try:
            coal_rev = float(row.get("Coal Share of Revenue", 0))
        except:
            coal_rev = 0.0
        try:
            coal_power = float(row.get("Coal Share of Power Production", 0))
        except:
            coal_power = 0.0
        try:
            installed_cap = float(row.get("Installed Coal Power Capacity (MW)", 0))
        except:
            installed_cap = 0.0
        try:
            gen_val = float(row.get("Generation (Thermal Coal)", 0))
        except:
            gen_val = 0.0
        try:
            therm_val = float(row.get("Thermal Coal Mining", 0))
        except:
            therm_val = 0.0
        prod_str = str(row.get(">10MT / >5GW", "")).lower()
        expansion_text = str(row.get("expansion", "")).lower()
        # Mining
        if exclude_mining:
            if apply_mining_coal_rev and (coal_rev * 100) > mining_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}%")
            if exclude_mining_prod_mt and (">10mt" in prod_str) and (mining_prod_mt_threshold <= 10):
                reasons.append(f">10MT indicated (threshold {mining_prod_mt_threshold}MT)")
            if exclude_thermal_coal_mining and (therm_val > thermal_coal_mining_threshold):
                reasons.append(f"Thermal Coal Mining {therm_val:.2f}% > {thermal_coal_mining_threshold}%")
        # Power
        if exclude_power:
            if apply_power_coal_rev and (coal_rev * 100) > power_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}%")
            if exclude_power_prod_percent and (coal_power * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power*100:.2f}% > {power_prod_threshold_percent}%")
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            if exclude_generation_thermal and (gen_val > generation_thermal_threshold):
                reasons.append(f"Generation (Thermal Coal) {gen_val:.2f}% > {generation_thermal_threshold}%")
        # Services
        if exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")
        # Global expansion
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

##############################################
# Helper functions for output adjustments
##############################################
def rename_ur_columns(df):
    """Rename UR identification columns with U_ prefix."""
    mapping = {"Company": "U_Company", "BB Ticker": "U_BB Ticker",
               "ISIN equity": "U_ISIN equity", "LEI": "U_LEI"}
    df = df.copy()
    df.rename(columns=mapping, inplace=True)
    return df

def add_empty_ur_columns(df):
    df = df.copy()
    for col in ["U_Company", "U_BB Ticker", "U_ISIN equity", "U_LEI"]:
        if col not in df.columns:
            df[col] = ""
    return df

def add_empty_sp_columns(df):
    df = df.copy()
    for col in ["SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI"]:
        if col not in df.columns:
            df[col] = ""
    return df

##############################################
# 9. MAIN STREAMLIT APP
##############################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Merged & Excluded", layout="wide")
    st.title("Coal Exclusion Filter")

    # Sidebar: File & Sheet Settings
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1", key="sp_sheet")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024", key="ur_sheet")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"], key="sp_file")
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"], key="ur_file")
    st.sidebar.markdown("---")

    # Sidebar: Mining Thresholds
    with st.sidebar.expander("Mining Thresholds", expanded=True):
        apply_mining_coal_rev = st.checkbox("Apply Mining: Max coal revenue threshold?", value=True, key="mining_rev")
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0, key="mining_rev_thresh")
        exclude_mining_prod_mt = st.checkbox("Exclude if >10MT indicated?", value=True, key="mining_10mt")
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0, key="mining_mt_thresh")
        exclude_mining_prod_gw = st.checkbox("Exclude if >5GW indicated?", value=True, key="mining_5gw")
        mining_prod_threshold_gw = st.number_input("Mining: Max production (GW)", value=5.0, key="mining_gw_thresh")
        exclude_thermal_coal_mining = st.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False, key="mining_thermal")
        thermal_coal_mining_threshold = st.number_input("Max allowed Thermal Coal Mining (%)", value=20.0, key="mining_thermal_thresh")
        exclude_metallurgical_coal_mining = st.checkbox("Exclude if Metallurgical Coal Mining > threshold?", value=False, key="mining_meta")
        metallurgical_coal_mining_threshold = st.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0, key="mining_meta_thresh")

    # Sidebar: Power Thresholds
    with st.sidebar.expander("Power Thresholds", expanded=True):
        apply_power_coal_rev = st.checkbox("Apply Power: Max coal revenue threshold?", value=True, key="power_rev")
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0, key="power_rev_thresh")
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True, key="power_prod")
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0, key="power_prod_thresh")
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True, key="power_cap")
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0, key="power_cap_thresh")
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False, key="power_gen")
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0, key="power_gen_thresh")

    # Sidebar: Services Thresholds
    with st.sidebar.expander("Services Thresholds", expanded=False):
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False, key="serv_rev")
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0, key="serv_rev_thresh")

    # Sidebar: Global Expansion
    with st.sidebar.expander("Global Expansion Exclusion", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansions_global = st.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[], key="global_exp")

    st.sidebar.markdown("---")
    start_time = time.time()

    if st.sidebar.button("Run", key="run_button"):
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

        # Merge UR into SP using custom matching (if any one key matches, merge)
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)

        # Group records:
        merged_sp = merged_df.copy()  # Merged (similar) companies
        unmatched_sp = sp_df[~sp_df.index.isin(merged_df.index)].copy()  # Unmatched SP records
        unmatched_ur = ur_only_df.copy()  # Unmatched UR records

        # S&P Only: unmatched SP with nonzero in key fields
        sp_only = unmatched_sp[
            (pd.to_numeric(unmatched_sp["Thermal Coal Mining"], errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(unmatched_sp["Generation (Thermal Coal)"], errors='coerce').fillna(0) > 0)
        ].copy()

        # Define desired output columns (exact order)
        desired_cols = [
            "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
            "Coal Industry Sector", "U_Company", ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)", "Coal Share of Power Production",
            "Coal Share of Revenue", "expansion", "Generation (Thermal Coal)",
            "Thermal Coal Mining", "U_BB Ticker", "U_ISIN equity", "U_LEI",
            "Excluded", "Exclusion Reasons"
        ]

        # Apply threshold filtering to each group using parameters
        params = {
            "mining_coal_rev_threshold": mining_coal_rev_threshold,
            "exclude_mining_revenue": apply_mining_coal_rev,
            "exclude_mining_prod_mt": exclude_mining_prod_mt,
            "mining_prod_mt_threshold": mining_prod_mt_threshold,
            "exclude_mining_prod_gw": exclude_mining_prod_gw,
            "mining_prod_threshold_gw": mining_prod_threshold_gw,
            "exclude_thermal_coal_mining": exclude_thermal_coal_mining,
            "thermal_coal_mining_threshold": thermal_coal_mining_threshold,
            "exclude_metallurgical_coal_mining": exclude_metallurgical_coal_mining,
            "metallurgical_coal_mining_threshold": metallurgical_coal_mining_threshold,
            "power_coal_rev_threshold": power_coal_rev_threshold,
            "exclude_power_revenue": apply_power_coal_rev,
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
        def compute_exclusion(row, **params):
            reasons = []
            try:
                coal_rev = float(row.get("Coal Share of Revenue", 0))
            except:
                coal_rev = 0.0
            try:
                coal_power = float(row.get("Coal Share of Power Production", 0))
            except:
                coal_power = 0.0
            try:
                installed_cap = float(row.get("Installed Coal Power Capacity (MW)", 0))
            except:
                installed_cap = 0.0
            try:
                gen_val = float(row.get("Generation (Thermal Coal)", 0))
            except:
                gen_val = 0.0
            try:
                therm_val = float(row.get("Thermal Coal Mining", 0))
            except:
                therm_val = 0.0
            prod_str = str(row.get(">10MT / >5GW", "")).lower()
            expansion = str(row.get("expansion", "")).lower()
            if params["exclude_mining_revenue"] and (coal_rev * 100) > params["mining_coal_rev_threshold"]:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {params['mining_coal_rev_threshold']}%")
            if params["exclude_mining_prod_mt"] and (">10mt" in prod_str) and (params["mining_prod_mt_threshold"] <= 10):
                reasons.append(f">10MT indicated (threshold {params['mining_prod_mt_threshold']}MT)")
            if therm_val > params["thermal_coal_mining_threshold"]:
                reasons.append(f"Thermal Coal Mining {therm_val:.2f}% > {params['thermal_coal_mining_threshold']}%")
            if params["exclude_power_revenue"] and (coal_rev * 100) > params["power_coal_rev_threshold"]:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {params['power_coal_rev_threshold']}%")
            if params["exclude_power_prod_percent"] and (coal_power * 100) > params["power_prod_threshold_percent"]:
                reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold_percent']}%")
            if installed_cap > params["capacity_threshold_mw"]:
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold_mw']}MW")
            if gen_val > params["generation_thermal_threshold"]:
                reasons.append(f"Generation (Thermal Coal) {gen_val:.2f}% > {params['generation_thermal_threshold']}%")
            if params["exclude_services_rev"] and (coal_rev * 100) > params["services_rev_threshold"]:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {params['services_rev_threshold']}% (Services)")
            if params["expansions_global"]:
                for kw in params["expansions_global"]:
                    if kw.lower() in expansion:
                        reasons.append(f"Expansion matched '{kw}'")
                        break
            return pd.Series([len(reasons) > 0, "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])
        
        # Apply threshold filtering to groups:
        merged_filtered = merged_sp.apply(lambda row: compute_exclusion(row, **params), axis=1)
        merged_sp["Excluded"] = merged_filtered["Excluded"]
        merged_sp["Exclusion Reasons"] = merged_filtered["Exclusion Reasons"]
        
        sp_unmatched_filtered = sp_only.apply(lambda row: compute_exclusion(row, **params), axis=1)
        sp_only["Excluded"] = sp_unmatched_filtered["Excluded"]
        sp_only["Exclusion Reasons"] = sp_unmatched_filtered["Exclusion Reasons"]
        
        ur_unmatched_filtered = unmatched_ur.apply(lambda row: compute_exclusion(row, **params), axis=1)
        unmatched_ur["Excluded"] = ur_unmatched_filtered["Excluded"]
        unmatched_ur["Exclusion Reasons"] = ur_unmatched_filtered["Exclusion Reasons"]

        # Build output groups:
        # Excluded: union of all records (merged, unmatched SP, unmatched UR) that are excluded.
        excluded_final = pd.concat([merged_sp[merged_sp["Excluded"] == True],
                                    sp_only[sp_only["Excluded"] == True],
                                    unmatched_ur[unmatched_ur["Excluded"] == True]], ignore_index=True)
        # Retained: only merged (similar) companies that passed thresholds.
        retained_merged = merged_sp[merged_sp["Excluded"] == False].copy()
        # S&P Only: retained unmatched SP records.
        sp_retained = sp_only[sp_only["Excluded"] == False].copy()
        # UR Only: retained unmatched UR records.
        ur_retained = unmatched_ur[unmatched_ur["Excluded"] == False].copy()

        # Adjust output columns (only these columns in exact order)
        output_cols = [
            "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
            "Coal Industry Sector", "U_Company", ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)", "Coal Share of Power Production",
            "Coal Share of Revenue", "expansion", "Generation (Thermal Coal)",
            "Thermal Coal Mining", "U_BB Ticker", "U_ISIN equity", "U_LEI",
            "Excluded", "Exclusion Reasons"
        ]

        # For SP records, add empty UR columns if missing.
        def add_empty_ur_cols(df):
            df = df.copy()
            for col in ["U_Company", "U_BB Ticker", "U_ISIN equity", "U_LEI"]:
                if col not in df.columns:
                    df[col] = ""
            return df

        sp_retained = add_empty_ur_cols(sp_retained)
        retained_merged = add_empty_ur_cols(retained_merged)

        # For UR records, rename identification columns and add empty SP columns.
        def rename_ur_columns(df):
            df = df.copy()
            mapping = {"Company": "U_Company", "BB Ticker": "U_BB Ticker",
                       "ISIN equity": "U_ISIN equity", "LEI": "U_LEI"}
            df.rename(columns=mapping, inplace=True)
            return df

        def add_empty_sp_cols(df):
            df = df.copy()
            for col in ["SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI"]:
                if col not in df.columns:
                    df[col] = ""
            return df

        ur_retained = rename_ur_columns(ur_retained)
        ur_retained = add_empty_sp_cols(ur_retained)

        # For excluded UR records:
        excluded_ur = excluded_final[excluded_final["SP_ENTITY_NAME"].isna()].copy()
        if not excluded_ur.empty:
            excluded_ur = rename_ur_columns(excluded_ur)
            excluded_ur = add_empty_sp_cols(excluded_ur)
        excluded_sp = excluded_final[excluded_final["SP_ENTITY_NAME"].notna()].copy()
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)

        # Finally, select only the desired columns in order.
        excluded_final = select_output_columns(excluded_final, output_cols)
        retained_merged = select_output_columns(retained_merged, output_cols)
        sp_retained = select_output_columns(sp_retained, output_cols)
        ur_retained = select_output_columns(ur_retained, output_cols)

        # Final output: four sheets in order:
        # 1. Excluded Companies, 2. Retained Companies, 3. S&P Only, 4. Urgewald Only.
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

if __name__ == "__main__":
    main()

