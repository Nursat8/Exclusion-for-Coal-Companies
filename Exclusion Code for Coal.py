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
            # For UR, skip renaming "Parent Company" to "Company"
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(pat.lower().strip() in col.lower() for pat in patterns):
                df.rename(columns={col: final_name}, inplace=True)
                used_cols.add(col)
                break
    return df

##############################################
# 3. REORDER OUTPUT COLUMNS
##############################################
def reorder_for_excel_custom(df, desired_order):
    """
    Ensure the DataFrame has exactly the desired columns in order.
    For any missing column, add it as an empty column.
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
            # Omit Metallurgical Coal Mining.
            "Coal Share of Revenue":     ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector":      ["coal industry sector", "industry sector"],
            ">10MT / >5GW":              [">10mt", ">5gw"],
            "expansion":                 ["expansion"],
        }
        sp_df = fuzzy_rename_columns(sp_df, rename_map_sp)
        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()

##############################################
# 5. LOAD URGEWALD (SINGLE HEADER)
##############################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")
        full_df = pd.DataFrame(data)
        header = full_df.iloc[0].fillna("")
        # Exclude any column named "parent company"
        filtered_header = [col for col in header if str(col).strip().lower() != "parent company"]
        ur_df = full_df.iloc[1:].reset_index(drop=True)
        ur_df = ur_df.loc[:, header.str.strip().str.lower() != "parent company"]
        ur_df.columns = filtered_header
        ur_df = make_columns_unique(ur_df)
        rename_map = {
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

##############################################
# 7. OPTIMIZED MERGE OF URGEWALD INTO SPGLOBAL
##############################################
def merge_ur_into_sp_opt(sp_df, ur_df):
    # Compute normalized keys for SP
    sp_df["norm_isin"] = sp_df["SP_ISIN"].astype(str).apply(normalize_key)
    sp_df["norm_lei"] = sp_df["SP_LEI"].astype(str).apply(normalize_key)
    sp_df["norm_name"] = sp_df["SP_ENTITY_NAME"].astype(str).apply(normalize_key)
    # Compute normalized keys for UR
    ur_df["norm_isin"] = ur_df["ISIN equity"].astype(str).apply(normalize_key)
    ur_df["norm_lei"] = ur_df["LEI"].astype(str).apply(normalize_key)
    ur_df["norm_company"] = ur_df["Company"].astype(str).apply(normalize_key)
    
    # Build dictionaries from SP records (take first occurrence for each key)
    dict_isin = {}
    dict_lei = {}
    dict_name = {}
    for idx, row in sp_df.iterrows():
        if row["norm_isin"]:
            if row["norm_isin"] not in dict_isin:
                dict_isin[row["norm_isin"]] = idx
        if row["norm_lei"]:
            if row["norm_lei"] not in dict_lei:
                dict_lei[row["norm_lei"]] = idx
        if row["norm_name"]:
            if row["norm_name"] not in dict_name:
                dict_name[row["norm_name"]] = idx
    
    # Prepare list to mark which UR records did not merge
    ur_not_merged = []
    # For each UR record, try to find a match in SP
    for idx, ur_row in ur_df.iterrows():
        found_index = None
        if ur_row["norm_isin"] and ur_row["norm_isin"] in dict_isin:
            found_index = dict_isin[ur_row["norm_isin"]]
        elif ur_row["norm_lei"] and ur_row["norm_lei"] in dict_lei:
            found_index = dict_lei[ur_row["norm_lei"]]
        elif ur_row["norm_company"] and ur_row["norm_company"] in dict_name:
            found_index = dict_name[ur_row["norm_company"]]
        
        if found_index is not None:
            # Merge non-empty UR fields into the matching SP record
            for col, val in ur_row.items():
                if col.startswith("norm_"):
                    continue
                if (col not in sp_df.loc[found_index] or pd.isna(sp_df.loc[found_index, col]) 
                    or str(sp_df.loc[found_index, col]).strip() == ""):
                    sp_df.loc[found_index, col] = val
            sp_df.loc[found_index, "Merged"] = True
        else:
            ur_not_merged.append(ur_row)
    
    # Mark SP rows with no match as not merged
    sp_df["Merged"] = sp_df.get("Merged", False)
    merged_df = sp_df.copy()
    ur_only_df = pd.DataFrame(ur_not_merged)
    
    # Drop temporary normalized columns
    for col in ["norm_isin", "norm_lei", "norm_name"]:
        if col in merged_df.columns:
            merged_df.drop(columns=[col], inplace=True)
    for col in ["norm_isin", "norm_lei", "norm_company"]:
        if col in ur_only_df.columns:
            ur_only_df.drop(columns=[col], inplace=True)
    return merged_df, ur_only_df

##############################################
# 8. FILTER COMPANIES (Thresholds & Exclusion Logic)
##############################################
# For UR records, Level 1 thresholds apply; then, if UR Level 2 is enabled, any UR record with revenue above that threshold is excluded.
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
    
    # Check source: SP record if SP_ENTITY_NAME exists; else, UR record.
    if str(row.get("SP_ENTITY_NAME", "")).strip():
        # S&P thresholds for mining:
        if "mining" in str(row.get("Coal Industry Sector", "")).lower():
            if params["sp_mining_checkbox"] and (coal_rev * 100) > params["sp_mining_threshold"]:
                reasons.append(f"SP Mining revenue {coal_rev*100:.2f}% > {params['sp_mining_threshold']}%")
            if params["exclude_mt"] and (">10mt" in prod_str):
                reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
        # S&P thresholds for power:
        if ("power" in str(row.get("Coal Industry Sector", "")).lower() or 
            "generation" in str(row.get("Coal Industry Sector", "")).lower()):
            if params["sp_power_checkbox"] and (coal_rev * 100) > params["sp_power_threshold"]:
                reasons.append(f"SP Power revenue {coal_rev*100:.2f}% > {params['sp_power_threshold']}%")
            if params["exclude_power_prod"] and (coal_power * 100) > params["power_prod_threshold"]:
                reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold']}%")
            if params["exclude_capacity"] and (installed_cap > params["capacity_threshold"]):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold']}MW")
    else:
        # UR record: Apply Level 1 thresholds per sector.
        if "mining" in str(row.get("Coal Industry Sector", "")).lower():
            if params["ur_mining_checkbox"] and (coal_rev * 100) > params["ur_mining_threshold"]:
                reasons.append(f"UR Mining revenue {coal_rev*100:.2f}% > {params['ur_mining_threshold']}%")
            if params["exclude_mt"] and (">10mt" in prod_str):
                reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
        if ("power" in str(row.get("Coal Industry Sector", "")).lower() or 
            "generation" in str(row.get("Coal Industry Sector", "")).lower()):
            if params["ur_power_checkbox"] and (coal_rev * 100) > params["ur_power_threshold"]:
                reasons.append(f"UR Power revenue {coal_rev*100:.2f}% > {params['ur_power_threshold']}%")
            if params["exclude_power_prod"] and (coal_power * 100) > params["power_prod_threshold"]:
                reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold']}%")
            if params["exclude_capacity"] and (installed_cap > params["capacity_threshold"]):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold']}MW")
        # Additionally, apply a global UR Level 2 threshold if enabled (applies to all UR records regardless of sector)
        if params.get("ur_level2_checkbox", False) and (coal_rev * 100) > params.get("ur_level2_threshold", 0):
            reasons.append(f"UR Level 2 revenue {coal_rev*100:.2f}% > {params['ur_level2_threshold']}%")
    
    # Global expansion check (applies to both)
    if params["expansion_exclude"]:
        for kw in params["expansion_exclude"]:
            if kw.lower() in expansion:
                reasons.append(f"Expansion matched '{kw}'")
                break
    return pd.Series([len(reasons) > 0, "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])

##############################################
# 9. OUTPUT ADJUSTMENT FUNCTIONS
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
# 10. MAIN STREAMLIT APP
##############################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Merged & Excluded", layout="wide")
    st.title("Coal Exclusion Filter")
    
    # Sidebar: File & Sheet Settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1", key="sp_sheet")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024", key="ur_sheet")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"], key="sp_file")
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"], key="ur_file")
    st.sidebar.markdown("---")
    
    # Sidebar: Mining Section
    with st.sidebar.expander("Mining", expanded=True):
        ur_mining_checkbox = st.checkbox("Urgewald: Exclude if thermal coal revenue > threshold (mining)", value=True, key="ur_mining")
        ur_mining_threshold = st.number_input("UR Mining: Level 1 threshold (%)", value=5.0, key="ur_mining_threshold")
        sp_mining_checkbox = st.checkbox("S&P: Exclude if thermal coal revenue > threshold (mining)", value=False, key="sp_mining")
        sp_mining_threshold = st.number_input("S&P Mining: Threshold (%)", value=15.0, key="sp_mining_threshold")
        exclude_mt = st.checkbox("Exclude if >10MT indicated", value=True, key="exclude_mt")
        mt_threshold = st.number_input("Max production (MT) threshold", value=10.0, key="mt_threshold")
    
    # Sidebar: Power Section
    with st.sidebar.expander("Power", expanded=True):
        ur_power_checkbox = st.checkbox("Urgewald: Exclude if thermal coal revenue > threshold (power)", value=False, key="ur_power")
        ur_power_threshold = st.number_input("UR Power: Level 1 threshold (%)", value=20.0, key="ur_power_threshold")
        sp_power_checkbox = st.checkbox("S&P: Exclude if thermal coal revenue > threshold (power)", value=False, key="sp_power")
        sp_power_threshold = st.number_input("S&P Power: Threshold (%)", value=20.0, key="sp_power_threshold")
        exclude_power_prod = st.checkbox("Exclude if > % production threshold", value=True, key="exclude_power_prod")
        power_prod_threshold = st.number_input("Max coal power production (%)", value=20.0, key="power_prod_threshold")
        exclude_capacity = st.checkbox("Exclude if > capacity (MW) threshold", value=True, key="exclude_capacity")
        capacity_threshold = st.number_input("Max installed capacity (MW)", value=10000.0, key="capacity_threshold")
    
    # Sidebar: UR Exclusion Level 2 (global for UR records)
    with st.sidebar.expander("UR Exclusion Level 2", expanded=True):
        ur_level2_checkbox = st.checkbox("Apply UR Level 2 exclusion", value=True, key="ur_level2_checkbox")
        ur_level2_threshold = st.number_input("UR Level 2 revenue threshold (%)", value=6.0, key="ur_level2_threshold")
    
    # Sidebar: Expansion Section
    with st.sidebar.expander("Expansion", expanded=False):
        expansion_exclude = st.multiselect("Exclude if expansion plans present", 
                                            ["mining", "infrastructure", "power", "subsidiary of a coal developer"],
                                            default=[], key="expansion_exclude")
    
    st.sidebar.markdown("---")
    start_time = time.time()
    
    if st.sidebar.button("Run", key="run_button"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return
        
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
        
        # Merge UR into SP using the optimized function
        sp_df, ur_df = merge_ur_into_sp_opt(sp_df, ur_df)
        
        # Split groups:
        merged_sp = sp_df[sp_df["Merged"] == True].copy()   # Merged (similar) companies
        unmatched_sp = sp_df[sp_df["Merged"] == False].copy() # Unmatched SP records
        unmatched_ur = ur_df[ur_df["Merged"] == False].copy()   # Unmatched UR records
        for group in [merged_sp, unmatched_sp, unmatched_ur]:
            if "Merged" in group.columns:
                group.drop(columns=["Merged"], inplace=True)
        
        # S&P Only: Unmatched SP records with nonzero in key mining/power fields
        sp_only = unmatched_sp[
            (pd.to_numeric(unmatched_sp["Thermal Coal Mining"], errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(unmatched_sp["Generation (Thermal Coal)"], errors='coerce').fillna(0) > 0)
        ].copy()
        
        # Prepare threshold parameters
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
            "expansion_exclude": expansion_exclude,
            # Global UR Level 2 threshold
            "ur_level2_checkbox": ur_level2_checkbox,
            "ur_level2_threshold": ur_level2_threshold
        }
        
        # Compute threshold exclusions using compute_exclusion:
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
            # Determine source: SP record if SP_ENTITY_NAME exists; otherwise, UR.
            if str(row.get("SP_ENTITY_NAME", "")).strip():
                # S&P thresholds for mining:
                if "mining" in str(row.get("Coal Industry Sector", "")).lower():
                    if params["sp_mining_checkbox"] and (coal_rev * 100) > params["sp_mining_threshold"]:
                        reasons.append(f"SP Mining revenue {coal_rev*100:.2f}% > {params['sp_mining_threshold']}%")
                    if params["exclude_mt"] and (">10mt" in prod_str):
                        reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
                # S&P thresholds for power:
                if ("power" in str(row.get("Coal Industry Sector", "")).lower() or 
                    "generation" in str(row.get("Coal Industry Sector", "")).lower()):
                    if params["sp_power_checkbox"] and (coal_rev * 100) > params["sp_power_threshold"]:
                        reasons.append(f"SP Power revenue {coal_rev*100:.2f}% > {params['sp_power_threshold']}%")
                    if params["exclude_power_prod"] and (coal_power * 100) > params["power_prod_threshold"]:
                        reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold']}%")
                    if params["exclude_capacity"] and (installed_cap > params["capacity_threshold"]):
                        reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold']}MW")
            else:
                # UR record: Apply UR Level 1 thresholds
                if "mining" in str(row.get("Coal Industry Sector", "")).lower():
                    if params["ur_mining_checkbox"] and (coal_rev * 100) > params["ur_mining_threshold"]:
                        reasons.append(f"UR Mining revenue {coal_rev*100:.2f}% > {params['ur_mining_threshold']}%")
                    if params["exclude_mt"] and (">10mt" in prod_str):
                        reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
                if ("power" in str(row.get("Coal Industry Sector", "")).lower() or 
                    "generation" in str(row.get("Coal Industry Sector", "")).lower()):
                    if params["ur_power_checkbox"] and (coal_rev * 100) > params["ur_power_threshold"]:
                        reasons.append(f"UR Power revenue {coal_rev*100:.2f}% > {params['ur_power_threshold']}%")
                    if params["exclude_power_prod"] and (coal_power * 100) > params["power_prod_threshold"]:
                        reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold']}%")
                    if params["exclude_capacity"] and (installed_cap > params["capacity_threshold"]):
                        reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold']}MW")
                # Now apply the separate UR Level 2 threshold globally (for all UR records)
                if params.get("ur_level2_checkbox", False) and (coal_rev * 100) > params.get("ur_level2_threshold", 0):
                    reasons.append(f"UR Level 2 revenue {coal_rev*100:.2f}% > {params['ur_level2_threshold']}%")
            if params["expansion_exclude"]:
                for kw in params["expansion_exclude"]:
                    if kw.lower() in expansion:
                        reasons.append(f"Expansion matched '{kw}'")
                        break
            return pd.Series([len(reasons) > 0, "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])
        
        merged_filtered = merged_sp.apply(lambda row: compute_exclusion(row, **params), axis=1)
        merged_sp["Excluded"] = merged_filtered["Excluded"]
        merged_sp["Exclusion Reasons"] = merged_filtered["Exclusion Reasons"]
        
        sp_filtered = sp_only.apply(lambda row: compute_exclusion(row, **params), axis=1)
        sp_only["Excluded"] = sp_filtered["Excluded"]
        sp_only["Exclusion Reasons"] = sp_filtered["Exclusion Reasons"]
        
        ur_filtered = unmatched_ur.apply(lambda row: compute_exclusion(row, **params), axis=1)
        unmatched_ur["Excluded"] = ur_filtered["Excluded"]
        unmatched_ur["Exclusion Reasons"] = ur_filtered["Exclusion Reasons"]
        
        # Build output groups:
        excluded_final = pd.concat([merged_sp[merged_sp["Excluded"] == True],
                                    sp_only[sp_only["Excluded"] == True],
                                    unmatched_ur[unmatched_ur["Excluded"] == True]], ignore_index=True)
        retained_merged = merged_sp[merged_sp["Excluded"] == False].copy()
        sp_retained = sp_only[sp_only["Excluded"] == False].copy()
        ur_retained = unmatched_ur[unmatched_ur["Excluded"] == False].copy()
        
        # Adjust output columns exactly as required.
        output_cols = ["SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
                       "Coal Industry Sector", "U_Company", ">10MT / >5GW",
                       "Installed Coal Power Capacity (MW)", "Coal Share of Power Production",
                       "Coal Share of Revenue", "expansion", "Generation (Thermal Coal)",
                       "Thermal Coal Mining", "U_BB Ticker", "U_ISIN equity", "U_LEI",
                       "Excluded", "Exclusion Reasons"]
        
        def add_empty_ur_cols(df):
            df = df.copy()
            for col in ["U_Company", "U_BB Ticker", "U_ISIN equity", "U_LEI"]:
                if col not in df.columns:
                    df[col] = ""
            return df
        
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
        
        sp_retained = add_empty_ur_cols(sp_retained)
        retained_merged = add_empty_ur_cols(retained_merged)
        ur_retained = rename_ur_columns(ur_retained)
        ur_retained = add_empty_sp_cols(ur_retained)
        
        excluded_sp = excluded_final[excluded_final["SP_ENTITY_NAME"].notna()].copy()
        excluded_ur = excluded_final[excluded_final["SP_ENTITY_NAME"].isna()].copy()
        if not excluded_ur.empty:
            excluded_ur = rename_ur_columns(excluded_ur)
            excluded_ur = add_empty_sp_cols(excluded_ur)
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)
        for df in [excluded_final, retained_merged, sp_retained, ur_retained]:
            for col in output_cols:
                if col not in df.columns:
                    df[col] = ""
            df = df[output_cols]
        
        # Final output: Four sheets in order: Excluded Companies, Retained Companies, S&P Only, Urgewald Only.
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
