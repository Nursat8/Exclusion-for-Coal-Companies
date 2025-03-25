import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import time
import re
import io

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
            # Skip 'Parent Company' => 'Company'
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(p.lower().strip() in col.lower() for p in patterns):
                df.rename(columns={col: final_name}, inplace=True)
                used_cols.add(col)
                break
    return df

##############################################
# 3. NORMALIZE KEY FOR MERGING
##############################################
def normalize_key(s):
    s = s.lower()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    return s.strip()

##############################################
# 4. MERGE URGEWALD INTO SP (OPTIMIZED)
##############################################
def merge_ur_into_sp_opt(sp_df, ur_df):
    """
    Merges UR data into SP if any one key (ISIN, LEI, or Name) matches.
    Builds dictionaries for quick lookup, merges nonempty UR fields into SP.
    Unmatched UR records go to ur_only_df.
    """
    sp_df = sp_df.copy()
    ur_df = ur_df.copy()
    # Convert to object dtype to avoid dtype issues
    sp_df = sp_df.astype(object)
    ur_df = ur_df.astype(object)

    # Create normalized keys for SP
    sp_df["norm_isin"] = sp_df.get("SP_ISIN","").astype(str).apply(normalize_key)
    sp_df["norm_lei"]  = sp_df.get("SP_LEI","").astype(str).apply(normalize_key)
    sp_df["norm_name"] = sp_df.get("SP_ENTITY_NAME","").astype(str).apply(normalize_key)

    # Create normalized keys for UR
    for col in ["ISIN equity", "LEI", "Company"]:
        if col not in ur_df.columns:
            ur_df[col] = ""
    ur_df["norm_isin"]    = ur_df["ISIN equity"].astype(str).apply(normalize_key)
    ur_df["norm_lei"]     = ur_df["LEI"].astype(str).apply(normalize_key)
    ur_df["norm_company"] = ur_df["Company"].astype(str).apply(normalize_key)

    # Build dictionaries
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
            # Merge non-empty UR fields into SP
            for col, val in ur_row.items():
                if col.startswith("norm_"):
                    continue
                if (col not in sp_df.columns) or pd.isna(sp_df.loc[found_index, col]) or str(sp_df.loc[found_index, col]).strip() == "":
                    sp_df.loc[found_index, col] = str(val)
            sp_df.loc[found_index, "Merged"] = True
        else:
            ur_not_merged.append(ur_row)

    sp_df["Merged"] = sp_df.get("Merged", False)
    merged_df = sp_df.copy()
    ur_only_df = pd.DataFrame(ur_not_merged)

    # Drop norm columns
    for c in ["norm_isin","norm_lei","norm_name"]:
        if c in merged_df.columns:
            merged_df.drop(columns=[c], inplace=True)
    for c in ["norm_isin","norm_lei","norm_company"]:
        if c in ur_only_df.columns:
            ur_only_df.drop(columns=[c], inplace=True)

    return merged_df, ur_only_df

##############################################
# 5. THRESHOLD FILTERING
##############################################
def compute_exclusion(row, **params):
    reasons = []
    # For S&P "Thermal Coal Mining" or "Generation (Thermal Coal)" we interpret them as the % for mining or power
    # For Urgewald "Coal Share of Revenue" is used for mining/power
    try:
        # For Urgewald we use "Coal Share of Revenue" * 100 as the threshold
        # For S&P we directly use "Thermal Coal Mining" or "Generation (Thermal Coal)" as the threshold
        # For 'Coal Share of Power Production' we interpret as a fraction
        coal_rev = float(row.get("Coal Share of Revenue", 0))  # for UR
    except:
        coal_rev = 0.0
    try:
        coal_power = float(row.get("Coal Share of Power Production", 0))  # fraction for UR
    except:
        coal_power = 0.0
    try:
        installed_cap = float(row.get("Installed Coal Power Capacity (MW)", 0))
    except:
        installed_cap = 0.0
    try:
        sp_mining_val = float(row.get("Thermal Coal Mining", 0))  # fraction for S&P
    except:
        sp_mining_val = 0.0
    try:
        sp_power_val = float(row.get("Generation (Thermal Coal)", 0))  # fraction for S&P
    except:
        sp_power_val = 0.0

    # For the >10MT / >5GW check
    prod_str = str(row.get(">10MT / >5GW", "")).lower()
    # For expansions
    expansion = str(row.get("expansion", "")).lower()

    # Check if SP or UR record
    is_sp = bool(str(row.get("SP_ENTITY_NAME","")).strip())  # If non-empty => SP
    sector = str(row.get("Coal Industry Sector","")).lower()

    if is_sp:
        # S&P record
        # Mining => "Thermal Coal Mining" => sp_mining_val
        if "mining" in sector:
            if params["sp_mining_checkbox"] and (sp_mining_val * 100) > params["sp_mining_threshold"]:
                reasons.append(f"SP Mining revenue {sp_mining_val*100:.2f}% > {params['sp_mining_threshold']}%")
        # Power => "Generation (Thermal Coal)" => sp_power_val
        if ("power" in sector or "generation" in sector):
            if params["sp_power_checkbox"] and (sp_power_val * 100) > params["sp_power_threshold"]:
                reasons.append(f"SP Power revenue {sp_power_val*100:.2f}% > {params['sp_power_threshold']}%")
    else:
        # UR record
        # Mining => "Coal Share of Revenue" => coal_rev
        if "mining" in sector:
            if params["ur_mining_checkbox"] and (coal_rev * 100) > params["ur_mining_threshold"]:
                reasons.append(f"UR Mining revenue {coal_rev*100:.2f}% > {params['ur_mining_threshold']}%")
            # Exclude if >10MT indicated
            if params["exclude_mt"] and (">10mt" in prod_str):
                reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
        # Power => "Coal Share of Revenue" => coal_rev
        if ("power" in sector or "generation" in sector):
            if params["ur_power_checkbox"] and (coal_rev * 100) > params["ur_power_threshold"]:
                reasons.append(f"UR Power revenue {coal_rev*100:.2f}% > {params['ur_power_threshold']}%")
            # Exclude if > capacity (MW) threshold
            if params["exclude_capacity"] and (installed_cap > params["capacity_threshold"]):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {params['capacity_threshold']}MW")
            # Exclude if > % production threshold => coal_power
            if params["exclude_power_prod"] and (coal_power * 100) > params["power_prod_threshold"]:
                reasons.append(f"Coal power production {coal_power*100:.2f}% > {params['power_prod_threshold']}%")

        # UR Exclusion Level 2 => apply to all UR records
        if params["ur_level2_checkbox"] and (coal_rev * 100) > params["ur_level2_threshold"]:
            reasons.append(f"UR Level 2 revenue {coal_rev*100:.2f}% > {params['ur_level2_threshold']}%")

    # expansions
    if params["expansion_exclude"]:
        for kw in params["expansion_exclude"]:
            if kw.lower() in expansion:
                reasons.append(f"Expansion matched '{kw}'")
                break

    # Return final
    return pd.Series([len(reasons) > 0, "; ".join(reasons)], index=["Excluded","Exclusion Reasons"])

##############################################
# 9. RENAME & ADD COLUMNS
##############################################
def rename_ur_columns(df):
    """Rename UR columns to have U_ prefix for final output."""
    mapping = {
        "Company": "U_Company",
        "BB Ticker": "U_BB Ticker",
        "ISIN equity": "U_ISIN equity",
        "LEI": "U_LEI"
    }
    df = df.copy()
    for old, new in mapping.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)
    return df

def add_empty_ur_columns(df):
    df = df.copy()
    for col in ["U_Company","U_BB Ticker","U_ISIN equity","U_LEI"]:
        if col not in df.columns:
            df[col] = ""
    return df

def add_empty_sp_columns(df):
    df = df.copy()
    for col in ["SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI"]:
        if col not in df.columns:
            df[col] = ""
    return df

##############################################
# 10. STREAMLIT MAIN
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
        # S&P mining
        sp_mining_checkbox = st.checkbox("S&P: Exclude if thermal coal revenue > threshold (mining)", value=True, key="sp_mining")
        sp_mining_threshold = st.number_input("S&P Mining Threshold (%)", value=15.0, key="sp_mining_threshold")
        # UR mining
        ur_mining_checkbox = st.checkbox("Urgewald: Exclude if thermal coal revenue > threshold (mining)", value=True, key="ur_mining")
        ur_mining_threshold = st.number_input("UR Mining: Level 1 threshold (%)", value=5.0, key="ur_mining_threshold")
        # >10MT
        exclude_mt = st.checkbox("Exclude if >10MT indicated", value=True, key="exclude_mt")
        mt_threshold = st.number_input("Max production (MT) threshold", value=10.0, key="mt_threshold")
    
    # Sidebar: Power Section
    with st.sidebar.expander("Power", expanded=True):
        # S&P power
        sp_power_checkbox = st.checkbox("S&P: Exclude if thermal coal revenue > threshold (power)", value=True, key="sp_power")
        sp_power_threshold = st.number_input("S&P Power Threshold (%)", value=20.0, key="sp_power_threshold")
        # UR power
        ur_power_checkbox = st.checkbox("Urgewald: Exclude if thermal coal revenue > threshold (power)", value=True, key="ur_power")
        ur_power_threshold = st.number_input("UR Power: Level 1 threshold (%)", value=20.0, key="ur_power_threshold")
        # exclude power production
        exclude_power_prod = st.checkbox("Exclude if > % production threshold", value=True, key="exclude_power_prod")
        power_prod_threshold = st.number_input("Max coal power production (%)", value=20.0, key="power_prod_threshold")
        # exclude capacity
        exclude_capacity = st.checkbox("Exclude if > capacity (MW) threshold", value=True, key="exclude_capacity")
        capacity_threshold = st.number_input("Max installed capacity (MW)", value=10000.0, key="capacity_threshold")
    
    # Sidebar: UR Exclusion Level 2
    with st.sidebar.expander("UR Exclusion Level 2", expanded=True):
        ur_level2_checkbox = st.checkbox("Apply UR Level 2 exclusion", value=True, key="ur_level2_checkbox")
        ur_level2_threshold = st.number_input("UR Level 2 revenue threshold (%)", value=6.0, key="ur_level2_threshold")
    
    # Sidebar: Expansion
    with st.sidebar.expander("Expansion", expanded=False):
        expansion_exclude = st.multiselect("Exclude if expansion plans present", 
            ["mining","infrastructure","power","subsidiary of a coal developer"],
            default=[], key="expansion_exclude")
    
    st.sidebar.markdown("---")
    start_time = time.time()
    
    if st.sidebar.button("Run", key="run_button"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return
        
        # LOAD SP
        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or not loaded.")
            return
        sp_df = make_columns_unique(sp_df)
        
        # LOAD UR
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return
        ur_df = make_columns_unique(ur_df)
        
        # MERGE
        sp_df, ur_df = merge_ur_into_sp_opt(sp_df, ur_df)
        
        # SPLIT
        merged_sp = sp_df[sp_df.get("Merged",False) == True].copy()
        unmatched_sp = sp_df[sp_df.get("Merged",False) == False].copy()
        unmatched_ur = ur_df[ur_df.get("Merged",False) == False].copy()
        for group in [merged_sp, unmatched_sp, unmatched_ur]:
            if "Merged" in group.columns:
                group.drop(columns=["Merged"], inplace=True)
        
        # S&P Only: from unmatched SP => those with nonzero in "Thermal Coal Mining" or "Generation (Thermal Coal)"
        sp_only = unmatched_sp[
            (pd.to_numeric(unmatched_sp.get("Thermal Coal Mining","0"), errors='coerce').fillna(0) > 0) |
            (pd.to_numeric(unmatched_sp.get("Generation (Thermal Coal)","0"), errors='coerce').fillna(0) > 0)
        ].copy()
        
        # Prepare threshold parameters
        params = {
            # S&P (mining/power)
            "sp_mining_checkbox":   sp_mining_checkbox,
            "sp_mining_threshold":  sp_mining_threshold,
            "sp_power_checkbox":    sp_power_checkbox,
            "sp_power_threshold":   sp_power_threshold,
            # UR (mining/power)
            "ur_mining_checkbox":   ur_mining_checkbox,
            "ur_mining_threshold":  ur_mining_threshold,
            "ur_power_checkbox":    ur_power_checkbox,
            "ur_power_threshold":   ur_power_threshold,
            # Common
            "exclude_mt":           exclude_mt,
            "mt_threshold":         mt_threshold,
            "exclude_power_prod":   exclude_power_prod,
            "power_prod_threshold": power_prod_threshold,
            "exclude_capacity":     exclude_capacity,
            "capacity_threshold":   capacity_threshold,
            "ur_level2_checkbox":   ur_level2_checkbox,
            "ur_level2_threshold":  ur_level2_threshold,
            "expansion_exclude":    expansion_exclude
        }
        
        # THRESHOLD FILTER
        merged_filtered = merged_sp.apply(lambda row: compute_exclusion(row, **params), axis=1)
        merged_sp["Excluded"] = merged_filtered["Excluded"]
        merged_sp["Exclusion Reasons"] = merged_filtered["Exclusion Reasons"]
        
        sp_filtered = sp_only.apply(lambda row: compute_exclusion(row, **params), axis=1)
        sp_only["Excluded"] = sp_filtered["Excluded"]
        sp_only["Exclusion Reasons"] = sp_filtered["Exclusion Reasons"]
        
        ur_filtered = unmatched_ur.apply(lambda row: compute_exclusion(row, **params), axis=1)
        unmatched_ur["Excluded"] = ur_filtered["Excluded"]
        unmatched_ur["Exclusion Reasons"] = ur_filtered["Exclusion Reasons"]
        
        # BUILD OUTPUT
        excluded_final = pd.concat([
            merged_sp[merged_sp["Excluded"] == True],
            sp_only[sp_only["Excluded"] == True],
            unmatched_ur[unmatched_ur["Excluded"] == True]
        ], ignore_index=True)
        retained_merged = merged_sp[merged_sp["Excluded"] == False].copy()
        sp_retained = sp_only[sp_only["Excluded"] == False].copy()
        ur_retained = unmatched_ur[unmatched_ur["Excluded"] == False].copy()
        
        # RENAME UR columns to U_ prefix for final output
        def rename_ur_cols(df):
            mapping = {
                "Company": "U_Company",
                "BB Ticker": "U_BB Ticker",
                "ISIN equity": "U_ISIN equity",
                "LEI": "U_LEI"
            }
            df = df.copy()
            for old, new in mapping.items():
                if old in df.columns:
                    df.rename(columns={old: new}, inplace=True)
            return df
        
        def add_empty_ur_cols(df):
            df = df.copy()
            for c in ["U_Company","U_BB Ticker","U_ISIN equity","U_LEI"]:
                if c not in df.columns:
                    df[c] = ""
            return df
        
        def add_empty_sp_cols(df):
            df = df.copy()
            for c in ["SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI"]:
                if c not in df.columns:
                    df[c] = ""
            return df
        
        # Adjust retained_merged & sp_retained to have empty UR columns
        retained_merged = add_empty_ur_cols(retained_merged)
        sp_retained = add_empty_ur_cols(sp_retained)
        # Rename & add empty SP for UR records
        ur_retained = rename_ur_cols(ur_retained)
        ur_retained = add_empty_sp_cols(ur_retained)
        
        # Excluded: separate SP vs UR
        excluded_sp = excluded_final[excluded_final.get("SP_ENTITY_NAME","").notna()].copy()
        excluded_ur = excluded_final[excluded_final.get("SP_ENTITY_NAME","").isna()].copy()
        if not excluded_ur.empty:
            excluded_ur = rename_ur_cols(excluded_ur)
            excluded_ur = add_empty_sp_cols(excluded_ur)
        excluded_final = pd.concat([excluded_sp, excluded_ur], ignore_index=True)
        
        final_cols = [
            "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
            "Coal Industry Sector", "U_Company", ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)", "Coal Share of Power Production",
            "Coal Share of Revenue", "expansion", "Generation (Thermal Coal)",
            "Thermal Coal Mining", "U_BB Ticker", "U_ISIN equity", "U_LEI",
            "Excluded", "Exclusion Reasons"
        ]
        
        def finalize_cols(df):
            df = df.copy()
            for col in final_cols:
                if col not in df.columns:
                    df[col] = ""
            return df[final_cols]
        
        excluded_final = finalize_cols(excluded_final)
        retained_merged = finalize_cols(retained_merged)
        sp_retained = finalize_cols(sp_retained)
        ur_retained = finalize_cols(ur_retained)
        
        # WRITE OUTPUT to Excel in memory
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
