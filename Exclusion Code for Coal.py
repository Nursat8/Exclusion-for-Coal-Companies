import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import time
import re
import io

##############################################
# Helper: robust conversion to float, auto-detects US vs EU format
##############################################
def to_float(val):
    s = str(val).strip().replace(" ", "")
    if s == "" or s.lower() in ("nan", "none"):
        return 0.0
    # if both separators present, decide which is decimal:
    if "." in s and "," in s:
        # if comma comes after the last dot, treat comma as decimal sep
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # dot is decimal, commas are thousands
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        # if there's exactly one comma and it looks like a decimal (1 or 2 digits after)
        if len(parts) == 2 and len(parts[1]) in (1,2):
            s = s.replace(",", ".")
        else:
            # otherwise assume commas are thousands sep
            s = s.replace(",", "")
    # else: only dots (or neither) → dot is decimal or plain integer
    try:
        return float(s)
    except ValueError:
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
            "Coal Share of Revenue":     ["coal share of revenue"],
            "Coal Share of Power Production": ["coal share of power production"],
            "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
            "Coal Industry Sector":      ["coal industry sector", "industry sector"],
            ">10MT / >5GW":              [">10mt", ">5gw"],
            "expansion":                 ["expansion"],
        }
        sp_df = fuzzy_rename_columns(sp_df, rename_map_sp)
        sp_df = sp_df.astype(object)

        # convert coal-related columns to floats
        for col in [
            "Thermal Coal Mining",
            "Generation (Thermal Coal)",
            "Coal Share of Revenue",
            "Coal Share of Power Production",
            "Installed Coal Power Capacity (MW)"
        ]:
            if col in sp_df.columns:
                sp_df[col] = sp_df[col].apply(to_float)

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

        # convert coal-related columns to floats
        for col in [
            "Thermal Coal Mining",
            "Generation (Thermal Coal)",
            "Coal Share of Revenue",
            "Coal Share of Power Production",
            "Installed Coal Power Capacity (MW)"
        ]:
            if col in ur_df.columns:
                ur_df[col] = ur_df[col].apply(to_float)

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
                # if the column doesn't exist in SP, or its value is null/empty, overwrite it
                if (col not in sp_df.columns) \
                   or pd.isnull(sp_df.loc[found_index, col]) \
                   or str(sp_df.loc[found_index, col]).strip() == "":
                    sp_df.loc[found_index, col] = val
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
    reasons = []

    sp_mining_val = row.get("Thermal Coal Mining", 0)
    sp_power_val  = row.get("Generation (Thermal Coal)", 0)
    _rev = row.get("Coal Share of Revenue", 0)
    ur_coal_rev = _rev if _rev > 1 else _rev * 100
    raw_coal_power = row.get("Coal Share of Power Production", 0)
    _pp = raw_coal_power
    raw_coal_power = _pp if _pp > 1 else _pp * 100

    ur_installed_cap = row.get("Installed Coal Power Capacity (MW)", 0)

    prod_str = str(row.get(">10MT / >5GW", "")).strip().lower()
    expansion_str = str(row.get("expansion", "")).strip().lower()
    is_sp = bool(str(row.get("SP_ENTITY_NAME", "")).strip())
    sector = str(row.get("Coal Industry Sector", "")).strip().lower()

    # 1) >10MT check
    if params["exclude_mt"] and "10mt" in prod_str:
        reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")

    # 2) Capacity
    if params["exclude_capacity"] and ur_installed_cap > params["capacity_threshold"]:
        reasons.append(f"Installed capacity {ur_installed_cap:.2f}MW > {params['capacity_threshold']}MW")

    # 3) Power production
    if params["exclude_power_prod"] and (raw_coal_power * 100) > params["power_prod_threshold"]:
        reasons.append(f"Coal power production {(raw_coal_power*100):.2f}% > {params['power_prod_threshold']}%")

    if is_sp:
        # SP mining
        if "mining" in sector and params["sp_mining_checkbox"] and sp_mining_val > params["sp_mining_threshold"]:
            reasons.append(f"SP Mining revenue {sp_mining_val:.2f}% > {params['sp_mining_threshold']}%")
        # SP power
        if ("power" in sector or "generation" in sector) and params["sp_power_checkbox"] and sp_power_val > params["sp_power_threshold"]:
            reasons.append(f"SP Power revenue {sp_power_val:.2f}% > {params['sp_power_threshold']}%")
    else:
        # UR mining-only
        if ("mining" in sector) and not ("power" in sector or "generation" in sector) and params["ur_mining_checkbox"] and (ur_coal_rev*100) > params["ur_mining_threshold"]:
            reasons.append(f"UR Mining revenue {(ur_coal_rev*100):.2f}% > {params['ur_mining_threshold']}%")
        # UR power-only
        if ("power" in sector or "generation" in sector) and not "mining" in sector and params["ur_power_checkbox"] and (ur_coal_rev*100) > params["ur_power_threshold"]:
            reasons.append(f"UR Power revenue {(ur_coal_rev*100):.2f}% > {params['ur_power_threshold']}%")
        # UR level2 applies always
    if params["ur_level2_checkbox"] and ur_coal_rev > params["ur_level2_threshold"]:
        reasons.append(f"UR Level 2 revenue {ur_coal_rev:.2f}% > {params['ur_level2_threshold']}%")
        
    # 7) expansion
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

    # Sidebar settings
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", "Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", "GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"] )
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"] )
    st.sidebar.markdown("---")

    with st.sidebar.expander("Mining", expanded=True):
        ur_mining_checkbox = st.checkbox("Urgewald: Exclude if thermal coal mining revenue > threshold", False)
        ur_mining_threshold  = st.number_input("UR Mining: Level 1 threshold (%)", value=5.0)
        sp_mining_checkbox   = st.checkbox("S&P: Exclude if thermal coal mining revenue > threshold", True)
        sp_mining_threshold  = st.number_input("S&P Mining Threshold (%)", value=5.0)
        exclude_mt           = st.checkbox("Exclude if > MT threshold", True)
        mt_threshold         = st.number_input("Max production (MT) threshold", value=10.0)

    with st.sidebar.expander("Power", expanded=True):
        ur_power_checkbox    = st.checkbox("Urgewald: Exclude if thermal coal generation revenue > threshold", False)
        ur_power_threshold   = st.number_input("UR Power: Level 1 threshold (%)", value=20.0)
        sp_power_checkbox    = st.checkbox("S&P: Exclude if thermal coal generation revenue > threshold", True)
        sp_power_threshold   = st.number_input("S&P Power Threshold (%)", value=20.0)
        exclude_power_prod   = st.checkbox("Exclude if > % production threshold", True)
        power_prod_threshold = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity     = st.checkbox("Exclude if > capacity (MW) threshold", True)
        capacity_threshold   = st.number_input("Max installed capacity (MW)", value=10000.0)

    with st.sidebar.expander("Urgewald Level 2", expanded=False):
        ur_level2_checkbox   = st.checkbox("Apply UR Level 2 exclusion", False)
        ur_level2_threshold  = st.number_input("UR Level 2 revenue threshold (%)", value=10.0)

    with st.sidebar.expander("Exclude expansions", expanded=False):
        expansions_possible  = ["mining","infrastructure","power","subsidiary of a coal developer"]
        expansion_exclude    = st.multiselect("Exclude if expansion text contains", expansions_possible, [])

    st.sidebar.markdown("---")
    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        sp_df = load_spglobal(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or not loaded.")
            return
        ur_df = load_urgewald(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or not loaded.")
            return

        merged_sp, ur_only_df = merge_ur_into_sp_opt(sp_df, ur_df)
       
        merged_sp["Merged"]    = merged_sp["Merged"].fillna(False).astype(bool)
        ur_only_df["Merged"]   = ur_only_df["Merged"].fillna(False).astype(bool)

       
        numeric_cols = [
            "Thermal Coal Mining",
            "Generation (Thermal Coal)",
            "Coal Share of Revenue",
            "Coal Share of Power Production",
            "Installed Coal Power Capacity (MW)"
        ]
        for df in (merged_sp, ur_only_df):
            for c in numeric_cols:
                if c in df.columns:
                    df[c] = df[c].apply(to_float)



        merged_sp["Merged"]    = merged_sp["Merged"].fillna(False).astype(bool)
        ur_only_df["Merged"]   = ur_only_df["Merged"].fillna(False).astype(bool)


        sp_merged   = merged_sp[ merged_sp["Merged"] ]
        sp_unmerged = merged_sp[ ~merged_sp["Merged"] ]
        ur_unmerged = ur_only_df[ ~ur_only_df["Merged"] ]

        merged_sp, ur_only_df = merge_ur_into_sp_opt(sp_df, ur_df)
        merged_sp["Merged"]  = merged_sp["Merged"].fillna(False).astype(bool)
        ur_only_df["Merged"] = ur_only_df["Merged"].fillna(False).astype(bool)

        sp_merged   = merged_sp[ merged_sp["Merged"] ]
        sp_unmerged = merged_sp[ ~merged_sp["Merged"] ]
        ur_unmerged = ur_only_df[ ~ur_only_df["Merged"] ]




        for df in [sp_merged, sp_unmerged, ur_unmerged]:
            if "Merged" in df.columns:
                df.drop(columns=["Merged"], inplace=True)

        sp_only = sp_unmerged[(sp_unmerged["Thermal Coal Mining"]>0)|(sp_unmerged["Generation (Thermal Coal)"]>0)]

        params = {k: v for k,v in locals().items() if k.endswith("checkbox") or k.endswith("threshold") or k in ["exclude_mt","mt_threshold","exclude_power_prod","power_prod_threshold","exclude_capacity","capacity_threshold","expansion_exclude"]}

        def apply_filter(df):
            if df.empty: return df.assign(Excluded=False, **{"Exclusion Reasons":""})
            filt = df.apply(lambda row: compute_exclusion(row, **params), axis=1, result_type="expand")
            df["Excluded"] = filt["Excluded"]
            df["Exclusion Reasons"] = filt["Exclusion Reasons"]
            return df

        sp_merged   = apply_filter(sp_merged)
        sp_only     = apply_filter(sp_only)
        ur_unmerged = apply_filter(ur_unmerged)

        excluded_final   = pd.concat([sp_merged[sp_merged.Excluded], sp_only[sp_only.Excluded], ur_unmerged[ur_unmerged.Excluded]])
        retained_merged  = sp_merged[~sp_merged.Excluded]
        sp_retained      = sp_only[~sp_only.Excluded]
        ur_retained      = ur_unmerged[~ur_unmerged.Excluded]

        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Coal Industry Sector","Company",">10MT / >5GW","Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production","Coal Share of Revenue","expansion",
            "Generation (Thermal Coal)","Thermal Coal Mining","BB Ticker","ISIN equity","LEI",
            "Excluded","Exclusion Reasons"
        ]
        def finalize_cols(df):
            for c in final_cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[final_cols]
            if "BB Ticker" in df.columns:
                df["BB Ticker"] = df["BB Ticker"].astype(str).str.replace(r'\s*Equity','',regex=True)
            return df

        excluded_final  = finalize_cols(excluded_final)
        retained_merged = finalize_cols(retained_merged)
        sp_retained     = finalize_cols(sp_retained)
        ur_retained     = finalize_cols(ur_retained)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_final.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_merged.to_excel(writer, sheet_name="Retained Companies", index=False)
            sp_retained.to_excel(writer, sheet_name="S&P Only", index=False)
            ur_retained.to_excel(writer, sheet_name="Urgewald Only", index=False)

        st.subheader("Results Summary")
        st.write(f"Excluded Companies: {len(excluded_final)}")
        st.write(f"Retained Companies (Merged & Retained): {len(retained_merged)}")
        st.write(f"S&P Only (Unmatched, Retained): {len(sp_retained)}")
        st.write(f"Urgewald Only (Unmatched, Retained): {len(ur_retained)}")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="Coal_Companies_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
