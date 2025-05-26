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
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in (1,2):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0

##############################################
# 1. MAKE COLUMNS UNIQUE
##############################################
def make_columns_unique(df):
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
        sp_df = fuzzy_rename_columns(sp_df, rename_map_sp).astype(object)

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
        full_df = pd.DataFrame(data)
        if full_df.empty:
            raise ValueError("Urgewald file is empty.")
        header = full_df.iloc[0].fillna("")
        keep = header.str.strip().str.lower() != "parent company"
        ur_df = full_df.iloc[1:].reset_index(drop=True).loc[:, keep]
        filtered_header = [col for col in header if str(col).strip().lower() != "parent company"]
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
        ur_df = fuzzy_rename_columns(ur_df, rename_map_ur).astype(object)

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
# 6. MERGE URGEWALD INTO SP
##############################################
def merge_ur_into_sp_opt(sp_df, ur_df):
    sp = sp_df.copy().astype(object)
    ur = ur_df.copy().astype(object)

    sp["norm_isin"] = sp.get("SP_ISIN","").astype(str).apply(normalize_key)
    sp["norm_lei"]  = sp.get("SP_LEI","").astype(str).apply(normalize_key)
    sp["norm_name"] = sp.get("SP_ENTITY_NAME","").astype(str).apply(normalize_key)
    for col in ["ISIN equity","LEI","Company"]:
        if col not in ur:
            ur[col] = ""
    ur["norm_isin"]    = ur["ISIN equity"].astype(str).apply(normalize_key)
    ur["norm_lei"]     = ur["LEI"].astype(str).apply(normalize_key)
    ur["norm_company"] = ur["Company"].astype(str).apply(normalize_key)

    isin_map, lei_map, name_map = {}, {}, {}
    for idx, r in sp.iterrows():
        if r["norm_isin"]:    isin_map.setdefault(r["norm_isin"], idx)
        if r["norm_lei"]:     lei_map.setdefault(r["norm_lei"], idx)
        if r["norm_name"]:    name_map.setdefault(r["norm_name"], idx)

    ur_not = []
    for _, r in ur.iterrows():
        found = None
        if r["norm_isin"] in isin_map:
            found = isin_map[r["norm_isin"]]
        elif r["norm_lei"] in lei_map:
            found = lei_map[r["norm_lei"]]
        elif r["norm_company"] in name_map:
            found = name_map[r["norm_company"]]
        if found is not None:
            for c,v in r.items():
                if c.startswith("norm_"): continue
                if (c not in sp.columns) or pd.isnull(sp.loc[found,c]) or str(sp.loc[found,c]).strip()=="":
                    sp.loc[found,c] = v
            sp.loc[found,"Merged"] = True
        else:
            ur_not.append(r)

    if "Merged" not in sp:
        sp["Merged"] = False
    merged_sp = sp.copy()
    ur_only   = pd.DataFrame(ur_not)
    for c in ["norm_isin","norm_lei","norm_name"]:
        if c in merged_sp:
            merged_sp.drop(columns=[c], inplace=True)
    for c in ["norm_isin","norm_lei","norm_company"]:
        if c in ur_only:
            ur_only.drop(columns=[c], inplace=True)
    if "Merged" not in ur_only:
        ur_only["Merged"] = False

    return merged_sp, ur_only

##############################################
# 7. THRESHOLD FILTERING
##############################################
def compute_exclusion(row, **params):
    reasons = []

    # --- S&P percentages ---
    sp_min = row.get("Thermal Coal Mining", 0)
    sp_pow = row.get("Generation (Thermal Coal)", 0)
    # assume values already as % (e.g. 5.0)
    # --- UR percentage ---
    _rev = row.get("Coal Share of Revenue", 0)
    ur_rev_pct = _rev if _rev > 1 else _rev * 100
    # boilerplate
    prod_str = str(row.get(">10MT / >5GW","")).lower()
    cap      = row.get("Installed Coal Power Capacity (MW)",0)
    expansion_str = str(row.get("expansion","")).lower()
    is_sp    = bool(str(row.get("SP_ENTITY_NAME","")).strip())
    sector   = str(row.get("Coal Industry Sector","")).lower()

    # 1) >10MT
    if params["exclude_mt"] and "10mt" in prod_str:
        reasons.append(f">10MT indicated (threshold {params['mt_threshold']}MT)")
    # 2) capacity
    if params["exclude_capacity"] and cap > params["capacity_threshold"]:
        reasons.append(f"Installed capacity {cap:.0f}MW > {params['capacity_threshold']}MW")

    if is_sp:
        # S&P Level 1
        if params["sp_mining_checkbox"] and sp_min > params["sp_mining_threshold"]:
            reasons.append(f"SP Mining revenue {sp_min:.2f}% > {params['sp_mining_threshold']}%")
        if params["sp_power_checkbox"] and sp_pow > params["sp_power_threshold"]:
            reasons.append(f"SP Power revenue {sp_pow:.2f}% > {params['sp_power_threshold']}%")
        # S&P Level 2 (combined)
        if params["sp_level2_checkbox"]:
            combo_sp = sp_min + sp_pow
            if combo_sp > params["sp_level2_threshold"]:
                reasons.append(f"SP Level 2 combined {combo_sp:.2f}% > {params['sp_level2_threshold']}%")
    else:
        # UR Level 1 mining‐only
        if params["ur_mining_checkbox"] and ("mining" in sector) and not ("power" in sector) and ur_rev_pct > params["ur_mining_threshold"]:
            reasons.append(f"UR Mining revenue {ur_rev_pct:.2f}% > {params['ur_mining_threshold']}%")
        # UR Level 1 power‐only
        if params["ur_power_checkbox"] and ("power" in sector) and not ("mining" in sector) and ur_rev_pct > params["ur_power_threshold"]:
            reasons.append(f"UR Power revenue {ur_rev_pct:.2f}% > {params['ur_power_threshold']}%")
        # UR Mixed Level 1 (mining & power)
        if params["ur_mixed_checkbox"] and ("mining" in sector) and ("power" in sector) and ur_rev_pct > params["ur_mixed_threshold"]:
            reasons.append(f"UR Mixed mining&power revenue {ur_rev_pct:.2f}% > {params['ur_mixed_threshold']}%")
        # UR Level 2 (overall share)
        if params["ur_level2_checkbox"] and ur_rev_pct > params["ur_level2_threshold"]:
            reasons.append(f"UR Level 2 revenue {ur_rev_pct:.2f}% > {params['ur_level2_threshold']}%")

    # 7) expansion
    for kw in params["expansion_exclude"]:
        if kw.lower() in expansion_str:
            reasons.append(f"Expansion matched '{kw}'")
            break

    return pd.Series([bool(reasons), "; ".join(reasons)],
                     index=["Excluded","Exclusion Reasons"])

##############################################
# 8. STREAMLIT MAIN
##############################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – Merged & Excluded", layout="wide")
    st.title("Coal Exclusion Filter")

    # Sidebar
    st.sidebar.header("File & Sheet Settings")
    sp_sheet   = st.sidebar.text_input("SPGlobal Sheet Name","Sheet1")
    ur_sheet   = st.sidebar.text_input("Urgewald Sheet Name","GCEL 2024")
    sp_file    = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file    = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])
    st.sidebar.markdown("---")

    with st.sidebar.expander("Mining", expanded=True):
        ur_mining_checkbox  = st.checkbox("UR: Exclude mining-only", False)
        ur_mining_threshold = st.number_input("UR Mining threshold (%)", value=5.0)
        sp_mining_checkbox  = st.checkbox("SP: Exclude mining-only", True)
        sp_mining_threshold = st.number_input("SP Mining threshold (%)", value=5.0)
        exclude_mt          = st.checkbox("Exclude >10MT", True)
        mt_threshold        = st.number_input("MT threshold", value=10.0)

    with st.sidebar.expander("Power", expanded=True):
        ur_power_checkbox  = st.checkbox("UR: Exclude power-only", False)
        ur_power_threshold = st.number_input("UR Power threshold (%)", value=20.0)
        sp_power_checkbox  = st.checkbox("SP: Exclude power-only", True)
        sp_power_threshold = st.number_input("SP Power threshold (%)", value=20.0)
        exclude_power_prod = st.checkbox("Exclude > power %", True)
        power_prod_threshold = st.number_input("Power % threshold", value=20.0)
        exclude_capacity   = st.checkbox("Exclude > capacity (MW)", True)
        capacity_threshold = st.number_input("Capacity threshold (MW)", value=10000.0)

    with st.sidebar.expander("UR Mixed Level 1", expanded=False):
        ur_mixed_checkbox  = st.checkbox("UR: Exclude mining & power", False)
        ur_mixed_threshold = st.number_input("UR Mixed threshold (%)", value=25.0)

    with st.sidebar.expander("UR Level 2", expanded=False):
        ur_level2_checkbox  = st.checkbox("UR: Apply Level 2", False)
        ur_level2_threshold = st.number_input("UR Level 2 threshold (%)", value=10.0)

    with st.sidebar.expander("SP Level 2", expanded=False):
        sp_level2_checkbox  = st.checkbox("SP: Apply Level 2", False)
        sp_level2_threshold = st.number_input("SP Level 2 threshold (%)", value=10.0)

    with st.sidebar.expander("Exclude expansions", expanded=False):
        expansions_possible = ["mining","infrastructure","power","subsidiary of a coal developer"]
        expansion_exclude   = st.multiselect("Exclude if expansion text contains", expansions_possible, [])

    st.sidebar.markdown("---")
    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please upload both files")
            return

        sp_df = load_spglobal(sp_file, sp_sheet)
        ur_df = load_urgewald(ur_file, ur_sheet)
        if sp_df.empty or ur_df.empty:
            st.warning("Error loading data")
            return

        merged_sp, ur_only = merge_ur_into_sp_opt(sp_df, ur_df)
        # ensure numeric columns are floats
        for df in (merged_sp, ur_only):
            df["Merged"] = df.get("Merged", False).fillna(False).astype(bool)
            for col in [
                "Thermal Coal Mining","Generation (Thermal Coal)",
                "Coal Share of Revenue","Coal Share of Power Production",
                "Installed Coal Power Capacity (MW)"
            ]:
                if col in df:
                    df[col] = df[col].apply(to_float)

        sp_merged   = merged_sp[ merged_sp["Merged"] ]
        sp_unmerged = merged_sp[ ~merged_sp["Merged"] ]
        ur_unmerged = ur_only[ ~ur_only["Merged"] ]

        sp_only = sp_unmerged[
            (sp_unmerged["Thermal Coal Mining"]>0) |
            (sp_unmerged["Generation (Thermal Coal)"]>0)
        ]

        # build params dict
        params = {
            # UR L1 mining & power mixed
            "ur_mixed_checkbox":    ur_mixed_checkbox,
            "ur_mixed_threshold":   ur_mixed_threshold,
            # UR L1 mining-only & power-only
            "ur_mining_checkbox":   ur_mining_checkbox,
            "ur_mining_threshold":  ur_mining_threshold,
            "ur_power_checkbox":    ur_power_checkbox,
            "ur_power_threshold":   ur_power_threshold,
            # SP L1 mining-only & power-only
            "sp_mining_checkbox":   sp_mining_checkbox,
            "sp_mining_threshold":  sp_mining_threshold,
            "sp_power_checkbox":    sp_power_checkbox,
            "sp_power_threshold":   sp_power_threshold,
            # >10MT / power-production / capacity
            "exclude_mt":           exclude_mt,
            "mt_threshold":         mt_threshold,
            "exclude_power_prod":   exclude_power_prod,
            "power_prod_threshold": power_prod_threshold,
            "exclude_capacity":     exclude_capacity,
            "capacity_threshold":   capacity_threshold,
            # UR L2
            "ur_level2_checkbox":   ur_level2_checkbox,
            "ur_level2_threshold":  ur_level2_threshold,
            # SP L2 combined
            "sp_level2_checkbox":   sp_level2_checkbox,
            "sp_level2_threshold":  sp_level2_threshold,
            # expansions
            "expansion_exclude":    expansion_exclude,
        }

        def apply_filter(df):
            if df.empty:
                return df.assign(Excluded=False, **{"Exclusion Reasons":""})
            filt = df.apply(lambda r: compute_exclusion(r, **params),
                            axis=1, result_type="expand")
            df["Excluded"] = filt["Excluded"]
            df["Exclusion Reasons"] = filt["Exclusion Reasons"]
            return df

        sp_merged   = apply_filter(sp_merged)
        sp_only     = apply_filter(sp_only)
        ur_unmerged = apply_filter(ur_unmerged)

        excluded_final  = pd.concat([
            sp_merged[sp_merged.Excluded],
            sp_only[sp_only.Excluded],
            ur_unmerged[ur_unmerged.Excluded]
        ], ignore_index=True)
        retained_merged = sp_merged[~sp_merged.Excluded]
        sp_retained     = sp_only[~sp_only.Excluded]
        ur_retained     = ur_unmerged[~ur_unmerged.Excluded]

        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Coal Industry Sector","Company",">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production","Coal Share of Revenue","expansion",
            "Generation (Thermal Coal)","Thermal Coal Mining",
            "BB Ticker","ISIN equity","LEI","Excluded","Exclusion Reasons"
        ]
        def finalize(df):
            for c in final_cols:
                if c not in df:
                    df[c] = ""
            df = df[final_cols]
            if "BB Ticker" in df:
                df["BB Ticker"] = df["BB Ticker"].astype(str).str.replace(r'\s*Equity','',regex=True)
            return df

        excluded_final  = finalize(excluded_final)
        retained_merged = finalize(retained_merged)
        sp_retained     = finalize(sp_retained)
        ur_retained     = finalize(ur_retained)

        # download
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
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
            data=out.getvalue(),
            file_name="Coal_Companies_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
