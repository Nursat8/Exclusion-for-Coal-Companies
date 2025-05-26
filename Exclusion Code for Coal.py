import streamlit as st
import pandas as pd
import openpyxl
import re
import io

##############################################
# Helper: robust conversion to float, auto-detects US vs EU format
##############################################
def to_float(val):
    s = str(val).strip().replace(" ", "")
    if not s or s.lower() in ("nan", "none"):
        return 0.0
    if "." in s and "," in s:
        # decide which is decimal sep
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
# Unique & fuzzy-rename helpers
##############################################
def make_columns_unique(df):
    seen, new = {}, []
    for c in df.columns:
        if c not in seen:
            seen[c] = 0
            new.append(c)
        else:
            seen[c] += 1
            new.append(f"{c}_{seen[c]}")
    df.columns = new
    return df

def fuzzy_rename_columns(df, rename_map):
    used = set()
    for target, patterns in rename_map.items():
        for c in df.columns:
            low = c.lower()
            if c in used: continue
            if target=="Company" and low=="parent company": continue
            if any(p.lower() in low for p in patterns):
                df.rename(columns={c: target}, inplace=True)
                used.add(c)
                break
    return df

def normalize_key(s):
    return re.sub(r'[^\w]', '', str(s).lower().strip())

##############################################
# Load SPGlobal (multi-header)
##############################################
def load_spglobal(f, sheet_name="Sheet1"):
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb[sheet_name]
    raw = list(ws.values)
    df = pd.DataFrame(raw)
    if len(df) < 6:
        st.error("SPGlobal needs ≥6 rows"); return pd.DataFrame()
    r5, r6 = df.iloc[4].fillna(""), df.iloc[5].fillna("")
    cols = []
    for i in range(df.shape[1]):
        t, b = str(r5[i]).strip(), str(r6[i]).strip()
        hdr = t or ""
        if b and b.lower() not in hdr.lower():
            hdr = (hdr + " " + b).strip()
        cols.append(hdr)
    data = df.iloc[6:].reset_index(drop=True)
    data.columns = cols
    data = make_columns_unique(data)
    rename_map = {
        "SP_ENTITY_NAME":  ["sp entity name","s&p entity name","entity name"],
        "SP_ENTITY_ID":    ["sp entity id","entity id"],
        "SP_COMPANY_ID":   ["sp company id","company id"],
        "SP_ISIN":         ["sp isin","isin code"],
        "SP_LEI":          ["sp lei","lei code"],
        "Generation (Thermal Coal)": ["generation (thermal coal)"],
        "Thermal Coal Mining":       ["thermal coal mining"],
        "Coal Share of Revenue":     ["coal share of revenue"],
        "Coal Share of Power Production": ["coal share of power production"],
        "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
        "Coal Industry Sector":      ["coal industry sector","industry sector"],
        ">10MT / >5GW":              ["10mt",">5gw"],
        "expansion":                 ["expansion"],
    }
    data = fuzzy_rename_columns(data, rename_map).astype(object)
    for c in ["Thermal Coal Mining","Generation (Thermal Coal)",
              "Coal Share of Revenue","Coal Share of Power Production",
              "Installed Coal Power Capacity (MW)"]:
        if c in data:
            data[c] = data[c].apply(to_float)
    return data

##############################################
# Load Urgewald (single header)
##############################################
def load_urgewald(f, sheet_name="GCEL 2024"):
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb[sheet_name]
    raw = list(ws.values)
    df = pd.DataFrame(raw)
    if df.empty:
        st.error("Urgewald empty"); return pd.DataFrame()
    header = df.iloc[0].fillna("")
    keep = header.str.strip().str.lower() != "parent company"
    data = df.iloc[1:].reset_index(drop=True).loc[:, keep]
    data.columns = [h for h, k in zip(header, keep) if k]
    data = make_columns_unique(data)
    rename_map = {
        "Company":        ["company","issuer name"],
        "ISIN equity":    ["isin equity","isin(eq)","isin eq"],
        "LEI":            ["lei","lei code"],
        "BB Ticker":      ["bb ticker","bloomberg ticker"],
        "Coal Industry Sector":["coal industry sector","industry sector"],
        ">10MT / >5GW":   ["10mt",">5gw"],
        "expansion":      ["expansion","expansion text"],
        "Coal Share of Power Production":["coal share of power production"],
        "Coal Share of Revenue":["coal share of revenue"],
        "Installed Coal Power Capacity (MW)":["installed coal power capacity"],
        "Generation (Thermal Coal)":["generation (thermal coal)"],
        "Thermal Coal Mining":["thermal coal mining"],
    }
    data = fuzzy_rename_columns(data, rename_map).astype(object)
    for c in ["Thermal Coal Mining","Generation (Thermal Coal)",
              "Coal Share of Revenue","Coal Share of Power Production",
              "Installed Coal Power Capacity (MW)"]:
        if c in data:
            data[c] = data[c].apply(to_float)
    return data

##############################################
# Merge Urgewald → SPGlobal
##############################################
def merge_ur_into_sp_opt(sp_df, ur_df):
    sp = sp_df.copy().astype(object)
    ur = ur_df.copy().astype(object)

    # build normalized keys in SP
    sp["norm_isin"] = sp.get("SP_ISIN","").astype(str).map(normalize_key)
    sp["norm_lei"]  = sp.get("SP_LEI","").astype(str).map(normalize_key)
    sp["norm_name"] = sp.get("SP_ENTITY_NAME","").astype(str).map(normalize_key)
    # ensure UR columns exist
    for col in ["ISIN equity","LEI","Company"]:
        ur.setdefault(col, "")
    ur["norm_isin"]    = ur["ISIN equity"].astype(str).map(normalize_key)
    ur["norm_lei"]     = ur["LEI"].astype(str).map(normalize_key)
    ur["norm_company"] = ur["Company"].astype(str).map(normalize_key)

    # lookup maps
    i_map = sp[sp.norm_isin!=""].set_index("norm_isin").index.to_series().to_dict()
    l_map = sp[sp.norm_lei!=""].set_index("norm_lei").index.to_series().to_dict()
    n_map = sp[sp.norm_name!=""].set_index("norm_name").index.to_series().to_dict()

    sp["Merged"] = False
    remainder = []
    for _, row in ur.iterrows():
        tgt = i_map.get(row.norm_isin) or l_map.get(row.norm_lei) or n_map.get(row.norm_company)
        if tgt is not None:
            for c, v in row.items():
                if c.startswith("norm_"): continue
                if c not in sp.columns or pd.isna(sp.at[tgt, c]) or not str(sp.at[tgt, c]).strip():
                    sp.at[tgt, c] = v
            sp.at[tgt, "Merged"] = True
        else:
            remainder.append(row)

    ur_only = pd.DataFrame(remainder)
    ur_only["Merged"] = False

    # drop norm cols
    for c in ["norm_isin","norm_lei","norm_name"]:
        sp.drop(columns=[c], inplace=True, errors="ignore")
    for c in ["norm_isin","norm_lei","norm_company"]:
        ur_only.drop(columns=[c], inplace=True, errors="ignore")

    return sp, ur_only

##############################################
# Exclusion logic
##############################################
def compute_exclusion(row, **p):
    reasons = []

    # SP percentages
    sp_min = row.get("Thermal Coal Mining", 0)
    sp_pow = row.get("Generation (Thermal Coal)", 0)
    sp_min_pct = sp_min * 100 if sp_min <= 1 else sp_min
    sp_pow_pct = sp_pow * 100 if sp_pow <= 1 else sp_pow

    # UR percentages
    ur_rev = row.get("Coal Share of Revenue", 0)
    ur_rev_pct = ur_rev * 100 if ur_rev <= 1 else ur_rev
    ur_pp = row.get("Coal Share of Power Production", 0)
    ur_pp_pct = ur_pp * 100 if ur_pp <= 1 else ur_pp

    # capacity & production flag
    cap = row.get("Installed Coal Power Capacity (MW)", 0)
    prod_str = str(row.get(">10MT / >5GW", "")).lower()
    exp_str = str(row.get("expansion", "")).lower()
    is_sp = bool(str(row.get("SP_ENTITY_NAME", "")).strip())
    sector = str(row.get("Coal Industry Sector", "")).lower()

    # 1) >10MT
    if p["exclude_mt"] and "10mt" in prod_str:
        reasons.append(f">10MT indicated (thr {p['mt_threshold']}MT)")

    # 2) capacity
    if p["exclude_capacity"] and cap > p["capacity_threshold"]:
        reasons.append(f"Cap {cap:.0f}MW > {p['capacity_threshold']}")

    # 3) UR power-production
    if p["exclude_power_prod"] and ur_pp_pct > p["power_prod_threshold"]:
        reasons.append(f"UR power prod {ur_pp_pct:.1f}% > {p['power_prod_threshold']}%")

    if is_sp:
        # S&P Level-1
        if p["sp_mining_checkbox"] and sp_min_pct > p["sp_mining_threshold"]:
            reasons.append(f"SP mining {sp_min_pct:.1f}% > {p['sp_mining_threshold']}%")
        if p["sp_power_checkbox"] and sp_pow_pct > p["sp_power_threshold"]:
            reasons.append(f"SP power  {sp_pow_pct:.1f}% > {p['sp_power_threshold']}%")
        # S&P Level-2
        if p["sp_level2_checkbox"]:
            combo = sp_min_pct + sp_pow_pct
            if combo > p["sp_level2_threshold"]:
                reasons.append(f"SP lvl2 combo {combo:.1f}% > {p['sp_level2_threshold']}%")
    else:
        # UR Level-1: mining-only
        if p["ur_mining_checkbox"] and "mining" in sector and "power" not in sector and ur_rev_pct > p["ur_mining_threshold"]:
            reasons.append(f"UR mining {ur_rev_pct:.1f}% > {p['ur_mining_threshold']}%")
        # UR Level-1: power-only
        if p["ur_power_checkbox"] and any(k in sector for k in ("power","generation")) and "mining" not in sector and ur_rev_pct > p["ur_power_threshold"]:
            reasons.append(f"UR power  {ur_rev_pct:.1f}% > {p['ur_power_threshold']}%")
        # UR Level-2
        if p["ur_level2_checkbox"] and ur_rev_pct > p["ur_level2_threshold"]:
            reasons.append(f"UR lvl2 rev {ur_rev_pct:.1f}% > {p['ur_level2_threshold']}%")

    # expansion
    for kw in p["expansion_exclude"]:
        if kw.lower() in exp_str:
            reasons.append(f"Expansion '{kw}'")
            break

    return pd.Series([bool(reasons), "; ".join(reasons)],
                     index=["Excluded","Exclusion Reasons"])

##############################################
# Streamlit UI
##############################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # file inputs
    sp_sheet = st.sidebar.text_input("SPGlobal sheet", "Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald sheet", "GCEL 2024")
    sp_file  = st.sidebar.file_uploader("SPGlobal (.xlsx)", type="xlsx")
    ur_file  = st.sidebar.file_uploader("Urgewald  (.xlsx)", type="xlsx")

    # thresholds
    with st.sidebar.expander("Mining"):
        ur_min_cb  = st.checkbox("UR mining L1", value=False)
        ur_min_thr = st.number_input("UR mining % thr", value=5.0)
        sp_min_cb  = st.checkbox("SP  mining L1", value=True)
        sp_min_thr = st.number_input("SP  mining % thr", value=5.0)
        excl_mt    = st.checkbox("Exclude >10MT", value=True)
        mt_thr     = st.number_input("10MT thr", value=10.0)

    with st.sidebar.expander("Power"):
        ur_pow_cb  = st.checkbox("UR power  L1", value=False)
        ur_pow_thr = st.number_input("UR power % thr", value=20.0)
        sp_pow_cb  = st.checkbox("SP  power  L1", value=True)
        sp_pow_thr = st.number_input("SP  power  % thr", value=20.0)
        excl_pp    = st.checkbox("Exclude >% prod", value=True)
        pp_thr     = st.number_input("Prod % thr", value=20.0)
        excl_cap   = st.checkbox("Exclude >capacity", value=True)
        cap_thr    = st.number_input("Cap MW thr", value=10000.0)

    with st.sidebar.expander("Urgewald Level 2"):
        ur_l2_cb  = st.checkbox("UR L2", value=False)
        ur_l2_thr = st.number_input("UR L2 % thr", value=10.0)

    with st.sidebar.expander("S&P Level 2"):
        sp_l2_cb  = st.checkbox("SP  L2", value=False)
        sp_l2_thr = st.number_input("SP  L2 % thr", value=10.0)

    with st.sidebar.expander("Exclude expansions"):
        exp_excl  = st.multiselect("Keywords", ["mining","infrastructure","power","subsidiary of a coal developer"], [])

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Upload both files!"); return

        sp_df = load_spglobal(sp_file, sp_sheet)
        ur_df = load_urgewald(ur_file, ur_sheet)
        if sp_df.empty or ur_df.empty:
            st.warning("Load error"); return

        merged_sp, ur_only = merge_ur_into_sp_opt(sp_df, ur_df)
        # ensure booleans & numeric recast
        for df in (merged_sp, ur_only):
            df["Merged"] = df.get("Merged", False).fillna(False).astype(bool)
            for c in ["Thermal Coal Mining","Generation (Thermal Coal)",
                      "Coal Share of Revenue","Coal Share of Power Production",
                      "Installed Coal Power Capacity (MW)"]:
                if c in df:
                    df[c] = df[c].apply(to_float)

        sp_merged   = merged_sp[ merged_sp["Merged"] ]
        sp_unmerged = merged_sp[ ~merged_sp["Merged"] ]
        ur_unmerged = ur_only[ ~ur_only["Merged"] ]

        # S&P-only for L1
        sp_only = sp_unmerged[
            (sp_unmerged["Thermal Coal Mining"]>0) |
            (sp_unmerged["Generation (Thermal Coal)"]>0)
        ]

        params = dict(
            ur_mining_checkbox=ur_min_cb, ur_mining_threshold=ur_min_thr,
            ur_power_checkbox =ur_pow_cb, ur_power_threshold =ur_pow_thr,
            sp_mining_checkbox=sp_min_cb, sp_mining_threshold=sp_min_thr,
            sp_power_checkbox =sp_pow_cb, sp_power_threshold =sp_pow_thr,
            exclude_mt        =excl_mt,   mt_threshold      =mt_thr,
            exclude_power_prod=excl_pp,   power_prod_threshold=pp_thr,
            exclude_capacity  =excl_cap,  capacity_threshold =cap_thr,
            ur_level2_checkbox=ur_l2_cb,  ur_level2_threshold=ur_l2_thr,
            sp_level2_checkbox=sp_l2_cb,  sp_level2_threshold=sp_l2_thr,
            expansion_exclude =exp_excl,
        )

        def apply_filter(df):
            if df.empty:
                return df.assign(Excluded=False,**{"Exclusion Reasons":""})
            out = df.apply(lambda r: compute_exclusion(r, **params), axis=1)
            df[["Excluded","Exclusion Reasons"]] = out
            return df

        sp_m = apply_filter(sp_merged)
        sp_o = apply_filter(sp_only)
        ur_o = apply_filter(ur_unmerged)

        excluded_final  = pd.concat([sp_m[sp_m.Excluded], sp_o[sp_o.Excluded], ur_o[ur_o.Excluded]], ignore_index=True)
        retained_merged = sp_m[~sp_m.Excluded]
        sp_retained     = sp_o[~sp_o.Excluded]
        ur_retained     = ur_o[~ur_o.Excluded]

        # Write out
        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Coal Industry Sector","Company",">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production","Coal Share of Revenue","expansion",
            "Generation (Thermal Coal)","Thermal Coal Mining",
            "BB Ticker","ISIN equity","LEI",
            "Excluded","Exclusion Reasons"
        ]
        def finalize(df):
            for c in final_cols:
                if c not in df: df[c]=""
            df = df[final_cols]
            if "BB Ticker" in df:
                df["BB Ticker"] = df["BB Ticker"].str.replace(r'\s*Equity',"",regex=True)
            return df

        excl, retm, reto, retu = map(finalize, [excluded_final, retained_merged, sp_retained, ur_retained])

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            excl.to_excel(w, "Excluded Companies", index=False)
            retm.to_excel(w, "Retained Companies", index=False)
            reto.to_excel(w, "S&P Only", index=False)
            retu.to_excel(w, "Urgewald Only", index=False)

        st.download_button("Download Filtered Results", out.getvalue(),
                           "Coal_Companies_Output.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.success(f"Excluded: {len(excl)}, Retained Merged: {len(retm)}, S&P Only Retained: {len(reto)}, UR Only Retained: {len(retu)}")

if __name__ == "__main__":
    main()

