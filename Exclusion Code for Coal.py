import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import re
import io

# ──────────────────────────────────────────────────────────────────────────────
# Helper ─ robust float conversion (handles “1.234,5” & “1,234.5”)
# ──────────────────────────────────────────────────────────────────────────────
def to_float(val):
    s = str(val).strip().replace(" ", "")
    if s in ("", "nan", "none"):
        return 0.0
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", ".") if len(parts) == 2 and len(parts[1]) in (1, 2) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ───────── column utilities ───────────────────────────────────────────────────
def make_columns_unique(df):
    seen, new_cols = {}, []
    for c in df.columns:
        seen[c] = seen.get(c, 0) + 1
        new_cols.append(c if seen[c] == 1 else f"{c}_{seen[c]-1}")
    df.columns = new_cols
    return df


def fuzzy_rename_columns(df, rename_map):
    used = set()
    for final_name, pats in rename_map.items():
        for col in df.columns:
            if col in used:
                continue
            if final_name == "Company" and col.strip().lower() == "parent company":
                continue
            if any(p.lower() in col.lower() for p in pats):
                df.rename(columns={col: final_name}, inplace=True)
                used.add(col)
                break
    return df


def normalize_key(s: str) -> str:
    return re.sub(r"[^\w\s]", "", re.sub(r"\s+", " ", s.lower())).strip()


# ───────── Loaders ────────────────────────────────────────────────────────────
def load_spglobal(file, sheet="Sheet1"):
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[sheet]
    df = pd.DataFrame(list(ws.values))
    if len(df) < 6:
        st.error("SPGlobal sheet too short"); return pd.DataFrame()

    hdr1, hdr2 = df.iloc[4].fillna(""), df.iloc[5].fillna("")
    cols = []
    for a, b in zip(hdr1, hdr2):
        col = str(a).strip()
        if b and b.lower() not in col.lower():
            col = f"{col} {b}".strip()
        cols.append(col)
    sp = df.iloc[6:].reset_index(drop=True)
    sp.columns = cols
    sp = make_columns_unique(sp)

    rename_map = {
        "SP_ENTITY_NAME":  ["entity name"],
        "SP_ENTITY_ID":    ["entity id"],
        "SP_COMPANY_ID":   ["company id"],
        "SP_ISIN":         ["isin"],
        "SP_LEI":          ["lei"],
        "Generation (Thermal Coal)": ["generation (thermal coal)"],
        "Thermal Coal Mining":       ["thermal coal mining"],
        "Coal Share of Revenue":     ["coal share of revenue"],
        "Coal Share of Power Production": ["coal share of power production"],
        "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
        "Coal Industry Sector":      ["industry sector"],
        ">10MT / >5GW":              [">10mt", ">5gw"],
        "expansion":                 ["expansion"],
    }
    sp = fuzzy_rename_columns(sp, rename_map).astype(object)

    for c in [
        "Thermal Coal Mining", "Generation (Thermal Coal)",
        "Coal Share of Revenue", "Coal Share of Power Production",
        "Installed Coal Power Capacity (MW)",
    ]:
        if c in sp:
            sp[c] = sp[c].apply(to_float)
    return sp


def load_urgewald(file, sheet="GCEL 2024"):
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[sheet]
    df = pd.DataFrame(list(ws.values))
    if df.empty:
        st.error("Urgewald sheet empty"); return pd.DataFrame()

    header = df.iloc[0].fillna("")
    keep = header.str.strip().str.lower() != "parent company"
    ur = df.iloc[1:].reset_index(drop=True).loc[:, keep]
    ur.columns = [c for c in header if str(c).strip().lower() != "parent company"]
    ur = make_columns_unique(ur)

    rename_map = {
        "Company": ["company", "issuer name"],
        "ISIN equity": ["isin equity", "isin(eq)", "isin eq"],
        "LEI": ["lei"],
        "BB Ticker": ["bb ticker"],
        "Coal Industry Sector": ["industry sector"],
        ">10MT / >5GW": [">10mt", ">5gw"],
        "expansion": ["expansion"],
        "Coal Share of Power Production": ["coal share of power production"],
        "Coal Share of Revenue": ["coal share of revenue"],
        "Installed Coal Power Capacity (MW)": ["installed coal power capacity"],
        "Generation (Thermal Coal)": ["generation (thermal coal)"],
        "Thermal Coal Mining": ["thermal coal mining"],
    }
    ur = fuzzy_rename_columns(ur, rename_map).astype(object)

    for c in [
        "Thermal Coal Mining", "Generation (Thermal Coal)",
        "Coal Share of Revenue", "Coal Share of Power Production",
        "Installed Coal Power Capacity (MW)",
    ]:
        if c in ur:
            ur[c] = ur[c].apply(to_float)
    return ur


# ───────── Merge ──────────────────────────────────────────────────────────────
def merge_ur_into_sp(sp, ur):
    sp = sp.copy().astype(object)
    ur = ur.copy().astype(object)

    for key, sp_col, ur_col in [
        ("norm_isin",  "SP_ISIN",     "ISIN equity"),
        ("norm_lei",   "SP_LEI",      "LEI"),
        ("norm_name",  "SP_ENTITY_NAME", "Company"),
    ]:
        sp[key] = sp.get(sp_col, "").astype(str).apply(normalize_key)
        ur[key] = ur.get(ur_col, "").astype(str).apply(normalize_key)

    maps = [{k: i for i, k in enumerate(sp[key]) if k} for key in ["norm_isin", "norm_lei", "norm_name"]]
    ur_not = []
    for _, r in ur.iterrows():
        found = next((maps[i][r[k]] for i, k in enumerate(["norm_isin", "norm_lei", "norm_name"]) if r[k] in maps[i]), None)
        if found is not None:
            for c, v in r.items():
                if c.startswith("norm_"): continue
                if c not in sp or pd.isna(sp.at[found, c]) or str(sp.at[found, c]).strip() == "":
                    sp.at[found, c] = v
            sp.at[found, "Merged"] = True
        else:
            ur_not.append(r)

    sp["Merged"] = sp.get("Merged", False).fillna(False)
    ur_only = pd.DataFrame(ur_not)
    ur_only["Merged"] = False
    for col in [c for c in sp.columns if c.startswith("norm_")]:
        sp.drop(columns=col, inplace=True, errors="ignore")
        ur_only.drop(columns=col, inplace=True, errors="ignore")
    return sp, ur_only


# ───────── Exclusion Logic ────────────────────────────────────────────────────
def cmp(val, thr, ge):  # comparison helper
    return val >= thr if ge else val > thr


def sym(ge):            # symbol helper
    return "≥" if ge else ">"


def compute_exclusion(row, p):
    reasons = []

    # numeric
    sp_min = row.get("Thermal Coal Mining", 0.0)
    sp_pow = row.get("Generation (Thermal Coal)", 0.0)

    ur_rev = row.get("Coal Share of Revenue", 0.0)
    ur_rev = ur_rev if ur_rev > 1 else ur_rev * 100

    ur_pp = row.get("Coal Share of Power Production", 0.0)
    ur_pp = ur_pp if ur_pp > 1 else ur_pp * 100

    # misc
    cap = row.get("Installed Coal Power Capacity (MW)", 0.0)
    prod_flag = "10mt" in str(row.get(">10MT / >5GW", "")).lower()
    expansion_text = str(row.get("expansion", "")).lower()

    # sector parsing
    sector_raw = str(row.get("Coal Industry Sector", "")).lower()
    mining_kw, power_kw = ("mining", "extraction", "producer"), ("power", "generation", "utility", "electric")
    tokens = [t.strip() for t in re.split(r"[;,/]|(?:\s*\n\s*)", sector_raw) if t.strip()]
    mining = [t for t in tokens if any(k in t for k in mining_kw)]
    power  = [t for t in tokens if any(k in t for k in power_kw)]
    others = [t for t in tokens if t not in mining + power]

    is_mining_only = bool(mining) and not power and not others
    is_power_only  = bool(power)  and not mining and not others
    is_mixed       = bool(mining) and bool(power) and not others

    # Toggles
    if p["exclude_mt"] and prod_flag:
        reasons.append(">10 MT indicator")

    if p["exclude_capacity"] and cmp(cap, p["capacity_threshold"], p["capacity_ge"]):
        reasons.append(f"Installed capacity {cap:.0f} MW {sym(p['capacity_ge'])} {p['capacity_threshold']} MW")

    if p["exclude_power_prod"] and cmp(ur_pp, p["power_prod_threshold"], p["power_prod_ge"]):
        reasons.append(f"Coal power production {ur_pp:.2f}% {sym(p['power_prod_ge'])} {p['power_prod_threshold']}%")

    # SP rules
    if p["sp_mining_checkbox"] and cmp(sp_min, p["sp_mining_threshold"], p["sp_mining_ge"]):
        reasons.append(f"SP mining revenue {sp_min:.2f}% {sym(p['sp_mining_ge'])} {p['sp_mining_threshold']}%")
    if p["sp_power_checkbox"] and cmp(sp_pow, p["sp_power_threshold"], p["sp_power_ge"]):
        reasons.append(f"SP power revenue {sp_pow:.2f}% {sym(p['sp_power_ge'])} {p['sp_power_threshold']}%")
    if p["sp_level2_checkbox"]:
        combo = sp_min + sp_pow
        if cmp(combo, p["sp_level2_threshold"], p["sp_level2_ge"]):
            reasons.append(f"SP level-2 combined {combo:.2f}% {sym(p['sp_level2_ge'])} {p['sp_level2_threshold']}%")

    # UR rules
    if is_mining_only and p["ur_mining_checkbox"] and cmp(ur_rev, p["ur_mining_threshold"], p["ur_mining_ge"]):
        reasons.append(f"UR mining revenue {ur_rev:.2f}% {sym(p['ur_mining_ge'])} {p['ur_mining_threshold']}%")
    if is_power_only and p["ur_power_checkbox"] and cmp(ur_rev, p["ur_power_threshold"], p["ur_power_ge"]):
        reasons.append(f"UR power revenue {ur_rev:.2f}% {sym(p['ur_power_ge'])} {p['ur_power_threshold']}%")
    if is_mixed and p["ur_mixed_checkbox"] and cmp(ur_rev, p["ur_mixed_threshold"], p["ur_mixed_ge"]):
        reasons.append(f"UR mixed revenue {ur_rev:.2f}% {sym(p['ur_mixed_ge'])} {p['ur_mixed_threshold']}%")
    if p["ur_level2_checkbox"] and cmp(ur_rev, p["ur_level2_threshold"], p["ur_level2_ge"]):
        reasons.append(f"UR level-2 revenue {ur_rev:.2f}% {sym(p['ur_level2_ge'])} {p['ur_level2_threshold']}%")

    # expansion
    if any(kw.lower() in expansion_text for kw in p["expansion_exclude"]):
        reasons.append("Expansion flag")

    return pd.Series([bool(reasons), "; ".join(reasons)], index=["Excluded", "Exclusion Reasons"])


# ───────── Streamlit UI ───────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Coal Exclusion Filter – per-threshold ≥ toggles", layout="wide")
    st.title("Coal Exclusion Filter")

    # files
    st.sidebar.header("Files & Sheets")
    sp_sheet = st.sidebar.text_input("SPGlobal sheet", "Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald sheet", "GCEL 2024")
    sp_file  = st.sidebar.file_uploader("SPGlobal .xlsx", type=["xlsx"])
    ur_file  = st.sidebar.file_uploader("Urgewald .xlsx", type=["xlsx"])

    # helper for numeric + ≥ checkbox
    def num_ge(label, default, key):
        col1, col2 = st.columns([3, 1])
        with col1:
            val = st.number_input(label, value=default, key=f"{key}_val")
        with col2:
            ge  = st.checkbox("≥", value=False, key=f"{key}_ge")
        return val, ge

    # Mining
    with st.sidebar.expander("Mining thresholds", True):
        ur_mining_cb = st.checkbox("UR: apply mining-only rule", False)
        ur_mining_thr, ur_mining_ge = num_ge("UR Mining threshold (%)", 5.0, "urmin")
        sp_mining_cb = st.checkbox("SP: apply mining-only rule", True)
        sp_mining_thr, sp_mining_ge = num_ge("SP Mining threshold (%)", 5.0, "spmin")

    # Power
    with st.sidebar.expander("Power thresholds", True):
        ur_power_cb = st.checkbox("UR: apply power-only rule", False)
        ur_power_thr, ur_power_ge = num_ge("UR Power threshold (%)", 20.0, "urpow")
        sp_power_cb = st.checkbox("SP: apply power-only rule", True)
        sp_power_thr, sp_power_ge = num_ge("SP Power threshold (%)", 20.0, "sppow")

        power_prod_thr, power_prod_ge = num_ge("Coal power production %", 20.0, "cpp")
        capacity_thr, capacity_ge = num_ge("Installed capacity (MW)", 10_000.0, "cap")
        exclude_mt = st.checkbox("Flag >10 MT indicator", True)

    # Mixed + Level-2
    with st.sidebar.expander("Mixed & Level-2", False):
        ur_mixed_cb = st.checkbox("UR: mixed mining+power rule", False)
        ur_mixed_thr, ur_mixed_ge = num_ge("UR Mixed threshold (%)", 25.0, "urmix")
        ur_L2_cb = st.checkbox("UR: Level-2", False)
        ur_L2_thr, ur_L2_ge = num_ge("UR Level-2 threshold (%)", 10.0, "url2")
        sp_L2_cb = st.checkbox("SP: Level-2 (mining+power)", False)
        sp_L2_thr, sp_L2_ge = num_ge("SP Level-2 threshold (%)", 10.0, "spl2")

    # Expansion
    with st.sidebar.expander("Expansion keywords"):
        expansion_exclude = st.text_input("Comma-separated keywords", "mining,infrastructure,power").split(",")

    run = st.sidebar.button("Run")
    if not run:
        st.stop()

    if not sp_file or not ur_file:
        st.warning("Please upload both files")
        st.stop()

    # Load & merge
    sp_df = load_spglobal(sp_file, sp_sheet)
    ur_df = load_urgewald(ur_file, ur_sheet)
    if sp_df.empty or ur_df.empty:
        st.warning("One of the sheets could not be loaded")
        st.stop()

    merged_sp, ur_only = merge_ur_into_sp(sp_df, ur_df)
    for df in (merged_sp, ur_only):
        df["Merged"] = df.get("Merged", False).fillna(False)

    sp_merged   = merged_sp[ merged_sp["Merged"] ]
    sp_unmerged = merged_sp[~merged_sp["Merged"] ]
    ur_unmerged = ur_only  [~ur_only ["Merged"] ]

    sp_only = sp_unmerged[
        (sp_unmerged["Thermal Coal Mining"] > 0) |
        (sp_unmerged["Generation (Thermal Coal)"] > 0)
    ]

    # param dict
    params = dict(
        # MT / capacity / power-prod
        exclude_mt=True,
        exclude_capacity=True,
        capacity_threshold=capacity_thr,
        capacity_ge=capacity_ge,
        exclude_power_prod=True,
        power_prod_threshold=power_prod_thr,
        power_prod_ge=power_prod_ge,

        # SP
        sp_mining_checkbox=sp_mining_cb,
        sp_mining_threshold=sp_mining_thr,
        sp_mining_ge=sp_mining_ge,
        sp_power_checkbox=sp_power_cb,
        sp_power_threshold=sp_power_thr,
        sp_power_ge=sp_power_ge,
        sp_level2_checkbox=sp_L2_cb,
        sp_level2_threshold=sp_L2_thr,
        sp_level2_ge=sp_L2_ge,

        # UR
        ur_mining_checkbox=ur_mining_cb,
        ur_mining_threshold=ur_mining_thr,
        ur_mining_ge=ur_mining_ge,
        ur_power_checkbox=ur_power_cb,
        ur_power_threshold=ur_power_thr,
        ur_power_ge=ur_power_ge,
        ur_mixed_checkbox=ur_mixed_cb,
        ur_mixed_threshold=ur_mixed_thr,
        ur_mixed_ge=ur_mixed_ge,
        ur_level2_checkbox=ur_L2_cb,
        ur_level2_threshold=ur_L2_thr,
        ur_level2_ge=ur_L2_ge,

        # expansion
        expansion_exclude=[k.strip() for k in expansion_exclude if k.strip()],
    )

    # apply filter
    def apply(df):
        if df.empty:
            return df.assign(Excluded=False, **{"Exclusion Reasons": ""})
        res = df.apply(lambda r: compute_exclusion(r, params), axis=1, result_type="expand")
        df["Excluded"], df["Exclusion Reasons"] = res["Excluded"], res["Exclusion Reasons"]
        return df

    sp_merged, sp_only, ur_unmerged = map(apply, (sp_merged, sp_only, ur_unmerged))

    excluded = pd.concat([d[d.Excluded] for d in (sp_merged, sp_only, ur_unmerged)], ignore_index=True)
    retained_merged = sp_merged[~sp_merged.Excluded]
    sp_retained     = sp_only  [~sp_only .Excluded]
    ur_retained     = ur_unmerged[~ur_unmerged.Excluded]

    cols = [
        "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
        "Coal Industry Sector","Company",">10MT / >5GW",
        "Installed Coal Power Capacity (MW)",
        "Coal Share of Power Production","Coal Share of Revenue","expansion",
        "Generation (Thermal Coal)","Thermal Coal Mining",
        "BB Ticker","ISIN equity","LEI","Excluded","Exclusion Reasons",
    ]
    for df in (excluded, retained_merged, sp_retained, ur_retained):
        for c in cols:
            if c not in df:
                df[c] = ""
        df = df[cols]  # reorder

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        excluded       .to_excel(w, "Excluded Companies", index=False)
        retained_merged.to_excel(w, "Retained Companies", index=False)
        sp_retained    .to_excel(w, "S&P Only", index=False)
        ur_retained    .to_excel(w, "Urgewald Only", index=False)

    st.subheader("Results")
    st.write(f"Excluded: **{len(excluded)}**")
    st.write(f"Retained (Merged): **{len(retained_merged)}**")
    st.write(f"S&P-only retained: **{len(sp_retained)}**")
    st.write(f"Urgewald-only retained: **{len(ur_retained)}**")

    st.download_button(
        "Download Excel",
        data=buf.getvalue(),
        file_name="Coal_Companies_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
