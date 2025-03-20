import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

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

def load_spglobal_data(file, sheet_name):
    try:
        # If row #5 and row #6 in Excel are both header rows, do header=[4,5].
        df = pd.read_excel(file, sheet_name=sheet_name, header=[4,5])
        # Flatten the 2-level header:
        df.columns = [
            " ".join(str(x).strip() for x in col if x not in (None, ""))
            for col in df.columns
        ]
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(x).strip() for x in col if x is not None)
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading Urgewald sheet '{sheet_name}': {e}")
        return None

def find_column(df, must_keywords, exclude_keywords=None):
    if exclude_keywords is None:
        exclude_keywords = []
    for col in df.columns:
        col_lower = col.lower()
        if all(mk.lower() in col_lower for mk in must_keywords):
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue
            return col
    return None

def rename_with_prefix(df, prefix):
    df.columns = [f"{prefix}{c}" for c in df.columns]
    return df

def merge_sp_ur(sp_df, ur_df):
    merged = pd.merge(
        left=sp_df, 
        right=ur_df, 
        how="outer",
        left_on=["SP_ISIN", "SP_LEI", "SP_Company"],
        right_on=["UR_ISIN", "UR_LEI", "UR_Company"]
    )
    return merged

def coalesce_cols(df, sp_col, ur_col, out_col):
    if sp_col not in df.columns:
        df[sp_col] = np.nan
    if ur_col not in df.columns:
        df[ur_col] = np.nan
    df[out_col] = df[sp_col].fillna(df[ur_col])

def filter_companies(
    df,
    mining_prod_mt_threshold,
    power_rev_threshold,
    power_prod_threshold_percent,
    capacity_threshold_mw,
    services_rev_threshold,
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    expansions_global,
    column_mapping,
    gen_coal_col,
    thermal_mining_col,
    metallurgical_mining_col,
    gen_thermal_coal_threshold,
    thermal_coal_mining_threshold,
    metallurgical_coal_mining_threshold,
    exclude_generation_thermal_coal,
    exclude_thermal_coal_mining,
    exclude_metallurgical_coal_mining
):
    exclusion_flags = []
    exclusion_reasons = []

    sec_col  = column_mapping["sector_col"]
    rev_col  = column_mapping["coal_rev_col"]
    power_col= column_mapping["coal_power_col"]
    cap_col  = column_mapping["capacity_col"]
    prod_col = column_mapping["production_col"]
    exp_col  = column_mapping["expansion_col"]

    for idx, row in df.iterrows():
        reasons = []
        sector_raw = str(row.get(sec_col, "")).strip()
        s_lc = sector_raw.lower()

        is_mining   = ("mining" in s_lc)
        is_power    = ("power" in s_lc) or ("generation" in s_lc)
        is_services = ("services" in s_lc)

        coal_share = pd.to_numeric(row.get(rev_col, 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(power_col, 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(cap_col, 0), errors="coerce") or 0.0

        prod_val = str(row.get(prod_col, "")).lower()
        exp_text = str(row.get(exp_col, "")).lower().strip()

        # Mining (sector-based)
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # Power (sector-based)
        if is_power and exclude_power:
            if exclude_power_rev and (coal_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            if exclude_capacity_mw and (installed_capacity > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold_mw}MW")

        # Services (sector-based)
        if is_services and exclude_services:
            if exclude_services_rev and (coal_share * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {services_rev_threshold}% (Services)")

        # Global expansions
        if expansions_global:
            for choice in expansions_global:
                if choice.lower() in exp_text:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break

        # Direct numeric columns
        if exclude_generation_thermal_coal:
            val_gen = pd.to_numeric(row.get(gen_coal_col, 0), errors="coerce") or 0.0
            if val_gen > gen_thermal_coal_threshold:
                reasons.append(f"Generation (Thermal Coal) {val_gen:.2f} > {gen_thermal_coal_threshold}")

        if exclude_thermal_coal_mining:
            val_thermal = pd.to_numeric(row.get(thermal_mining_col, 0), errors="coerce") or 0.0
            if val_thermal > thermal_coal_mining_threshold:
                reasons.append(f"Thermal Coal Mining {val_thermal:.2f} > {thermal_coal_mining_threshold}")

        if exclude_metallurgical_coal_mining:
            val_met = pd.to_numeric(row.get(metallurgical_mining_col, 0), errors="coerce") or 0.0
            if val_met > metallurgical_coal_mining_threshold:
                reasons.append(f"Metallurgical Coal Mining {val_met:.2f} > {metallurgical_coal_mining_threshold}")

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter (SP + Urgewald Merge)")

    st.sidebar.header("File & Sheet Settings")
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    urgewald_sheet = st.sidebar.text_input("Urgewald GCEL Sheet Name", value="GCEL 2024")
    urgewald_file  = st.sidebar.file_uploader("Upload Urgewald GCEL Excel file", type=["xlsx"])

    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT for Mining", value=True)

    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed coal power capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded", value=True)

    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded", value=False)

    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansion text contains any of these",
        expansions_possible,
        default=[]
    )

    st.sidebar.header("SPGlobal Coal Sectors (Direct Numeric)")
    gen_thermal_coal_threshold = st.sidebar.number_input("Generation (Thermal Coal) Threshold", value=0.05)
    exclude_generation_thermal_coal = st.sidebar.checkbox("Exclude Generation (Thermal Coal)?", value=True)
    thermal_coal_mining_threshold = st.sidebar.number_input("Thermal Coal Mining Threshold", value=0.05)
    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude Thermal Coal Mining?", value=True)
    metallurgical_coal_mining_threshold = st.sidebar.number_input("Metallurgical Coal Mining Threshold", value=0.05)
    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude Metallurgical Coal Mining?", value=True)

    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please upload both SPGlobal and Urgewald GCEL files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        if sp_df is None:
            return

        # 2) Load Urgewald
        ur_df = load_urgewald_data(urgewald_file, urgewald_sheet)
        if ur_df is None:
            return

        # 3) Make sure we have the columns as strings if we intend to merge on them:
        #    We have not renamed anything yet, so let's rename first to ensure
        #    the columns exist and are properly named before we coerce them to string.

        # Identify columns in SP
        sp_company_col = find_column(sp_df, ["company"]) or "Company"
        sp_isin_col    = find_column(sp_df, ["isin"]) or "ISIN"
        sp_lei_col     = find_column(sp_df, ["lei"])  or "LEI"
        sp_gen_coal_col = find_column(sp_df, ["generation","thermal","coal"]) or "GenerationThermalCoal_SP"
        sp_thermal_mining_col = find_column(sp_df, ["thermal","coal","mining"]) or "ThermalCoalMining_SP"
        sp_met_mining_col = find_column(sp_df, ["metallurgical","coal","mining"]) or "MetallurgicalCoalMining_SP"

        rename_map_sp = {
            sp_company_col: "SP_Company",
            sp_isin_col: "SP_ISIN",
            sp_lei_col: "SP_LEI",
            sp_gen_coal_col: "SP_GenThermal",
            sp_thermal_mining_col: "SP_ThermalMining",
            sp_met_mining_col: "SP_MetMining"
        }
        sp_df.rename(columns=rename_map_sp, inplace=True, errors="ignore")

        # Identify columns in UR
        ur_company_col = find_column(ur_df, ["company"]) or "Company"
        ur_isin_col    = find_column(ur_df, ["isin"]) or "ISIN"
        ur_lei_col     = find_column(ur_df, ["lei"])  or "LEI"
        ur_gen_coal_col = find_column(ur_df, ["generation","thermal","coal"]) or "GenerationThermalCoal_UR"
        ur_thermal_mining_col = find_column(ur_df, ["thermal","coal","mining"]) or "ThermalCoalMining_UR"
        ur_met_mining_col = find_column(ur_df, ["metallurgical","coal","mining"]) or "MetallurgicalCoalMining_UR"

        rename_map_ur = {
            ur_company_col: "UR_Company",
            ur_isin_col: "UR_ISIN",
            ur_lei_col: "UR_LEI",
            ur_gen_coal_col: "UR_GenThermal",
            ur_thermal_mining_col: "UR_ThermalMining",
            ur_met_mining_col: "UR_MetMining"
        }
        ur_df.rename(columns=rename_map_ur, inplace=True, errors="ignore")

        # Now we can safely coerce them to strings, if they exist:
        for col in ["SP_Company", "SP_ISIN", "SP_LEI"]:
            if col in sp_df.columns:
                sp_df[col] = sp_df[col].astype(str)

        for col in ["UR_Company", "UR_ISIN", "UR_LEI"]:
            if col in ur_df.columns:
                ur_df[col] = ur_df[col].astype(str)

        # 4) Merge
        merged_df = merge_sp_ur(sp_df, ur_df)

        # 5) Coalesce numeric columns
        coalesce_cols(merged_df, "SP_GenThermal", "UR_GenThermal", "Generation (Thermal Coal)")
        coalesce_cols(merged_df, "SP_ThermalMining", "UR_ThermalMining", "Thermal Coal Mining")
        coalesce_cols(merged_df, "SP_MetMining", "UR_MetMining", "Metallurgical Coal Mining")

        # 6) Build or coalesce sector
        #    (skipping details, same as before)
        merged_df["Sector"] = np.nan  # etc. or do a coalesce approach

        # 7) Filter
        column_mapping = {
            "sector_col":    "Sector",
            "coal_rev_col":  "SP_Coal Share of Revenue",   # or coalesced
            "coal_power_col":"SP_Coal Share of Power Production",
            "capacity_col":  "SP_Installed Coal Power Capacity (MW)",
            "production_col":"SP_>10MT / >5GW",
            "expansion_col": "SP_Expansion",
            "company_col":   "SP_Company",
            "ticker_col":    "SP_BB Ticker",
            "isin_col":      "SP_ISIN",
            "lei_col":       "SP_LEI"
        }

        filtered_df = filter_companies(
            merged_df,
            mining_prod_mt_threshold,
            power_rev_threshold,
            power_prod_threshold_percent,
            capacity_threshold_mw,
            services_rev_threshold,
            exclude_mining,
            exclude_power,
            exclude_services,
            exclude_mining_prod_mt,
            exclude_power_rev,
            exclude_power_prod_percent,
            exclude_capacity_mw,
            exclude_services_rev,
            expansions_global,
            column_mapping,
            gen_coal_col="Generation (Thermal Coal)",
            thermal_mining_col="Thermal Coal Mining",
            metallurgical_mining_col="Metallurgical Coal Mining",
            gen_thermal_coal_threshold=gen_thermal_coal_threshold,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            exclude_generation_thermal_coal=exclude_generation_thermal_coal,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining
        )

        # 8) Final output
        excluded_df = filtered_df[filtered_df["Excluded"] == True].copy()
        retained_df = filtered_df[filtered_df["Excluded"] == False].copy()
        no_data_df  = filtered_df[filtered_df["Sector"].isna()].copy()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data (Sector)", index=False)

        st.subheader("Statistics")
        st.write(f"Total merged rows: {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"No data in 'Sector': {len(no_data_df)}")

        st.subheader("Excluded (preview)")
        st.dataframe(excluded_df.head(50))

        st.download_button(
            label="Download Merged Results",
            data=output.getvalue(),
            file_name="merged_filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()

