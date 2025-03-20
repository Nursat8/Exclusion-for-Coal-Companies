import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

##############################
# 1) MAKE COLUMNS UNIQUE
##############################
def make_columns_unique(df):
    """
    Ensures DataFrame df has uniquely named columns by appending a suffix 
    (e.g., '_1', '_2') if duplicates exist. This prevents 'InvalidIndexError'
    when concatenating or merging.
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

##############################
# 2) LOAD FUNCTIONS
##############################
def load_spglobal_data(file, sheet_name):
    """
    Load the SPGlobal file, skipping the first 5 rows so row #6 is treated as headers.
    Flatten columns if there's a multi-level header, then ensure columns are unique.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=[4,5])
        df.columns = [
            " ".join(str(x).strip() for x in col if x not in (None, ""))
            for col in df.columns
        ]
# Now df has combined headers from row #5 and row #6

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
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    """
    Load the Urgewald GCEL file normally (assuming row #1 is header).
    Flatten columns if multi-level, then ensure columns are unique.
    """
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

##############################
# 3) HELPERS: find_column + rename
##############################
def find_column(df, must_keywords, exclude_keywords=None):
    """
    Finds a column in df that contains all 'must_keywords' (case-insensitive)
    and excludes any column containing any 'exclude_keywords'.
    Returns the first matching column name or None if no match.
    """
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
    """
    Renames all columns in df to prefix + original_col,
    e.g. "SP_" or "UR_". 
    This helps keep them distinct after merging.
    """
    df.columns = [f"{prefix}{c}" for c in df.columns]
    return df

##############################
# 4) MERGE & COALESCE
##############################

# Ensure the columns used in the merge are all strings
for col in ["SP_Company", "SP_ISIN", "SP_LEI"]:
    if col in sp_df.columns:
        sp_df[col] = sp_df[col].astype(str)

for col in ["UR_Company", "UR_ISIN", "UR_LEI"]:
    if col in ur_df.columns:
        ur_df[col] = ur_df[col].astype(str)

# Now do the merge
merged_df = pd.merge(
    sp_df,
    ur_df,
    how="outer",
    left_on=["SP_Company", "SP_ISIN", "SP_LEI"],
    right_on=["UR_Company", "UR_ISIN", "UR_LEI"]
)

def merge_sp_ur(sp_df, ur_df):
    """
    Merge sp_df and ur_df on (ISIN, LEI, Company) with a full outer join.
    Returns the merged DataFrame.
    """
    merged = pd.merge(
        left=sp_df, 
        right=ur_df, 
        how="outer",
        left_on=["SP_ISIN", "SP_LEI", "SP_Company"],
        right_on=["UR_ISIN", "UR_LEI", "UR_Company"]
    )
    return merged

def coalesce_cols(df, sp_col, ur_col, out_col):
    """
    If sp_col is non-null, use that. Otherwise use ur_col.
    Creates a new column df[out_col].
    """
    # If the columns don't exist, fill with NaN
    if sp_col not in df.columns:
        df[sp_col] = np.nan
    if ur_col not in df.columns:
        df[ur_col] = np.nan

    df[out_col] = df[sp_col].fillna(df[ur_col])

##############################
# 5) FILTER COMPANIES
##############################
def filter_companies(
    df,
    # Mining thresholds
    mining_prod_mt_threshold,
    # Power thresholds
    power_rev_threshold,
    power_prod_threshold_percent,
    capacity_threshold_mw,
    # Services thresholds
    services_rev_threshold,
    # Exclusion toggles
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    # Global expansions
    expansions_global,
    column_mapping,
    # Direct numeric columns + thresholds + toggles
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
    """
    Apply your existing toggles & thresholds. 
    We detect "mining", "power", "services" from the sector 
    to exclude based on the user toggles. 
    We also exclude if the direct numeric columns for Generation, 
    Thermal Coal Mining, or Met Coal Mining exceed thresholds.
    """
    exclusion_flags = []
    exclusion_reasons = []

    sec_col    = column_mapping["sector_col"]
    rev_col    = column_mapping["coal_rev_col"]
    power_col  = column_mapping["coal_power_col"]
    cap_col    = column_mapping["capacity_col"]
    prod_col   = column_mapping["production_col"]
    exp_col    = column_mapping["expansion_col"]

    for idx, row in df.iterrows():
        reasons = []
        sector_raw = str(row.get(sec_col, "")).strip()
        s_lc = sector_raw.lower()

        # Basic sector detection
        is_mining    = ("mining" in s_lc)
        # We'll treat "generation" as power:
        is_power     = ("power" in s_lc) or ("generation" in s_lc)
        is_services  = ("services" in s_lc)

        # Numeric fields
        coal_share         = pd.to_numeric(row.get(rev_col, 0), errors="coerce") or 0.0
        coal_power_share   = pd.to_numeric(row.get(power_col, 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(cap_col, 0), errors="coerce") or 0.0

        prod_val = str(row.get(prod_col, "")).lower()
        exp_text = str(row.get(exp_col, "")).lower().strip()

        # ===== MINING (sector-based) =====
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # ===== POWER (sector-based) =====
        if is_power and exclude_power:
            if exclude_power_rev and (coal_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            if exclude_capacity_mw and (installed_capacity > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold_mw}MW")

        # ===== SERVICES (sector-based) =====
        if is_services and exclude_services:
            if exclude_services_rev and (coal_share * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {services_rev_threshold}% (Services)")

        # ===== GLOBAL EXPANSIONS (text) =====
        if expansions_global:
            for choice in expansions_global:
                if choice.lower() in exp_text:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break

        # ===== DIRECT NUMERIC COLUMNS (coalesced) =====
        # Generation (Thermal Coal)
        if exclude_generation_thermal_coal:
            val_gen = pd.to_numeric(row.get(gen_coal_col, 0), errors="coerce") or 0.0
            if val_gen > gen_thermal_coal_threshold:
                reasons.append(f"Generation (Thermal Coal) {val_gen:.2f} > {gen_thermal_coal_threshold}")

        # Thermal Coal Mining
        if exclude_thermal_coal_mining:
            val_thermal = pd.to_numeric(row.get(thermal_mining_col, 0), errors="coerce") or 0.0
            if val_thermal > thermal_coal_mining_threshold:
                reasons.append(f"Thermal Coal Mining {val_thermal:.2f} > {thermal_coal_mining_threshold}")

        # Metallurgical Coal Mining
        if exclude_metallurgical_coal_mining:
            val_met = pd.to_numeric(row.get(metallurgical_mining_col, 0), errors="coerce") or 0.0
            if val_met > metallurgical_coal_mining_threshold:
                reasons.append(f"Metallurgical Coal Mining {val_met:.2f} > {metallurgical_coal_mining_threshold}")

        # If any reasons => excluded
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

##############################
# 6) MAIN STREAMLIT APP
##############################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter (SP + Urgewald Merge)")

    # ============ FILE & SHEET =============
    st.sidebar.header("File & Sheet Settings")

    # 1) SPGlobal file
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    # 2) Urgewald file
    urgewald_sheet = st.sidebar.text_input("Urgewald GCEL Sheet Name", value="GCEL 2024")
    urgewald_file  = st.sidebar.file_uploader("Upload Urgewald GCEL Excel file", type=["xlsx"])

    # ============ MINING THRESHOLDS ============
    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT for Mining", value=True)

    # ============ POWER THRESHOLDS ============
    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)

    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded", value=True)

    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded", value=True)

    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed coal power capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded", value=True)

    # ============ SERVICES THRESHOLDS ============
    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded", value=False)

    # ============ GLOBAL EXPANSION EXCLUSION ============
    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansion text contains any of these",
        expansions_possible,
        default=[]
    )

    # ============ SPGlobal Coal Sectors (Direct Numeric) ============
    st.sidebar.header("SPGlobal Coal Sectors (Direct Numeric)")
    gen_thermal_coal_threshold = st.sidebar.number_input("Generation (Thermal Coal) Threshold", value=0.05)
    exclude_generation_thermal_coal = st.sidebar.checkbox("Exclude Generation (Thermal Coal)?", value=True)

    thermal_coal_mining_threshold = st.sidebar.number_input("Thermal Coal Mining Threshold", value=0.05)
    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude Thermal Coal Mining?", value=True)

    metallurgical_coal_mining_threshold = st.sidebar.number_input("Metallurgical Coal Mining Threshold", value=0.05)
    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude Metallurgical Coal Mining?", value=True)

    # ============ RUN FILTER =============
    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please upload both SPGlobal and Urgewald GCEL files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        if sp_df is None:
            return
        # Identify likely columns for key fields in SP
        sp_company_col = find_column(sp_df, ["company"]) or "Company"
        sp_isin_col    = find_column(sp_df, ["isin"]) or "ISIN"
        sp_lei_col     = find_column(sp_df, ["lei"])  or "LEI"
        sp_gen_coal_col = find_column(sp_df, ["generation","thermal","coal"]) or "GenerationThermalCoal_SP"
        sp_thermal_mining_col = find_column(sp_df, ["thermal","coal","mining"]) or "ThermalCoalMining_SP"
        sp_met_mining_col = find_column(sp_df, ["metallurgical","coal","mining"]) or "MetallurgicalCoalMining_SP"

        # Rename them with a "SP_" prefix so we keep them separate
        rename_map_sp = {}
        rename_map_sp[sp_company_col] = "SP_Company"
        rename_map_sp[sp_isin_col]    = "SP_ISIN"
        rename_map_sp[sp_lei_col]     = "SP_LEI"

        # For the 3 numeric columns, rename them similarly
        rename_map_sp[sp_gen_coal_col] = "SP_GenThermal"
        rename_map_sp[sp_thermal_mining_col] = "SP_ThermalMining"
        rename_map_sp[sp_met_mining_col] = "SP_MetMining"

        sp_df.rename(columns=rename_map_sp, inplace=True, errors="ignore")

        # 2) Load Urgewald
        ur_df = load_urgewald_data(urgewald_file, urgewald_sheet)
        if ur_df is None:
            return
        # Identify likely columns for key fields in UR
        ur_company_col = find_column(ur_df, ["company"]) or "Company"
        ur_isin_col    = find_column(ur_df, ["isin"]) or "ISIN"
        ur_lei_col     = find_column(ur_df, ["lei"])  or "LEI"
        ur_gen_coal_col = find_column(ur_df, ["generation","thermal","coal"]) or "GenerationThermalCoal_UR"
        ur_thermal_mining_col = find_column(ur_df, ["thermal","coal","mining"]) or "ThermalCoalMining_UR"
        ur_met_mining_col = find_column(ur_df, ["metallurgical","coal","mining"]) or "MetallurgicalCoalMining_UR"

        # Rename with "UR_" prefix
        rename_map_ur = {}
        rename_map_ur[ur_company_col] = "UR_Company"
        rename_map_ur[ur_isin_col]    = "UR_ISIN"
        rename_map_ur[ur_lei_col]     = "UR_LEI"
        rename_map_ur[ur_gen_coal_col] = "UR_GenThermal"
        rename_map_ur[ur_thermal_mining_col] = "UR_ThermalMining"
        rename_map_ur[ur_met_mining_col] = "UR_MetMining"

        ur_df.rename(columns=rename_map_ur, inplace=True, errors="ignore")

        # 3) Merge them on (ISIN, LEI, Company) => full outer
        merged_df = merge_sp_ur(sp_df, ur_df)

        # 4) Coalesce the 3 numeric columns into final columns
        #    If SP has a non-null value, use that, else use UR's value
        coalesce_cols(merged_df, "SP_GenThermal", "UR_GenThermal", "Generation (Thermal Coal)")
        coalesce_cols(merged_df, "SP_ThermalMining", "UR_ThermalMining", "Thermal Coal Mining")
        coalesce_cols(merged_df, "SP_MetMining", "UR_MetMining", "Metallurgical Coal Mining")

        # 5) Make a "Sector" column by coalescing "SP_Coal Industry Sector" 
        #    or "UR_Coal Industry Sector" or something similar if you want:
        #    For demonstration, let's just do partial matches for a sector column:
        #    If you want to prefer SP's sector over UR's, do the same coalesce approach.
        #    We'll guess there's a column named "SP_Coal Industry Sector", "UR_Coal Industry Sector"
        #    or fallback to something else. Adjust as needed.
        sp_sector_col = find_column(merged_df, ["sp_coal industry sector"]) or "SP_Sector"
        ur_sector_col = find_column(merged_df, ["ur_coal industry sector"]) or "UR_Sector"
        if sp_sector_col not in merged_df.columns:
            merged_df[sp_sector_col] = np.nan
        if ur_sector_col not in merged_df.columns:
            merged_df[ur_sector_col] = np.nan
        merged_df["Sector"] = merged_df[sp_sector_col].fillna( merged_df[ur_sector_col] )

        # 6) Similarly, you can coalesce a "Coal Share of Revenue" column, etc.
        #    For brevity, we'll just keep them separate (SP_ vs. UR_). 
        #    In the final filter, pick whichever you want. 
        #    For example, let's define:
        #        rev_col = "SP_Coal Share of Revenue" if that column exists, else "UR_Coal Share of Revenue", etc.
        #    Or do a coalesce if you want. 
        #    For demonstration, let's just guess:
        rev_col_sp = find_column(merged_df, ["sp_coal share of revenue"]) 
        rev_col_ur = find_column(merged_df, ["ur_coal share of revenue"]) 
        if not rev_col_sp and not rev_col_ur:
            # If neither found, create dummy:
            merged_df["CoalShareRevenue_final"] = 0.0
        else:
            # If both exist, coalesce
            sp_col = rev_col_sp or "tmp_sp_rev"
            ur_col = rev_col_ur or "tmp_ur_rev"
            if sp_col not in merged_df.columns:
                merged_df[sp_col] = np.nan
            if ur_col not in merged_df.columns:
                merged_df[ur_col] = np.nan
            merged_df["CoalShareRevenue_final"] = merged_df[sp_col].fillna( merged_df[ur_col] )

        # 7) Now apply your filter
        column_mapping = {
            "sector_col":    "Sector",  # We coalesced to "Sector"
            # We'll just use "CoalShareRevenue_final" as your coal_rev_col
            "coal_rev_col":  "CoalShareRevenue_final",
            # If you want to do the same for "Coal Share of Power Production", do so here
            "coal_power_col": "SP_Coal Share of Power Production",  # or a coalesce to UR...
            "capacity_col":   "SP_Installed Coal Power Capacity (MW)",
            "production_col": "SP_>10MT / >5GW",
            "expansion_col":  "SP_Expansion",
            "company_col":    "SP_Company",  # or coalesce if you prefer
            "ticker_col":     "SP_BB Ticker",
            "isin_col":       "SP_ISIN",
            "lei_col":        "SP_LEI"
        }

        filtered_df = filter_companies(
            df=merged_df,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            power_rev_threshold=power_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            services_rev_threshold=services_rev_threshold,
            # Toggles
            exclude_mining=exclude_mining,
            exclude_power=exclude_power,
            exclude_services=exclude_services,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            exclude_power_rev=exclude_power_rev,
            exclude_power_prod_percent=exclude_power_prod_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global,
            column_mapping=column_mapping,
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

        # 8) Build final sheets
        #    Show both SP_ and UR_ columns plus the final coalesced columns for Generation, etc.
        #    Adjust as you like.
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
