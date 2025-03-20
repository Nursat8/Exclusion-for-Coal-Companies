import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

##############################
# DATA-LOADING FUNCTIONS
##############################

def load_spglobal_data(file, sheet_name):
    """
    Load the SPGlobal file, skipping the first 5 rows so row #6 is treated as headers.
    In pandas, header=5 => row #6 in Excel is the header line.
    Flatten columns if there's a multi-level header.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=5)

        # Flatten multi-level columns if needed
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                ' '.join(str(x).strip() for x in col if x not in (None, ''))
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]

        return df
    except Exception as e:
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    """
    Load the Urgewald GCEL file normally (assuming row #1 is header).
    Flatten columns if multi-level.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                ' '.join(str(x).strip() for x in col if x not in (None, ''))
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]

        return df
    except Exception as e:
        st.error(f"Error loading Urgewald sheet '{sheet_name}': {e}")
        return None

##############################
# HELPER: find_column
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
        col_lower = col.lower().strip()
        # must_keywords must ALL appear
        if all(mk.lower() in col_lower for mk in must_keywords):
            # exclude_keywords must NOT appear
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue
            return col
    return None

##############################
# HELPERS FOR COAL SHARES
##############################

def fill_coal_shares(df, sector_col, coal_rev_col):
    """
    Populate 'Generation (Thermal Coal) Share', 'Thermal Coal Mining Share',
    and 'Metallurgical Coal Mining Share' based on the sector text plus
    a numeric column (e.g. 'Coal Share of Revenue' or something similar).

    We make the matching conditions a bit more flexible so that
    'generation' or 'power' with 'thermal' triggers Generation (Thermal Coal).
    """
    df["Generation (Thermal Coal) Share"] = np.nan
    df["Thermal Coal Mining Share"] = np.nan
    df["Metallurgical Coal Mining Share"] = np.nan

    for i, row in df.iterrows():
        sector_text = str(row.get(sector_col, "")).strip().lower()
        # The numeric "coal share" to fill in
        # (Change this if you want to fill from another column.)
        numeric_val = pd.to_numeric(row.get(coal_rev_col, 0), errors="coerce") or 0.0

        # More flexible matches for "Generation (Thermal Coal)"
        # e.g. sector might say "thermal generation", "coal-fired power generation", etc.
        # We check for any form of "gen"/"power" + "thermal" + "coal"
        # Adjust as needed for your dataset.
        if ("thermal" in sector_text or "coal-fired" in sector_text) and ("gen" in sector_text or "power" in sector_text):
            df.at[i, "Generation (Thermal Coal) Share"] = numeric_val

        # Thermal Coal Mining
        # We look for "thermal" + "coal" + "mining"
        if "thermal" in sector_text and "coal" in sector_text and "mining" in sector_text:
            df.at[i, "Thermal Coal Mining Share"] = numeric_val

        # Metallurgical Coal Mining
        # We look for "metallurgical" + "coal" + "mining"
        if "metallurgical" in sector_text and "coal" in sector_text and "mining" in sector_text:
            df.at[i, "Metallurgical Coal Mining Share"] = numeric_val

    return df

##############################
# CORE: filter_companies
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
    # SPGlobal coal sector thresholds + toggles
    gen_thermal_coal_threshold,
    thermal_coal_mining_threshold,
    metallurgical_coal_mining_threshold,
    exclude_generation_thermal_coal,
    exclude_thermal_coal_mining,
    exclude_metallurgical_coal_mining
):
    """
    Apply exclusion criteria to filter companies based on thresholds.

    1) Toggles for excluding Mining, Power, Services.
    2) Single expansions_global for expansions.
    3) PARTIAL MATCH for SPGlobal sectors using keywords:
       - "generation" + "thermal"
       - "thermal" + "coal" + "mining"
       - "metallurgical" + "coal" + "mining"
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
        is_power     = ("power"  in s_lc) or ("generation" in s_lc)
        is_services  = ("services" in s_lc)

        # Numeric fields
        coal_share         = pd.to_numeric(row.get(rev_col, 0), errors="coerce") or 0.0
        coal_power_share   = pd.to_numeric(row.get(power_col, 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(cap_col, 0), errors="coerce") or 0.0

        prod_val = str(row.get(prod_col, "")).lower()
        exp_text = str(row.get(exp_col, "")).lower().strip()

        # ===== MINING =====
        if is_mining and exclude_mining:
            # Check "production >10mt"
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # ===== POWER =====
        if is_power and exclude_power:
            # If we are excluding for coal share % of revenue
            if exclude_power_rev and (coal_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {power_rev_threshold}% (Power)")

            # If we are excluding for coal share % of power production
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            # If we are excluding for installed capacity
            if exclude_capacity_mw and (installed_capacity > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold_mw}MW")

        # ===== SERVICES =====
        if is_services and exclude_services:
            if exclude_services_rev and (coal_share * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {services_rev_threshold}% (Services)")

        # ===== GLOBAL EXPANSIONS =====
        if expansions_global:
            for choice in expansions_global:
                if choice.lower() in exp_text:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break

        # ===== SPGlobal COAL SECTORS (PARTIAL MATCH) =====
        # Generation (Thermal Coal)
        if exclude_generation_thermal_coal:
            # Searching for both 'generation' & 'thermal' 
            # (or at least the synonyms we consider for "power" or "thermal")
            if ("generation" in s_lc or "power" in s_lc) and "thermal" in s_lc:
                if coal_share > gen_thermal_coal_threshold:
                    reasons.append(f"Generation (Thermal Coal) {coal_share:.2f} > {gen_thermal_coal_threshold}")

        # Thermal Coal Mining
        if exclude_thermal_coal_mining:
            if "thermal" in s_lc and "coal" in s_lc and "mining" in s_lc:
                if coal_share > thermal_coal_mining_threshold:
                    reasons.append(f"Thermal Coal Mining {coal_share:.2f} > {thermal_coal_mining_threshold}")

        # Metallurgical Coal Mining
        if exclude_metallurgical_coal_mining:
            if "metallurgical" in s_lc and "coal" in s_lc and "mining" in s_lc:
                if coal_share > metallurgical_coal_mining_threshold:
                    reasons.append(f"Metallurgical Coal Mining {coal_share:.2f} > {metallurgical_coal_mining_threshold}")

        # If any reasons found => excluded
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

##############################
# DEDUPLICATION HELPER
##############################

def deduplicate_any(df, name_col, isin_col, lei_col):
    """
    Remove duplicates if ANY of name, ISIN, or LEI match. 
    We implement "OR" logic by removing duplicates in steps:
       1) drop duplicates on 'name_col'
       2) drop duplicates on 'isin_col'
       3) drop duplicates on 'lei_col'
    so that if any two rows share the same name or the same ISIN 
    or the same LEI, only the first row is kept.

    This is a simple approach to approximate "OR" logic deduplication.
    """
    # If the column doesn't exist or is fully NaN, skip that step
    # to avoid dropping everything as duplicates of NaN.
    for col in [name_col, isin_col, lei_col]:
        if col in df.columns:
            if not df[col].isnull().all():
                df.drop_duplicates(subset=[col], keep="first", inplace=True)
    return df

##############################
# MAIN STREAMLIT APP
##############################

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # ============ FILE & SHEET =============
    st.sidebar.header("File & Sheet Settings")

    # 1) SPGlobal file (row #6 => header=5)
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    # 2) Urgewald GCEL file (row #1 => normal read)
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

    # ============ SPGlobal Coal Sectors (Partial Matches) ============
    st.sidebar.header("SPGlobal Coal Sectors")
    gen_thermal_coal_threshold = st.sidebar.number_input("Generation (Thermal Coal) Threshold", value=15.0)
    exclude_generation_thermal_coal = st.sidebar.checkbox("Exclude Generation (Thermal Coal)?", value=True)

    thermal_coal_mining_threshold = st.sidebar.number_input("Thermal Coal Mining Threshold", value=20.0)
    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude Thermal Coal Mining?", value=True)

    metallurgical_coal_mining_threshold = st.sidebar.number_input("Metallurgical Coal Mining Threshold", value=25.0)
    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude Metallurgical Coal Mining?", value=True)

    # ============ RUN FILTER =============
    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please upload both SPGlobal and Urgewald GCEL files.")
            return

        # Load each file
        spglobal_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        urgewald_df = load_urgewald_data(urgewald_file, urgewald_sheet)

        if spglobal_df is None or urgewald_df is None:
            return

        ########################################################
        # 1) Identify columns in each file for: company, ISIN, LEI, sector, etc.
        ########################################################
        def find_co(df, *keywords, exclude=None):
            return find_column(df, list(keywords), exclude_keywords=exclude or [])

        # -------- SPGlobal --------
        sp_company_col = find_co(spglobal_df, "company")
        if not sp_company_col:
            sp_company_col = "Company_SP"  # fallback
            spglobal_df[sp_company_col] = np.nan

        sp_isin_col = find_co(spglobal_df, "isin")
        if not sp_isin_col:
            sp_isin_col = "ISIN_SP"
            spglobal_df[sp_isin_col] = np.nan

        sp_lei_col = find_co(spglobal_df, "lei")
        if not sp_lei_col:
            sp_lei_col = "LEI_SP"
            spglobal_df[sp_lei_col] = np.nan

        sp_sector_col = find_co(spglobal_df, "industry", "sector")
        if not sp_sector_col:
            sp_sector_col = "Sector_SP"
            spglobal_df[sp_sector_col] = np.nan

        sp_coal_rev_col = find_co(spglobal_df, "coal", "share", "revenue")
        if not sp_coal_rev_col:
            sp_coal_rev_col = "CoalShareRev_SP"
            spglobal_df[sp_coal_rev_col] = 0.0

        sp_coal_power_col = find_co(spglobal_df, "coal", "share", "power")
        if not sp_coal_power_col:
            sp_coal_power_col = "CoalSharePower_SP"
            spglobal_df[sp_coal_power_col] = 0.0

        sp_capacity_col = find_co(spglobal_df, "installed", "coal", "power", "capacity")
        if not sp_capacity_col:
            sp_capacity_col = "InstalledCoalCap_SP"
            spglobal_df[sp_capacity_col] = 0.0

        sp_prod_col = find_co(spglobal_df, "10mt", "5gw")
        if not sp_prod_col:
            sp_prod_col = "Production_SP"
            spglobal_df[sp_prod_col] = ""

        sp_expansion_col = find_co(spglobal_df, "expansion")
        if not sp_expansion_col:
            sp_expansion_col = "Expansion_SP"
            spglobal_df[sp_expansion_col] = ""

        # Rename to a standardized set of columns so we can combine easily:
        spglobal_df.rename(columns={
            sp_company_col:    "Company",
            sp_isin_col:       "ISIN",
            sp_lei_col:        "LEI",
            sp_sector_col:     "Sector",
            sp_coal_rev_col:   "CoalShareRevenue",
            sp_coal_power_col: "CoalSharePower",
            sp_capacity_col:   "InstalledCoalCapacity",
            sp_prod_col:       "Production",
            sp_expansion_col:  "Expansion"
        }, inplace=True)

        # -------- Urgewald --------
        urw_company_col = find_co(urgewald_df, "company")
        if not urw_company_col:
            urw_company_col = "Company_URW"
            urgewald_df[urw_company_col] = np.nan

        urw_isin_col = find_co(urgewald_df, "isin", "equity")
        if not urw_isin_col:
            urw_isin_col = "ISIN_URW"
            urgewald_df[urw_isin_col] = np.nan

        urw_lei_col = find_co(urgewald_df, "lei")
        if not urw_lei_col:
            urw_lei_col = "LEI_URW"
            urgewald_df[urw_lei_col] = np.nan

        urw_sector_col = find_co(urgewald_df, "industry", "sector")
        if not urw_sector_col:
            urw_sector_col = "Sector_URW"
            urgewald_df[urw_sector_col] = np.nan

        urw_coal_rev_col = find_co(urgewald_df, "coal", "share", "revenue")
        if not urw_coal_rev_col:
            urw_coal_rev_col = "CoalShareRev_URW"
            urgewald_df[urw_coal_rev_col] = 0.0

        urw_coal_power_col = find_co(urgewald_df, "coal", "share", "power")
        if not urw_coal_power_col:
            urw_coal_power_col = "CoalSharePower_URW"
            urgewald_df[urw_coal_power_col] = 0.0

        urw_capacity_col = find_co(urgewald_df, "installed", "coal", "power", "capacity")
        if not urw_capacity_col:
            urw_capacity_col = "InstalledCoalCap_URW"
            urgewald_df[urw_capacity_col] = 0.0

        urw_prod_col = find_co(urgewald_df, "10mt", "5gw")
        if not urw_prod_col:
            urw_prod_col = "Production_URW"
            urgewald_df[urw_prod_col] = ""

        urw_expansion_col = find_co(urgewald_df, "expansion")
        if not urw_expansion_col:
            urw_expansion_col = "Expansion_URW"
            urgewald_df[urw_expansion_col] = ""

        urgewald_df.rename(columns={
            urw_company_col:    "Company",
            urw_isin_col:       "ISIN",
            urw_lei_col:        "LEI",
            urw_sector_col:     "Sector",
            urw_coal_rev_col:   "CoalShareRevenue",
            urw_coal_power_col: "CoalSharePower",
            urw_capacity_col:   "InstalledCoalCapacity",
            urw_prod_col:       "Production",
            urw_expansion_col:  "Expansion"
        }, inplace=True)

        ########################################################
        # 2) Combine both dataframes
        ########################################################
        combined_df = pd.concat([spglobal_df, urgewald_df], ignore_index=True)

        ########################################################
        # 3) Deduplicate using the "OR" logic on (Company, ISIN, LEI).
        #    If any match => duplicate => drop.
        ########################################################

        # Just call our deduplicate function:
        combined_df = deduplicate_any(combined_df,
                                      name_col="Company",
                                      isin_col="ISIN",
                                      lei_col="LEI")

        ########################################################
        # 4) Fill columns for "Generation (Thermal Coal) Share" etc.
        #    based on flexible matches in the 'Sector' column
        ########################################################
        # We'll fill from combined_df["CoalShareRevenue"] if we see relevant keywords in 'Sector'.
        combined_df = fill_coal_shares(
            df=combined_df,
            sector_col="Sector",
            coal_rev_col="CoalShareRevenue"  # you can change if you want a different numeric col
        )

        ########################################################
        # 5) Filter out companies based on user threshold settings
        ########################################################
        column_mapping = {
            "sector_col":    "Sector",
            "company_col":   "Company",
            "coal_rev_col":  "CoalShareRevenue",
            "coal_power_col": "CoalSharePower",
            "capacity_col":  "InstalledCoalCapacity",
            "production_col": "Production",
            "ticker_col":    "",  # Not strictly found in the above logic, but keep placeholders
            "isin_col":      "ISIN",
            "lei_col":       "LEI",
            "expansion_col": "Expansion"
        }

        filtered_df = filter_companies(
            combined_df,
            # Mining threshold
            mining_prod_mt_threshold,
            # Power thresholds
            power_rev_threshold,
            power_prod_threshold_percent,
            capacity_threshold_mw,
            # Services threshold
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
            expansions_global,
            column_mapping,
            # SPGlobal thresholds + toggles
            gen_thermal_coal_threshold,
            thermal_coal_mining_threshold,
            metallurgical_coal_mining_threshold,
            exclude_generation_thermal_coal,
            exclude_thermal_coal_mining,
            exclude_metallurgical_coal_mining
        )

        ########################################################
        # 6) Build final output sheets
        ########################################################
        excluded_cols = [
            "Company",
            "Production",
            "InstalledCoalCapacity",
            "CoalShareRevenue",
            "Sector",
            "ISIN",
            "LEI",
            "Generation (Thermal Coal) Share",
            "Thermal Coal Mining Share",
            "Metallurgical Coal Mining Share",
            "Exclusion Reasons"
        ]
        retained_cols = [
            "Company",
            "Production",
            "InstalledCoalCapacity",
            "CoalShareRevenue",
            "Sector",
            "ISIN",
            "LEI",
            "Generation (Thermal Coal) Share",
            "Thermal Coal Mining Share",
            "Metallurgical Coal Mining Share"
        ]

        excluded_df = filtered_df[ filtered_df["Excluded"] == True ].copy()
        retained_df = filtered_df[ filtered_df["Excluded"] == False ].copy()
        no_data_df  = filtered_df[ filtered_df["Sector"].isna() ].copy()

        excluded_df = excluded_df[excluded_cols]
        retained_df = retained_df[retained_cols]
        no_data_df  = no_data_df[retained_cols]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # Statistics
        st.subheader("Statistics")
        st.write(f"Total companies (after dedup): {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"Companies with no sector data: {len(no_data_df)}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded_df)

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
