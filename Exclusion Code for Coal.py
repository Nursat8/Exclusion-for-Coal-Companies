import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

##############################
# UTILITY: MAKE COLUMNS UNIQUE
##############################
def make_columns_unique(df):
    """
    Ensures DataFrame df has uniquely named columns by appending a suffix 
    (e.g., '_1', '_2') if duplicates exist. This prevents 'InvalidIndexError'
    when concatenating.
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
# DATA-LOADING FUNCTIONS
##############################
def load_spglobal_data(file, sheet_name):
    """
    Load the SPGlobal file, skipping the first 5 rows so row #6 is treated as headers.
    Flatten columns if there's a multi-level header and make columns unique.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=5)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                ' '.join(str(x).strip() for x in col if x not in (None, ''))
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
    Flatten columns if multi-level and make columns unique.
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

        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading Urgewald sheet '{sheet_name}': {e}")
        return None

##############################
# HELPER: find_column
##############################
def find_column(df, must_keywords, exclude_keywords=None):
    """
    Finds a column in df that contains ALL 'must_keywords' (case-insensitive)
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
    Apply exclusion criteria to filter companies based on thresholds and toggles:

    1) Toggles for excluding Mining, Power, Services.
    2) expansions_global for expansions text search.
    3) Partial-match logic for 'Generation (Thermal Coal)', 'Thermal Coal Mining', 
       'Metallurgical Coal Mining' based on sector text.
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
        # We'll also treat "generation" as power:
        is_power     = ("power" in s_lc) or ("generation" in s_lc)
        is_services  = ("services" in s_lc)

        # Numeric fields
        coal_share         = pd.to_numeric(row.get(rev_col, 0), errors="coerce") or 0.0
        coal_power_share   = pd.to_numeric(row.get(power_col, 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(cap_col, 0), errors="coerce") or 0.0

        # production col is often text like ">10mt", so let's just check text
        prod_val = str(row.get(prod_col, "")).lower()
        # expansions
        exp_text = str(row.get(exp_col, "")).lower().strip()

        # ===== MINING =====
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # ===== POWER =====
        if is_power and exclude_power:
            if exclude_power_rev and (coal_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

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
            # searching for "thermal" + ("power" or "generation")
            if ("thermal" in s_lc) and (("power" in s_lc) or ("generation" in s_lc)):
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

        # 1) Load each file
        spglobal_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        urgewald_df = load_urgewald_data(urgewald_file, urgewald_sheet)

        if spglobal_df is None or urgewald_df is None:
            return

        # 2) Simply concatenate - NO dedup, so we keep ALL rows
        combined_df = pd.concat([spglobal_df, urgewald_df], ignore_index=True)

        # 3) Identify relevant columns
        def find_co(*kw):
            return find_column(combined_df, list(kw), exclude_keywords=["parent"])

        company_col    = find_co("company") or "Company"
        ticker_col     = find_co("bb","ticker") or "BB Ticker"
        isin_col       = find_co("isin","equity") or "ISIN equity"
        lei_col        = find_co("lei") or "LEI"
        sector_col     = find_co("industry","sector") or "Sector"
        coal_rev_col   = find_co("coal","share","revenue") or "Coal Share of Revenue"
        coal_power_col = find_co("coal","share","power") or "Coal Share of Power Production"
        capacity_col   = find_co("installed","coal","power","capacity") or "Installed Coal Power Capacity (MW)"
        production_col = find_co("10mt","5gw") or ">10MT / >5GW"
        expansion_col  = find_co("expansion") or "Expansion"

        # 4) Create columns for final output
        combined_df["Generation (Thermal Coal) Share"] = np.nan
        combined_df["Thermal Coal Mining Share"]       = np.nan
        combined_df["Metallurgical Coal Mining Share"] = np.nan

        # 5) Fill them if partial match of 'Sector'
        for i, row in combined_df.iterrows():
            sec_val = str(row.get(sector_col, "")).strip().lower()
            # Use the 'coal_rev_col' as the numeric value to store in these share columns
            # (If you actually want to read from a different numeric column, adjust here)
            val = pd.to_numeric(row.get(coal_rev_col, 0), errors="coerce") or 0.0

            # Generation (Thermal Coal)
            if ("thermal" in sec_val) and (("generation" in sec_val) or ("power" in sec_val)):
                combined_df.at[i, "Generation (Thermal Coal) Share"] = val

            # Thermal Coal Mining
            if "thermal" in sec_val and "coal" in sec_val and "mining" in sec_val:
                combined_df.at[i, "Thermal Coal Mining Share"] = val

            # Metallurgical Coal Mining
            if "metallurgical" in sec_val and "coal" in sec_val and "mining" in sec_val:
                combined_df.at[i, "Metallurgical Coal Mining Share"] = val

        # 6) Filter according to your thresholds & toggles
        column_mapping = {
            "sector_col":    sector_col,
            "company_col":   company_col,
            "coal_rev_col":  coal_rev_col,
            "coal_power_col": coal_power_col,
            "capacity_col":  capacity_col,
            "production_col": production_col,
            "ticker_col":    ticker_col,
            "isin_col":      isin_col,
            "lei_col":       lei_col,
            "expansion_col": expansion_col
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

        # 7) Build final sheets
        excluded_cols = [
            company_col,
            production_col,
            capacity_col,
            coal_rev_col,
            sector_col,
            ticker_col,
            isin_col,
            lei_col,
            "Generation (Thermal Coal) Share",
            "Thermal Coal Mining Share",
            "Metallurgical Coal Mining Share",
            "Exclusion Reasons"
        ]
        retained_cols = [
            company_col,
            production_col,
            capacity_col,
            coal_rev_col,
            sector_col,
            ticker_col,
            isin_col,
            lei_col,
            "Generation (Thermal Coal) Share",
            "Thermal Coal Mining Share",
            "Metallurgical Coal Mining Share"
        ]

        excluded_df = filtered_df[filtered_df["Excluded"] == True][excluded_cols]
        retained_df = filtered_df[filtered_df["Excluded"] == False][retained_cols]

        # "No Data" means missing sector info
        no_data_df = filtered_df[filtered_df[sector_col].isna()][retained_cols]

        # 8) Write final results to an in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # 9) Show stats & data
        st.subheader("Statistics")
        st.write(f"Total companies (all rows combined): {len(combined_df)}")
        st.write(f"Total companies after filter: {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"Companies with no data in '{sector_col}': {len(no_data_df)}")

        # Show a preview of final
        st.subheader("Excluded Companies")
        st.dataframe(excluded_df.head(50))

        # 10) Download button
        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
