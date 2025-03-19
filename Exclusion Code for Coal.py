import streamlit as st
import pandas as pd
import io
import numpy as np  # We use numpy to fill columns with NaN, etc.

def load_data(file, sheet_name):
    """Load the Excel file and extract the specified sheet."""
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        return df
    except Exception as e:
        st.error(f"Error loading sheet '{sheet_name}': {e}")
        return None

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
        # must contain all must_keywords
        if all(mk.lower() in col_lower for mk in must_keywords):
            # skip if any exclude keywords found
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue
            return col
    return None

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
    exclude_generation_thermal_coal,    # NEW
    exclude_thermal_coal_mining,        # NEW
    exclude_metallurgical_coal_mining   # NEW
):
    """
    Apply exclusion criteria to filter companies based on thresholds.

    1) We have toggles for excluding Mining, Power, Services, plus
       toggles for excluding each SPGlobal sector individually.
    2) We already constructed columns "Generation (Thermal Coal) Share",
       "Thermal Coal Mining Share", and "Metallurgical Coal Mining Share"
       outside this function for reporting. We still rely on 'sector_col'
       for the actual check.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip()

        is_mining = ("mining" in sector.lower())
        is_power  = ("power"  in sector.lower())
        is_services = ("services" in sector.lower())

        # Example numeric fields (as in original script)
        coal_share = pd.to_numeric(row.get(column_mapping["coal_rev_col"], 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping["coal_power_col"], 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping["capacity_col"], 0), errors="coerce") or 0.0

        production_val = str(row.get(column_mapping["production_col"], "")).lower()
        expansion_text = str(row.get(column_mapping["expansion_col"], "")).lower().strip()

        # =========== MINING ===========
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in production_val:
                reasons.append(f"Mining production suggests >10MT vs threshold {mining_prod_mt_threshold}MT")

        # =========== POWER ===========
        if is_power and exclude_power:
            if exclude_power_rev and (coal_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            if exclude_capacity_mw and (installed_capacity > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold_mw}MW")

        # =========== SERVICES ===========
        if is_services and exclude_services:
            if exclude_services_rev and (coal_share * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_share*100:.2f}% > {services_rev_threshold}% (Services)")

        # =========== GLOBAL EXPANSION ================
        if expansions_global:
            for choice in expansions_global:
                if choice.lower() in expansion_text:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break  # exclude once we find any match

        # =========== SPGlobal SECTOR EXCLUSIONS (NEW) ===========
        # user toggles each sector's exclusion
        if sector == "Generation (Thermal Coal)" and exclude_generation_thermal_coal:
            if coal_share > gen_thermal_coal_threshold:
                reasons.append(f"{sector} {coal_share:.2f} > {gen_thermal_coal_threshold}")
        elif sector == "Thermal Coal Mining" and exclude_thermal_coal_mining:
            if coal_share > thermal_coal_mining_threshold:
                reasons.append(f"{sector} {coal_share:.2f} > {thermal_coal_mining_threshold}")
        elif sector == "Metallurgical Coal Mining" and exclude_metallurgical_coal_mining:
            if coal_share > metallurgical_coal_mining_threshold:
                reasons.append(f"{sector} {coal_share:.2f} > {metallurgical_coal_mining_threshold}")

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # ============ FILE & SHEET =============
    st.sidebar.header("File & Sheet Settings")

    # Two file uploads for SPGlobal & Urgewald GCEL
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="SPGlobalSheet")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

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

    # ============ GLOBAL EXPANSION EXCLUSION =============
    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansion text contains any of these",
        expansions_possible,
        default=[]
    )

    # ============ SPGlobal Coal Sectors =============
    st.sidebar.header("SPGlobal Coal Sectors")
    # user can define threshold AND turn off the exclusion
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

        # Load each file into a DataFrame
        spglobal_df = load_data(spglobal_file, spglobal_sheet)
        urgewald_df = load_data(urgewald_file, urgewald_sheet)

        if spglobal_df is None or urgewald_df is None:
            return

        # Combine them
        combined_df = pd.concat([spglobal_df, urgewald_df], ignore_index=True)
        # Remove exact duplicates of all columns
        combined_df.drop_duplicates(inplace=True)

        # Identify the key columns automatically
        def find_co(*kw): 
            return find_column(combined_df, list(kw), exclude_keywords=["parent"])
        company_col = find_co("company") or "Company"

        expansion_col = find_column(combined_df, ["expansion"]) or "expansion"

        sector_col = find_column(combined_df, ["industry","sector"]) or "Coal Industry Sector"
        coal_rev_col = find_column(combined_df, ["coal","share","revenue"]) or "Coal Share of Revenue"
        coal_power_col = find_column(combined_df, ["coal","share","power"]) or "Coal Share of Power Production"
        capacity_col   = find_column(combined_df, ["installed","coal","power","capacity"]) or "Installed Coal Power Capacity (MW)"
        production_col = find_column(combined_df, ["10mt","5gw"]) or ">10MT / >5GW"
        ticker_col     = find_column(combined_df, ["bb","ticker"]) or "BB Ticker"
        isin_col       = find_column(combined_df, ["isin","equity"]) or "ISIN equity"
        lei_col        = find_column(combined_df, ["lei"]) or "LEI"

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

        # -- Create extra columns for final report: "Generation (Thermal Coal) Share", etc.
        # We'll fill them with the same numeric data that 'coal_rev_col' contains,
        # but only if that row's sector matches.
        combined_df["Generation (Thermal Coal) Share"] = np.nan
        combined_df["Thermal Coal Mining Share"] = np.nan
        combined_df["Metallurgical Coal Mining Share"] = np.nan

        # Fill them conditionally
        for i, row in combined_df.iterrows():
            s = str(row.get(sector_col, "")).strip()
            val = pd.to_numeric(row.get(coal_rev_col, 0), errors="coerce") or 0.0
            if s == "Generation (Thermal Coal)":
                combined_df.at[i, "Generation (Thermal Coal) Share"] = val
            elif s == "Thermal Coal Mining":
                combined_df.at[i, "Thermal Coal Mining Share"] = val
            elif s == "Metallurgical Coal Mining":
                combined_df.at[i, "Metallurgical Coal Mining Share"] = val

        # Now apply the filter
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
            # GLOBAL expansions
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

        # =========== OUTPUT SHEETS ===========
        excluded_cols = [
            company_col,
            production_col,
            capacity_col,
            coal_rev_col,
            sector_col,
            ticker_col,
            isin_col,
            lei_col,
            "Generation (Thermal Coal) Share",   # show newly added columns
            "Thermal Coal Mining Share",
            "Metallurgical Coal Mining Share",
            "Exclusion Reasons"
        ]
        excluded_df = filtered_df[filtered_df["Excluded"] == True][excluded_cols]

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
        retained_df = filtered_df[filtered_df["Excluded"] == False][retained_cols]

        no_data_df = filtered_df[filtered_df[sector_col].isna()][retained_cols]

        import openpyxl
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # =========== STATS =============
        st.subheader("Statistics")
        st.write(f"Total companies (after merge & dedup): {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"Companies with no data: {len(no_data_df)}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded_df)

        st.download_button(
            "Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
