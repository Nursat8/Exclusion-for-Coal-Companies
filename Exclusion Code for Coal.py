import streamlit as st
import pandas as pd
import io

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
    # -- REMOVED "Mining: Max coal revenue (%)" => no more `mining_rev_threshold`
    mining_prod_mt_threshold,               # keep only for demonstration
    # -- REMOVED "Mining: Max production threshold (GW)"
    # Power thresholds
    power_rev_threshold,                    # keep to show we left old code alone
    power_prod_threshold_percent,
    # -- REMOVED "Power: Max production threshold (MT)"
    # -- REMOVED "Power: Max production threshold (GW)"
    capacity_threshold_mw,
    # Services thresholds
    services_rev_threshold,
    # -- REMOVED "Services: Max production threshold (MT)"
    # -- REMOVED "Services: Max production threshold (GW)"
    # Exclusion toggles
    exclude_mining,
    exclude_power,
    exclude_services,
    # -- REMOVED booleans: exclude_mining_rev, exclude_mining_prod_gw, exclude_power_prod_mt, exclude_power_prod_gw,
    #                      exclude_services_prod_mt, exclude_services_prod_gw
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    # -- REMOVED expansions_mining, expansions_power, expansions_services
    # ADDED single expansions param:
    expansions_global,  # MODIFIED
    column_mapping,
    # ADDED three new thresholds for SPGlobal sectors:
    gen_thermal_coal_threshold,        # MODIFIED
    thermal_coal_mining_threshold,     # MODIFIED
    metallurgical_coal_mining_threshold  # MODIFIED
):
    """
    Apply exclusion criteria to filter companies based on thresholds.
    REMOVED expansions logic from each sector, now there's one expansions_global.
    ADDED logic for dynamic SPGlobal sectors: Generation (Thermal Coal), Thermal Coal Mining, Metallurgical Coal Mining.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip()

        # Identify membership from your existing code:
        is_mining = ("mining" in sector.lower())
        is_power  = ("power"  in sector.lower())
        is_services = ("services" in sector.lower())

        # Example: user wants to define threshold for each row's "coal share" (or some numeric measure)
        # We'll call it 'coal_share_col' just as an example:
        coal_share = pd.to_numeric(row.get(column_mapping["coal_rev_col"], 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping["coal_power_col"], 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping["capacity_col"], 0), errors="coerce") or 0.0

        # Production text
        production_val = str(row.get(column_mapping["production_col"], "")).lower()

        # expansions text
        expansion_text = str(row.get(column_mapping["expansion_col"], "")).lower().strip()

        # =========== MINING ===========
        if is_mining and exclude_mining:
            # We removed the mining_rev_threshold logic => no revenue check here
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

        # =========== GLOBAL EXPANSIONS (NEW) ===========
        # removed expansions_{sector} checks, replaced with one check:
        if expansions_global:  # user picked expansions to exclude
            for choice in expansions_global:
                if choice.lower() in expansion_text:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break

        # =========== DYNAMIC SPGlobal Sectors (NEW) ===========
        # If the 'sector' is exactly "Generation (Thermal Coal)"
        # or "Thermal Coal Mining" or "Metallurgical Coal Mining",
        # compare to user threshold. We'll re-use `coal_share` for demonstration.
        if sector == "Generation (Thermal Coal)":
            if coal_share > gen_thermal_coal_threshold:
                reasons.append(f"{sector}: {coal_share:.2f} > {gen_thermal_coal_threshold} threshold")

        elif sector == "Thermal Coal Mining":
            if coal_share > thermal_coal_mining_threshold:
                reasons.append(f"{sector}: {coal_share:.2f} > {thermal_coal_mining_threshold} threshold")

        elif sector == "Metallurgical Coal Mining":
            if coal_share > metallurgical_coal_mining_threshold:
                reasons.append(f"{sector}: {coal_share:.2f} > {metallurgical_coal_mining_threshold} threshold")

        # If reasons => excluded
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
    sheet_name = st.sidebar.text_input("Sheet name to check", value="GCEL 2024")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    # ============ MINING THRESHOLDS ============
    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    # -- REMOVED "Mining: Max coal revenue (%)"
    # -- REMOVED "Exclude if mining rev threshold exceeded"
    # Keep only "Mining: Max production threshold (MT)" for demonstration:
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT for Mining", value=True)
    # -- REMOVED "Mining: Max production threshold (GW)" & its checkbox

    # ============ POWER THRESHOLDS ============
    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded", value=True)
    # -- REMOVED "Power: Max production threshold (MT)" & its checkbox
    # -- REMOVED "Power: Max production threshold (GW)" & its checkbox
    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed coal power capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded", value=True)

    # ============ SERVICES THRESHOLDS ============
    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded", value=False)
    # -- REMOVED "Services: Max production threshold (MT)" & checkbox
    # -- REMOVED "Services: Max production threshold (GW)" & checkbox

    # ============ GLOBAL EXPANSION EXCLUSION (NEW) ============
    st.sidebar.header("Global Expansion Exclusion")  # MODIFIED
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansion text contains any of these",
        expansions_possible,
        default=[]
    )

    # ============ NEW: SPGlobal SECTORS & THRESHOLDS ============
    st.sidebar.header("SPGlobal Coal Sectors (New)")  # MODIFIED
    gen_thermal_coal_threshold = st.sidebar.number_input("Generation (Thermal Coal) Threshold", value=15.0)
    thermal_coal_mining_threshold = st.sidebar.number_input("Thermal Coal Mining Threshold", value=20.0)
    metallurgical_coal_mining_threshold = st.sidebar.number_input("Metallurgical Coal Mining Threshold", value=25.0)

    # ============ RUN FILTER =============
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file, sheet_name)
        if df is None:
            return

        # For 'company', skip "parent" to avoid 'Parent Company'
        def find_co(*kw): 
            return find_column(df, list(kw), exclude_keywords=["parent"])
        company_col = find_co("company") or "Company"

        # also find 'expansion' column
        expansion_col = find_column(df, ["expansion"]) or "expansion"

        column_mapping = {
            "sector_col":     find_column(df, ["industry","sector"]) or "Coal Industry Sector",
            "company_col":    company_col,
            "coal_rev_col":   find_column(df, ["coal","share","revenue"]) or "Coal Share of Revenue",
            "coal_power_col": find_column(df, ["coal","share","power"]) or "Coal Share of Power Production",
            "capacity_col":   find_column(df, ["installed","coal","power","capacity"]) or "Installed Coal Power Capacity (MW)",
            "production_col": find_column(df, ["10mt","5gw"]) or ">10MT / >5GW",
            "ticker_col":     find_column(df, ["bb","ticker"]) or "BB Ticker",
            "isin_col":       find_column(df, ["isin","equity"]) or "ISIN equity",
            "lei_col":        find_column(df, ["lei"]) or "LEI",
            "expansion_col":  expansion_col
        }

        filtered_df = filter_companies(
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
            # GLOBAL expansions (NEW)
            expansions_global,  # MODIFIED
            column_mapping,
            # The three new SPGlobal thresholds:
            gen_thermal_coal_threshold,         # MODIFIED
            thermal_coal_mining_threshold,      # MODIFIED
            metallurgical_coal_mining_threshold # MODIFIED
        )

        # =========== OUTPUT SHEETS ===========
        excluded_cols = [
            column_mapping["company_col"],
            column_mapping["production_col"],
            column_mapping["capacity_col"],
            column_mapping["coal_rev_col"],
            column_mapping["sector_col"],
            column_mapping["ticker_col"],
            column_mapping["isin_col"],
            column_mapping["lei_col"],
            "Exclusion Reasons"
        ]
        excluded_df = filtered_df[filtered_df["Excluded"] == True][excluded_cols]

        retained_cols = [
            column_mapping["company_col"],
            column_mapping["production_col"],
            column_mapping["capacity_col"],
            column_mapping["coal_rev_col"],
            column_mapping["sector_col"],
            column_mapping["ticker_col"],
            column_mapping["isin_col"],
            column_mapping["lei_col"]
        ]
        retained_df = filtered_df[filtered_df["Excluded"] == False][retained_cols]
        no_data_df = df[df[column_mapping["sector_col"]].isna()][retained_cols]

        import openpyxl
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # =========== STATS =============
        st.subheader("Statistics")
        st.write(f"Total companies: {len(df)}")
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
