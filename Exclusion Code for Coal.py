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
            # skip if any exclude keyword
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue
            return col
    return None

def filter_companies(
    df,
    # Mining
    mining_rev_threshold,
    mining_prod_mt_threshold,
    mining_prod_gw_threshold,
    # Power
    power_rev_threshold,
    power_prod_threshold_percent,
    power_prod_mt_threshold,
    power_prod_gw_threshold,
    capacity_threshold_mw,
    # Services
    services_rev_threshold,
    services_prod_mt_threshold,
    services_prod_gw_threshold,
    # Exclusion toggles
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_rev,
    exclude_mining_prod_mt,
    exclude_mining_prod_gw,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_power_prod_mt,
    exclude_power_prod_gw,
    exclude_capacity_mw,
    exclude_services_rev,
    exclude_services_prod_mt,
    exclude_services_prod_gw,
    column_mapping
):
    """
    Apply exclusion criteria to filter companies based on thresholds.
    Decimal values in data are fractions-of-1.0, so multiply by 100 for rev/production % checks.
    Uses the 'production_col' text (e.g., '>10MT / >5GW') to detect if it's >10MT or >5GW
    for demonstration of (MT)/(GW) logic.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip().lower()

        # Identify membership
        is_mining = ("mining" in sector)
        is_power  = ("power" in sector)
        is_services = ("services" in sector)

        # Extract numeric fields
        coal_rev = pd.to_numeric(row.get(column_mapping["coal_rev_col"], 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping["coal_power_col"], 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping["capacity_col"], 0), errors="coerce") or 0.0

        # Check text from production column
        production_val = str(row.get(column_mapping["production_col"], "")).lower()

        # =========== MINING ===========
        if is_mining and exclude_mining:
            # => Mining Revenue check
            if exclude_mining_rev and (coal_rev * 100) > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {mining_rev_threshold}% (Mining)")

            # => Mining Production (MT)
            if exclude_mining_prod_mt:
                if ">10mt" in production_val:
                    reasons.append(f"Mining production suggests >10MT vs threshold {mining_prod_mt_threshold}MT")

            # => Mining Production (GW)
            if exclude_mining_prod_gw:
                if ">5gw" in production_val:
                    reasons.append(f"Mining production suggests >5GW vs threshold {mining_prod_gw_threshold}GW")

        # =========== POWER ===========
        if is_power and exclude_power:
            # => Power Revenue check
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {power_rev_threshold}% (Power)")

            # => Power Production (%)
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share * 100:.2f}% > {power_prod_threshold_percent}%")

            # => Power Production (MT)
            if exclude_power_prod_mt:
                if ">10mt" in production_val:
                    reasons.append(f"Power production suggests >10MT vs threshold {power_prod_mt_threshold}MT")

            # => Power Production (GW)
            if exclude_power_prod_gw:
                if ">5gw" in production_val:
                    reasons.append(f"Power production suggests >5GW vs threshold {power_prod_gw_threshold}GW")

            # => Max capacity (MW)
            if exclude_capacity_mw and installed_capacity > capacity_threshold_mw:
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold_mw}MW")

        # =========== SERVICES ===========
        if is_services and exclude_services:
            # => Services Revenue check
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {services_rev_threshold}% (Services)")

            # => Services Production (MT)
            if exclude_services_prod_mt:
                if ">10mt" in production_val:
                    reasons.append(f"Services production suggests >10MT vs threshold {services_prod_mt_threshold}MT")

            # => Services Production (GW)
            if exclude_services_prod_gw:
                if ">5gw" in production_val:
                    reasons.append(f"Services production suggests >5GW vs threshold {services_prod_gw_threshold}GW")

        # If reasons found => excluded
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # ---- FILE & SHEET
    st.sidebar.header("File & Sheet Settings")
    sheet_name = st.sidebar.text_input("Sheet name to check", value="GCEL 2024")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_rev_threshold = st.sidebar.number_input("Mining: Max coal revenue (%)", value=5.0)
    exclude_mining_rev = st.sidebar.checkbox("Exclude if mining rev threshold exceeded", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT for Mining", value=True)
    mining_prod_gw_threshold = st.sidebar.number_input("Mining: Max production threshold (GW)", value=5.0)
    exclude_mining_prod_gw = st.sidebar.checkbox("Exclude if > GW for Mining", value=True)

    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded", value=True)
    power_prod_mt_threshold = st.sidebar.number_input("Power: Max production threshold (MT)", value=10.0)
    exclude_power_prod_mt = st.sidebar.checkbox("Exclude if > MT for Power", value=True)
    power_prod_gw_threshold = st.sidebar.number_input("Power: Max production threshold (GW)", value=5.0)
    exclude_power_prod_gw = st.sidebar.checkbox("Exclude if > GW for Power", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed coal power capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded", value=True)

    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded", value=False)
    services_prod_mt_threshold = st.sidebar.number_input("Services: Max production threshold (MT)", value=10.0)
    exclude_services_prod_mt = st.sidebar.checkbox("Exclude if > MT for Services", value=True)
    services_prod_gw_threshold = st.sidebar.number_input("Services: Max production threshold (GW)", value=5.0)
    exclude_services_prod_gw = st.sidebar.checkbox("Exclude if > GW for Services", value=True)

    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file, sheet_name)
        if df is None:
            return

        # For 'company', skip "parent" to avoid Parent Company
        def find_co(*kw): return find_column(df, list(kw), exclude_keywords=["parent"])
        company_col = find_co("company") or "Company"

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
        }

        # Filter
        filtered_df = filter_companies(
            df,
            # Mining
            mining_rev_threshold,
            mining_prod_mt_threshold,
            mining_prod_gw_threshold,
            # Power
            power_rev_threshold,
            power_prod_threshold_percent,
            power_prod_mt_threshold,
            power_prod_gw_threshold,
            capacity_threshold_mw,
            # Services
            services_rev_threshold,
            services_prod_mt_threshold,
            services_prod_gw_threshold,
            # Exclusion toggles
            exclude_mining,
            exclude_power,
            exclude_services,
            exclude_mining_rev,
            exclude_mining_prod_mt,
            exclude_mining_prod_gw,
            exclude_power_rev,
            exclude_power_prod_percent,
            exclude_power_prod_mt,
            exclude_power_prod_gw,
            exclude_capacity_mw,
            exclude_services_rev,
            exclude_services_prod_mt,
            exclude_services_prod_gw,
            column_mapping
        )

        # EXCLUDED
        excluded_cols = [
            column_mapping["company_col"],
            column_mapping["production_col"],
            column_mapping["capacity_col"],
            column_mapping["coal_rev_col"],
            column_mapping["sector_col"],
            column_mapping["ticker_col"],
            column_mapping["isin_col"],
            column_mapping["lei_col"],
            "Exclusion Reasons",
        ]
        excluded_df = filtered_df[filtered_df["Excluded"] == True][excluded_cols]

        # RETAINED
        retained_cols = [
            column_mapping["company_col"],
            column_mapping["production_col"],
            column_mapping["capacity_col"],
            column_mapping["coal_rev_col"],
            column_mapping["sector_col"],
            column_mapping["ticker_col"],
            column_mapping["isin_col"],
            column_mapping["lei_col"],
        ]
        retained_df = filtered_df[filtered_df["Excluded"] == False][retained_cols]

        # NO-DATA
        no_data_df = df[df[column_mapping["sector_col"]].isna()][retained_cols]

        # Write final 3 sheets to Excel
        import openpyxl
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # Statistics
        st.subheader("Statistics")
        st.write(f"Total companies: {len(df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"Companies with no data: {len(no_data_df)}")

        # Show Excluded in UI
        st.subheader("Excluded Companies")
        st.dataframe(excluded_df)

        # Download
        st.download_button(
            "Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
