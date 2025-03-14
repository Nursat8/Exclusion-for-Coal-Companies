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
    Finds a column in df that contains all must_keywords (case-insensitive)
    and excludes any column that contains any exclude_keywords.
    Returns the first matching column name or None if no match.
    """
    if exclude_keywords is None:
        exclude_keywords = []

    for col in df.columns:
        col_lower = col.lower().strip()
        # Must contain all must_keywords
        if all(mk.lower() in col_lower for mk in must_keywords):
            # Exclude if any exclude_keywords match
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue  # Skip this column
            return col
    return None

def filter_companies(
    df,
    mining_rev_threshold,
    power_rev_threshold,
    services_rev_threshold,
    power_prod_threshold,
    mining_prod_threshold,
    capacity_threshold,
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_rev,
    exclude_mining_prod,
    exclude_power_rev,
    exclude_power_prod,
    exclude_capacity,
    exclude_services_rev,
    column_mapping
):
    """
    Apply exclusion criteria to filter companies based on thresholds.
    Decimal values in the data are fractions-of-1.0, so multiply by 100 for comparison.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip().lower()

        # Identify sector membership
        is_mining = ("mining" in sector)
        is_power = ("power" in sector)
        is_services = ("services" in sector)

        # Extract numeric fields
        coal_rev = pd.to_numeric(row.get(column_mapping["coal_rev_col"], 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping["coal_power_col"], 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping["capacity_col"], 0), errors="coerce") or 0.0

        # -- MINING --
        if is_mining and exclude_mining:
            if exclude_mining_rev and (coal_rev * 100) > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {mining_rev_threshold}% (Mining)")
            if exclude_mining_prod:
                production_val = str(row.get(column_mapping["production_col"], "")).lower()
                if ">10mt" in production_val:
                    reasons.append("Company listed as '>10Mt' producer (Mining)")

        # -- POWER --
        if is_power and exclude_power:
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {power_rev_threshold}% (Power)")
            if exclude_power_prod and (coal_power_share * 100) > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share * 100:.2f}% > {power_prod_threshold}%")
            if exclude_capacity and (installed_capacity > capacity_threshold):
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold}MW")

        # -- SERVICES --
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {services_rev_threshold}% (Services)")

        # Exclude if reasons found
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # 1) File & Sheet
    st.sidebar.header("File & Sheet Settings")
    sheet_name = st.sidebar.text_input("Sheet name to check", value="GCEL 2024")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    # 2) Thresholds
    st.sidebar.header("Sector Thresholds")

    with st.sidebar.expander("Mining Settings", expanded=True):
        exclude_mining = st.checkbox("Enable Exclusion for Mining", value=True)
        mining_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=5.0)
        exclude_mining_rev = st.checkbox("Enable Mining Revenue Exclusion", value=True)
        mining_prod_threshold = st.number_input("Mining: Max production threshold (MT)", value=10.0)
        exclude_mining_prod = st.checkbox("Enable Mining Production Exclusion", value=True)

    with st.sidebar.expander("Power Settings", expanded=True):
        exclude_power = st.checkbox("Enable Exclusion for Power", value=True)
        power_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_rev = st.checkbox("Enable Power Revenue Exclusion", value=True)
        power_prod_threshold = st.number_input("Power: Max coal power production (%)", value=20.0)
        exclude_power_prod = st.checkbox("Enable Power Production Exclusion", value=True)
        capacity_threshold = st.number_input("Power: Max installed coal power capacity (MW)", value=10000.0)
        exclude_capacity = st.checkbox("Enable Power Capacity Exclusion", value=True)

    with st.sidebar.expander("Services Settings", expanded=True):
        exclude_services = st.checkbox("Enable Exclusion for Services", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)
        exclude_services_rev = st.checkbox("Enable Services Revenue Exclusion", value=False)

    # 3) Filter
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file, sheet_name)
        if df is None:
            return

        # 4) Column detection, exclude "parent" for company
        company_column = find_column(df, must_keywords=["company"], exclude_keywords=["parent"]) \
            or "Company"

        column_mapping = {
            "sector_col": (
                find_column(df, ["industry", "sector"]) or "Coal Industry Sector"
            ),
            "company_col": company_column,
            "coal_rev_col": (
                find_column(df, ["coal", "share", "revenue"]) or "Coal Share of Revenue"
            ),
            "coal_power_col": (
                find_column(df, ["coal", "share", "power"]) or "Coal Share of Power Production"
            ),
            "capacity_col": (
                find_column(df, ["installed", "coal", "power", "capacity"])
                or "Installed Coal Power Capacity (MW)"
            ),
            "production_col": (
                find_column(df, ["10mt", "5gw"]) or ">10MT / >5GW"
            ),
            "ticker_col": (
                find_column(df, ["bb", "ticker"]) or "BB Ticker"
            ),
            "isin_col": (
                find_column(df, ["isin", "equity"]) or "ISIN equity"
            ),
            "lei_col": (
                find_column(df, ["lei"]) or "LEI"
            ),
        }

        # 5) Filter
        filtered_df = filter_companies(
            df,
            mining_rev_threshold,
            power_rev_threshold,
            services_rev_threshold,
            power_prod_threshold,
            mining_prod_threshold,
            capacity_threshold,
            exclude_mining,
            exclude_power,
            exclude_services,
            exclude_mining_rev,
            exclude_mining_prod,
            exclude_power_rev,
            exclude_power_prod,
            exclude_capacity,
            exclude_services_rev,
            column_mapping
        )

        # Show partial result
        st.subheader("Excluded Companies")
        excluded_df = filtered_df[filtered_df["Excluded"] == True]
        st.dataframe(excluded_df)

if __name__ == "__main__":
    main()
