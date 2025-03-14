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
    and excludes any column that contains any of the exclude_keywords.
    Returns the first matching column name or None if no match.
    """
    if exclude_keywords is None:
        exclude_keywords = []

    for col in df.columns:
        col_lower = col.lower().strip()

        # Must contain all these
        if all(mk.lower() in col_lower for mk in must_keywords):
            # Must not contain any exclude keywords
            if any(ek.lower() in col_lower for ek in exclude_keywords):
                continue  # skip this col
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
    Decimal values in the data are treated as fraction-of-1.0, so multiply by 100 for comparison.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip().lower()

        # Identify sector membership
        is_mining = "mining" in sector
        is_power = "power" in sector
        is_services = "services" in sector

        # Extract numeric fields
        coal_rev = pd.to_numeric(row.get(column_mapping["coal_rev_col"], 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping["coal_power_col"], 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping["capacity_col"], 0), errors="coerce") or 0.0

        # ---- MINING ----
        if is_mining and exclude_mining:
            if exclude_mining_rev and coal_rev * 100 > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {mining_rev_threshold}% (Mining)")
            if exclude_mining_prod:
                production_val = str(row.get(column_mapping["production_col"], "")).lower()
                if ">10mt" in production_val:
                    reasons.append("Company listed as '>10Mt' producer (Mining)")

        # ---- POWER ----
        if is_power and exclude_power:
            if exclude_power_rev and coal_rev * 100 > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {power_rev_threshold}% (Power)")
            if exclude_power_prod and coal_power_share * 100 > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share * 100:.2f}% > {power_prod_threshold}%")
            if exclude_capacity and installed_capacity > capacity_threshold:
                reasons.append(f"Installed coal power capacity {installed_capacity:.2f}MW > {capacity_threshold}MW")

        # ---- SERVICES ----
        if is_services and exclude_services:
            if exclude_services_rev and coal_rev * 100 > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100:.2f}% > {services_rev_threshold}% (Services)")

        # If we found reasons, the row is excluded
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # =============== FILE & SHEET Settings ===============
    st.sidebar.header("File & Sheet Settings")
    sheet_name = st.sidebar.text_input("Sheet name to check", value="GCEL 2024")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    # =============== Sector Thresholds ===============
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

    # =============== RUN ===============
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file, sheet_name)
        if df is None:
            return

        # Dynamically detect columns
        column_mapping = {
            "sector_col": find_column(df, ["industry", "sector"]) or "Coal Industry Sector",
            # EXCLUDE 'parent' here so we don't pick "Parent Company" columns
            "company_col": find_column(df, ["company"]) or "Company",
            "coal_rev_col": find_column(df, ["coal", "share", "revenue"]) or "Coal Share of Revenue",
            "coal_power_col": find_column(df, ["coal", "share", "power"]) or "Coal Share of Power Production",
            "capacity_col": find_column(df, ["installed", "coal", "power", "capacity"]) or "Installed Coal Power Capacity (MW)",
            "production_col": find_column(df, ["10mt", "5gw"]) or ">10MT / >5GW",
            "ticker_col": find_column(df, ["bb", "ticker"]) or "BB Ticker",
            "isin_col": find_column(df, ["isin", "equity"]) or "ISIN equity",
            "lei_col": find_column(df, ["lei"]) or "LEI",
        }

        # Filter
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

        # Build slices
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
        no_data_df = df[df[column_mapping["sector_col"]].isna()][retained_cols]

        # Save to Excel in-memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        # Stats
        st.subheader("Statistics")
        st.write(f"Total companies: {len(df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Retained companies: {len(retained_df)}")
        st.write(f"Companies with no data: {len(no_data_df)}")

        # Excluded preview
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
