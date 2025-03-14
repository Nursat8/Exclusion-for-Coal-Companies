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

def find_column(df, keywords, exclude_keywords=[]):
    """Finds a column in df that contains the given keywords but not any of the exclude_keywords (case-insensitive)."""
    for col in df.columns:
        col_lower = col.lower().strip()
        if all(kw.lower() in col_lower for kw in keywords) and not any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
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
    """Apply exclusion criteria to filter companies based on thresholds."""
    exclusion_flags = []
    exclusion_reasons = []

    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping["sector_col"], "")).strip().lower()

        # Check sector membership (if sector is blank or "/"/"na"/"ni", we skip logic)
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

    # ---------------- FILE & SHEET Settings ----------------
    st.sidebar.header("File & Sheet Settings")
    sheet_name = st.sidebar.text_input("Sheet name to check", value="GCEL 2024")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    # ---------------- Process & Filter ----------------
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file, sheet_name)
        if df is None:
            return  # error displayed already

        # Dynamically find columns
        column_mapping = {
            "sector_col": find_column(df, ["industry", "sector"]),
            "company_col": find_column(df, ["company"], exclude_keywords=["parent"]),  # Ensure we exclude "Parent Company"
            "coal_rev_col": find_column(df, ["coal", "share", "revenue"]),
            "coal_power_col": find_column(df, ["coal", "share", "power"]),
            "capacity_col": find_column(df, ["installed", "coal", "power", "capacity"]),
            "production_col": find_column(df, ["10mt", "5gw"]),
            "ticker_col": find_column(df, ["bb", "ticker"]),
            "isin_col": find_column(df, ["isin", "equity"]),
            "lei_col": find_column(df, ["lei"])
        }

        # Ensure all required columns are found, or set default
        for key, default in {
            "sector_col": "Coal Industry Sector",
            "company_col": "Company",
            "coal_rev_col": "Coal Share of Revenue",
            "coal_power_col": "Coal Share of Power Production",
            "capacity_col": "Installed Coal Power Capacity (MW)",
            "production_col": ">10MT / >5GW",
            "ticker_col": "BB Ticker",
            "isin_col": "ISIN equity",
            "lei_col": "LEI"
        }.items():
            column_mapping[key] = column_mapping.get(key) or default

        # Filter the data
        filtered_df = filter_companies(
            df,
            5.0,  # Default mining_rev_threshold
            20.0, # Default power_rev_threshold
            10.0, # Default services_rev_threshold
            20.0, # Default power_prod_threshold
            10.0, # Default mining_prod_threshold
            10000.0, # Default capacity_threshold
            True, True, False, True, True, True, True, True, True, column_mapping
        )

        # Show filtered results
        st.subheader("Excluded Companies")
        st.dataframe(filtered_df[filtered_df["Excluded"] == True])

if __name__ == "__main__":
    main()
