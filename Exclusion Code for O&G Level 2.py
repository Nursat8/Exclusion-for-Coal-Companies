import streamlit as st
import pandas as pd

def load_data(file):
    """Load the Excel file and extract the GCEL 2024 sheet."""
    df = pd.read_excel(file, sheet_name="GCEL 2024")
    return df

def find_column(df, keywords):
    """Finds a column in df that contains the given keywords (case-insensitive)."""
    for col in df.columns:
        col_lower = col.lower().strip()
        if all(kw in col_lower for kw in keywords):
            return col
    return None  # Return None if no matching column is found

def filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, prod_threshold, capacity_threshold):
    """Apply exclusion criteria to filter companies."""
    exclusion_reasons = []
    exclusion_flags = []
    
    # Dynamically find columns
    company_col = find_column(df, ["company"]) or "Company"
    sector_col = find_column(df, ["coal", "industry", "sector"]) or "Coal Industry Sector"
    coal_rev_col = find_column(df, ["coal", "share", "revenue"]) or "Coal Share of Revenue"
    coal_power_col = find_column(df, ["coal", "share", "power"]) or "Coal Share of Power Production"
    capacity_col = find_column(df, ["installed", "coal", "power", "capacity"]) or "Installed Coal Power Capacity (MW)"
    production_col = find_column(df, [">10mt", ">5gw"]) or ">10MT / >5GW"
    ticker_col = find_column(df, ["bb", "ticker"]) or "BB Ticker"
    isin_col = find_column(df, ["isin", "equity"]) or "ISIN equity"
    lei_col = find_column(df, ["lei"]) or "LEI"
    
    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(sector_col, "")).strip().lower()
        
        # Skip if no valid sector data
        if sector in ["", "ni", "na", "/"]:
            exclusion_flags.append(False)
            exclusion_reasons.append("")
            continue
        
        # Identify if the company is in Power or Mining
        is_mining = "mining" in sector
        is_power = "power" in sector
        
        # Extract numerical values, safely handling missing or non-numeric entries
        coal_rev = pd.to_numeric(row.get(coal_rev_col, 0), errors="coerce") or 0
        coal_power_share = pd.to_numeric(row.get(coal_power_col, 0), errors="coerce") or 0
        installed_capacity = pd.to_numeric(row.get(capacity_col, 0), errors="coerce") or 0
        
        # Handle ">10MT / >5GW" column
        production_val = str(row.get(production_col, "")).strip().lower()
        is_large_producer = ">10mt" in production_val  # Exclude if true
        
        # Mining criteria
        if is_mining:
            if coal_rev > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {mining_rev_threshold}% (Mining)")
            if is_large_producer:
                reasons.append("Company listed as '>10Mt' producer (Mining)")
        
        # Power criteria
        if is_power:
            if coal_rev > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {power_rev_threshold}% (Power)")
            if coal_power_share > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share}% > {power_prod_threshold}%")
            if is_large_producer:
                reasons.append("Company listed as '>10Mt' producer (Power)")
            if installed_capacity > capacity_threshold:
                reasons.append(f"Installed coal power capacity {installed_capacity}MW > {capacity_threshold}MW")
        
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")
    
    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    
    return df

def main():
    st.title("Coal Exclusion Filter")
    
    # Move interface to the right using sidebar
    st.sidebar.header("Settings")
    mining_rev_threshold = st.sidebar.number_input("Mining: Max coal revenue (%)", value=5.0)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    power_prod_threshold = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    prod_threshold = st.sidebar.number_input("Max production/capacity threshold (e.g., 10MT, 5GW)", value=10.0)
    capacity_threshold = st.sidebar.number_input("Max installed coal power capacity (MW)", value=10000.0)
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file)
        filtered_df = filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, prod_threshold, capacity_threshold)
        excluded_df = filtered_df[filtered_df["Excluded"] == True]
        non_excluded_df = filtered_df[filtered_df["Excluded"] == False]
        
        # Display statistics
        st.subheader("Statistics")
        st.write(f"Total companies: {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Non-excluded companies: {len(non_excluded_df)}")
        
        # Ensure only existing columns are selected for display
        available_columns = [col for col in [company_col, sector_col, coal_rev_col, coal_power_col, capacity_col, ticker_col, isin_col, lei_col, "Exclusion Reasons"] if col in excluded_df.columns]
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded_df[available_columns])
        
        available_columns_non_excluded = [col for col in [company_col, sector_col, coal_rev_col, coal_power_col, capacity_col, ticker_col, isin_col, lei_col] if col in non_excluded_df.columns]
        
        st.subheader("Non-Excluded Companies")
        st.dataframe(non_excluded_df[available_columns_non_excluded])
        
        # Allow download of full results
        st.download_button("Download Results", data=filtered_df.to_csv(index=False), file_name="filtered_results.csv")

if __name__ == "__main__":
    main()
