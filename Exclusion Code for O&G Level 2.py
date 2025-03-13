import streamlit as st
import pandas as pd

def load_data(file):
    """Load the Excel file and extract the GCEL 2024 sheet."""
    df = pd.read_excel(file, sheet_name="GCEL 2024")
    return df

def filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, mining_prod_threshold, power_prod_threshold_mt, capacity_threshold):
    """Apply exclusion criteria to filter companies."""
    exclusion_reasons = []
    exclusion_flags = []
    
    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get("Coal Industry Sector", "")).strip().lower()
        
        # Skip if no valid sector data
        if sector in ["", "ni", "na", "/"]:
            exclusion_flags.append(False)
            exclusion_reasons.append("")
            continue
        
        # Identify if the company is in Mining, Power, or Services
        is_mining = "mining" in sector
        is_power = "power" in sector
        is_services = "services" in sector
        
        # Extract numerical values, safely handling missing or non-numeric entries
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0
        installed_capacity = pd.to_numeric(row.get("Installed Coal Power Capacity\n(MW)", 0), errors="coerce") or 0
        
        # Handle ">10MT / >5GW" column for mining only
        production_val = str(row.get(">10MT / >5GW", "")).strip().lower()
        is_large_mining_producer = is_mining and ">10mt" in production_val  # Exclude if true
        is_large_power_producer = is_power and ">10mt" in production_val  # Exclude if true
        
        # Mining criteria
        if is_mining:
            if coal_rev > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {mining_rev_threshold}% (Mining)")
            if is_large_mining_producer:
                reasons.append("Company listed as '>10Mt' producer (Mining)")
        
        # Power criteria
        if is_power:
            if coal_rev > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {power_rev_threshold}% (Power)")
            if coal_power_share > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share}% > {power_prod_threshold}%")
            if is_large_power_producer:
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
    mining_prod_threshold = st.sidebar.number_input("Mining: Max production threshold (e.g., 10MT)", value=10.0)
    power_prod_threshold_mt = st.sidebar.number_input("Power: Max production threshold (e.g., 10MT)", value=10.0)
    capacity_threshold = st.sidebar.number_input("Max installed coal power capacity (MW)", value=10000.0)
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file)
        
        filtered_df = filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, mining_prod_threshold, power_prod_threshold_mt, capacity_threshold)
        excluded_df = filtered_df[filtered_df["Excluded"] == True]
        non_excluded_df = filtered_df[filtered_df["Excluded"] == False]
        
        # Display statistics
        st.subheader("Statistics")
        st.write(f"Total companies: {len(filtered_df)}")
        st.write(f"Excluded companies: {len(excluded_df)}")
        st.write(f"Non-excluded companies: {len(non_excluded_df)}")
        
        # Ensure only existing columns are selected for display
        selected_columns = [col for col in ["Company", "Coal Industry Sector", "Coal Share of Revenue", "Coal Share of Power Production", "Installed Coal Power Capacity\n(MW)", ">10MT / >5GW", "BB Ticker", "ISIN equity", "LEI", "Exclusion Reasons"] if col in excluded_df.columns]
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded_df[selected_columns])
        
        selected_columns_non_excluded = [col for col in ["Company", "Coal Industry Sector", "Coal Share of Revenue", "Coal Share of Power Production", "Installed Coal Power Capacity\n(MW)", ">10MT / >5GW", "BB Ticker", "ISIN equity", "LEI"] if col in non_excluded_df.columns]
        
        st.subheader("Non-Excluded Companies")
        st.dataframe(non_excluded_df[selected_columns_non_excluded])
        
        # Allow download of full results
        st.download_button("Download Results", data=filtered_df.to_csv(index=False), file_name="filtered_results.csv")

if __name__ == "__main__":
    main()
