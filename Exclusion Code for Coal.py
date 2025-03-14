import streamlit as st
import pandas as pd

def load_data(file):
    """Load the Excel file and extract the GCEL 2024 sheet."""
    df = pd.read_excel(file, sheet_name="GCEL 2024")
    return df

def filter_companies(df, selected_sectors, mining_rev_threshold, power_rev_threshold, services_rev_threshold, power_prod_threshold, mining_prod_threshold, capacity_threshold,
                      exclude_mining, exclude_power, exclude_services,
                      exclude_mining_rev, exclude_mining_prod, exclude_power_rev, exclude_power_prod, exclude_capacity, exclude_services_rev):
    """Apply exclusion criteria to filter companies based on selected sectors and thresholds."""
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
        
        # Identify sector membership
        is_mining = "mining" in sector and "Mining" in selected_sectors
        is_power = "power" in sector and "Power" in selected_sectors
        is_services = "services" in sector and "Services" in selected_sectors
        
        # Extract numerical values, safely handling missing or non-numeric entries
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0
        installed_capacity = pd.to_numeric(row.get("Installed Coal Power Capacity\n(MW)", 0), errors="coerce") or 0
        
        # Apply thresholds based on selected sectors
        if is_mining and exclude_mining:
            if exclude_mining_rev and coal_rev > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {mining_rev_threshold}% (Mining)")
            if exclude_mining_prod and "10mt" in str(row.get(">10MT / >5GW", "")).lower():
                reasons.append("Company listed as '>10Mt' producer (Mining)")
        
        if is_power and exclude_power:
            if exclude_power_rev and coal_rev > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {power_rev_threshold}% (Power)")
            if exclude_power_prod and coal_power_share > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share}% > {power_prod_threshold}%")
            if exclude_capacity and installed_capacity > capacity_threshold:
                reasons.append(f"Installed coal power capacity {installed_capacity}MW > {capacity_threshold}MW")
        
        if is_services and exclude_services:
            if exclude_services_rev and coal_rev > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {services_rev_threshold}% (Services)")
        
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")
    
    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")
    st.sidebar.header("Settings")
    
    # Multiple sector selection
    selected_sectors = st.sidebar.multiselect("Select Sectors", ["Mining", "Power", "Services"], default=["Mining", "Power", "Services"])
    
    st.sidebar.markdown("---")
    
    # Organizing thresholds per sector with better structured UI
    with st.sidebar.expander("Mining Settings"):
        exclude_mining = st.checkbox("Enable Exclusion for Mining", value=True)
        mining_rev_threshold = st.number_input("Max coal revenue (%)", value=5.0, key="mining_rev")
        exclude_mining_rev = st.checkbox("Enable Mining Revenue Exclusion", value=True, key="exclude_mining_rev")
        mining_prod_threshold = st.number_input("Max production threshold (e.g., 10MT)", value=10.0, key="mining_prod")
        exclude_mining_prod = st.checkbox("Enable Mining Production Exclusion", value=True, key="exclude_mining_prod")
    
    with st.sidebar.expander("Power Settings"):
        exclude_power = st.checkbox("Enable Exclusion for Power", value=True)
        power_rev_threshold = st.number_input("Max coal revenue (%)", value=20.0, key="power_rev")
        exclude_power_rev = st.checkbox("Enable Power Revenue Exclusion", value=True, key="exclude_power_rev")
        power_prod_threshold = st.number_input("Max coal power production (%)", value=20.0, key="power_prod")
        exclude_power_prod = st.checkbox("Enable Power Production Exclusion", value=True, key="exclude_power_prod")
        capacity_threshold = st.number_input("Max installed coal power capacity (MW)", value=10000.0, key="capacity")
        exclude_capacity = st.checkbox("Enable Power Capacity Exclusion", value=True, key="exclude_capacity")
    
    with st.sidebar.expander("Services Settings"):
        exclude_services = st.checkbox("Enable Exclusion for Services", value=False, key="exclude_services")
        services_rev_threshold = st.number_input("Max coal revenue (%)", value=10.0, key="services_rev")
        exclude_services_rev = st.checkbox("Enable Services Revenue Exclusion", value=False, key="exclude_services_rev")
    
    st.sidebar.markdown("---")
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file and st.sidebar.button("Run"):
        df = load_data(uploaded_file)
        
        filtered_df = filter_companies(df, selected_sectors, mining_rev_threshold, power_rev_threshold, services_rev_threshold, power_prod_threshold, mining_prod_threshold, capacity_threshold,
                                       exclude_mining, exclude_power, exclude_services,
                                       exclude_mining_rev, exclude_mining_prod, exclude_power_rev, exclude_power_prod, exclude_capacity, exclude_services_rev)
        excluded_df = filtered_df[filtered_df["Excluded"] == True][["Company", "BB Ticker", "ISIN equity", "LEI", "Coal Share of Revenue", "Coal Share of Power Production", "Installed Coal Power Capacity\n(MW)", ">10MT / >5GW", "Exclusion Reasons"]]
        
        # Display statistics
        st.subheader("Statistics")
        st.metric("Total Companies", len(filtered_df))
        st.metric("Excluded Companies", len(excluded_df))
        st.metric("Non-Excluded Companies", len(filtered_df) - len(excluded_df))
        
        # Display results
        st.subheader("Excluded Companies")
        st.dataframe(excluded_df)
        
        # Allow download of full results
        st.download_button("Download Results", data=excluded_df.to_csv(index=False), file_name="filtered_results.csv")

if __name__ == "__main__":
    main()
