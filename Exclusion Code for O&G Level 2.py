import streamlit as st
import pandas as pd

def load_data(file):
    """Load the Excel file and extract the GCEL 2024 sheet."""
    df = pd.read_excel(file, sheet_name="GCEL 2024")
    return df

def filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, prod_threshold):
    """Apply exclusion criteria to filter companies."""
    exclusion_reasons = []
    exclusion_flags = []
    
    for _, row in df.iterrows():
        reasons = []
        sector = str(row["Coal Industry Sector"]).strip().lower()
        coal_rev = row.get("Coal Share of Revenue", 0) or 0
        coal_power_share = row.get("Coal Share of Power Production", 0) or 0
        
        # Convert production_val to numeric safely
        production_val = pd.to_numeric(row.get(">10MT / >5GW", 0), errors="coerce") or 0
        
        # Mining criteria
        if "mining" in sector:
            if coal_rev > mining_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {mining_rev_threshold}% (Mining)")
            if production_val > prod_threshold:
                reasons.append(f"Production/Capacity {production_val} > {prod_threshold} (Mining)")
        
        # Power criteria
        elif "power" in sector:
            if coal_rev > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {power_rev_threshold}% (Power)")
            if coal_power_share > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share}% > {power_prod_threshold}%")
            if production_val > prod_threshold:
                reasons.append(f"Production/Capacity {production_val} > {prod_threshold} (Power)")
        
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")
    
    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    
    return df

def main():
    st.title("Coal Exclusion Filter")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    
    if uploaded_file:
        df = load_data(uploaded_file)
        
        # User-defined thresholds
        mining_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=5.0)
        power_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        power_prod_threshold = st.number_input("Power: Max coal power production (%)", value=20.0)
        prod_threshold = st.number_input("Max production/capacity threshold (e.g., 10MT, 5GW)", value=10.0)
        
        if st.button("Apply Filters"):
            filtered_df = filter_companies(df, mining_rev_threshold, power_rev_threshold, power_prod_threshold, prod_threshold)
            excluded_df = filtered_df[filtered_df["Excluded"] == True]
            non_excluded_df = filtered_df[filtered_df["Excluded"] == False]
            
            st.subheader("Excluded Companies")
            st.dataframe(excluded_df[["Company", "BB Ticker", "ISIN equity", "LEI", "Exclusion Reasons"]])
            
            st.subheader("Non-Excluded Companies")
            st.dataframe(non_excluded_df[["Company", "BB Ticker", "ISIN equity", "LEI"]])
            
            st.download_button("Download Results", data=filtered_df.to_csv(index=False), file_name="filtered_results.csv")

if __name__ == "__main__":
    main()
