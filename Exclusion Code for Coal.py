import streamlit as st
import pandas as pd
import io

def load_data(file):
    """Load the Excel file and extract the GCEL 2024 sheet."""
    df = pd.read_excel(file, sheet_name="GCEL 2024")
    return df

def find_column(df, keywords):
    """Finds a column in df that contains the given keywords (case-insensitive)."""
    for col in df.columns:
        col_lower = col.lower().strip()
        if all(kw.lower() in col_lower for kw in keywords):
            return col
    return None  # Return None if no matching column is found

def filter_companies(df, mining_rev_threshold, power_rev_threshold, services_rev_threshold, power_prod_threshold, mining_prod_threshold, capacity_threshold,
                      exclude_mining, exclude_power, exclude_services,
                      exclude_mining_rev, exclude_mining_prod, exclude_power_rev, exclude_power_prod, exclude_capacity, exclude_services_rev,
                      column_mapping):
    """Apply exclusion criteria to filter companies based on thresholds."""
    exclusion_reasons = []
    exclusion_flags = []
    
    for _, row in df.iterrows():
        reasons = []
        sector = str(row.get(column_mapping.get("sector_col", ""), "")).strip().lower()
        
        # Identify sector membership
        is_mining = "mining" in sector
        is_power = "power" in sector
        is_services = "services" in sector
        
        # Extract numerical values, safely handling missing or non-numeric entries
        coal_rev = pd.to_numeric(row.get(column_mapping.get("coal_rev_col", ""), 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get(column_mapping.get("coal_power_col", ""), 0), errors="coerce") or 0.0
        installed_capacity = pd.to_numeric(row.get(column_mapping.get("capacity_col", ""), 0), errors="coerce") or 0.0
        
        # Apply thresholds based on selected sectors
        if is_mining and exclude_mining:
            if exclude_mining_rev and coal_rev * 100 > mining_rev_threshold:  # Convert decimal to percentage
                reasons.append(f"Coal share of revenue {coal_rev * 100}% > {mining_rev_threshold}% (Mining)")
            if exclude_mining_prod and ">10mt" in str(row.get(column_mapping.get("production_col", ""), "")).lower():
                reasons.append("Company listed as '>10Mt' producer (Mining)")
        
        if is_power and exclude_power:
            if exclude_power_rev and coal_rev * 100 > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100}% > {power_rev_threshold}% (Power)")
            if exclude_power_prod and coal_power_share * 100 > power_prod_threshold:
                reasons.append(f"Coal share of power production {coal_power_share * 100}% > {power_prod_threshold}%")
            if exclude_capacity and installed_capacity > capacity_threshold:
                reasons.append(f"Installed coal power capacity {installed_capacity}MW > {capacity_threshold}MW")
        
        if is_services and exclude_services:
            if exclude_services_rev and coal_rev * 100 > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev * 100}% > {services_rev_threshold}% (Services)")
        
        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")
    
    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    
    return df

def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")
    st.sidebar.header("Settings")
    
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        df = load_data(uploaded_file)
        
        # Dynamically detect column names
        column_mapping = {
            "sector_col": find_column(df, ["industry", "sector"]) or "Coal Industry Sector",
            "company_col": find_column(df, ["company"]) or "Company",
            "coal_rev_col": find_column(df, ["coal", "share", "revenue"]) or "Coal Share of Revenue",
            "coal_power_col": find_column(df, ["coal", "share", "power"]) or "Coal Share of Power Production",
            "capacity_col": find_column(df, ["installed", "coal", "power", "capacity"]) or "Installed Coal Power Capacity (MW)",
            "production_col": find_column(df, ["10mt", "5gw"]) or ">10MT / >5GW",
            "ticker_col": find_column(df, ["bb", "ticker"]) or "BB Ticker",
            "isin_col": find_column(df, ["isin", "equity"]) or "ISIN equity",
            "lei_col": find_column(df, ["lei"]) or "LEI"
        }
        
        if st.sidebar.button("Run"):
            filtered_df = filter_companies(df, 5.0, 20.0, 10.0, 20.0, 10.0, 10000.0,
                                           True, True, True,
                                           True, True, True, True, True, True,
                                           column_mapping)
            
            excluded_df = filtered_df[filtered_df["Excluded"] == True][[column_mapping["company_col"], column_mapping["production_col"], column_mapping["capacity_col"],
                                                                          column_mapping["coal_rev_col"], column_mapping["sector_col"], column_mapping["ticker_col"],
                                                                          column_mapping["isin_col"], column_mapping["lei_col"], "Exclusion Reasons"]]
            retained_df = filtered_df[filtered_df["Excluded"] == False]
            no_data_df = df[df[column_mapping["sector_col"]].isna()]
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
                retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
                no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)
                writer.close()
            
            st.subheader("Statistics")
            st.write(f"Total companies: {len(df)}")
            st.write(f"Excluded companies: {len(excluded_df)}")
            st.write(f"Retained companies: {len(retained_df)}")
            st.write(f"Companies with no data: {len(no_data_df)}")
            
            st.subheader("Excluded Companies")
            st.dataframe(excluded_df)
            
            st.download_button("Download Results", data=output.getvalue(), file_name="filtered_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
