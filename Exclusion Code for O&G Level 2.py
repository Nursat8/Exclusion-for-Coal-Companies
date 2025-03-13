import streamlit as st
import pandas as pd

def find_column_by_keywords(df, keywords, required=True):
    """
    Find a column in df whose header contains *all* the given keywords.
    If required=True and no match is found, raises an error.
    """
    # Convert keywords to lowercase for robust comparison
    keywords = [kw.lower().strip() for kw in keywords]
    for col in df.columns:
        col_lower = col.lower().strip()
        # Check if *all* keywords appear in the column name
        if all(kw in col_lower for kw in keywords):
            return col
    if required:
        raise ValueError(f"Could not find a column matching keywords: {keywords}")
    return None

def main():
    st.title("Coal Exclusion Filter")

    # ---- Step 1: File Upload ----
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    if uploaded_file is None:
        st.stop()

    # Read Excel into DataFrame
    # If your file has multiple sheets, specify sheet_name or use a selector.
    df = pd.read_excel(uploaded_file)

    # ---- Step 2: Let users adjust thresholds in the UI ----
    st.subheader("Set Thresholds")
    mining_revenue_threshold = st.number_input(
        "Mining: Max allowed coal share of revenue (%)",
        min_value=0.0,
        max_value=100.0,
        value=5.0,
        step=0.1
    )
    power_revenue_threshold = st.number_input(
        "Power: Max allowed coal share of revenue (%)",
        min_value=0.0,
        max_value=100.0,
        value=20.0,
        step=0.1
    )
    power_prod_share_threshold = st.number_input(
        "Power: Max allowed coal share of power production (%)",
        min_value=0.0,
        max_value=100.0,
        value=20.0,
        step=0.1
    )
    production_threshold = st.number_input(
        "Max allowed production or capacity threshold (e.g., 10 for 10MT, 5 for 5GW)",
        min_value=0.0,
        value=10.0,
        step=0.1
    )

    # ---- Step 3: Dynamically find columns by keywords ----
    # Adjust the keyword lists below to match your real column names.
    try:
        coal_sector_col  = find_column_by_keywords(df, ["coal", "industry", "sector"])
        coal_rev_col     = find_column_by_keywords(df, ["coal", "share", "revenue"])
        coal_power_col   = find_column_by_keywords(df, ["coal", "share", "power"])  # For power production share
        production_col   = find_column_by_keywords(df, [">10mt", ">5gw"], required=False)
        if not production_col:
            # If you have a better keyword or different name for the production/capacity column, adjust here
            # or set `required=True` if you must have it.
            production_col = find_column_by_keywords(df, ["production"], required=False)

        company_col      = find_column_by_keywords(df, ["company"])
        bb_ticker_col    = find_column_by_keywords(df, ["bb", "ticker"], required=False)
        isin_col         = find_column_by_keywords(df, ["isin"], required=False)
        lei_col          = find_column_by_keywords(df, ["lei"], required=False)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    # If some columns might be missing, fill them with empty data
    # so we avoid errors in the next steps.
    for possible_col in [bb_ticker_col, isin_col, lei_col, production_col, coal_power_col]:
        if possible_col and possible_col not in df.columns:
            df[possible_col] = None

    # ---- Step 4: Create a function to determine exclusion and reason(s) ----
    def determine_exclusion(row):
        sector = str(row[coal_sector_col]).strip().lower()
        
        # Safely get numeric values (if cell is blank or not a number, treat as 0.0)
        def safe_value(val):
            try:
                return float(val)
            except:
                return 0.0
        
        coal_rev = safe_value(row[coal_rev_col])
        coal_power_share = 0.0
        if coal_power_col:
            coal_power_share = safe_value(row[coal_power_col])
        production_val = 0.0
        if production_col:
            production_val = safe_value(row[production_col])
        
        reasons = []

        # MINING logic
        if "mining" in sector:
            # Exclude if coal_rev > mining_revenue_threshold or production_val > production_threshold
            if coal_rev > mining_revenue_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {mining_revenue_threshold}% (Mining)")
            if production_val > production_threshold:
                reasons.append(f"Production/Capacity {production_val} > {production_threshold} (Mining)")

        # POWER logic
        elif "power" in sector:
            # Exclude if any of these hold:
            #  - coal_rev > power_revenue_threshold
            #  - coal_power_share > power_prod_share_threshold
            #  - production_val > production_threshold
            #  - (You mentioned “coal share of revenue > 5% or …” in the prompt – adapt if needed)
            if coal_rev > power_revenue_threshold:
                reasons.append(f"Coal share of revenue {coal_rev}% > {power_revenue_threshold}% (Power)")
            if coal_power_share > power_prod_share_threshold:
                reasons.append(f"Coal share of power {coal_power_share}% > {power_prod_share_threshold}%")
            if production_val > production_threshold:
                reasons.append(f"Production/Capacity {production_val} > {production_threshold} (Power)")
            
            # If you specifically also want to check "coal_rev > 5%", uncomment below:
            # if coal_rev > 5:
            #     reasons.append(f"Coal share of revenue {coal_rev}% > 5% (Power)")

        # Return a tuple: (is_excluded, reason_string)
        if len(reasons) > 0:
            return True, "; ".join(reasons)
        else:
            return False, ""

    # ---- Step 5: Apply the exclusion logic to each row ----
    exclusion_flags = []
    exclusion_reasons = []
    for _, row in df.iterrows():
        excluded, reason = determine_exclusion(row)
        exclusion_flags.append(excluded)
        exclusion_reasons.append(reason)

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons

    # ---- Step 6: Create final output with only needed columns ----
    # Make sure to handle if bb_ticker_col, isin_col, lei_col, etc. are None
    final_cols = [col for col in [
        company_col,
        bb_ticker_col,
        isin_col,
        lei_col,
        coal_sector_col,
        coal_rev_col,
        coal_power_col,
        production_col,
        "Excluded",
        "Exclusion Reasons"
    ] if col is not None]

    output_df = df[final_cols].copy()

    # Rename columns for clarity
    rename_map = {
        company_col:        "Company",
        bb_ticker_col:      "BB Ticker",
        isin_col:           "ISIN Equity",
        lei_col:            "LEI",
        coal_sector_col:    "Coal Industry Sector",
        coal_rev_col:       "Coal Share of Revenue",
    }
    if coal_power_col:
        rename_map[coal_power_col] = "Coal Share of Power Production"
    if production_col:
        rename_map[production_col] = "Production/Capacity Column"

    output_df.rename(columns=rename_map, inplace=True)

    # ---- Step 7: Display the results ----
    st.subheader("Excluded Companies")
    excluded_df = output_df[output_df["Excluded"] == True]
    st.write(excluded_df)

    st.subheader("Non-Excluded Companies")
    non_excluded_df = output_df[output_df["Excluded"] == False]
    st.write(non_excluded_df)

    # Optionally, let user download results
    st.download_button(
        label="Download Full Results (CSV)",
        data=output_df.to_csv(index=False),
        file_name="exclusion_results.csv",
        mime="text/csv"
    )

if __name__ == "__main__":
    main()
