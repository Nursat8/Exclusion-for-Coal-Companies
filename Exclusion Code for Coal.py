import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

############################################################
# 1) MAKE COLUMNS UNIQUE
############################################################
def make_columns_unique(df):
    """
    Renames duplicate column names by appending a suffix (_1, _2, etc.)
    so we don't get 'InvalidIndexError' on concat.
    """
    seen = {}
    new_cols = []
    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    df.columns = new_cols
    return df

############################################################
# 2) LOAD FUNCTIONS
############################################################
def load_spglobal_data(file, sheet_name):
    """
    SPGlobal has multi-header lines in row #4 and #5 (1-based).
    => we pass header=[3,4] to read them as a multi-level header.
    Then flatten.
    Adjust if your file actually has them at a different place.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=[3,4])
        # Flatten multi-level columns
        df.columns = [
            " ".join(str(x).strip() for x in col if x not in (None, ""))
            for col in df.columns
        ]
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    """
    Urgewald has a single header in row #1 => header=0.
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=0)
        # Flatten if multi-level
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(x).strip() for x in col if x not in (None, ""))
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]

        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading Urgewald sheet '{sheet_name}': {e}")
        return None

############################################################
# 3) FILTERING LOGIC
############################################################
def filter_companies(
    df,
    # Mining thresholds
    mining_prod_mt_threshold,
    # Power thresholds
    power_rev_threshold,   # for coal revenue
    power_prod_threshold_percent,
    capacity_threshold_mw,
    # Services thresholds
    services_rev_threshold,
    # Exclusion toggles
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    # expansions
    expansions_global
):
    """
    Applies your toggles/thresholds. We'll read from columns:
      - 'Coal Share of Revenue' (for power_rev_threshold, services_rev_threshold)
      - 'Coal Industry Sector' or something for "mining/power/services" checks
      - '>10MT / >5GW' for text-based production checks
      - 'Installed Coal Power Capacity(MW)' for capacity threshold
      - 'Coal Share of Power Production' for power_prod_threshold_percent
      - 'expansion' text for expansions
      - 'Generation (Thermal Coal)', 'Thermal Coal Mining', 'Metallurgical Coal Mining'
        if you want direct numeric checks. (We can incorporate them if you need.)
    """
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []

        # We detect sector from 'Coal Industry Sector' if it exists
        sector_val = str(row.get("Coal Industry Sector", "")).lower()
        is_mining    = ("mining" in sector_val)
        is_power     = ("power" in sector_val) or ("generation" in sector_val)
        is_services  = ("services" in sector_val)

        # expansions
        expansion_text = str(row.get("expansion", "")).lower()

        # Production text for mining
        prod_val = str(row.get(">10MT / >5GW", "")).lower()

        # Numeric fields
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap    = pd.to_numeric(row.get("Installed Coal Power Capacity(MW)", 0), errors="coerce") or 0.0

        ######################
        # 1) Mining Exclusion
        ######################
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs threshold={mining_prod_mt_threshold}MT")

        ######################
        # 2) Power Exclusion
        ######################
        if is_power and exclude_power:
            # exclude if coal_rev > power_rev_threshold
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {power_rev_threshold}% (Power)")

            # exclude if coal_power_share > power_prod_threshold_percent
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            # exclude if installed_cap > capacity_threshold_mw
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")

        ######################
        # 3) Services Exclusion
        ######################
        if is_services and exclude_services:
            # if you want to exclude if coal_rev>services_rev_threshold
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")

        ######################
        # 4) expansions
        ######################
        if expansions_global:
            for keyword in expansions_global:
                if keyword.lower() in expansion_text:
                    reasons.append(f"Expansion matched '{keyword}'")
                    break

        ######################
        # 5) If you want direct numeric checks for Generation (Thermal Coal),
        #    Thermal Coal Mining, Metallurgical Coal Mining, incorporate them here
        #    For example:
        #    gen_val = pd.to_numeric(row.get("Generation (Thermal Coal)",0), errors="coerce") or 0.0
        #    if gen_val > some_threshold => ...
        ######################

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df

############################################################
# 4) MAIN STREAMLIT APP
############################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    st.sidebar.header("File & Sheet Settings")
    # SPGlobal
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    # Urgewald
    urgewald_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    urgewald_file  = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if >10MT for Mining?", value=True)

    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power?", value=True)
    power_rev_threshold = st.sidebar.number_input("Max coal revenue (%) for power", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded?", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded?", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Max installed capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded?", value=True)

    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services?", value=False)
    services_rev_threshold = st.sidebar.number_input("Max services rev threshold (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded?", value=False)

    st.sidebar.header("Expansions")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansion text contains any of these",
        expansions_possible,
        default=[]
    )

    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # =============== LOAD SPGlobal ===============
        sp_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        if sp_df is None:
            return
        st.write("SPGlobal columns:", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        # =============== LOAD Urgewald ===============
        ur_df = load_urgewald_data(urgewald_file, urgewald_sheet)
        if ur_df is None:
            return
        st.write("Urgewald columns:", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # =============== RENAME SPGlobal columns ===============
        # Adjust these to match the columns actually found in sp_df.columns
        rename_map_sp = {
            "SP_ENTITY_NAME": "SP_ENTITY_NAME",
            "SP_ENTITY_ID":   "SP_ENTITY_ID",
            "SP_COMPANY_ID":  "SP_COMPANY_ID",
            "SP_ISIN":        "SP_ISIN",
            "SP_LEI":         "SP_LEI",
            "Generation (Thermal Coal)":  "Generation (Thermal Coal)",
            "Thermal Coal Mining":        "Thermal Coal Mining",
            "Metallurgical Coal Mining":  "Metallurgical Coal Mining",
            # etc. if needed
        }
        for old_col, new_col in rename_map_sp.items():
            if old_col in sp_df.columns:
                sp_df.rename(columns={old_col: new_col}, inplace=True)

        # =============== RENAME Urgewald columns ===============
        # Suppose Urgewald has these exact columns:
        rename_map_ur = {
            ">10MT / >5GW": ">10MT / >5GW",
            "expansion": "expansion",
            "Company": "Company",
            "Installed Coal Power Capacity(MW)": "Installed Coal Power Capacity(MW)",
            "Coal Share of Power Production": "Coal Share of Power Production",
            "Coal Share of Revenue": "Coal Share of Revenue",
            "BB Ticker": "BB Ticker",
            "ISIN equity": "ISIN equity",
            "LEI": "LEI"
        }
        for old_col, new_col in rename_map_ur.items():
            if old_col in ur_df.columns:
                ur_df.rename(columns={old_col: new_col}, inplace=True)

        # =============== CONCATENATE ===============
        combined_df = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write("Combined columns:", combined_df.columns.tolist())
        st.dataframe(combined_df.head(5))

        # =============== FILTER ===============
        filtered_df = filter_companies(
            df=combined_df,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            power_rev_threshold=power_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            services_rev_threshold=services_rev_threshold,
            exclude_mining=exclude_mining,
            exclude_power=exclude_power,
            exclude_services=exclude_services,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            exclude_power_rev=exclude_power_rev,
            exclude_power_prod_percent=exclude_power_prod_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            exclude_services_rev=exclude_services_rev,
            expansions_global=expansions_global
        )

        # =============== BUILD OUTPUT SHEETS ===============
        excluded_df = filtered_df[filtered_df["Excluded"] == True].copy()
        retained_df = filtered_df[filtered_df["Excluded"] == False].copy()
        no_data_df  = filtered_df[filtered_df["Coal Industry Sector"].isna()] if "Coal Industry Sector" in filtered_df.columns else pd.DataFrame()

        # For final, pick some columns
        excluded_cols = [
            "SP_ENTITY_NAME",
            "SP_ENTITY_ID",
            "SP_COMPANY_ID",
            "SP_ISIN",
            "SP_LEI",
            "Company",  # from Urgewald
            "BB Ticker",
            "ISIN equity",
            "LEI",  # from Urgewald
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity(MW)",
            "Coal Share of Power Production",
            "Coal Share of Revenue",
            "expansion",
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            "Exclusion Reasons"
        ]
        # Create missing columns if needed:
        for c in excluded_cols:
            if c not in excluded_df.columns:
                excluded_df[c] = np.nan
            if c not in retained_df.columns:
                retained_df[c] = np.nan
            if not no_data_df.empty and c not in no_data_df.columns:
                no_data_df[c] = np.nan

        # Reorder
        excluded_df = excluded_df[excluded_cols]
        retained_df = retained_df[excluded_cols]
        if not no_data_df.empty:
            no_data_df = no_data_df[excluded_cols]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, "Excluded Companies", index=False)
            retained_df.to_excel(writer, "Retained Companies", index=False)
            if not no_data_df.empty:
                no_data_df.to_excel(writer, "No Data Companies", index=False)

        # =============== STATS & DOWNLOAD ===============
        st.subheader("Statistics")
        st.write(f"Total rows after concat: {len(filtered_df)}")
        st.write(f"Excluded: {len(excluded_df)}")
        st.write(f"Retained: {len(retained_df)}")
        if not no_data_df.empty:
            st.write(f"No-data (missing sector): {len(no_data_df)}")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
