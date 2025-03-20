import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

##############################
# MAKE COLUMNS UNIQUE
##############################
def make_columns_unique(df):
    """
    Ensures DataFrame df has uniquely named columns by appending a suffix 
    (e.g., '_1', '_2') if duplicates exist. This prevents 'InvalidIndexError'
    when concatenating.
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

##############################
# DATA-LOADING FUNCTIONS
##############################
def load_spglobal_data(file, sheet_name):
    """
    Load the SPGlobal file, skipping the first 5 rows so row #6 is used as column headers.
    """
    try:
        # row #6 => header=5 (0-based)
        df = pd.read_excel(file, sheet_name=sheet_name, header=5)
        # If you have a two-row header, you'd do header=[4,5] instead.

        if isinstance(df.columns, pd.MultiIndex):
            # Flatten a multi-level header if present
            df.columns = [
                " ".join(str(x).strip() for x in col if x not in (None, ""))
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]

        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    """
    Load the Urgewald GCEL file using row #1 as headers (header=0 in pandas).
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=0)

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

##############################
# COLUMN-FINDING HELPER
##############################
def find_column(df, must_keywords):
    """
    Returns the first column name whose .lower() contains all must_keywords (also .lower()).
    or None if not found.
    """
    must_keywords = [mk.lower() for mk in must_keywords]
    for col in df.columns:
        low = col.lower()
        if all(mk in low for mk in must_keywords):
            return col
    return None

##############################
# FILTER LOGIC (same as before)
##############################
def filter_companies(
    df,
    mining_prod_mt_threshold,
    power_rev_threshold,
    power_prod_threshold_percent,
    capacity_threshold_mw,
    services_rev_threshold,
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    expansions_global,
    # We'll just read everything from final column names to keep it simple
    # and store Exclusion Reason in "exclusion reason"
):
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []
        # We detect Mining/Power/Services from "Coal Industry Sector" if it exists
        sector_val = str(row.get("Coal Industry Sector", "")).lower()
        # expansions from "expansion" col
        expansion_txt = str(row.get("expansion", "")).lower()

        # numeric fields
        col_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        col_installed   = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0
        # production col is text
        prod_val = str(row.get(">10MT / >5GW", "")).lower()

        # Mining detection
        if "mining" in sector_val and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # Power detection
        if (("power" in sector_val) or ("generation" in sector_val)) and exclude_power:
            # If user wants to exclude by power revenue, you might have a "Coal Share of Revenue" as well:
            # not listed in final columns though. If you do, handle similarly.
            if exclude_power_prod_percent and (col_power_share * 100) > power_rev_threshold:
                reasons.append(f"Coal share of power production {col_power_share*100:.2f}% > {power_rev_threshold}%")

            if exclude_capacity_mw and col_installed > capacity_threshold_mw:
                reasons.append(f"Installed coal power capacity {col_installed:.2f}MW > {capacity_threshold_mw}MW")

        # Services detection
        if "services" in sector_val and exclude_services:
            if exclude_services_rev:
                # If you had "Coal Share of Revenue" you'd do the check here
                pass

        # expansions
        if expansions_global:
            for choice in expansions_global:
                if choice.lower() in expansion_txt:
                    reasons.append(f"Expansion plan matched '{choice}'")
                    break

        # Direct numeric checks on "Generation (Thermal Coal), Thermal Coal Mining, Metallurgical Coal Mining"
        gen_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_val   = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # Suppose thresholds are fractional (0.05 => 5%)
        # If the user meant absolute % or something, adjust accordingly
        if gen_val > power_prod_threshold_percent and exclude_power_prod_percent:
            reasons.append(f"Generation (Thermal Coal) {gen_val:.2f} > {power_prod_threshold_percent}")

        # Similarly for thermal coal
        if therm_val > power_prod_threshold_percent and exclude_power_prod_percent:
            reasons.append(f"Thermal Coal Mining {therm_val:.2f} > {power_prod_threshold_percent}")

        # Similarly for metallurgical
        if met_val > power_prod_threshold_percent and exclude_power_prod_percent:
            reasons.append(f"Metallurgical Coal Mining {met_val:.2f} > {power_prod_threshold_percent}")

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["exclusion reason"] = exclusion_reasons
    return df

##############################
# MAIN STREAMLIT APP
##############################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter: SPGlobal + Urgewald")

    # File inputs
    st.sidebar.header("File & Sheet Settings")
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    urgewald_sheet = st.sidebar.text_input("Urgewald GCEL Sheet Name", value="GCEL 2024")
    urgewald_file  = st.sidebar.file_uploader("Upload Urgewald GCEL Excel file", type=["xlsx"])

    # Basic thresholds
    st.sidebar.header("Exclusion Toggles")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    exclude_power  = st.sidebar.checkbox("Exclude Power", value=True)
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if >10MT for Mining?", value=True)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded?", value=True)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded?", value=True)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded?", value=False)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded?", value=True)

    # Numeric thresholds
    st.sidebar.header("Numeric Thresholds")
    mining_prod_mt_threshold = st.sidebar.number_input(">10MT threshold", value=10.0)
    power_rev_threshold = st.sidebar.number_input("Power rev threshold (%)", value=20.0)
    power_prod_threshold_percent = st.sidebar.number_input("Power prod threshold (%)", value=20.0)
    capacity_threshold_mw = st.sidebar.number_input("Installed capacity threshold (MW)", value=10000.0)
    services_rev_threshold = st.sidebar.number_input("Services rev threshold (%)", value=10.0)

    # expansions
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansions text contains any of these",
        expansions_possible,
        default=[]
    )

    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load each file
        sp_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        ur_df = load_urgewald_data(urgewald_file, urgewald_sheet)
        if sp_df is None or ur_df is None:
            return

        # 2) Rename columns in SP to match your final schema
        #    We'll guess actual columns in SP. Adjust these to your real names:
        rename_map_sp = {
            "Company": "company name",  # or "SP_ENTITY_NAME"? Adjust as needed
            "Entity Name": "SP_ENTITY_NAME",
            "Entity ID":   "SP_ENTITY_ID",
            "Company ID":  "SP_COMPANY_ID",
            "ISIN":        "SP_ISIN",
            "LEI":         "SP_LEI",
            "BB Ticker":   "BB Ticker",
            "Coal Industry Sector": "Coal Industry Sector",
            ">10MT / >5GW": ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)": "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production": "Coal Share of Power Production",
            "Expansion": "expansion",
            "Generation (Thermal Coal)":     "Generation (Thermal Coal)",
            "Thermal Coal Mining":           "Thermal Coal Mining",
            "Metallurgical Coal Mining":     "Metallurgical Coal Mining",
        }
        # rename columns if found:
        for old_col, new_col in rename_map_sp.items():
            if old_col in sp_df.columns:
                sp_df.rename(columns={old_col: new_col}, inplace=True)

        # 3) Rename columns in Urgewald similarly
        #    If Urgewald has "ISIN equity" => rename to "ISIN equity" (the user wants that in final).
        rename_map_ur = {
            "Company": "company name",  # or keep as "company name" if you want
            "ISIN equity": "ISIN equity", 
            "Coal Industry Sector": "Coal Industry Sector",
            ">10MT / >5GW": ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)": "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production": "Coal Share of Power Production",
            "Expansion": "expansion",
            "Generation (Thermal Coal)":     "Generation (Thermal Coal)",
            "Thermal Coal Mining":           "Thermal Coal Mining",
            "Metallurgical Coal Mining":     "Metallurgical Coal Mining",
            "LEI": "LEI"  # so we see it as a separate col if you want
        }
        for old_col, new_col in rename_map_ur.items():
            if old_col in ur_df.columns:
                ur_df.rename(columns={old_col: new_col}, inplace=True)

        # 4) Concatenate so we see all rows from SP + Urgewald
        combined_df = pd.concat([sp_df, ur_df], ignore_index=True)

        # 5) Filter them
        filtered_df = filter_companies(
            combined_df,
            mining_prod_mt_threshold,
            power_rev_threshold,
            power_prod_threshold_percent,
            capacity_threshold_mw,
            services_rev_threshold,
            exclude_mining,
            exclude_power,
            exclude_services,
            exclude_mining_prod_mt,
            exclude_power_rev,
            exclude_power_prod_percent,
            exclude_capacity_mw,
            exclude_services_rev,
            expansions_global
        )

        # 6) Build final Excluded / Retained / NoData
        excluded_df = filtered_df[filtered_df["Excluded"] == True].copy()
        retained_df = filtered_df[filtered_df["Excluded"] == False].copy()
        # "No Data" if missing a sector
        no_data_df  = filtered_df[filtered_df["Coal Industry Sector"].isna()].copy()

        # 7) We ONLY want these columns in the excluded file, in this order:
        final_cols = [
            "company name",
            "SP_ENTITY_NAME",
            "SP_ENTITY_ID",
            "SP_COMPANY_ID",
            "SP_ISIN",
            "SP_LEI",
            "BB Ticker",
            "ISIN equity",
            "LEI",
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "expansion",
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            "exclusion reason"
        ]
        # If any of these don't exist, create them so we can output a blank column:
        for c in final_cols:
            if c not in excluded_df.columns:
                excluded_df[c] = np.nan
            if c not in retained_df.columns:
                retained_df[c] = np.nan
            if c not in no_data_df.columns:
                no_data_df[c] = np.nan

        # Reorder columns
        excluded_df = excluded_df[final_cols]
        retained_df = retained_df[final_cols]
        no_data_df  = no_data_df[final_cols]

        # 8) Write out to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df.to_excel(writer, sheet_name="No Data Companies", index=False)

        st.subheader("Statistics")
        st.write(f"Total rows: {len(filtered_df)}")
        st.write(f"Excluded: {len(excluded_df)}")
        st.write(f"Retained: {len(retained_df)}")
        st.write(f"No-data (missing sector): {len(no_data_df)}")

        st.subheader("Excluded (Preview)")
        st.dataframe(excluded_df.head(30))

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
