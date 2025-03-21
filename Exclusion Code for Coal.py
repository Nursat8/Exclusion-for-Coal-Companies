import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
import re

################################################
# MAKE COLUMNS UNIQUE
################################################
def make_columns_unique(df):
    """
    If there are duplicate column names, append _1, _2, etc. to make them unique.
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

################################################
# UTILITY: Fuzzy-rename columns to a known label
################################################
def fuzzy_rename_columns(df, rename_map):
    """
    rename_map: { "Annual Coal Production (in million metric tons)": ["annual coal production"] }
    For each "target_column" in rename_map keys,
    we'll search all df.columns to see if any column's lower-case name
    *contains* any of the strings in the rename_map[target_column].
    If so, we rename that column to target_column.

    Example:
      fuzzy_rename_columns(df, {
         "Annual Coal Production (in million metric tons)": ["annual coal production"]
      })
    """
    col_lower_map = {}
    for col_name in df.columns:
        col_lower_map[str(col_name).lower()] = col_name
    used = set()  # track which df columns we've renamed so we don't rename multiple times

    for target_col, patterns in rename_map.items():
        # Look for any column that matches ANY of the patterns
        for c in df.columns:
            if c not in used:
                c_lower = c.lower()
                for pat in patterns:
                    if pat in c_lower:
                        df.rename(columns={c: target_col}, inplace=True)
                        used.add(c)  # mark as used
                        break

    return df

################################################
# LOAD SPGLOBAL DYNAMICALLY:
#   - Row 5 (index=4) => IDs
#   - Row 6 (index=5) => e.g. Generation (Thermal Coal), Thermal Coal Mining, ...
#   - Data from row 7 onward
################################################
def load_spglobal_dynamic(file, sheet_name="Sheet1"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]

        data = list(ws.values)  # list of tuples per row
        full_df = pd.DataFrame(data)

        if len(full_df) < 6:
            raise ValueError("SPGlobal sheet does not have enough rows (need >= 6).")

        # row 4 => Excel row 5 has SP_ENTITY_NAME, SP_ENTITY_ID, ...
        # row 5 => Excel row 6 has Generation (Thermal Coal), etc.
        row_5 = full_df.iloc[4].fillna("")
        row_6 = full_df.iloc[5].fillna("")

        # Combine row 5 and 6 into single header
        final_col_names = []
        for col_idx in range(full_df.shape[1]):
            top_val = str(row_5[col_idx]).strip()
            bot_val = str(row_6[col_idx]).strip()
            combined_name = top_val if top_val else ""
            if bot_val and (bot_val.lower() not in combined_name.lower()):
                combined_name = (combined_name + " " + bot_val).strip() if combined_name else bot_val
            final_col_names.append(combined_name.strip())

        sp_data_df = full_df.iloc[6:].reset_index(drop=True)  # data from row 7
        sp_data_df.columns = final_col_names
        sp_data_df = make_columns_unique(sp_data_df)

        # Additional rename if you want to strip "SP_ESG_BUS_INVOLVE_REV_PCT"
        # e.g. "SP_ESG_BUS_INVOLVE_REV_PCT Generation (Thermal Coal)" => "Generation (Thermal Coal)"
        rename_map_sp = {
            "SP_ESG_BUS_INVOLVE_REV_PCT Generation (Thermal Coal)": "Generation (Thermal Coal)",
            "SP_ESG_BUS_INVOLVE_REV_PCT Thermal Coal Mining": "Thermal Coal Mining",
            "SP_ESG_BUS_INVOLVE_REV_PCT Metallurgical Coal Mining": "Metallurgical Coal Mining",
        }
        for old_col, new_col in rename_map_sp.items():
            if old_col in sp_data_df.columns:
                sp_data_df.rename(columns={old_col: new_col}, inplace=True)

        # Fuzzy rename for consistency
        fuzzy_map = {
            "SP_ENTITY_NAME":            ["sp_entity_name"],
            "SP_ENTITY_ID":              ["sp_entity_id"],
            "SP_COMPANY_ID":             ["sp_company_id"],
            "SP_ISIN":                   ["sp_isin"],
            "SP_LEI":                    ["sp_lei"],
            "Generation (Thermal Coal)": ["generation (thermal coal)"],
            "Thermal Coal Mining":       ["thermal coal mining"],
            "Metallurgical Coal Mining": ["metallurgical coal mining"],
        }
        sp_data_df = fuzzy_rename_columns(sp_data_df, fuzzy_map)

        return sp_data_df

    except Exception as e:
        st.error(f"Error loading SPGlobal data dynamically: {e}")
        return None

################################################
# LOAD URGEWALD DYNAMICALLY:
#   - Row 1 (index=0) => header
#   - data from row 2 onward
#   - Then fuzzy rename to standard column names so the filter logic can find them
################################################
def load_urgewald_data(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]

        data = list(ws.values)
        full_df = pd.DataFrame(data)

        if len(full_df) < 1:
            raise ValueError("Urgewald sheet does not have enough rows.")

        new_header = full_df.iloc[0].fillna("")
        ur_data_df = full_df.iloc[1:].reset_index(drop=True)
        ur_data_df.columns = new_header
        ur_data_df = make_columns_unique(ur_data_df)

        # Fuzzy rename to unify columns:
        rename_map_ug = {
            # Identity columns
            "Company":                   ["company"],
            "ISIN equity":               ["isin equity"],
            "LEI":                       ["lei"],
            "BB Ticker":                 ["bb ticker"],
            # Sectors
            "Coal Industry Sector":      ["coal industry sector"],
            # Production & capacity
            "Annual Coal Production (in million metric tons)": ["annual coal production"],
            "Installed Coal Power Capacity (MW)":              ["installed coal power capacity"],
            ">10MT / >5GW":                                    [">10mt", ">5gw"],  # If you need it
            # Expansions
            "expansion":                ["expansion"],
            # Shares
            "Coal Share of Power Production": ["coal share of power production"],
            "Coal Share of Revenue":         ["coal share of revenue"],
        }
        ur_data_df = fuzzy_rename_columns(ur_data_df, rename_map_ug)

        return ur_data_df

    except Exception as e:
        st.error(f"Error loading Urgewald file: {e}")
        return None

################################################
# REMOVE DUPLICATES (OR logic)
################################################
def remove_duplicates_or(df):
    """
    Remove duplicates if ANY match (case-insensitive):
    (SP_ENTITY_NAME vs Company) OR (SP_ISIN vs ISIN equity) OR (SP_LEI vs LEI).
    We'll unify each row's name/isin/lei, then do 3 passes of drop_duplicates.
    """
    df["_key_name_"] = df.apply(lambda r: unify_name(r), axis=1)
    df["_key_isin_"] = df.apply(lambda r: unify_isin(r), axis=1)
    df["_key_lei_"]  = df.apply(lambda r: unify_lei(r), axis=1)

    def drop_dups_on_key(data, key):
        data.loc[data[key].isna() | (data[key]==""), key] = np.nan
        data.drop_duplicates(subset=[key], keep="first", inplace=True)

    drop_dups_on_key(df, "_key_name_")
    drop_dups_on_key(df, "_key_isin_")
    drop_dups_on_key(df, "_key_lei_")

    df.drop(columns=["_key_name_","_key_isin_","_key_lei_"], inplace=True, errors="ignore")
    return df

def unify_name(r):
    sp_name = str(r.get("SP_ENTITY_NAME","")).strip().lower()
    ur_name = str(r.get("Company","")).strip().lower()
    return sp_name if sp_name else (ur_name if ur_name else None)

def unify_isin(r):
    sp_isin = str(r.get("SP_ISIN","")).strip().lower()
    ur_isin = str(r.get("ISIN equity","")).strip().lower()
    return sp_isin if sp_isin else (ur_isin if ur_isin else None)

def unify_lei(r):
    sp_lei = str(r.get("SP_LEI","")).strip().lower()
    ur_lei = str(r.get("LEI","")).strip().lower()
    return sp_lei if sp_lei else (ur_lei if ur_lei else None)


################################################
# FILTER COMPANIES
#    Now includes:
#      Generation (Thermal Coal),
#      Thermal Coal Mining,
#      Metallurgical Coal Mining
#    with separate thresholds
################################################
def filter_companies(
    df,
    mining_prod_mt_threshold,
    power_rev_threshold,
    power_prod_threshold_percent,
    capacity_threshold_mw,
    services_rev_threshold,
    # new business involvement thresholds
    generation_thermal_threshold,
    thermal_coal_mining_threshold,
    metallurgical_coal_mining_threshold,
    # Exclusion toggles
    exclude_mining,
    exclude_power,
    exclude_services,
    exclude_mining_prod_mt,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    exclude_services_rev,
    exclude_generation_thermal,
    exclude_thermal_coal_mining,
    exclude_metallurgical_coal_mining,
    # expansions
    expansions_global
):
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []

        sector_val = str(row.get("Coal Industry Sector", "")).lower()
        is_mining   = ("mining" in sector_val)
        is_power    = ("power" in sector_val) or ("generation" in sector_val)
        is_services = ("service" in sector_val)

        # expansions text
        expansion_text = str(row.get("expansion","")).lower()

        # Numeric columns
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0
        annual_coal_prod = pd.to_numeric(row.get("Annual Coal Production (in million metric tons)", 0), errors="coerce") or 0.0

        # New coal-involvement columns (percent). If blank or non-numeric => 0
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # MINING logic
        if is_mining and exclude_mining:
            # If user wants to exclude if annual prod > threshold:
            if exclude_mining_prod_mt and (annual_coal_prod > mining_prod_mt_threshold):
                reasons.append(
                    f"Annual coal production {annual_coal_prod:.2f}MT > {mining_prod_mt_threshold}MT"
                )

        # POWER logic
        if is_power and exclude_power:
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(
                    f"Coal share of revenue {coal_rev*100:.2f}% > {power_rev_threshold}% (Power)"
                )
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(
                    f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%"
                )
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(
                    f"Installed coal power capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW"
                )

        # SERVICES logic
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(
                    f"Coal share of revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)"
                )

        # Additional thresholds for the 3 columns:
        #   Generation (Thermal Coal), Thermal Coal Mining, Metallurgical Coal Mining
        # Each is presumably a % figure. If user wants to exclude if it exceeds threshold, do so.
        # Because any exceedance might exclude the company, we do an "or" approach here.
        if exclude_generation_thermal and (gen_thermal_val * 100) > generation_thermal_threshold:
            reasons.append(f"Generation (Thermal Coal) {gen_thermal_val*100:.2f}% > {generation_thermal_threshold}%")

        if exclude_thermal_coal_mining and (therm_mining_val * 100) > thermal_coal_mining_threshold:
            reasons.append(f"Thermal Coal Mining {therm_mining_val*100:.2f}% > {thermal_coal_mining_threshold}%")

        if exclude_metallurgical_coal_mining and (met_coal_val * 100) > metallurgical_coal_mining_threshold:
            reasons.append(f"Metallurgical Coal Mining {met_coal_val*100:.2f}% > {metallurgical_coal_mining_threshold}%")

        # expansions logic
        if expansions_global:
            for kw in expansions_global:
                if kw.lower() in expansion_text:
                    reasons.append(f"Expansion matched '{kw}'")
                    break

        exclusion_flags.append(bool(reasons))
        exclusion_reasons.append("; ".join(reasons) if reasons else "")

    df["Excluded"] = exclusion_flags
    df["Exclusion Reasons"] = exclusion_reasons
    return df


################################################
# MAIN STREAMLIT APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter (Dynamic Columns)", layout="wide")
    st.title("Coal Exclusion Filter with Separate Thresholds & Urgewald Data Inclusion")

    # FILES
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    # THRESHOLD TOGGLES
    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining (Sector)", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max annual production (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if Annual Coal Production > threshold?", value=True)

    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power (Sector)?", value=True)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded?", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded?", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed capacity (MW)", value=5000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity exceeded?", value=True)

    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services (Sector)?", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded?", value=False)

    # Business involvement thresholds
    st.sidebar.header("Business Involvement Thresholds (%)")
    exclude_generation_thermal = st.sidebar.checkbox("Exclude if 'Generation (Thermal Coal)' exceeds threshold?", value=False)
    generation_thermal_threshold = st.sidebar.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude if 'Thermal Coal Mining' exceeds threshold?", value=False)
    thermal_coal_mining_threshold = st.sidebar.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)

    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude if 'Metallurgical Coal Mining' exceeds threshold?", value=False)
    metallurgical_coal_mining_threshold = st.sidebar.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    # expansions
    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansions text contains any of these keywords:",
        expansions_possible,
        default=[]
    )

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal_dynamic(sp_file, sp_sheet)
        if sp_df is None or sp_df.empty:
            st.error("SPGlobal DataFrame is empty or could not be loaded.")
            return
        st.write("SPGlobal columns =>", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        # 2) Load Urgewald
        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df is None or ur_df.empty:
            st.error("Urgewald DataFrame is empty or could not be loaded.")
            return
        st.write("Urgewald columns =>", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # 3) Concatenate
        combined = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write(f"Combined shape => {combined.shape}")
        st.write("Combined columns =>", combined.columns.tolist())

        # 4) Remove duplicates (OR logic)
        deduped = remove_duplicates_or(combined.copy())
        st.write(f"After removing duplicates => {deduped.shape}")

        # 5) Filter
        filtered = filter_companies(
            df=deduped,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            power_rev_threshold=power_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            services_rev_threshold=services_rev_threshold,
            # new thresholds
            generation_thermal_threshold=generation_thermal_threshold,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            # toggles
            exclude_mining=exclude_mining,
            exclude_power=exclude_power,
            exclude_services=exclude_services,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            exclude_power_rev=exclude_power_rev,
            exclude_power_prod_percent=exclude_power_prod_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            exclude_services_rev=exclude_services_rev,
            exclude_generation_thermal=exclude_generation_thermal,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            expansions_global=expansions_global
        )

        # Separate excluded vs retained
        excluded_df = filtered[filtered["Excluded"] == True].copy()
        retained_df = filtered[filtered["Excluded"] == False].copy()

        # (Optional) "No Data" if there's no sector
        if "Coal Industry Sector" in filtered.columns:
            no_data_df = filtered[filtered["Coal Industry Sector"].isna()].copy()
        else:
            no_data_df = pd.DataFrame()

        # 6) Final columns for output
        final_cols = [
            # SP identifiers
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            # Urgewald columns
            "Company","ISIN equity","LEI","BB Ticker",
            # Sectors, expansions, etc.
            "Coal Industry Sector",
            "expansion",
            # Key metrics
            ">10MT / >5GW",
            "Annual Coal Production (in million metric tons)",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "Coal Share of Revenue",
            # Additional business involvement
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            # Exclusion
            "Excluded","Exclusion Reasons",
        ]

        # Ensure columns exist
        for c in final_cols:
            if c not in excluded_df.columns:
                excluded_df[c] = np.nan
            if c not in retained_df.columns:
                retained_df[c] = np.nan
            if not no_data_df.empty and c not in no_data_df.columns:
                no_data_df[c] = np.nan

        excluded_df = excluded_df[final_cols]
        retained_df = retained_df[final_cols]
        if not no_data_df.empty:
            no_data_df = no_data_df[final_cols]

        # 7) Output Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, "Excluded Companies", index=False)
            retained_df.to_excel(writer, "Retained Companies", index=False)
            if not no_data_df.empty:
                no_data_df.to_excel(writer, "No Data Companies", index=False)

        st.subheader("Statistics")
        st.write(f"Total after dedup: {len(filtered)}")
        st.write(f"Excluded: {len(excluded_df)}")
        st.write(f"Retained: {len(retained_df)}")
        if not no_data_df.empty:
            st.write(f"No data: {len(no_data_df)}")

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__=="__main__":
    main()
