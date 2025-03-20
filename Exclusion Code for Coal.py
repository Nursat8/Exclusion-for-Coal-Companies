import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

################################################################
# 1) MAKE COLUMNS UNIQUE
################################################################
def make_columns_unique(df):
    """If a DataFrame has duplicate column names, rename them by appending '_1','_2', etc."""
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

################################################################
# 2) LOAD SPGLOBAL
################################################################
def load_spglobal_data(sp_file, sheet_name):
    """
    Attempt to skip 'PEND' lines and read the correct row as your header.
    Try skiprows=2 or skiprows=3 or skiprows=4 etc. until you see
    'SP_ENTITY_NAME' etc. in df.columns.

    Below is an example with skiprows=2, header=0. Adjust if needed.
    """
    try:
        # If row #3 in Excel is #PEND, row #4 is #PEND, row #5 has SP_ENTITY_NAME...
        # Then possibly skiprows=3 => discards lines 1..3, next line is row #4 => that might still be #PEND.
        # If row #5 is your real header => skiprows=4, header=0 might be correct.
        # Tweak these until you see the correct columns in the debug print!
        df = pd.read_excel(sp_file, sheet_name=sheet_name, skiprows=4, header=0)

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
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

################################################################
# 3) LOAD URGEWALD
################################################################
def load_urgewald_data(ur_file, sheet_name):
    """
    Urgewald presumably has a normal 1-row header => header=0.
    Adjust skiprows if there's a blank line or title row above.
    """
    try:
        df = pd.read_excel(ur_file, sheet_name=sheet_name, header=0)

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

################################################################
# 4) REMOVE DUPLICATES (OR LOGIC) ON (NAME, ISIN, LEI)
################################################################
def remove_duplicates_or(df):
    """
    Remove duplicates if ANY of these match (case-insensitive):
     - (SP_ENTITY_NAME vs. Company)
     - (SP_ISIN vs. ISIN equity)
     - (SP_LEI vs. LEI)
    We'll do a 3-pass approach with drop_duplicates.

    1) unify name => if sp_entity_name is non-empty, that's the key, else if 'Company' is non-empty, that's the key
    2) unify isin => sp_isin if non-empty, else 'ISIN equity'
    3) unify lei => sp_lei if non-empty, else 'LEI'

    Then we do drop_duplicates on each key sequentially. That ensures "OR" logic.
    """
    df["_key_name_"] = df.apply(lambda r: unify_name(r), axis=1)
    df["_key_isin_"] = df.apply(lambda r: unify_isin(r), axis=1)
    df["_key_lei_"]  = df.apply(lambda r: unify_lei(r), axis=1)

    # We define a helper to do drop_duplicates on that key,
    # ignoring blanks so we don't treat empty as duplicates.
    def drop_dups_on_key(df, key):
        df.loc[df[key].isna() | (df[key]==""), key] = np.nan
        df.drop_duplicates(subset=[key], keep="first", inplace=True)
        return df

    drop_dups_on_key(df, "_key_name_")
    drop_dups_on_key(df, "_key_isin_")
    drop_dups_on_key(df, "_key_lei_")

    df.drop(columns=["_key_name_","_key_isin_","_key_lei_"], inplace=True)
    return df

def unify_name(row):
    sp_name = str(row.get("SP_ENTITY_NAME","")).strip().lower()
    ur_name = str(row.get("Company","")).strip().lower()
    if sp_name:
        return sp_name
    elif ur_name:
        return ur_name
    else:
        return None

def unify_isin(row):
    sp_isin = str(row.get("SP_ISIN","")).strip().lower()
    ur_isin = str(row.get("ISIN equity","")).strip().lower()
    if sp_isin:
        return sp_isin
    elif ur_isin:
        return ur_isin
    return None

def unify_lei(row):
    sp_lei = str(row.get("SP_LEI","")).strip().lower()
    ur_lei = str(row.get("LEI","")).strip().lower()
    if sp_lei:
        return sp_lei
    elif ur_lei:
        return ur_lei
    return None

################################################################
# 5) FILTERING LOGIC (restoring your toggles)
################################################################
def filter_companies(
    df,
    # Mining thresholds
    mining_prod_mt_threshold,
    # Power thresholds
    power_rev_threshold,
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
    Similar to your original partial-match logic for sector, expansions, etc.
    We'll read from possible columns: 'Coal Industry Sector', 'Coal Share of Power Production',
    'Installed Coal Power Capacity (MW)', 'Coal Share of Revenue', etc.

    We'll also do the partial match for 'Generation (Thermal Coal)', 'Thermal Coal Mining',
    'Metallurgical Coal Mining' if you want. For brevity, let's do the basic version.
    """
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []
        # Sector
        sector_val = str(row.get("Coal Industry Sector", "")).lower()

        is_mining    = ("mining" in sector_val)
        is_power     = ("power" in sector_val) or ("generation" in sector_val)
        is_services  = ("services" in sector_val)

        # expansions
        expansion_text = str(row.get("expansion","")).lower()

        # numeric fields
        prod_val = str(row.get(">10MT / >5GW","")).lower()
        # We might read 'Coal Share of Revenue'
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue",0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production",0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity(MW)",0), errors="coerce") or 0.0

        # ---------- MINING ----------
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # ---------- POWER ----------
        if is_power and exclude_power:
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")

        # ---------- SERVICES ----------
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")

        # ---------- expansions ----------
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

################################################################
# 6) MAIN STREAMLIT APP
################################################################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter + Dedup (OR)", layout="wide")
    st.title("Coal Exclusion Filter + Remove Duplicates (OR)")

    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file  = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    # Mining toggles
    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Max production threshold (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT for Mining?", value=True)

    # Power toggles
    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)
    power_rev_threshold = st.sidebar.number_input("Max coal revenue (%) for power", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded?", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power production % exceeded?", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Max installed coal power capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity threshold exceeded?", value=True)

    # Services
    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services?", value=False)
    services_rev_threshold = st.sidebar.number_input("Max services rev threshold (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded?", value=False)

    # expansions
    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect("Exclude if expansions text contains any of these",
                                               expansions_possible,
                                               default=[])

    # RUN
    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal_data(sp_file, sp_sheet)
        if sp_df is None:
            return
        st.write("SP columns =>", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        # 2) Load Urgewald
        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df is None:
            return
        st.write("UR columns =>", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # 3) Concat
        combined_df = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write(f"Combined shape => {combined_df.shape}")
        st.write("Combined columns =>", combined_df.columns.tolist())

        # 4) Remove duplicates using the OR logic
        dedup_df = remove_duplicates_or(combined_df.copy())
        st.write(f"After dedup => {dedup_df.shape}")

        # 5) Filter with your toggles
        filtered_df = filter_companies(
            df=dedup_df,
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

        # 6) Build final sheets
        excluded_df = filtered_df[filtered_df["Excluded"]==True].copy()
        retained_df = filtered_df[filtered_df["Excluded"]==False].copy()

        # "No Data" if missing 'Coal Industry Sector'
        if "Coal Industry Sector" in filtered_df.columns:
            no_data_df = filtered_df[filtered_df["Coal Industry Sector"].isna()].copy()
        else:
            no_data_df = pd.DataFrame()

        # 7) Reorder columns if you want a specific set. For example:
        # We'll just keep them all for demonstration. If you want a specific set, define them:
        final_cols = list(filtered_df.columns)
        # If you only want certain columns in a certain order => define them here and reindex each DF.

        # Output
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df[final_cols].to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df[final_cols].to_excel(writer, sheet_name="Retained Companies", index=False)
            no_data_df[final_cols].to_excel(writer, sheet_name="No Data Companies", index=False)

        st.subheader("Statistics")
        st.write(f"Total (after dedup): {len(filtered_df)}")
        st.write(f"Excluded: {len(excluded_df)}")
        st.write(f"Retained: {len(retained_df)}")
        st.write(f"No-data: {len(no_data_df)}")

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
