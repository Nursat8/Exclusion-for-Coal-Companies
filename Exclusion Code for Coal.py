import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

################################################
# MAKE COLUMNS UNIQUE
################################################
def make_columns_unique(df):
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
# AUTO-DETECT MULTI-HEADER FOR SPGLOBAL
################################################
def load_spglobal_autodetect(file, sheet_name):
    """
    1) Read the entire sheet with header=None so Pandas doesn't guess columns.
    2) Search for row containing 'SP_ENTITY_NAME'.
    3) That row + the next row become your 2-level headers. Then flatten.
    4) The subsequent rows become data.
    Adjust if your actual layout differs (maybe skip more blank lines).
    """
    try:
        full_df = pd.read_excel(file, sheet_name=sheet_name, header=None, dtype=str)

        # Find the row index that includes 'SP_ENTITY_NAME'
        row_index = None
        for i in range(len(full_df)):
            row_vals = [str(x).strip().lower() for x in full_df.iloc[i].tolist() if pd.notnull(x)]
            if any("sp_entity_name" in val for val in row_vals):
                row_index = i
                break

        if row_index is None:
            st.error("Could not find a row containing 'SP_ENTITY_NAME' in SPGlobal sheet.")
            return None

        # We'll treat row_index as top-level header, row_index+1 as second-level
        top_row = row_index
        second_row = row_index + 1
        header_rows = full_df.iloc[[top_row, second_row]].fillna("")

        # Build multi-level column tuples
        col_tuples = []
        for j in range(full_df.shape[1]):
            c1 = str(header_rows.iloc[0,j]).strip()
            c2 = str(header_rows.iloc[1,j]).strip()
            col_tuples.append((c1, c2))

        # Data starts from row (second_row+1)
        data_start = second_row + 1
        data_df = full_df.iloc[data_start:].reset_index(drop=True)

        # Assign the multi-index, then flatten
        multi_index = pd.MultiIndex.from_tuples(col_tuples)
        data_df.columns = multi_index
        data_df.columns = [
            " ".join(str(x).strip() for x in col if x not in (None, ""))
            for col in data_df.columns
        ]
        data_df = make_columns_unique(data_df)
        return data_df

    except Exception as e:
        st.error(f"Error auto-detecting SPGlobal multi-header: {e}")
        return None

################################################
# LOAD URGEWALD (SINGLE HEADER=0)
################################################
def load_urgewald_data(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name, header=0, dtype=str)
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
        st.error(f"Error loading Urgewald file: {e}")
        return None

################################################
# REMOVE DUPLICATES (OR LOGIC)
################################################
def remove_duplicates_or(df):
    """
    Remove duplicates if ANY match (case-insensitive):
    (SP_ENTITY_NAME vs Company) or (SP_ISIN vs ISIN equity) or (SP_LEI vs LEI).
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

    df.drop(columns=["_key_name_","_key_isin_","_key_lei_"], inplace=True)
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
################################################
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
    exclusion_flags = []
    exclusion_reasons = []

    for idx, row in df.iterrows():
        reasons = []
        sector_val = str(row.get("Coal Industry Sector","")).lower()
        is_mining   = ("mining" in sector_val)
        is_power    = ("power" in sector_val) or ("generation" in sector_val)
        is_services = ("services" in sector_val)

        expansion_text = str(row.get("expansion","")).lower()
        prod_val = str(row.get(">10MT / >5GW","")).lower()

        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue",0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production",0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)",0), errors="coerce") or 0.0

        # MINING
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and ">10mt" in prod_val:
                reasons.append(f"Mining production >10MT vs {mining_prod_mt_threshold}MT")

        # POWER
        if is_power and exclude_power:
            if exclude_power_rev and (coal_rev * 100) > power_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {power_rev_threshold}% (Power)")

            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")

            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed coal power capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")

        # SERVICES
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal share of revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")

        # expansions
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
    st.set_page_config(page_title="Coal Exclusion Filter (Rename Columns)", layout="wide")
    st.title("Coal Exclusion Filter with Multi-Header, Column Renames, Duplicates, etc.")

    # FILES
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    # TOGGLES
    st.sidebar.header("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Mining: Max production (MT)", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Exclude if > MT?", value=True)

    st.sidebar.header("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power?", value=True)
    power_rev_threshold = st.sidebar.number_input("Power: Max coal revenue (%)", value=20.0)
    exclude_power_rev = st.sidebar.checkbox("Exclude if power rev threshold exceeded?", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Power: Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Exclude if power prod % exceeded?", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Power: Max installed capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Exclude if capacity exceeded?", value=True)

    st.sidebar.header("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services?", value=False)
    services_rev_threshold = st.sidebar.number_input("Services: Max coal revenue (%)", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Exclude if services rev threshold exceeded?", value=False)

    # expansions
    st.sidebar.header("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect(
        "Exclude if expansions text contains any of these",
        expansions_possible,
        default=[]
    )

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SPGlobal with auto-detect
        sp_df = load_spglobal_autodetect(sp_file, sp_sheet)
        if sp_df is None:
            return
        st.write("SP columns =>", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        # 2) RENAME columns for simpler names, e.g. removing "SP_ESG_BUS_INVOLVE_REV_PCT "
        rename_map_sp = {
            # e.g. "SP_ESG_BUS_INVOLVE_REV_PCT Generation (Thermal Coal)" => "Generation (Thermal Coal)"
            "SP_ESG_BUS_INVOLVE_REV_PCT Generation (Thermal Coal)":     "Generation (Thermal Coal)",
            "SP_ESG_BUS_INVOLVE_REV_PCT Thermal Coal Mining":           "Thermal Coal Mining",
            "SP_ESG_BUS_INVOLVE_REV_PCT Metallurgical Coal Mining":     "Metallurgical Coal Mining",
            # Add more if needed (like "SP_ESG_BUS_INVOLVE_REV_PCT Tobacco" => "Tobacco", etc.)
        }
        for old_col, new_col in rename_map_sp.items():
            if old_col in sp_df.columns:
                sp_df.rename(columns={old_col: new_col}, inplace=True)

        st.write("SP columns after rename =>", sp_df.columns.tolist())

        # 3) Load Urgewald
        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df is None:
            return
        st.write("Urgewald columns =>", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # 4) Concatenate
        combined = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write(f"Combined shape => {combined.shape}")
        st.write("Combined columns =>", combined.columns.tolist())

        # 5) Remove duplicates (OR logic)
        deduped = remove_duplicates_or(combined.copy())
        st.write(f"After removing duplicates => {deduped.shape}")

        # 6) Filter
        filtered = filter_companies(
            df=deduped,
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

        excluded_df = filtered[filtered["Excluded"]==True].copy()
        retained_df = filtered[filtered["Excluded"]==False].copy()
        if "Coal Industry Sector" in filtered.columns:
            no_data_df = filtered[filtered["Coal Industry Sector"].isna()].copy()
        else:
            no_data_df = pd.DataFrame()

        # 7) Final columns for output
        final_cols = [
            "SP_ENTITY_NAME",
            "SP_ENTITY_ID",
            "SP_COMPANY_ID",
            "SP_ISIN",
            "SP_LEI",
            "BB Ticker",           # if it exists
            "ISIN equity",
            "LEI",
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "expansion",
            "Generation (Thermal Coal)",    # newly renamed
            "Thermal Coal Mining",          # newly renamed
            "Metallurgical Coal Mining",    # newly renamed
            "Exclusion Reasons"
        ]
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

        # 8) Output Excel
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
