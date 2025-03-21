import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from collections import deque

################################################
# MAKE COLUMNS UNIQUE
################################################
def make_columns_unique(df):
    """Append _1, _2, etc. to duplicate column names to avoid PyArrow issues."""
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
# REORDER COLUMNS FOR FINAL EXCEL
# Force:
#   "Company" in column G (7th),
#   "BB Ticker" in column AP (42nd),
#   "ISIN equity" in column AQ (43rd),
#   "LEI" in column AT (46th).
# Then move "Excluded" and "Exclusion Reasons" to the end.
################################################
def reorder_for_excel(df):
    desired_length = 46  # columns A..AT => 1..46
    placeholders = ["(placeholder)"] * desired_length

    # Force required columns
    placeholders[6]   = "Company"      # G => index=6
    placeholders[41]  = "BB Ticker"    # AP => index=41
    placeholders[42]  = "ISIN equity"  # AQ => index=42
    placeholders[45]  = "LEI"          # AT => index=45

    forced_positions = {6, 41, 42, 45}
    forced_cols = {"Company", "BB Ticker", "ISIN equity", "LEI"}

    all_cols = list(df.columns)
    remaining_cols = [c for c in all_cols if c not in forced_cols]

    idx_remain = 0
    for i in range(desired_length):
        if i not in forced_positions:
            if idx_remain < len(remaining_cols):
                placeholders[i] = remaining_cols[idx_remain]
                idx_remain += 1

    leftover = remaining_cols[idx_remain:]
    final_col_order = placeholders + leftover

    # Create empty columns for placeholders if needed
    for c in final_col_order:
        if c not in df.columns and c == "(placeholder)":
            df[c] = np.nan

    # Reorder
    df = df[final_col_order]
    # Drop placeholder columns that are entirely empty
    df = df.loc[:, ~((df.columns == "(placeholder)") & (df.isna().all()))]

    # Move "Excluded" & "Exclusion Reasons" to the very end if present
    cols = list(df.columns)
    for c in ["Excluded", "Exclusion Reasons"]:
        if c in cols:
            cols.remove(c)
            cols.append(c)
    df = df[cols]

    return df

################################################
# LOAD SPGLOBAL (MULTI-HEADER)
################################################
def load_spglobal_autodetect(file, sheet_name):
    """Load SPGlobal file with row5=IDs, row6=metrics, data from row7 onward."""
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if len(full_df) < 6:
            raise ValueError("Not enough rows for multi-header logic in SPGlobal.")
        row5 = full_df.iloc[4].fillna("")
        row6 = full_df.iloc[5].fillna("")
        final_cols = []
        for col_idx in range(full_df.shape[1]):
            top_val = str(row5[col_idx]).strip()
            bot_val = str(row6[col_idx]).strip()
            combined = top_val if top_val else ""
            if bot_val and bot_val.lower() not in combined.lower():
                combined = (combined + " " + bot_val).strip() if combined else bot_val
            final_cols.append(combined.strip())
        sp_df = full_df.iloc[6:].reset_index(drop=True)
        sp_df.columns = final_cols
        sp_df = make_columns_unique(sp_df)
        # Optional rename
        rename_map_sp = {
            "SP_ESG_BUS_INVOLVE_REV_PCT Generation (Thermal Coal)": "Generation (Thermal Coal)",
            "SP_ESG_BUS_INVOLVE_REV_PCT Thermal Coal Mining":       "Thermal Coal Mining",
            "SP_ESG_BUS_INVOLVE_REV_PCT Metallurgical Coal Mining": "Metallurgical Coal Mining",
        }
        for old_col, new_col in rename_map_sp.items():
            if old_col in sp_df.columns:
                sp_df.rename(columns={old_col: new_col}, inplace=True)
        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal: {e}")
        return pd.DataFrame()

################################################
# LOAD URGEWALD (SINGLE HEADER)
################################################
def load_urgewald_data(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald file is empty.")
        full_df = pd.DataFrame(data)
        header = full_df.iloc[0].fillna("")
        ur_df = full_df.iloc[1:].reset_index(drop=True)
        ur_df.columns = header
        ur_df = make_columns_unique(ur_df)
        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald file: {e}")
        return pd.DataFrame()

################################################
# UNIFY FUNCTIONS (NAME, ISIN, LEI)
################################################
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
# MERGE URGEWALD INTO SPGLOBAL
#   If row matches SP by (name OR isin OR lei), merge.
#   Otherwise, keep it as UR-only row.
################################################
def merge_ur_into_sp(sp_df, ur_df):
    sp_records = sp_df.copy().to_dict('records')
    merged_records = []
    ur_only_records = []
    for rec in sp_records:
        rec['Source'] = "SP"
        merged_records.append(rec)

    for _, ur_row in ur_df.iterrows():
        merged_flag = False
        ur_name = unify_name(ur_row)
        ur_isin = unify_isin(ur_row)
        ur_lei  = unify_lei(ur_row)

        for rec in merged_records:
            sp_name = unify_name(rec)
            sp_isin = unify_isin(rec)
            sp_lei  = unify_lei(rec)

            if ((sp_name and ur_name and sp_name == ur_name) or
                (sp_isin and ur_isin and sp_isin == ur_isin) or
                (sp_lei and ur_lei and sp_lei == ur_lei)):
                # Merge UR row's non-empty fields
                for k, v in ur_row.items():
                    if (k not in rec) or (rec[k] is None) or (str(rec[k]).strip() == ""):
                        rec[k] = v
                rec['Source'] = "SP+UR"
                merged_flag = True
                break

        if not merged_flag:
            new_rec = ur_row.to_dict()
            new_rec['Source'] = "UR"
            ur_only_records.append(new_rec)

    merged_df = pd.DataFrame(merged_records)
    ur_only_df = pd.DataFrame(ur_only_records)
    # Remove 'Source' column if not needed
    merged_df.drop(columns=['Source'], inplace=True, errors='ignore')
    ur_only_df.drop(columns=['Source'], inplace=True, errors='ignore')
    return merged_df, ur_only_df

################################################
# FILTER COMPANIES
#   Checks "sector" to apply logic for Mining/Power/Services.
#   Checks >10MT or >5GW from column ">10MT / >5GW".
################################################
def filter_companies(
    df,
    # Mining thresholds
    exclude_mining,
    mining_coal_rev_threshold,
    exclude_mining_prod_mt,
    mining_prod_mt_threshold,
    exclude_mining_prod_gw,
    mining_prod_threshold_GW,
    exclude_thermal_coal_mining,
    thermal_coal_mining_threshold,
    exclude_metallurgical_coal_mining,
    metallurgical_coal_mining_threshold,
    # Power thresholds
    exclude_power,
    power_coal_rev_threshold,
    exclude_power_prod_percent,
    power_prod_threshold_percent,
    exclude_capacity_mw,
    capacity_threshold_mw,
    exclude_generation_thermal,
    generation_thermal_threshold,
    # Services thresholds
    exclude_services,
    services_rev_threshold,
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
        is_services = ("service" in sector_val)

        expansion_text = str(row.get("expansion","")).lower()

        # numeric columns
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0

        # business involvement columns
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # check production string
        prod_str = str(row.get(">10MT / >5GW", "")).lower()

        # annual_coal_prod is not used in this version unless you want to add it
        # (But if you do, you can parse "Annual Coal Production (in million metric tons)".)

        ####### MINING #######
        if is_mining and exclude_mining:
            # check coal rev
            if (coal_rev * 100) > mining_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")

            # check >10MT / >5GW
            if exclude_mining_prod_mt:
                if ">10mt" in prod_str:  # means it's definitely above 10MT
                    if mining_prod_mt_threshold <= 10:
                        reasons.append(f"Production indicated as >10MT (threshold <= {mining_prod_mt_threshold}MT)")

            if exclude_mining_prod_gw:
                if ">5gw" in prod_str:  # means definitely above 5GW
                    if mining_prod_threshold_GW <= 5:
                        reasons.append(f"Production indicated as >5GW (threshold <= {mining_prod_threshold_GW}GW)")

            # check Thermal Coal Mining
            if exclude_thermal_coal_mining and (therm_mining_val > thermal_coal_mining_threshold):
                reasons.append(f"Thermal Coal Mining {therm_mining_val:.2f}% > {thermal_coal_mining_threshold}%")

            # check Metallurgical Coal Mining
            if exclude_metallurgical_coal_mining and (met_coal_val > metallurgical_coal_mining_threshold):
                reasons.append(f"Metallurgical Coal Mining {met_coal_val:.2f}% > {metallurgical_coal_mining_threshold}%")

        ####### POWER #######
        if is_power and exclude_power:
            # check coal rev
            if (coal_rev * 100) > power_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
            # check power production
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")
            # check installed capacity
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            # check Generation (Thermal Coal)
            if exclude_generation_thermal and (gen_thermal_val > generation_thermal_threshold):
                reasons.append(f"Generation (Thermal Coal) {gen_thermal_val:.2f}% > {generation_thermal_threshold}%")

        ####### SERVICES #######
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")

        ####### expansions #######
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
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter with >10MT / >5GW Logic & Improved Interface")

    # FILE & SHEET
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    st.sidebar.markdown("---")

    #### MINING ####
    with st.sidebar.expander("Mining Thresholds", expanded=True):
        exclude_mining = st.checkbox("Exclude Mining Sector?", value=True)
        mining_coal_rev_threshold = st.number_input("Mining: Max coal revenue (%)", value=15.0)
        exclude_mining_prod_mt = st.checkbox("Exclude if >10MT indicated?", value=True)
        mining_prod_mt_threshold = st.number_input("Mining: Max production (MT)", value=10.0)
        exclude_mining_prod_gw = st.checkbox("Exclude if >5GW indicated?", value=True)
        mining_prod_threshold_GW = st.number_input("Mining: Max production (GW)", value=5.0)
        exclude_thermal_coal_mining = st.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False)
        thermal_coal_mining_threshold = st.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)
        exclude_metallurgical_coal_mining = st.checkbox("Exclude if Metallurgical Coal Mining > threshold?", value=False)
        metallurgical_coal_mining_threshold = st.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    #### POWER ####
    with st.sidebar.expander("Power Thresholds", expanded=True):
        exclude_power = st.checkbox("Exclude Power Sector?", value=True)
        power_coal_rev_threshold = st.number_input("Power: Max coal revenue (%)", value=20.0)
        exclude_power_prod_percent = st.checkbox("Exclude if coal power production > threshold?", value=True)
        power_prod_threshold_percent = st.number_input("Max coal power production (%)", value=20.0)
        exclude_capacity_mw = st.checkbox("Exclude if installed capacity > threshold?", value=True)
        capacity_threshold_mw = st.number_input("Max installed capacity (MW)", value=10000.0)
        exclude_generation_thermal = st.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
        generation_thermal_threshold = st.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    #### SERVICES ####
    with st.sidebar.expander("Services Thresholds", expanded=False):
        exclude_services = st.checkbox("Exclude Services Sector?", value=False)
        services_rev_threshold = st.number_input("Services: Max coal revenue (%)", value=10.0)
        exclude_services_rev = st.checkbox("Exclude if services revenue > threshold?", value=False)

    #### EXPANSIONS ####
    with st.sidebar.expander("Global Expansion Exclusion", expanded=False):
        expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
        expansions_global = st.multiselect(
            "Exclude if expansion text contains any of these",
            expansions_possible,
            default=[]
        )

    st.sidebar.markdown("---")

    # RUN
    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load
        sp_df = load_spglobal_autodetect(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or could not be loaded.")
            return
        st.subheader("SPGlobal Data (first 5 rows)")
        st.dataframe(sp_df.head(5))

        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or could not be loaded.")
            return
        st.subheader("Urgewald Data (first 5 rows)")
        st.dataframe(ur_df.head(5))

        # 2) Merge
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)
        st.write(f"Merged dataset shape: {merged_df.shape}")
        st.write(f"UR-only dataset shape: {ur_only_df.shape}")

        # 3) Filter
        filtered_merged = filter_companies(
            df=merged_df,
            # Mining
            exclude_mining=exclude_mining,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_GW=mining_prod_threshold_GW,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            # Power
            exclude_power=exclude_power,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services
            exclude_services=exclude_services,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # expansions
            expansions_global=expansions_global
        )

        filtered_ur_only = filter_companies(
            df=ur_only_df,
            # Mining
            exclude_mining=exclude_mining,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_GW=mining_prod_threshold_GW,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            # Power
            exclude_power=exclude_power,
            power_coal_rev_threshold=power_coal_rev_threshold,
            exclude_power_prod_percent=exclude_power_prod_percent,
            power_prod_threshold_percent=power_prod_threshold_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_generation_thermal=exclude_generation_thermal,
            generation_thermal_threshold=generation_thermal_threshold,
            # Services
            exclude_services=exclude_services,
            services_rev_threshold=services_rev_threshold,
            exclude_services_rev=exclude_services_rev,
            # expansions
            expansions_global=expansions_global
        )

        # 4) Separate Excluded / Retained
        excluded_df = filtered_merged[filtered_merged["Excluded"]==True].copy()
        retained_df = filtered_merged[filtered_merged["Excluded"]==False].copy()

        # 5) Final columns
        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Company","ISIN equity","LEI","BB Ticker",
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "Coal Share of Revenue",
            "expansion",
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            "Excluded","Exclusion Reasons"
        ]
        def ensure_cols(df_):
            for c in final_cols:
                if c not in df_.columns:
                    df_[c] = np.nan
            return df_

        excluded_df = ensure_cols(excluded_df)[final_cols]
        retained_df = ensure_cols(retained_df)[final_cols]
        filtered_ur_only = ensure_cols(filtered_ur_only)[final_cols]

        # 6) Reorder so that "Company"=G, "BB Ticker"=AP, "ISIN equity"=AQ, "LEI"=AT, then Exclusion at the end
        excluded_df = reorder_for_excel(excluded_df)
        retained_df = reorder_for_excel(retained_df)
        filtered_ur_only = reorder_for_excel(filtered_ur_only)

        # 7) Output
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, "Excluded Companies", index=False)
            retained_df.to_excel(writer, "Retained Companies", index=False)
            filtered_ur_only.to_excel(writer, "Urgewald Only", index=False)

        st.subheader("Results Summary")
        st.write(f"Merged total: {len(filtered_merged)}")
        st.write(f"Excluded (merged): {len(excluded_df)}")
        st.write(f"Retained (merged): {len(retained_df)}")
        st.write(f"UR-only: {len(filtered_ur_only)}")

        st.download_button(
            label="Download Filtered Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
