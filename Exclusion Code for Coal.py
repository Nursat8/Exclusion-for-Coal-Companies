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
    """Append _1, _2, etc. to duplicate column names."""
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
#   "Company" in col G (7th),
#   "BB Ticker" in col AP (42nd),
#   "ISIN equity" in col AQ (43rd),
#   "LEI" in col AT (46th),
# and then move "Excluded" and "Exclusion Reasons" to the very end.
################################################
def reorder_for_excel(df):
    desired_length = 46  # Force positions for columns A..AT
    placeholders = ["(placeholder)"] * desired_length

    # Force required columns at fixed positions (0-indexed)
    placeholders[6]   = "Company"      # Column G (7th)
    placeholders[41]  = "BB Ticker"    # Column AP (42nd)
    placeholders[42]  = "ISIN equity"  # Column AQ (43rd)
    placeholders[45]  = "LEI"          # Column AT (46th)

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

    # For any placeholder not in df, create an empty column.
    for c in final_col_order:
        if c not in df.columns and c == "(placeholder)":
            df[c] = np.nan

    df = df[final_col_order]
    # Drop any placeholder columns that are completely empty.
    df = df.loc[:, ~((df.columns == "(placeholder)") & (df.isna().all()))]
    # Finally, move "Excluded" and "Exclusion Reasons" to the end.
    cols = list(df.columns)
    if "Excluded" in cols and "Exclusion Reasons" in cols:
        cols = [c for c in cols if c not in ["Excluded", "Exclusion Reasons"]] + ["Excluded", "Exclusion Reasons"]
        df = df[cols]
    return df

################################################
# LOAD SPGLOBAL WITH AUTO-DETECTION OF MULTI-HEADER
# Row 5 (index=4): ID columns (e.g., SP_ENTITY_NAME)
# Row 6 (index=5): Additional fields (e.g., Generation (Thermal Coal))
# Data starts at row 7 (index=6)
################################################
def load_spglobal_autodetect(file, sheet_name):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)
        if len(full_df) < 6:
            raise ValueError("Not enough rows in SPGlobal file.")
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
        # Rename to simplify headers if needed.
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
# LOAD URGEWALD (Single header in row 1)
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
        st.error(f"Error loading Urgewald: {e}")
        return pd.DataFrame()

################################################
# UNIFY FUNCTIONS FOR MATCHING
################################################
def unify_name(r):
    sp_name = str(r.get("SP_ENTITY_NAME", "")).strip().lower()
    ur_name = str(r.get("Company", "")).strip().lower()
    return sp_name if sp_name else (ur_name if ur_name else None)

def unify_isin(r):
    sp_isin = str(r.get("SP_ISIN", "")).strip().lower()
    ur_isin = str(r.get("ISIN equity", "")).strip().lower()
    return sp_isin if sp_isin else (ur_isin if ur_isin else None)

def unify_lei(r):
    sp_lei = str(r.get("SP_LEI", "")).strip().lower()
    ur_lei = str(r.get("LEI", "")).strip().lower()
    return sp_lei if sp_lei else (ur_lei if ur_lei else None)

################################################
# MERGE URGEWALD INTO SPGLOBAL
# For each UR row, if it matches (by OR logic on name/ISIN/LEI)
# any SPGlobal row, merge non‑empty values into that SP record.
# Otherwise, keep the UR row separately.
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
        for rec in merged_records:
            if ((unify_name(rec) and unify_name(ur_row) and unify_name(rec) == unify_name(ur_row)) or
                (unify_isin(rec) and unify_isin(ur_row) and unify_isin(rec) == unify_isin(ur_row)) or
                (unify_lei(rec) and unify_lei(ur_row) and unify_lei(rec) == unify_lei(ur_row))):
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
    merged_df.drop(columns=['Source'], inplace=True, errors='ignore')
    ur_only_df.drop(columns=['Source'], inplace=True, errors='ignore')
    return merged_df, ur_only_df

################################################
# FILTER COMPANIES (Exclusion Logic with Sector-Specific Thresholds)
################################################
def filter_companies(
    df,
    # Mining thresholds:
    mining_prod_mt_threshold,      # Numeric threshold for annual production (MT)
    mining_coal_rev_threshold,       # Max allowed coal revenue (%) for mining
    exclude_mining,
    exclude_mining_prod_mt,
    exclude_mining_prod_gw,          # Checkbox for GW check
    mining_prod_threshold_GW,        # Numeric threshold (GW) for mining production
    # Power thresholds:
    power_coal_rev_threshold,        # Max allowed coal revenue (%) for power
    power_prod_threshold_percent,    # Max allowed coal power production (%)
    capacity_threshold_mw,           # Installed capacity (MW)
    exclude_power,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    # Services thresholds:
    services_rev_threshold,
    exclude_services,
    exclude_services_rev,
    # Business involvement thresholds:
    generation_thermal_threshold,    # For power: Generation (Thermal Coal) (%)
    exclude_generation_thermal,
    thermal_coal_mining_threshold,   # For mining: Thermal Coal Mining (%)
    exclude_thermal_coal_mining,
    metallurgical_coal_mining_threshold,  # For mining: Metallurgical Coal Mining (%)
    exclude_metallurgical_coal_mining,
    # Global expansion:
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
        expansion_text = str(row.get("expansion", "")).lower()

        # Numeric values
        coal_rev = pd.to_numeric(row.get("Coal Share of Revenue", 0), errors="coerce") or 0.0
        coal_power_share = pd.to_numeric(row.get("Coal Share of Power Production", 0), errors="coerce") or 0.0
        installed_cap = pd.to_numeric(row.get("Installed Coal Power Capacity (MW)", 0), errors="coerce") or 0.0
        annual_coal_prod = pd.to_numeric(row.get("Annual Coal Production (in million metric tons)", 0), errors="coerce") or 0.0
        # Business involvement values (already percentages, do not multiply)
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0
        # Production string from ">10MT / >5GW" column
        prod_str = str(row.get(">10MT / >5GW", "")).upper()

        # --- MINING checks ---
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt:
                if annual_coal_prod > mining_prod_mt_threshold:
                    reasons.append(f"Annual production {annual_coal_prod:.2f}MT > {mining_prod_mt_threshold}MT")
                # Check text indicator for >10MT:
                if "10MT" in prod_str:  # Accepts ">10MT" or "10MT"
                    if mining_prod_mt_threshold < 10:
                        reasons.append(f"Production indicator '>10MT' found (allowed threshold {mining_prod_mt_threshold}MT)")
            if exclude_mining_prod_gw:
                # Check text indicator for >5GW:
                if "5GW" in prod_str:  # Accepts ">5GW" or "5GW"
                    if mining_prod_threshold_GW < 5:
                        reasons.append(f"Production indicator '>5GW' found (allowed threshold {mining_prod_threshold_GW}GW)")
            if (coal_rev * 100) > mining_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {mining_coal_rev_threshold}% (Mining)")
            if exclude_thermal_coal_mining and (therm_mining_val) > thermal_coal_mining_threshold:
                reasons.append(f"Thermal Coal Mining {therm_mining_val:.2f}% > {thermal_coal_mining_threshold}% (Mining)")
            if exclude_metallurgical_coal_mining and (met_coal_val) > metallurgical_coal_mining_threshold:
                reasons.append(f"Metallurgical Coal Mining {met_coal_val:.2f}% > {metallurgical_coal_mining_threshold}% (Mining)")

        # --- POWER checks ---
        if is_power and exclude_power:
            if (coal_rev * 100) > power_coal_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {power_coal_rev_threshold}% (Power)")
            if exclude_power_prod_percent and (coal_power_share * 100) > power_prod_threshold_percent:
                reasons.append(f"Coal power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%")
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(f"Installed capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW")
            if exclude_generation_thermal and (gen_thermal_val) > generation_thermal_threshold:
                reasons.append(f"Generation (Thermal Coal) {gen_thermal_val:.2f}% > {generation_thermal_threshold}% (Power)")

        # --- SERVICES checks ---
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev * 100) > services_rev_threshold:
                reasons.append(f"Coal revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)")

        # --- Global expansion ---
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
    st.set_page_config(page_title="Coal Exclusion Filter (Merged SP & Urgewald)", layout="wide")
    st.title("Coal Exclusion Filter: Merged SPGlobal & Urgewald with Sector-Specific Production Thresholds")

    # --- File & Sheet Settings ---
    st.sidebar.subheader("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    # --- Mining Thresholds ---
    st.sidebar.subheader("Mining Thresholds")
    exclude_mining = st.sidebar.checkbox("Exclude Mining", value=True)
    mining_prod_mt_threshold = st.sidebar.number_input("Max production (MT) for Mining", value=10.0)
    exclude_mining_prod_mt = st.sidebar.checkbox("Apply MT production threshold?", value=True)
    exclude_mining_prod_gw = st.sidebar.checkbox("Apply GW production threshold?", value=True)
    mining_prod_threshold_GW = st.sidebar.number_input("Max production (GW) for Mining", value=5.0)
    mining_coal_rev_threshold = st.sidebar.number_input("Max coal revenue (%) for Mining", value=15.0)

    # --- Power Thresholds ---
    st.sidebar.subheader("Power Thresholds")
    exclude_power = st.sidebar.checkbox("Exclude Power", value=True)
    power_prod_threshold_percent = st.sidebar.number_input("Max coal power production (%)", value=20.0)
    exclude_power_prod_percent = st.sidebar.checkbox("Apply power production threshold?", value=True)
    capacity_threshold_mw = st.sidebar.number_input("Max installed capacity (MW)", value=10000.0)
    exclude_capacity_mw = st.sidebar.checkbox("Apply capacity threshold?", value=True)
    power_coal_rev_threshold = st.sidebar.number_input("Max coal revenue (%) for Power", value=20.0)

    # --- Services Thresholds ---
    st.sidebar.subheader("Services Thresholds")
    exclude_services = st.sidebar.checkbox("Exclude Services", value=False)
    services_rev_threshold = st.sidebar.number_input("Max coal revenue (%) for Services", value=10.0)
    exclude_services_rev = st.sidebar.checkbox("Apply services revenue threshold?", value=False)

    # --- Business Involvement (Sector-specific) ---
    st.sidebar.subheader("Business Involvement Thresholds (%)")
    # For Power:
    exclude_generation_thermal = st.sidebar.checkbox("Exclude Power if Generation (Thermal Coal) > threshold", value=False)
    generation_thermal_threshold = st.sidebar.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)
    # For Mining:
    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude Mining if Thermal Coal Mining > threshold", value=False)
    thermal_coal_mining_threshold = st.sidebar.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)
    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude Mining if Metallurgical Coal Mining > threshold", value=False)
    metallurgical_coal_mining_threshold = st.sidebar.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    # --- Global Expansion ---
    st.sidebar.subheader("Global Expansion Exclusion")
    expansions_possible = ["mining", "infrastructure", "power", "subsidiary of a coal developer"]
    expansions_global = st.sidebar.multiselect("Exclude if expansion text contains any of these", expansions_possible, default=[])

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # Load files
        sp_df = load_spglobal_autodetect(sp_file, sp_sheet)
        if sp_df.empty:
            st.warning("SPGlobal data is empty or could not be loaded.")
            return
        st.write("SPGlobal columns:", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df.empty:
            st.warning("Urgewald data is empty or could not be loaded.")
            return
        st.write("Urgewald columns:", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # Merge UR into SP; keep UR-only separately.
        merged_df, ur_only_df = merge_ur_into_sp(sp_df, ur_df)
        st.write(f"After merging: {merged_df.shape[0]} merged companies")
        st.dataframe(merged_df.head(5))
        st.write(f"UR-only companies: {ur_only_df.shape[0]}")
        st.dataframe(ur_only_df.head(5))

        # Apply filtering
        filtered_merged = filter_companies(
            df=merged_df,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining=exclude_mining,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_GW=mining_prod_threshold_GW,
            power_coal_rev_threshold=power_coal_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_power=exclude_power,
            exclude_power_prod_percent=exclude_power_prod_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            services_rev_threshold=services_rev_threshold,
            exclude_services=exclude_services,
            exclude_services_rev=exclude_services_rev,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_generation_thermal=exclude_generation_thermal,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            expansions_global=expansions_global
        )
        filtered_ur_only = filter_companies(
            df=ur_only_df,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            mining_coal_rev_threshold=mining_coal_rev_threshold,
            exclude_mining=exclude_mining,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            exclude_mining_prod_gw=exclude_mining_prod_gw,
            mining_prod_threshold_GW=mining_prod_threshold_GW,
            power_coal_rev_threshold=power_coal_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_power=exclude_power,
            exclude_power_prod_percent=exclude_power_prod_percent,
            exclude_capacity_mw=exclude_capacity_mw,
            services_rev_threshold=services_rev_threshold,
            exclude_services=exclude_services,
            exclude_services_rev=exclude_services_rev,
            generation_thermal_threshold=generation_thermal_threshold,
            exclude_generation_thermal=exclude_generation_thermal,
            thermal_coal_mining_threshold=thermal_coal_mining_threshold,
            exclude_thermal_coal_mining=exclude_thermal_coal_mining,
            metallurgical_coal_mining_threshold=metallurgical_coal_mining_threshold,
            exclude_metallurgical_coal_mining=exclude_metallurgical_coal_mining,
            expansions_global=expansions_global
        )

        # Separate merged set into Excluded and Retained.
        excluded_df = filtered_merged[filtered_merged["Excluded"] == True].copy()
        retained_df = filtered_merged[filtered_merged["Excluded"] == False].copy()

        # Define final columns.
        final_cols = [
            "SP_ENTITY_NAME", "SP_ENTITY_ID", "SP_COMPANY_ID", "SP_ISIN", "SP_LEI",
            "Company", "ISIN equity", "LEI", "BB Ticker",
            "Coal Industry Sector",
            ">10MT / >5GW",
            "Installed Coal Power Capacity (MW)",
            "Coal Share of Power Production",
            "Coal Share of Revenue",
            "Annual Coal Production (in million metric tons)",
            "expansion",
            "Generation (Thermal Coal)",
            "Thermal Coal Mining",
            "Metallurgical Coal Mining",
            "Excluded", "Exclusion Reasons"
        ]
        def ensure_cols_exist(df_):
            for c in final_cols:
                if c not in df_.columns:
                    df_[c] = np.nan
            return df_
        excluded_df = ensure_cols_exist(excluded_df)[final_cols]
        retained_df = ensure_cols_exist(retained_df)[final_cols]
        filtered_ur_only = ensure_cols_exist(filtered_ur_only)[final_cols]

        # Reorder columns so that:
        # "Company" -> G, "BB Ticker" -> AP, "ISIN equity" -> AQ, "LEE" -> AT,
        # and "Excluded" & "Exclusion Reasons" are at the end.
        excluded_df = reorder_for_excel(excluded_df)
        retained_df = reorder_for_excel(retained_df)
        filtered_ur_only = reorder_for_excel(filtered_ur_only)

        # Write to Excel with three sheets:
        # "Excluded Companies" (merged), "Retained Companies" (merged), "Urgewald Only" (no match)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, sheet_name="Excluded Companies", index=False)
            retained_df.to_excel(writer, sheet_name="Retained Companies", index=False)
            filtered_ur_only.to_excel(writer, sheet_name="Urgewald Only", index=False)

        st.subheader("Statistics")
        st.write(f"Total merged companies: {len(filtered_merged)}")
        st.write(f"Excluded (merged): {len(excluded_df)}")
        st.write(f"Retained (merged): {len(retained_df)}")
        st.write(f"Urgewald-only companies: {len(filtered_ur_only)}")

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
