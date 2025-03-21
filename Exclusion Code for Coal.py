import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from collections import defaultdict, deque

################################################
# MAKE COLUMNS UNIQUE
################################################
def make_columns_unique(df):
    """
    If there are duplicate column names, append _1, _2, etc. to avoid PyArrow errors.
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
# REORDER COLUMNS FOR FINAL EXCEL
# Place "Company" in col G(7), "BB Ticker" in AP(42),
# "ISIN equity" in AQ(43), "LEI" in AT(46),
# keep all other columns around/after them.
################################################
def reorder_for_excel(df):
    desired_length = 46  # columns A..AT => 1..46
    placeholders = ["(placeholder)"] * desired_length

    # Force our 4 columns
    placeholders[6]   = "Company"      # G => index=6 (0-based)
    placeholders[41]  = "BB Ticker"    # AP => index=41
    placeholders[42]  = "ISIN equity"  # AQ => index=42
    placeholders[45]  = "LEI"          # AT => index=45

    forced_positions = {6, 41, 42, 45}
    forced_cols = {"Company","BB Ticker","ISIN equity","LEI"}

    all_cols = list(df.columns)
    # remove forced columns from all_cols so we don't place them twice
    remaining_cols = [c for c in all_cols if c not in forced_cols]

    # fill placeholders from left to right with remaining_cols
    idx_remain = 0
    for i in range(desired_length):
        if i not in forced_positions:
            if idx_remain < len(remaining_cols):
                placeholders[i] = remaining_cols[idx_remain]
                idx_remain += 1

    # leftover columns (beyond AT) get appended
    leftover = remaining_cols[idx_remain:]
    final_col_order = placeholders + leftover

    # build the final DataFrame
    reordered_cols = []
    for c in final_col_order:
        if c in df.columns:
            reordered_cols.append(c)
        else:
            # placeholder => create empty column
            df[c] = np.nan
            reordered_cols.append(c)

    return df[reordered_cols]

################################################
# LOAD SPGLOBAL
################################################
def load_spglobal_autodetect(file, sheet_name):
    """
    Reads an SPGlobal file with:
      - row 5 => index=4 => e.g. SP_ENTITY_NAME, SP_ENTITY_ID...
      - row 6 => index=5 => e.g. Generation (Thermal Coal)...
      - data from row 7 => index=6 onward
    """
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)

        if len(full_df) < 6:
            raise ValueError("SPGlobal sheet doesn't have enough rows for multi-header logic.")

        row5 = full_df.iloc[4].fillna("")
        row6 = full_df.iloc[5].fillna("")
        final_cols = []
        for col_idx in range(full_df.shape[1]):
            top_val = str(row5[col_idx]).strip()
            bot_val = str(row6[col_idx]).strip()
            combined = top_val if top_val else ""
            if bot_val and bot_val.lower() not in combined.lower():
                if combined:
                    combined += " " + bot_val
                else:
                    combined = bot_val
            final_cols.append(combined.strip())

        sp_df = full_df.iloc[6:].reset_index(drop=True)
        sp_df.columns = final_cols
        sp_df = make_columns_unique(sp_df)

        # optionally rename columns
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
# LOAD URGEWALD
################################################
def load_urgewald_data(file, sheet_name="GCEL 2024"):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald sheet is empty.")

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
# MERGE DUPLICATES (OR logic) WITHOUT DROPPING
# We produce ONE row for each "connected component"
# i.e. any 2 rows that share unify_name OR unify_isin OR unify_lei
# go into the same group, and we coalesce columns.
################################################
def merge_duplicates_or(df):
    """
    1) unify_name = unify_name(r)
       unify_isin = unify_isin(r)
       unify_lei  = unify_lei(r)
    2) We'll build adjacency: rows that share a non-empty unify name or unify isin or unify lei
       belong to the same connected component.
    3) For each connected component, produce ONE row by coalescing columns (first non-null).
    4) Return new DataFrame with these merged rows.
    """
    n = len(df)
    if n == 0:
        return df

    # Extract unify keys
    unify_name_list = []
    unify_isin_list = []
    unify_lei_list  = []

    for i in range(n):
        row = df.iloc[i]
        unify_name_list.append( unify_name(row) )
        unify_isin_list.append( unify_isin(row) )
        unify_lei_list.append(  unify_lei(row)  )

    # Build dicts => key => list of row indices
    dict_name = defaultdict(list)
    dict_isin = defaultdict(list)
    dict_lei  = defaultdict(list)

    for i in range(n):
        if unify_name_list[i]:
            dict_name[unify_name_list[i]].append(i)
        if unify_isin_list[i]:
            dict_isin[unify_isin_list[i]].append(i)
        if unify_lei_list[i]:
            dict_lei[unify_lei_list[i]].append(i)

    # adjacency list
    adj = [[] for _ in range(n)]
    for i in range(n):
        name_key = unify_name_list[i]
        if name_key and name_key in dict_name:
            for j in dict_name[name_key]:
                if j != i:
                    adj[i].append(j)

        isin_key = unify_isin_list[i]
        if isin_key and isin_key in dict_isin:
            for j in dict_isin[isin_key]:
                if j != i:
                    adj[i].append(j)

        lei_key = unify_lei_list[i]
        if lei_key and lei_key in dict_lei:
            for j in dict_lei[lei_key]:
                if j != i:
                    adj[i].append(j)

    # find connected components using BFS/DFS
    visited = [False]*n
    components = []
    for start_idx in range(n):
        if not visited[start_idx]:
            # BFS or DFS
            queue = deque([start_idx])
            visited[start_idx] = True
            group = [start_idx]
            while queue:
                cur = queue.popleft()
                for nxt in adj[cur]:
                    if not visited[nxt]:
                        visited[nxt] = True
                        queue.append(nxt)
                        group.append(nxt)
            components.append(group)

    # Coalesce columns for each component
    merged_rows = []
    all_columns = df.columns
    for comp in components:
        # gather all row data
        comp_data = df.iloc[comp]  # subset of rows
        # coalesce by taking the first non-null from left to right
        merged_dict = {}
        for c in all_columns:
            val = None
            for _, subrow in comp_data.iterrows():
                if pd.notnull(subrow.get(c)) and str(subrow.get(c)).strip() != "":
                    val = subrow.get(c)
                    break
            merged_dict[c] = val
        merged_rows.append(merged_dict)

    merged_df = pd.DataFrame(merged_rows, columns=all_columns)
    return merged_df

def unify_name(r):
    sp_name = str(r.get("SP_ENTITY_NAME","")).strip().lower()
    ur_name = str(r.get("Company","")).strip().lower()
    # If either is non-empty, that's the unify name
    # but for "OR" we want either one => we pick sp_name if it exists, else ur_name
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
# FILTER COMPANIES (same thresholds etc. as before)
################################################
def filter_companies(
    df,
    # Mining thresholds
    mining_prod_mt_threshold,
    exclude_mining,
    exclude_mining_prod_mt,
    # Power thresholds
    power_rev_threshold,
    power_prod_threshold_percent,
    capacity_threshold_mw,
    exclude_power,
    exclude_power_rev,
    exclude_power_prod_percent,
    exclude_capacity_mw,
    # Services thresholds
    services_rev_threshold,
    exclude_services,
    exclude_services_rev,
    # Additional coal involvement thresholds
    generation_thermal_threshold,
    exclude_generation_thermal,
    thermal_coal_mining_threshold,
    exclude_thermal_coal_mining,
    metallurgical_coal_mining_threshold,
    exclude_metallurgical_coal_mining,
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
        annual_coal_prod = pd.to_numeric(row.get("Annual Coal Production (in million metric tons)", 0), errors="coerce") or 0.0

        # 3 columns
        gen_thermal_val = pd.to_numeric(row.get("Generation (Thermal Coal)", 0), errors="coerce") or 0.0
        therm_mining_val = pd.to_numeric(row.get("Thermal Coal Mining", 0), errors="coerce") or 0.0
        met_coal_val = pd.to_numeric(row.get("Metallurgical Coal Mining", 0), errors="coerce") or 0.0

        # MINING
        if is_mining and exclude_mining:
            if exclude_mining_prod_mt and (annual_coal_prod > mining_prod_mt_threshold):
                reasons.append(
                    f"Annual coal production {annual_coal_prod:.2f}MT > {mining_prod_mt_threshold}MT"
                )

        # POWER
        if is_power and exclude_power:
            if exclude_power_rev and (coal_rev*100) > power_rev_threshold:
                reasons.append(
                    f"Coal share of revenue {coal_rev*100:.2f}% > {power_rev_threshold}% (Power)"
                )
            if exclude_power_prod_percent and (coal_power_share*100) > power_prod_threshold_percent:
                reasons.append(
                    f"Coal share of power production {coal_power_share*100:.2f}% > {power_prod_threshold_percent}%"
                )
            if exclude_capacity_mw and (installed_cap > capacity_threshold_mw):
                reasons.append(
                    f"Installed coal power capacity {installed_cap:.2f}MW > {capacity_threshold_mw}MW"
                )

        # SERVICES
        if is_services and exclude_services:
            if exclude_services_rev and (coal_rev*100) > services_rev_threshold:
                reasons.append(
                    f"Coal share of revenue {coal_rev*100:.2f}% > {services_rev_threshold}% (Services)"
                )

        # Additional thresholds
        if exclude_generation_thermal and (gen_thermal_val*100) > generation_thermal_threshold:
            reasons.append(f"Generation (Thermal Coal) {gen_thermal_val*100:.2f}% > {generation_thermal_threshold}%")
        if exclude_thermal_coal_mining and (therm_mining_val*100) > thermal_coal_mining_threshold:
            reasons.append(f"Thermal Coal Mining {therm_mining_val*100:.2f}% > {thermal_coal_mining_threshold}%")
        if exclude_metallurgical_coal_mining and (met_coal_val*100) > metallurgical_coal_mining_threshold:
            reasons.append(f"Metallurgical Coal Mining {met_coal_val*100:.2f}% > {metallurgical_coal_mining_threshold}%")

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
    st.set_page_config(page_title="Coal Exclusion Filter (Merged, No-Duplicate Rows)", layout="wide")
    st.title("Coal Exclusion Filter: Merge SP & Urgewald by OR Logic, Keep All Features")

    # FILES
    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    # THRESHOLD TOGGLES
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

    st.sidebar.header("Business Involvement Thresholds (%)")
    exclude_generation_thermal = st.sidebar.checkbox("Exclude if Generation (Thermal Coal) > threshold?", value=False)
    generation_thermal_threshold = st.sidebar.number_input("Max allowed Generation (Thermal Coal) (%)", value=20.0)

    exclude_thermal_coal_mining = st.sidebar.checkbox("Exclude if Thermal Coal Mining > threshold?", value=False)
    thermal_coal_mining_threshold = st.sidebar.number_input("Max allowed Thermal Coal Mining (%)", value=20.0)

    exclude_metallurgical_coal_mining = st.sidebar.checkbox("Exclude if Metallurgical Coal Mining > threshold?", value=False)
    metallurgical_coal_mining_threshold = st.sidebar.number_input("Max allowed Metallurgical Coal Mining (%)", value=20.0)

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal_autodetect(sp_file, sp_sheet)
        if sp_df is None or sp_df.empty:
            st.warning("SPGlobal data is empty or couldn't be loaded.")
            return
        st.write("SP columns =>", sp_df.columns.tolist())
        st.dataframe(sp_df.head(5))

        # 2) Load Urgewald
        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if ur_df is None or ur_df.empty:
            st.warning("Urgewald data is empty or couldn't be loaded.")
            return
        st.write("Urgewald columns =>", ur_df.columns.tolist())
        st.dataframe(ur_df.head(5))

        # 3) Concatenate (ALL rows from both data sets)
        combined = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write(f"Combined shape => {combined.shape}")
        st.write("Combined columns =>", combined.columns.tolist())

        # 4) Merge duplicates (OR) => single row for each matched company
        merged = merge_duplicates_or(combined.copy())
        st.write(f"After merging duplicates => {merged.shape}")
        st.write("Merged columns =>", merged.columns.tolist())
        st.dataframe(merged.head(5))

        # 5) Filter
        filtered = filter_companies(
            df=merged,
            mining_prod_mt_threshold=mining_prod_mt_threshold,
            exclude_mining=exclude_mining,
            exclude_mining_prod_mt=exclude_mining_prod_mt,
            power_rev_threshold=power_rev_threshold,
            power_prod_threshold_percent=power_prod_threshold_percent,
            capacity_threshold_mw=capacity_threshold_mw,
            exclude_power=exclude_power,
            exclude_power_rev=exclude_power_rev,
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

        # 6) Separate Excluded/Retained/NoData
        excluded_df = filtered[filtered["Excluded"] == True].copy()
        retained_df = filtered[filtered["Excluded"] == False].copy()

        if "Coal Industry Sector" in filtered.columns:
            no_data_df = filtered[filtered["Coal Industry Sector"].isna()].copy()
        else:
            no_data_df = pd.DataFrame()

        # 7) Final columns
        # Add or remove any columns you want to appear in the final Excel
        # We'll keep the same as before, for demonstration
        final_cols = [
            "SP_ENTITY_NAME","SP_ENTITY_ID","SP_COMPANY_ID","SP_ISIN","SP_LEI",
            "Company","ISIN equity","LEI","BB Ticker",
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
            "Excluded",
            "Exclusion Reasons"
        ]
        def ensure_cols(df_):
            for c in final_cols:
                if c not in df_.columns:
                    df_[c] = np.nan
            return df_

        excluded_df = ensure_cols(excluded_df)
        retained_df = ensure_cols(retained_df)
        no_data_df  = ensure_cols(no_data_df)

        excluded_df = excluded_df[final_cols]
        retained_df = retained_df[final_cols]
        if not no_data_df.empty:
            no_data_df = no_data_df[final_cols]

        # 8) Reorder so "Company" => G, "BB Ticker" => AP, "ISIN equity" => AQ, "LEI" => AT
        excluded_df = reorder_for_excel(excluded_df)
        retained_df = reorder_for_excel(retained_df)
        no_data_df  = reorder_for_excel(no_data_df)

        # 9) Output Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, "Excluded Companies", index=False)
            retained_df.to_excel(writer, "Retained Companies", index=False)
            if not no_data_df.empty:
                no_data_df.to_excel(writer, "No Data Companies", index=False)

        st.subheader("Statistics")
        st.write(f"Total merged: {len(filtered)}")
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

if __name__ == "__main__":
    main()
