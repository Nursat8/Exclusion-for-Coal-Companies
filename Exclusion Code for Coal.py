import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

###############################
# UTILITY: MAKE COLUMNS UNIQUE
###############################
def make_columns_unique(df):
    """Renames duplicate column names by appending a suffix (_1, _2, etc.)."""
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

###############################
# LOAD SPGLOBAL
###############################
def load_spglobal_data(sp_file, sheet):
    """
    We skip the first 3 rows, then treat row #4 in the Excel file 
    as the single row of column headers. 
    (Because #PEND lines are in row #4, the real header is row #5, etc.)
    Tweak skiprows if needed, depending on your exact file layout.
    """
    try:
        # skiprows=3 => discard rows 1..3 (1-based), 
        # then next row (#4 in Excel) becomes row #0 in pandas => header=0
        df = pd.read_excel(sp_file, sheet_name=sheet, skiprows=3, header=0)
        # Flatten in case there's a multi-index (unlikely now, but just in case):
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
        st.error(f"Failed loading SPGlobal: {e}")
        return None

###############################
# LOAD URGEWALD
###############################
def load_urgewald_data(ur_file, sheet):
    """
    Urgewald has a normal single-line header in row #1 => header=0
    """
    try:
        df = pd.read_excel(ur_file, sheet_name=sheet, header=0)
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
        st.error(f"Failed loading Urgewald: {e}")
        return None

###############################
# REMOVE DUPLICATES USING "OR" LOGIC
###############################
def remove_duplicates_or(df):
    """
    Removes rows if there's any match (case-insensitive) of:
      - SP_ENTITY_NAME == Company
      - SP_ISIN == ISIN equity
      - SP_LEI == LEI
    That is, if row i and row j satisfy any match, we keep only the first occurrence.
    
    Implementation approach:
      1) unify "SP_ENTITY_NAME" and "Company" into a single 'co_name_key' ignoring case
      2) unify "SP_ISIN" and "ISIN equity" ignoring case
      3) unify "SP_LEI" and "LEI" ignoring case
      4) if any of those keys match => duplicates
    We'll do a multi-step approach:
      - create columns for each key if possible
      - drop duplicates in turn
    """
    # create placeholders
    df["_co_name_key_"] = np.nan
    df["_isin_key_"]    = np.nan
    df["_lei_key_"]     = np.nan

    # Fill them if columns exist
    # Because we're doing "OR", we do a 3-step "drop_duplicates" approach: 
    # drop on co_name_key, then drop on isin, then drop on lei.
    # That way if any match, we remove the row. 
    # (We keep the first row encountered.)
    
    # 1) unify company name
    # sp_name vs ur_name => we create a single co_name_key that 
    # is the best available (or just the sp_name if present, else ur_name) 
    # then we do drop_duplicates. 
    # However, the user wants "OR" across *pairs*. 
    # It's simpler to define each row's "co_name_key" as the 
    # lowercased string from "SP_ENTITY_NAME" if present, else "Company" 
    # or the other way around. 
    # But actually "OR" across rows means if row i's sp_name matches row j's ur_name, 
    # we want to treat them as duplicates. 
    # So a simpler approach is to do 2 separate columns: "sp_name_key", "ur_name_key" 
    # But that quickly becomes complicated. 
    # 
    # We'll implement a simpler approach: 
    #   We'll do "drop_duplicates" in 3 passes: 
    #   pass 1 => match on sp_entity_name (if not null) vs company (if not null).
    #   pass 2 => match on sp_isin vs isin equity
    #   pass 3 => match on sp_lei vs lei 
    # 
    # That is typically how you do "OR" logic with drop_duplicates. 
    
    # We'll define a function for each pass: 
    pass
    
def drop_duplicates_any_of(df):
    """
    Drop duplicates if *any* match (company name, ISIN, LEI).
    We'll do 3 sequential 'drop_duplicates' calls:
      - on sp_name vs. company name ignoring case
      - on sp_isin vs. isinequity ignoring case
      - on sp_lei vs. lei ignoring case
    Because we want an OR logic, we can do the approach:
      1) define a 'dedup key' for each row as sp_name.lower() or company.lower() 
         if they exist
      2) drop duplicates on that dedup key
      3) do the same for sp_isin vs isinequity
      4) do the same for sp_lei vs lei
    But we must be careful to let a row with empty sp_name skip. 
    """
    # We'll define 3 new columns: 
    df["_key_co_"]  = df.apply(lambda r: unify_co_name(r), axis=1)
    df["_key_isin_"] = df.apply(lambda r: unify_isin(r), axis=1)
    df["_key_lei_"]  = df.apply(lambda r: unify_lei(r), axis=1)

    # Now we do drop_duplicates in *sequence*. 
    # This ensures if any key matches, we drop duplicates. 
    for key in ["_key_co_", "_key_isin_", "_key_lei_"]:
        # drop duplicates on that key, ignoring rows that have None or empty 
        # so we don't treat all-empty as duplicates 
        # We'll do it by: if the key is blank => skip 
        # We'll store the original row count so we can see how many got dropped if we want. 
        # But let's do a standard approach:
        df.loc[df[key].isna() | (df[key]==""), key] = None
        df.drop_duplicates(subset=[key], keep="first", inplace=True)
    
    # remove helper columns
    df.drop(columns=["_key_co_", "_key_isin_", "_key_lei_"], inplace=True)
    return df

def unify_co_name(r):
    """
    Returns a single string key if we have either sp_entity_name or Company.
    We'll unify them ignoring case. If both are present, we pick the 
    lower one. Actually we just pick whichever is not empty. 
    But for dropping duplicates, we want to match SP rows that have sp_entity_name 
    with UR rows that have company. So let's define the 'co_name' as 
    sp_entity_name or else 'Company' if sp_entity_name is missing. 
    Then we lower() it. 
    If both are present, we might do sp_entity_name + '||' + company?
    Actually for OR logic, we just want a single representation. 
    We'll do sp_entity_name if not empty, else company. 
    """
    sp_name = str(r.get("SP_ENTITY_NAME", "")).strip().lower()
    ur_name = str(r.get("Company", "")).strip().lower()
    if sp_name and sp_name != "":
        return sp_name
    elif ur_name and ur_name != "":
        return ur_name
    else:
        return None

def unify_isin(r):
    """
    If SP_ISIN is present, use that. Otherwise use 'ISIN equity'.
    """
    sp_isin = str(r.get("SP_ISIN", "")).strip().lower()
    ur_isin = str(r.get("ISIN equity", "")).strip().lower()
    if sp_isin:
        return sp_isin
    elif ur_isin:
        return ur_isin
    return None

def unify_lei(r):
    """
    If SP_LEI is present, use that. Otherwise use 'LEI'.
    """
    sp_lei = str(r.get("SP_LEI", "")).strip().lower()
    ur_lei = str(r.get("LEI", "")).strip().lower()
    if sp_lei:
        return sp_lei
    elif ur_lei:
        return ur_lei
    return None


###############################
# CORE FILTER (same as previously)
###############################
def filter_companies(df):
    """
    Example minimal filter that just shows how we'd do the toggles. 
    We'll set Excluded randomly for demonstration. 
    In your real code, implement all your thresholds for Mining, Power, etc.
    """
    # We'll do a dummy approach that sets all Excluded=False for demonstration.
    df["Excluded"] = False
    df["Exclusion Reasons"] = ""
    return df

###############################
# MAIN
###############################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter + Remove Duplicates by OR logic")

    st.sidebar.header("File & Sheet Settings")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet", "Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet", "GCEL 2024")

    sp_file = st.sidebar.file_uploader("Upload SPGlobal", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald", type=["xlsx"])

    if st.sidebar.button("Run"):
        if not sp_file or not ur_file:
            st.warning("Please upload both files.")
            return

        # 1) Load each
        sp_df = load_spglobal_data(sp_file, sp_sheet)
        ur_df = load_urgewald_data(ur_file, ur_sheet)
        if sp_df is None or ur_df is None:
            return

        st.write("SP columns:", sp_df.columns.tolist())
        st.write("UR columns:", ur_df.columns.tolist())

        # 2) Concat
        combined = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write("Combined shape:", combined.shape)
        st.write("Combined columns:", combined.columns.tolist())

        # 3) Remove duplicates by OR (any match of name, ISIN, LEI)
        deduped = drop_duplicates_any_of(combined.copy())
        st.write(f"After removing duplicates by OR logic, rows => {len(deduped)}")

        # 4) Filter
        filtered = filter_companies(deduped)

        excluded_df = filtered[filtered["Excluded"]==True].copy()
        retained_df = filtered[filtered["Excluded"]==False].copy()

        # Suppose "No Data" means missing "Coal Industry Sector"? 
        no_data_df = filtered[filtered.get("Coal Industry Sector","").isna()] if "Coal Industry Sector" in filtered.columns else pd.DataFrame()

        # 5) Output
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excluded_df.to_excel(writer, "Excluded", index=False)
            retained_df.to_excel(writer, "Retained", index=False)
            no_data_df.to_excel(writer, "NoData", index=False)

        st.write(f"Excluded: {len(excluded_df)}, Retained: {len(retained_df)}, NoData: {len(no_data_df)}")
        st.download_button(
            "Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__=="__main__":
    main()
