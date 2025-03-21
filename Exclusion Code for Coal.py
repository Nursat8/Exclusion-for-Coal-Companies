import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

################################################
# MAKE COLUMNS UNIQUE (avoid PyArrow duplicate error)
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
# LOAD SPGLOBAL
################################################
def load_spglobal(file, sheet_name="Sheet1"):
    """
    Example logic that:
      - Reads entire sheet with openpyxl
      - Uses row 5 (index=4) for ID columns
      - Uses row 6 (index=5) for coal-metric columns
      - Data from row 7 onward
    Adjust if your file differs.
    """
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        full_df = pd.DataFrame(data)

        if len(full_df) < 6:
            raise ValueError("SPGlobal sheet does not have enough rows (need >= 6).")

        row_5 = full_df.iloc[4].fillna("")
        row_6 = full_df.iloc[5].fillna("")

        final_col_names = []
        for col_idx in range(full_df.shape[1]):
            top_val = str(row_5[col_idx]).strip()
            bot_val = str(row_6[col_idx]).strip()
            combined_name = top_val if top_val else ""
            if bot_val and (bot_val.lower() not in combined_name.lower()):
                if combined_name:
                    combined_name += " " + bot_val
                else:
                    combined_name = bot_val
            final_col_names.append(combined_name.strip())

        sp_df = full_df.iloc[6:].reset_index(drop=True)
        sp_df.columns = final_col_names

        # Make columns unique to avoid arrow error
        sp_df = make_columns_unique(sp_df)
        return sp_df
    except Exception as e:
        st.error(f"Error loading SPGlobal data: {e}")
        return pd.DataFrame()

################################################
# LOAD URGEWALD
################################################
def load_urgewald(file, sheet_name="GCEL 2024"):
    """
    Example logic that:
      - Takes row 1 (index=0) as the header
      - Data from row 2 (index=1) onward
    """
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb[sheet_name]
        data = list(ws.values)
        if len(data) < 1:
            raise ValueError("Urgewald sheet does not have enough rows.")

        full_df = pd.DataFrame(data)
        new_header = full_df.iloc[0].fillna("")
        ur_df = full_df.iloc[1:].reset_index(drop=True)
        ur_df.columns = new_header

        # Make columns unique
        ur_df = make_columns_unique(ur_df)
        return ur_df
    except Exception as e:
        st.error(f"Error loading Urgewald file: {e}")
        return pd.DataFrame()

################################################
# DEDUPLICATION
################################################
def remove_duplicates_or(df):
    """
    Remove duplicates if ANY match (case-insensitive):
      (SP_ENTITY_NAME vs Company) OR
      (SP_ISIN vs ISIN equity) OR
      (SP_LEI vs LEI)
    """
    df["_key_name_"] = df.apply(lambda r: unify_name(r), axis=1)
    df["_key_isin_"] = df.apply(lambda r: unify_isin(r), axis=1)
    df["_key_lei_"]  = df.apply(lambda r: unify_lei(r), axis=1)

    def drop_dups_on_key(data, key):
        data.loc[data[key].isna() | (data[key] == ""), key] = np.nan
        data.drop_duplicates(subset=[key], keep="first", inplace=True)

    drop_dups_on_key(df, "_key_name_")
    drop_dups_on_key(df, "_key_isin_")
    drop_dups_on_key(df, "_key_lei_")

    df.drop(columns=["_key_name_","_key_isin_","_key_lei_"], inplace=True, errors="ignore")
    return df

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
# FILTERING (if you still want to exclude some rows)
# You can remove or customize if you want all companies
################################################
def filter_companies(df):
    """
    Example minimal filter that doesn't do anything except
    keep columns. If you truly want no filtering, remove it
    or always return "Excluded" = False.
    """
    df["Excluded"] = False
    df["Exclusion Reasons"] = ""
    return df

################################################
# BUILD A COLUMN LIST WITH PLACEHOLDERS
# so that "Company" ends up in col G,
# "BB Ticker" in col AP, "ISIN equity" in col AQ,
# "LEI" in col AT, while keeping all other columns.
################################################
def reorder_columns_with_placeholders(all_cols):
    """
    1) We'll build a list of exactly 46 columns from A..AT
    2) Insert placeholders for all positions except G(7), AP(42), AQ(43), AT(46).
    3) We'll fill them with the actual columns from all_cols if they exist.

    This means:
      A(1), B(2), C(3), D(4), E(5), F(6)
      G(7) => 'Company'
      H(8), I(9), ... AO(41)
      AP(42) => 'BB Ticker'
      AQ(43) => 'ISIN equity'
      AR(44), AS(45)
      AT(46) => 'LEI'
      (We can keep going if you have columns beyond AT, but let's keep it at 46 max.)
    Then we place all other columns in the next available slots that are placeholders.
    If we run out of placeholders, we just append them at the end.
    """

    # We'll create a 46-length list of placeholders
    desired_length = 46  # up to column AT
    placeholders = ["(placeholder)"] * desired_length

    # We know 1-based indexing for columns:
    #  7 => G
    #  42 => AP
    #  43 => AQ
    #  46 => AT
    # We'll map them:
    placeholders[7 - 1]  = "Company"      # G is index 6 in 0-based
    placeholders[42 - 1] = "BB Ticker"    # AP
    placeholders[43 - 1] = "ISIN equity"  # AQ
    placeholders[46 - 1] = "LEI"          # AT

    # We'll remove these special columns from all_cols if they exist
    forced_cols = {"Company","BB Ticker","ISIN equity","LEI"}
    remaining_cols = [c for c in all_cols if c not in forced_cols]

    # Now fill placeholders from left to right with remaining_cols
    filled_positions = set([6, 41, 42, 45])  # 0-based indices for the forced columns
    idx_remain = 0
    for i in range(desired_length):
        if i not in filled_positions:
            if idx_remain < len(remaining_cols):
                placeholders[i] = remaining_cols[idx_remain]
                idx_remain += 1

    # If we still have leftover columns not placed, append them after the 46th
    # so we truly "keep all features."
    leftover = remaining_cols[idx_remain:]
    final_column_order = placeholders + leftover

    return final_column_order

################################################
# MAIN APP
################################################
def main():
    st.set_page_config(page_title="Coal Exclusion - Keep All Features", layout="wide")
    st.title("Coal Exclusion: Include All Urgewald & SPGlobal, Keep All Columns")

    st.sidebar.header("File Inputs")
    sp_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    ur_sheet = st.sidebar.text_input("Urgewald Sheet Name", value="GCEL 2024")
    sp_file = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])
    ur_file = st.sidebar.file_uploader("Upload Urgewald Excel file", type=["xlsx"])

    run_button = st.sidebar.button("Run")

    if run_button:
        if not sp_file or not ur_file:
            st.warning("Please provide both SPGlobal and Urgewald files.")
            return

        # 1) Load SPGlobal
        sp_df = load_spglobal(sp_file, sp_sheet)
        st.subheader("SPGlobal Sample")
        st.dataframe(sp_df.head(5))

        # 2) Load Urgewald
        ur_df = load_urgewald(ur_file, ur_sheet)
        st.subheader("Urgewald Sample")
        st.dataframe(ur_df.head(5))

        # 3) Concatenate => keep all rows
        combined = pd.concat([sp_df, ur_df], ignore_index=True)
        st.write("Combined shape:", combined.shape)
        st.write("Combined columns:", list(combined.columns))

        # 4) Deduplicate => remove duplicates if name/isin/lei match
        deduped = remove_duplicates_or(combined.copy())
        st.write("After dedup =>", deduped.shape)

        # 5) Filter logic (or do nothing if you want all companies)
        final_df = filter_companies(deduped.copy())

        # 6) Reorder columns so that:
        #   - "Company" is col G
        #   - "BB Ticker" is col AP
        #   - "ISIN equity" is col AQ
        #   - "LEI" is col AT
        #   - keep all other columns
        all_cols = list(final_df.columns)
        final_col_order = reorder_columns_with_placeholders(all_cols)

        # Make sure we only include columns that actually exist in final_df
        # (placeholder) columns won't exist, so we fill them with blank
        actual_cols = []
        for c in final_col_order:
            if c in final_df.columns:
                actual_cols.append(c)
            else:
                # For placeholders, we add a dummy column
                new_col = c
                final_df[new_col] = np.nan
                actual_cols.append(new_col)

        final_output = final_df[actual_cols]

        st.write("Final output shape:", final_output.shape)
        st.write("Final columns:", list(final_output.columns))

        # 7) Output to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_output.to_excel(writer, sheet_name="All Companies", index=False)

        st.download_button(
            label="Download Results",
            data=output.getvalue(),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
