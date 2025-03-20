import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl

##############################
# UTILITY: MAKE COLUMNS UNIQUE
##############################
def make_columns_unique(df):
    """
    Ensures DataFrame df has uniquely named columns by appending a suffix 
    (e.g., '_1', '_2') if duplicates exist.
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
    try:
        # Skip first 5 rows => row #6 is header
        df = pd.read_excel(file, sheet_name=sheet_name, header=5)

        # Flatten multi-level columns if needed
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(x).strip() for x in col if x is not None)
                for col in df.columns
            ]
        else:
            df.columns = [str(c).strip() for c in df.columns]

        # Ensure columns are unique
        df = make_columns_unique(df)
        return df

    except Exception as e:
        st.error(f"Error loading SPGlobal sheet '{sheet_name}': {e}")
        return None

def load_urgewald_data(file, sheet_name):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(x).strip() for x in col if x is not None)
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
# FIND COLUMN BY KEYWORDS
##############################
def find_column(df, must_keywords, exclude_keywords=None):
    """
    Searches for a column in df whose header contains all must_keywords (case-insensitive),
    and does not contain any exclude_keywords.
    Returns the first match, or None.
    """
    if exclude_keywords is None:
        exclude_keywords = []

    for col in df.columns:
        col_lower = col.lower()
        if all(mk.lower() in col_lower for mk in must_keywords):
            if any(ex_kw.lower() in col_lower for ex_kw in exclude_keywords):
                continue
            return col
    return None

##############################
# FIND/RENAME COAL SHARE COLUMNS
##############################
def find_coal_share_column(df, label_keywords):
    """
    Looks for a column whose header includes all `label_keywords`.
    Returns the column name if found, else None.
    Example usage for 'Generation (Thermal Coal)': label_keywords = ["generation","thermal"].
    """
    return find_column(df, label_keywords)

def copy_or_create_column(df, source_col, target_col):
    """
    If source_col exists, copy it into target_col;
    else create target_col with NaN.
    """
    if source_col in df.columns:
        df[target_col] = pd.to_numeric(df[source_col], errors="coerce")
    else:
        df[target_col] = np.nan

##############################
# DUPLICATE REMOVAL
##############################
def remove_exact_duplicates(df, company_col, isin_col, lei_col):
    """
    Keep all unique rows, removing a row only if
    *company_col, isin_col, and lei_col* are all identical
    and non-empty in two or more rows.
    """
    # 1) If user columns do not exist or are all empty, skip the step
    must_exist = all(col in df.columns for col in [company_col, isin_col, lei_col])
    if not must_exist:
        return df  # can't deduplicate properly if columns missing

    # 2) We only consider rows that have actual (not NaN) values for all 3 columns
    #    So create a "dup_key" that is None if any is missing, else (company, isin, lei).
    df["_dup_key_"] = df.apply(
        lambda x: (
            str(x[company_col]).strip().lower(),
            str(x[isin_col]).strip().lower(),
            str(x[lei_col]).strip().lower()
        )
        if (
            pd.notnull(x[company_col]) and 
            pd.notnull(x[isin_col]) and
            pd.notnull(x[lei_col]) and
            str(x[company_col]).strip() != "" and
            str(x[isin_col]).strip() != "" and
            str(x[lei_col]).strip() != ""
        )
        else None,
        axis=1
    )

    # 3) Now drop duplicates on _dup_key_ (keeping the first)
    #    This ensures we only drop rows if they share a real 3-field key.
    df.drop_duplicates(subset=["_dup_key_"], keep="first", inplace=True)
    df.drop(columns=["_dup_key_"], inplace=True)
    return df

##############################
# MAIN STREAMLIT APP
##############################
def main():
    st.set_page_config(page_title="Coal Exclusion Filter", layout="wide")
    st.title("Coal Exclusion Filter")

    # ============ FILE & SHEET =============
    st.sidebar.header("File & Sheet Settings")

    # 1) SPGlobal file
    spglobal_sheet = st.sidebar.text_input("SPGlobal Sheet Name", value="Sheet1")
    spglobal_file  = st.sidebar.file_uploader("Upload SPGlobal Excel file", type=["xlsx"])

    # 2) Urgewald GCEL file
    urgewald_sheet = st.sidebar.text_input("Urgewald GCEL Sheet Name", value="GCEL 2024")
    urgewald_file  = st.sidebar.file_uploader("Upload Urgewald GCEL Excel file", type=["xlsx"])

    # Only proceed if both files exist
    if st.sidebar.button("Run"):
        if not spglobal_file or not urgewald_file:
            st.warning("Please upload both SPGlobal and Urgewald GCEL files.")
            return

        # Load data
        spglobal_df = load_spglobal_data(spglobal_file, spglobal_sheet)
        urgewald_df = load_urgewald_data(urgewald_file, urgewald_sheet)

        if spglobal_df is None or urgewald_df is None:
            return  # error loading => stop

        # Identify likely columns for "Company", "ISIN", "LEI"
        sp_company_col = find_column(spglobal_df, ["company"]) or "Company_SP"
        sp_isin_col    = find_column(spglobal_df, ["isin"]) or "ISIN_SP"
        sp_lei_col     = find_column(spglobal_df, ["lei"])  or "LEI_SP"

        urw_company_col = find_column(urgewald_df, ["company"]) or "Company_URW"
        urw_isin_col    = find_column(urgewald_df, ["isin"]) or "ISIN_URW"
        urw_lei_col     = find_column(urgewald_df, ["lei"])  or "LEI_URW"

        # For clarity, rename the found columns to the same standard
        # so we can easily concat them
        spglobal_df.rename(columns={
            sp_company_col: "Company",
            sp_isin_col:    "ISIN",
            sp_lei_col:     "LEI",
        }, inplace=True, errors="ignore")

        urgewald_df.rename(columns={
            urw_company_col: "Company",
            urw_isin_col:    "ISIN",
            urw_lei_col:     "LEI",
        }, inplace=True, errors="ignore")

        # Now combine
        combined_df = pd.concat([spglobal_df, urgewald_df], ignore_index=True)

        # Remove exact duplicates on the basis of (Company, ISIN, LEI)
        combined_df = remove_exact_duplicates(combined_df,
                                              company_col="Company",
                                              isin_col="ISIN",
                                              lei_col="LEI")

        # Make sure these columns exist at least:
        if "Company" not in combined_df.columns:
            combined_df["Company"] = np.nan
        if "ISIN" not in combined_df.columns:
            combined_df["ISIN"] = np.nan
        if "LEI" not in combined_df.columns:
            combined_df["LEI"] = np.nan

        # ============ FIND AND COPY the 3 Coal columns ============
        # For each DataFrame row, we want to fill these columns if they exist
        # in the original. Because we've combined them, let's just do
        # one pass of searching for them in the combined columns:

        # generation/thermal
        gen_col = find_coal_share_column(combined_df, ["generation", "thermal"]) \
                  or find_coal_share_column(combined_df, ["thermal", "gen"])
        copy_or_create_column(combined_df, gen_col, "Generation (Thermal Coal) Share")

        # thermal coal mining
        thermal_mining_col = find_coal_share_column(combined_df, ["thermal", "coal", "mining"])
        copy_or_create_column(combined_df, thermal_mining_col, "Thermal Coal Mining Share")

        # metallurgical coal mining
        metallurgical_mining_col = find_coal_share_column(combined_df, ["metallurgical", "coal", "mining"])
        copy_or_create_column(combined_df, metallurgical_mining_col, "Metallurgical Coal Mining Share")

        # If your files have different text in the headers for these columns,
        # you might need to adjust the above keyword lists or do multiple tries.

        # ============ Show the final DataFrame in Streamlit ============

        st.write(f"Total rows after combining: {len(combined_df)}")

        # Let user download the final
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            combined_df.to_excel(writer, sheet_name="All Companies", index=False)

        st.dataframe(combined_df.head(50))  # show first 50 rows for preview

        st.download_button(
            label="Download Combined Results",
            data=output.getvalue(),
            file_name="combined_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
