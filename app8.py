import io
from datetime import datetime
import pandas as pd
import streamlit as st

# Make sure you have:  pip install openpyxl xlsxwriter

st.set_page_config(page_title="Active Case Data Collection Filter", layout="wide")

st.title("📊 Active Case Data Collection Filter")

st.markdown(
    """
Upload your **case repository Excel (.xlsx)**, choose a **Case Start Date range**  
and this app will:
- keep **only active cases** (no Case End Date),
- keep **only cases where System DNC Status = "Allow Data Collection"**,
- keep cases whose **Case Start Date falls within** the selected range,
- keep cases where **Applicable Functions** includes any of:  
  *Engineering Research and Development, Procurement, Supply Chain, Manufacturing*,
- give you a **downloadable Excel** with the final list.

> 🔁 If your file is `.xls`, open it in Excel and **Save As ➜ Excel Workbook (.xlsx)**, then upload.
"""
)

# --- 1. File upload ---
uploaded_file = st.file_uploader(
    "Upload case repository Excel file (.xlsx only)",
    type=["xlsx"],  # enforce xlsx to avoid xlrd dependency issues
    help="File must contain at least: Case Code, Case Start Date, Case End Date, Applicable Functions, System DNC Status.",
)

# Exact function phrases we care about
TARGET_FUNCTIONS = [
    "Engineering Research and Development",
    "Procurement",
    "Supply Chain",
    "Manufacturing",
]


def read_excel_xlsx(file):
    """Read an .xlsx file using openpyxl only."""
    return pd.read_excel(file, engine="openpyxl")


if uploaded_file is not None:
    # --- 2. Read Excel ---
    try:
        df = read_excel_xlsx(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
        st.stop()

    st.subheader("📁 Preview of Uploaded Data")
    st.dataframe(df.head(20))

    # --- 3. Ensure required columns exist ---
    required_columns = [
        "Case Code",
        "Case Start Date",
        "Case End Date",
        "Applicable Functions",
        "System DNC Status",
    ]
    missing_cols = [c for c in required_columns if c not in df.columns]
    if missing_cols:
        st.error(
            f"❌ These required columns are missing from your file: {', '.join(missing_cols)}"
        )
        st.stop()

    # --- 4. Parse date columns safely ---
    def parse_date_series(series):
        return pd.to_datetime(series, errors="coerce")

    df["Case Start Date_dt"] = parse_date_series(df["Case Start Date"])
    df["Case End Date_dt"] = parse_date_series(df["Case End Date"])

    # --- 5. Sidebar filters ---
    st.sidebar.header("🔍 Filters")

    # Default Case Start Date range for calendar
    start_min = df["Case Start Date_dt"].min()
    start_max = df["Case Start Date_dt"].max()

    if pd.isna(start_min):
        start_min = datetime(2020, 1, 1)  # fallback if no dates
    if pd.isna(start_max):
        start_max = datetime.today()

    # Calendar-style selector
    date_range = st.sidebar.date_input(
        "Case Start Date range (calendar)",
        value=(start_min.date(), start_max.date()),
        help="Only cases whose Case Start Date falls within this range will be shortlisted.",
    )

    if isinstance(date_range, tuple) and len(date_range) == 2:
        range_start_date, range_end_date = date_range
    else:
        st.sidebar.error("Please select a valid start and end date.")
        st.stop()

    range_start = pd.to_datetime(range_start_date)
    range_end = pd.to_datetime(range_end_date)

    # --- 6. Only active cases (no Case End Date) ---
    active_mask = df["Case End Date_dt"].isna()

    # --- 7. Case Start Date within selected range ---
    start_in_range_mask = (
        df["Case Start Date_dt"].notna()
        & (df["Case Start Date_dt"] >= range_start)
        & (df["Case Start Date_dt"] <= range_end)
    )

    # --- 8. System DNC Status = "Allow Data Collection" ---
    dnc_mask = df["System DNC Status"].astype(str).str.strip().eq("Allow Data Collection")

    # --- 9. Applicable Functions contains any of the target functions ---
    df["Applicable Functions_str"] = df["Applicable Functions"].astype(str)

    func_mask = False
    for func in TARGET_FUNCTIONS:
        func_mask = func_mask | df["Applicable Functions_str"].str.contains(
            func, case=False, na=False
        )

    # --- 10. Combine all masks ---
    final_mask = active_mask & start_in_range_mask & dnc_mask & func_mask
    filtered_df = df.loc[final_mask].copy()

    # --- 11. Show results ---
    st.subheader("✅ Shortlisted Active Cases (Ready for Data Collection)")
    st.write(
        f"**Total cases matching filters:** {len(filtered_df)} "
        f"(out of {len(df)} rows in the uploaded file)"
    )

    if not filtered_df.empty:
        st.dataframe(filtered_df)

        # --- 12. Download filtered Excel ---
        def to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                dataframe.to_excel(writer, index=False, sheet_name="Active Cases")
            buffer.seek(0)
            return buffer.read()

        excel_bytes = to_excel_bytes(filtered_df)

        st.download_button(
            label="📥 Download filtered cases as Excel",
            data=excel_bytes,
            file_name=f"active_cases_{range_start_date}_to_{range_end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info(
            "No cases matched these filters. "
            "Try widening the Case Start Date range, or check if System DNC Status / functions values match the expected strings."
        )

else:
    st.info("👆 Please upload your case repository Excel (.xlsx) to get started.")
