
import streamlit as st
import pandas as pd
import io

# ▲── CONFIG ──────────────────────────────────────────────────────────────────
# 1) Define your filter hierarchy in the exact order you want them applied:
FILTER_HIERARCHY = [
    "PM Name",
    "Project Code",
    "Project Type",
    "Project Name",
    "Received From",
    "Start date",
    "End Date",
    "Priority",
    "Status"
]
# 2) List which categorical columns you want charts for:
CHART_COLUMNS = ["Status", "Priority", "Project Type"]
# ──────────────────────────────────────────────────────────────────────────────

@st.cache_data
def load_data():
    """
    Dynamically detect header row by matching expected headers,
    then load DataFrame with correct header.
    """
    # read raw data without header
    raw = pd.read_excel("data/master.xlsx", sheet_name=0, header=None)
    # find header row
    desired = [col.lower() for col in FILTER_HIERARCHY]
    header_row = None
    for idx, row in raw.iterrows():
        values = [str(x).strip().lower() for x in row.tolist()]
        if values[: len(desired)] == desired:
            header_row = idx
            break
    if header_row is None:
        raise ValueError("Header row not found in master.xlsx")
    # load again with proper header
    df = pd.read_excel(
        "data/master.xlsx",
        sheet_name=0,
        header=header_row,
        parse_dates=["Start date", "End Date"],
    )
    # drop rows before header
    df = df.copy()
    # clean column names
    df.columns = df.columns.map(str).str.strip()
    # drop unnamed junk
    df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", regex=True)]
    return df


def generate_excel_report(df: pd.DataFrame, chart_cols: list) -> bytes:
    """
    Creates an in-memory Excel file with:
      • Sheet 'Data' = filtered DataFrame
      • One sheet per chart column, each with table + embedded chart
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # export data
        df.to_excel(writer, sheet_name="Data", index=False)
        workbook = writer.book
        # build charts
        for col in chart_cols:
            if col not in df.columns:
                continue
            counts = df[col].value_counts().reset_index()
            counts.columns = [col, "Count"]
            sheet = col[:31]
            counts.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            chart = workbook.add_chart({"type": "column"})
            max_row = len(counts) + 1
            chart.add_series({
                "name": f"Count by {col}",
                "categories": [sheet, 1, 0, max_row, 0],
                "values": [sheet, 1, 1, max_row, 1],
            })
            chart.set_title({"name": f"Records by {col}"})
            chart.set_x_axis({"name": col})
            chart.set_y_axis({"name": "Count"})
            ws.insert_chart("D2", chart, {"x_scale": 1.3, "y_scale": 1.3})
        writer.save()
    return output.getvalue()


def main():
    st.set_page_config(page_title="Spine Dashboard", layout="wide")
    st.title("📊 Spine Monthly Tracker Dashboard")
    # load data
    try:
        df = load_data()
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return
    st.write("**Detected columns:**", df.columns.tolist())
    # apply filters
    st.sidebar.header("🔎 Apply Filters")
    mask = pd.Series(True, index=df.index)
    for col in FILTER_HIERARCHY:
        if col not in df.columns:
            st.sidebar.error(f"Column '{col}' not found.")
            return
        opts = sorted(df.loc[mask, col].dropna().unique())
        sel = st.sidebar.multiselect(col, options=opts, default=opts)
        mask &= df[col].isin(sel)
    df_filt = df[mask]
    # key metrics
    st.subheader("🔢 Key Metrics")
    st.metric("Total Records", len(df_filt))
    # charts
    st.subheader("📈 Charts")
    for col in CHART_COLUMNS:
        if col in df_filt.columns:
            st.markdown(f"**Records by {col}**")
            st.bar_chart(df_filt[col].value_counts())
    # data table
    st.subheader("📋 Filtered Data")
    st.dataframe(df_filt, use_container_width=True)
    # export
    if st.button("📥 Export Full Report to Excel"):
        report = generate_excel_report(df_filt, CHART_COLUMNS)
        st.download_button(
            label="Download .xlsx",
            data=report,
            file_name="Spine_Tracker_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
