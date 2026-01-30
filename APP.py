import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Internal Team KPI", layout="wide")

# ============================================================
# FUNCTION: AUTO-DETECT COLUMN SAFELY
# ============================================================
def find_column(columns, keyword1, keyword2=None):
    """
    Auto-detects a column that contains keyword1 (and optional keyword2)
    regardless of spaces, case, slashes, etc.
    """
    for col in columns:
        clean = col.replace(" ", "").replace("_", "").upper()
        if keyword2:
            if keyword1 in clean and keyword2 in clean:
                return col
        else:
            if keyword1 in clean:
                return col
    return None


# ============================================================
# FUNCTION: GENERATE KPI
# ============================================================
def generate_kpi(sales, mistake):

    # Clean columns
    sales.columns = sales.columns.str.strip().str.upper()
    mistake.columns = mistake.columns.str.strip().str.upper()

    # Detect AGENT column
    if "AGENT NAME" in sales.columns:
        sales.rename(columns={"AGENT NAME": "AGENT"}, inplace=True)
    if "PERSON" in mistake.columns:
        mistake.rename(columns={"PERSON": "AGENT"}, inplace=True)

    # Detect Date columns
    if "DATE" in sales.columns:
        sales.rename(columns={"DATE": "SO_DATE"}, inplace=True)
    if "DATE" in mistake.columns:
        mistake.rename(columns={"DATE": "MISTAKE_DATE"}, inplace=True)

    sales["SO_DATE"] = pd.to_datetime(sales["SO_DATE"], errors="coerce")
    mistake["MISTAKE_DATE"] = pd.to_datetime(mistake["MISTAKE_DATE"], errors="coerce")

    sales["MONTH"] = sales["SO_DATE"].dt.to_period("M").astype(str)
    mistake["MONTH"] = mistake["MISTAKE_DATE"].dt.to_period("M").astype(str)

    # AUTO-DETECT SO/BILL COLUMN
    so_bill_col = find_column(mistake.columns, "SO", "BILL")
    if so_bill_col is None:
        st.error("❌ Could NOT detect the SO/BILL column in Mistake File.")
        st.stop()

    # Only SO mistakes
    mistake_SO = mistake[
        mistake[so_bill_col].astype(str).str.upper().str.replace(" ", "") == "SO"
    ]

    # DAILY KPI
    daily_so = sales.groupby(["SO_DATE", "AGENT"]).size().reset_index(name="SO_COUNT")
    daily_mistake_count = (
        mistake_SO.groupby(["MISTAKE_DATE", "AGENT"])
        .size()
        .reset_index(name="MISTAKE_COUNT")
    )
    daily_mistake_points = (
        mistake_SO.groupby(["MISTAKE_DATE", "AGENT"])["NO OF MISTAKE"]
        .sum()
        .reset_index(name="TOTAL_MISTAKE_POINTS")
    )

    daily_mistake_all = pd.merge(
        daily_mistake_count, daily_mistake_points,
        on=["MISTAKE_DATE", "AGENT"], how="outer"
    ).fillna(0)

    daily_kpi = pd.merge(
        daily_so, daily_mistake_all,
        left_on=["SO_DATE", "AGENT"],
        right_on=["MISTAKE_DATE", "AGENT"],
        how="left"
    ).fillna(0)

    daily_kpi["KPI_COUNT_OF_MISTAKE_SO"] = (
        1 - (daily_kpi["MISTAKE_COUNT"] / daily_kpi["SO_COUNT"])
    ) * 100

    daily_kpi["KPI_NUMBER_OF_MISTAKES"] = (
        1 - (daily_kpi["TOTAL_MISTAKE_POINTS"] / daily_kpi["SO_COUNT"])
    ) * 100

    daily_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    # MONTHLY KPI
    monthly_so = sales.groupby(["MONTH", "AGENT"]).size().reset_index(name="SO_COUNT")
    monthly_mistake_count = (
        mistake_SO.groupby(["MONTH", "AGENT"])
        .size()
        .reset_index(name="MISTAKE_COUNT")
    )
    monthly_mistake_points = (
        mistake_SO.groupby(["MONTH", "AGENT"])["NO OF MISTAKE"]
        .sum()
        .reset_index(name="TOTAL_MISTAKE_POINTS")
    )

    monthly_mistake_all = pd.merge(
        monthly_mistake_count, monthly_mistake_points,
        on=["MONTH", "AGENT"], how="outer"
    ).fillna(0)

    monthly_kpi = pd.merge(
        monthly_so, monthly_mistake_all,
        on=["MONTH", "AGENT"], how="left"
    ).fillna(0)

    monthly_kpi["KPI_COUNT_OF_MISTAKE_SO"] = (
        1 - (monthly_kpi["MISTAKE_COUNT"] / monthly_kpi["SO_COUNT"])
    ) * 100

    monthly_kpi["KPI_NUMBER_OF_MISTAKES"] = (
        1 - (monthly_kpi["TOTAL_MISTAKE_POINTS"] / monthly_kpi["SO_COUNT"])
    ) * 100

    monthly_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    return daily_kpi, monthly_kpi


# ============================================================
# STREAMLIT UI
# ============================================================
st.title("📊 Internal Team KPI Dashboard")

col1, col2 = st.columns(2)

with col1:
    sales_file = st.file_uploader("📥 Upload Sales Order File", type=["xls", "xlsx"])

with col2:
    mistake_file = st.file_uploader("📥 Upload Mistake File", type=["xls", "xlsx"])

# ============================================================
# PROCESSING
# ============================================================
if sales_file and mistake_file:

    sales = pd.read_excel(sales_file)
    mistake = pd.read_excel(mistake_file)

    daily_kpi, monthly_kpi = generate_kpi(sales, mistake)
    st.success("Files processed successfully!")

    # Filters
    st.sidebar.header("🔍 Filters")
    view_type = st.sidebar.selectbox("View Mode", ["Daily KPI", "Monthly KPI"])

    agent_filter = st.sidebar.multiselect(
        "Filter by Agent Name",
        sorted(daily_kpi["AGENT"].unique())
    )

    if agent_filter:
        daily_kpi = daily_kpi[daily_kpi["AGENT"].isin(agent_filter)]
        monthly_kpi = monthly_kpi[monthly_kpi["AGENT"].isin(agent_filter)]

    # DATE RANGE FILTER
    if view_type == "Daily KPI":
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=(daily_kpi["SO_DATE"].min(), daily_kpi["SO_DATE"].max())
        )

        if len(date_range) == 2:
            start, end = map(pd.to_datetime, date_range)
            daily_kpi = daily_kpi[
                (daily_kpi["SO_DATE"] >= start) &
                (daily_kpi["SO_DATE"] <= end)
            ]

    # MONTH filter
    if view_type == "Monthly KPI":
        month_filter = st.sidebar.selectbox(
            "Select Month",
            sorted(monthly_kpi["MONTH"].unique())
        )
        if month_filter:
            monthly_kpi = monthly_kpi[monthly_kpi["MONTH"] == month_filter]

    # ============================================================
    # OUTPUT SECTION
    # ============================================================
    if view_type == "Daily KPI":
        st.subheader("📅 Daily KPI Table")
        st.dataframe(daily_kpi, use_container_width=True)

        # Bar chart: SO_COUNT vs Mistake SO
        st.markdown("### 📊 Sales Orders vs Mistake Sales Orders (Daily)")
        bar_data = daily_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum()
        st.bar_chart(bar_data)

        # Pie chart: KPI
        st.markdown("### 🥧 KPI (Count of Mistake SO) – Pie Chart")
        pie_data = daily_kpi.groupby("AGENT")["KPI_COUNT_OF_MISTAKE_SO"].mean().reset_index()
        pie_fig = px.pie(pie_data, values="KPI_COUNT_OF_MISTAKE_SO", names="AGENT")
        st.plotly_chart(pie_fig)

    else:
        st.subheader("📅 Monthly KPI Table")
        st.dataframe(monthly_kpi, use_container_width=True)

        # Bar chart (monthly)
        st.markdown("### 📊 Sales Orders vs Mistake Sales Orders (Monthly)")
        bar_data = monthly_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum()
        st.bar_chart(bar_data)

        # Pie chart (monthly)
        st.markdown("### 🥧 KPI – Pie Chart")
        pie_data = monthly_kpi.groupby("AGENT")["KPI_COUNT_OF_MISTAKE_SO"].mean().reset_index()
        pie_fig = px.pie(pie_data, values="KPI_COUNT_OF_MISTAKE_SO", names="AGENT")
        st.plotly_chart(pie_fig)

    # DOWNLOAD EXCEL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily_kpi.to_excel(writer, sheet_name="DAILY_SO_KPI", index=False)
        monthly_kpi.to_excel(writer, sheet_name="MONTHLY_SO_KPI", index=False)

    st.download_button(
        label="📥 Download KPI Excel Report",
        data=output.getvalue(),
        file_name="INTERNAL_TEAM_KPI.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("⬆️ Please upload both files to generate the KPI report.")
