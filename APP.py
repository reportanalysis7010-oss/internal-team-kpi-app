import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Internal Team KPI", layout="wide")

# ============================================================
# FUNCTION: AUTO-DETECT COLUMN SAFELY
# ============================================================
def find_column(columns, keyword1, keyword2=None):
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
# FUNCTION: KPI DONUT CHART (GREEN + RED)
# ============================================================
def create_kpi_pie(kpi_value):
    remaining = 100 - kpi_value

    df = pd.DataFrame({
        "Label": ["KPI Score", "Mistake Impact"],
        "Value": [kpi_value, remaining]
    })

    fig = px.pie(
        df,
        values="Value",
        names="Label",
        color="Label",
        color_discrete_map={"KPI Score": "green", "Mistake Impact": "red"},
        hole=0.55
    )
    fig.update_traces(textinfo="label+percent")
    fig.update_layout(title="KPI Breakdown (Green = Good Orders, Red = Mistakes)")
    return fig


# ============================================================
# FUNCTION: GENERATE KPI
# ============================================================
def generate_kpi(sales, mistake):

    sales.columns = sales.columns.str.strip().str.upper()
    mistake.columns = mistake.columns.str.strip().str.upper()

    if "AGENT NAME" in sales.columns:
        sales.rename(columns={"AGENT NAME": "AGENT"}, inplace=True)
    if "PERSON" in mistake.columns:
        mistake.rename(columns={"PERSON": "AGENT"}, inplace=True)

    if "DATE" in sales.columns:
        sales.rename(columns={"DATE": "SO_DATE"}, inplace=True)
    if "DATE" in mistake.columns:
        mistake.rename(columns={"DATE": "MISTAKE_DATE"}, inplace=True)

    sales["SO_DATE"] = pd.to_datetime(sales["SO_DATE"], errors="coerce")
    mistake["MISTAKE_DATE"] = pd.to_datetime(mistake["MISTAKE_DATE"], errors="coerce")

    sales["MONTH"] = sales["SO_DATE"].dt.to_period("M").astype(str)
    mistake["MONTH"] = mistake["MISTAKE_DATE"].dt.to_period("M").astype(str)

    # Detect SO/BILL Column
    so_bill_col = find_column(mistake.columns, "SO", "BILL")
    if so_bill_col is None:
        st.error("❌ Could NOT detect SO/BILL column in Mistake File.")
        st.stop()

    mistake_SO = mistake[mistake[so_bill_col].astype(str).str.upper().str.replace(" ", "") == "SO"]

    # ================= DAILY KPI =====================
    daily_so = sales.groupby(["SO_DATE", "AGENT"]).size().reset_index(name="SO_COUNT")
    daily_mc = mistake_SO.groupby(["MISTAKE_DATE", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    daily_mp = mistake_SO.groupby(["MISTAKE_DATE", "AGENT"])["NO OF MISTAKE"].sum().reset_index(name="TOTAL_MISTAKE_POINTS")

    daily_all = pd.merge(daily_mc, daily_mp, on=["MISTAKE_DATE", "AGENT"], how="outer").fillna(0)
    daily_kpi = pd.merge(
        daily_so, daily_all,
        left_on=["SO_DATE", "AGENT"],
        right_on=["MISTAKE_DATE", "AGENT"],
        how="left"
    ).fillna(0)

    daily_kpi["KPI_COUNT_OF_MISTAKE_SO"] = (
        1 - (daily_kpi["COUNT_OF_MISTAKE_SO"] / daily_kpi["SO_COUNT"])
    ) * 100

    daily_kpi["KPI_NUMBER_OF_MISTAKES"] = (
        1 - (daily_kpi["NUMBER_OF_MISTAKES"] / daily_kpi["SO_COUNT"])
    ) * 100

    # ================= MONTHLY KPI =====================
    monthly_so = sales.groupby(["MONTH", "AGENT"]).size().reset_index(name="SO_COUNT")
    monthly_mc = mistake_SO.groupby(["MONTH", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    monthly_mp = mistake_SO.groupby(["MONTH", "AGENT"])["NO OF MISTAKE"].sum().reset_index(name="TOTAL_MISTAKE_POINTS")

    monthly_all = pd.merge(monthly_mc, monthly_mp, on=["MONTH", "AGENT"], how="outer").fillna(0)
    monthly_kpi = pd.merge(monthly_so, monthly_all, on=["MONTH", "AGENT"], how="left").fillna(0)

    monthly_kpi["KPI_COUNT_OF_MISTAKE_SO"] = (
        1 - (monthly_kpi["MISTAKE_COUNT"] / monthly_kpi["SO_COUNT"])
    ) * 100

    monthly_kpi["KPI_NUMBER_OF_MISTAKES"] = (
        1 - (monthly_kpi["TOTAL_MISTAKE_POINTS"] / monthly_kpi["SO_COUNT"])
    ) * 100

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
# PROCESSING BLOCK
# ============================================================
if sales_file and mistake_file:

    sales = pd.read_excel(sales_file)
    mistake = pd.read_excel(mistake_file)

    daily_kpi, monthly_kpi = generate_kpi(sales, mistake)
    st.success("✔ Files processed successfully!")

    # ---------------- FILTERS ----------------
    st.sidebar.header("🔍 FILTERS")
    view_type = st.sidebar.selectbox("View Mode", ["Daily KPI", "Monthly KPI"])

    agent_filter = st.sidebar.multiselect(
        "Filter by Agent Name",
        sorted(daily_kpi["AGENT"].unique())
    )

    if agent_filter:
        daily_kpi = daily_kpi[daily_kpi["AGENT"].isin(agent_filter)]
        monthly_kpi = monthly_kpi[monthly_kpi["AGENT"].isin(agent_filter)]

    # ---------------- DAILY FILTER ----------------
    if view_type == "Daily KPI":
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=(daily_kpi["SO_DATE"].min(), daily_kpi["SO_DATE"].max())
        )
        if len(date_range) == 2:
            start, end = map(pd.to_datetime, date_range)
            daily_kpi = daily_kpi[(daily_kpi["SO_DATE"] >= start) &
                                  (daily_kpi["SO_DATE"] <= end)]

    # ---------------- MONTHLY FILTER ----------------
    if view_type == "Monthly KPI":
        month_filter = st.sidebar.selectbox(
            "Select Month",
            sorted(monthly_kpi["MONTH"].unique())
        )
        monthly_kpi = monthly_kpi[monthly_kpi["MONTH"] == month_filter]


    # ============================================================
    # DAILY KPI VIEW
    # ============================================================
    if view_type == "Daily KPI":

        st.subheader("📅 Daily KPI Table")
        st.dataframe(daily_kpi, use_container_width=True)

        # ---- DAILY BAR CHART ----
        st.markdown("### 📊 Sales Orders vs Mistake Sales Orders (Daily)")
        bar_data = daily_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum()
        st.bar_chart(bar_data)

        # ---- DAILY TREND CHART ----
        st.markdown("### 📈 Daily Trend – Number of Mistakes")
        trend_data = daily_kpi.groupby(["SO_DATE", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()

        fig_trend = px.line(
            trend_data,
            x="SO_DATE",
            y="NUMBER_OF_MISTAKES",
            color="AGENT",
            markers=True,
            title="Daily Number of Mistakes Trend"
        )
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---- KPI PIE ----
        if len(daily_kpi["AGENT"].unique()) == 1:
            kpi_val = float(daily_kpi["KPI_COUNT_OF_MISTAKE_SO"].mean())
            st.plotly_chart(create_kpi_pie(kpi_val))
        else:
            st.info("➡ Select ONE agent to view KPI Pie chart")


    # ============================================================
    # MONTHLY KPI VIEW — WITH DAILY TREND
    # ============================================================
    else:

        st.subheader("📅 Monthly KPI Table")
        st.dataframe(monthly_kpi, use_container_width=True)

        # ---- MONTHLY BAR ----
        st.markdown("### 📊 Sales Orders vs Mistake Orders (Monthly)")
        bar_data = monthly_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum()
        st.bar_chart(bar_data)

        # ---- DAILY TREND INSIDE MONTHLY ----
        st.markdown("### 📈 Daily Trend – Number of Mistakes (Filtered by Month)")

        selected_month = monthly_kpi["MONTH"].iloc[0]
        daily_filtered = daily_kpi[daily_kpi["MONTH"] == selected_month]

        trend_data = daily_filtered.groupby(["SO_DATE", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()

        fig_trend = px.line(
            trend_data,
            x="SO_DATE",
            y="NUMBER_OF_MISTAKES",
            color="AGENT",
            markers=True,
            title=f"Daily Mistake Trend for {selected_month}"
        )
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---- KPI PIE ----
        if len(monthly_kpi["AGENT"].unique()) == 1:
            kpi_val = float(monthly_kpi["KPI_COUNT_OF_MISTAKE_SO"].mean())
            st.plotly_chart(create_kpi_pie(kpi_val))
        else:
            st.info("➡ Select ONE agent to view KPI Pie chart")

    # ============================================================
    # DOWNLOAD EXCEL
    # ============================================================
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
    st.info("⬆️ Please upload both files to generate KPI report.")
