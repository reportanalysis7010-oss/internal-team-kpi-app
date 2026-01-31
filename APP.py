import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Internal Team KPI Dashboard", layout="wide")

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
def create_kpi_pie(kpi_score, kpi_fail, title):
    df = pd.DataFrame({
        "Label": ["SO_KPI_SCORE", "SO_KPI_FAIL"],
        "Value": [kpi_score, kpi_fail]
    })

    fig = px.pie(
        df,
        values="Value",
        names="Label",
        color="Label",
        color_discrete_map={
            "SO_KPI_SCORE": "green",
            "SO_KPI_FAIL": "red"
        },
        hole=0.55,
        title=title
    )

    fig.update_traces(textinfo="label+percent")
    return fig


# ============================================================
# FUNCTION: GENERATE KPI
# ============================================================
def generate_kpi(sales, mistake):

    sales.columns = sales.columns.str.strip().str.upper()
    mistake.columns = mistake.columns.str.strip().str.upper()

    # Rename columns
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

    # Detect SO/BILL
    so_bill_col = find_column(mistake.columns, "SO", "BILL")
    if so_bill_col is None:
        st.error("❌ SO/BILL column not found")
        st.stop()

    mistake_SO = mistake[mistake[so_bill_col].astype(str).str.upper().str.replace(" ", "") == "SO"]

    # ---------- DAILY KPI ----------
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

    # MAIN KPIs
    daily_kpi["SO_KPI_SCORE"] = (1 - (daily_kpi["MISTAKE_COUNT"] / daily_kpi["SO_COUNT"])) * 100
    daily_kpi["SO_KPI_FAIL"] = 100 - daily_kpi["SO_KPI_SCORE"]

    daily_kpi["MISTAKE_KPI_SCORE"] = (1 - (daily_kpi["TOTAL_MISTAKE_POINTS"] / daily_kpi["SO_COUNT"])) * 100
    daily_kpi["MISTAKE_KPI_FAIL"] = 100 - daily_kpi["MISTAKE_KPI_SCORE"]

    daily_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    daily_kpi["MONTH"] = daily_kpi["SO_DATE"].dt.to_period("M").astype(str)

    # ---------- MONTHLY KPI ----------
    monthly_so = sales.groupby(["MONTH", "AGENT"]).size().reset_index(name="SO_COUNT")
    monthly_mc = mistake_SO.groupby(["MONTH", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    monthly_mp = mistake_SO.groupby(["MONTH", "AGENT"])["NO OF MISTAKE"].sum().reset_index(name="TOTAL_MISTAKE_POINTS")

    monthly_all = pd.merge(monthly_mc, monthly_mp, on=["MONTH", "AGENT"], how="outer").fillna(0)
    monthly_kpi = pd.merge(monthly_so, monthly_all, on=["MONTH", "AGENT"], how="left").fillna(0)

    monthly_kpi["SO_KPI_SCORE"] = (1 - (monthly_kpi["MISTAKE_COUNT"] / monthly_kpi["SO_COUNT"])) * 100
    monthly_kpi["SO_KPI_FAIL"] = 100 - monthly_kpi["SO_KPI_SCORE"]

    monthly_kpi["MISTAKE_KPI_SCORE"] = (1 - (monthly_kpi["TOTAL_MISTAKE_POINTS"] / monthly_kpi["SO_COUNT"])) * 100
    monthly_kpi["MISTAKE_KPI_FAIL"] = 100 - monthly_kpi["MISTAKE_KPI_SCORE"]

    monthly_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    # --- REORDER COLUMNS ---
    desired_order = [
        "MONTH", "AGENT", "SO_COUNT",
        "COUNT_OF_MISTAKE_SO", "NUMBER_OF_MISTAKES",
        "SO_KPI_SCORE", "SO_KPI_FAIL",
        "MISTAKE_KPI_SCORE", "MISTAKE_KPI_FAIL"
    ]

    daily_kpi = daily_kpi[[c for c in desired_order if c in daily_kpi.columns]]
    monthly_kpi = monthly_kpi[[c for c in desired_order if c in monthly_kpi.columns]]

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

    # Filters
    st.sidebar.header("🔍 Filters")
    view_type = st.sidebar.selectbox("View Mode", ["Daily KPI", "Monthly KPI"])

    agent_filter = st.sidebar.multiselect("Agent Filter", sorted(daily_kpi["AGENT"].unique()))
    if agent_filter:
        daily_kpi = daily_kpi[daily_kpi["AGENT"].isin(agent_filter)]
        monthly_kpi = monthly_kpi[monthly_kpi["AGENT"].isin(agent_filter)]

    # DAILY FILTER
    if view_type == "Daily KPI":
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=(daily_kpi["MONTH"].min(), daily_kpi["MONTH"].max())
        )
        if len(date_range) == 2:
            start, end = map(pd.to_datetime, date_range)
            daily_kpi = daily_kpi[
                (pd.to_datetime(daily_kpi["MONTH"]) >= start) &
                (pd.to_datetime(daily_kpi["MONTH"]) <= end)
            ]

    # MONTHLY FILTER
    if view_type == "Monthly KPI":
        month_filter = st.sidebar.selectbox("Select Month", sorted(monthly_kpi["MONTH"].unique()))
        monthly_kpi = monthly_kpi[monthly_kpi["MONTH"] == month_filter]

    # ============================================================
    # DISPLAY SECTION
    # ============================================================

    # ---------- DAILY VIEW ----------
    if view_type == "Daily KPI":

        st.subheader("📅 Daily KPI Table")
        st.dataframe(daily_kpi, use_container_width=True)

        # BAR CHART
        st.markdown("### 📊 Daily SO vs Mistakes")
        fig_bar = px.bar(
            daily_kpi,
            x="AGENT",
            y=["SO_COUNT", "COUNT_OF_MISTAKE_SO"],
            barmode="group",
            title="Daily Sales Orders vs Mistake Orders",
            color_discrete_map={"SO_COUNT": "lightblue", "COUNT_OF_MISTAKE_SO": "red"}
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # TREND CHART
        st.markdown("### 📈 Daily Mistake Trend")
        trend = daily_kpi.groupby(["MONTH", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()

        fig_trend = px.line(trend, x="MONTH", y="NUMBER_OF_MISTAKES", color="AGENT", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # PIE CHARTS FOR EACH AGENT
        st.markdown("### 🥧 KPI Pie Charts")

        agents = daily_kpi["AGENT"].unique()
        cols = st.columns(3)

        for i, agent in enumerate(agents):
            data = daily_kpi[daily_kpi["AGENT"] == agent]
            score = float(data["SO_KPI_SCORE"].mean())
            fail = 100 - score
            fig = create_kpi_pie(score, fail, f"{agent} - Daily KPI")

            cols[i % 3].plotly_chart(fig)

    # ---------- MONTHLY VIEW ----------
    else:

        st.subheader("📅 Monthly KPI Table")
        st.dataframe(monthly_kpi, use_container_width=True)

        # BAR CHART
        st.markdown("### 📊 Monthly SO vs Mistakes")
        fig_bar = px.bar(
            monthly_kpi,
            x="AGENT",
            y=["SO_COUNT", "COUNT_OF_MISTAKE_SO"],
            barmode="group",
            title="Monthly Sales Orders vs Mistake Orders",
            color_discrete_map={"SO_COUNT": "lightblue", "COUNT_OF_MISTAKE_SO": "red"}
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # TREND CHART (daily inside selected month)
        st.markdown("### 📈 Daily Trend inside Month")
        trend = daily_kpi.groupby(["MONTH", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()

        fig_trend = px.line(trend, x="MONTH", y="NUMBER_OF_MISTAKES", color="AGENT", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # PIE CHARTS FOR EACH AGENT
        st.markdown("### 🥧 KPI Pie Charts (Monthly)")

        agents = monthly_kpi["AGENT"].unique()
        cols = st.columns(3)

        for i, agent in enumerate(agents):
            data = monthly_kpi[monthly_kpi["AGENT"] == agent]
            score = float(data["SO_KPI_SCORE"].mean())
            fail = 100 - score
            fig = create_kpi_pie(score, fail, f"{agent} - Monthly KPI")

            cols[i % 3].plotly_chart(fig)

    # ------------------------------------------
    # DOWNLOAD EXCEL
    # ------------------------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily_kpi.to_excel(writer, "DAILY_KPI", index=False)
        monthly_kpi.to_excel(writer, "MONTHLY_KPI", index=False)

    st.download_button(
        "📥 Download KPI Excel Report",
        output.getvalue(),
        "INTERNAL_TEAM_KPI.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


else:
    st.info("⬆️ Please upload both files to generate the KPI report.")
