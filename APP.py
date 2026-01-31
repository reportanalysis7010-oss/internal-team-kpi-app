import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

st.set_page_config(page_title="Internal Team KPI", layout="wide")


# ============================================================
# FUNCTION: AUTO-DETECT COLUMN
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
# FUNCTION: PIE CHART FOR EACH AGENT
# ============================================================
def create_agent_pie(agent, kpi_value):
    remaining = 100 - kpi_value

    df = pd.DataFrame({
        "Label": ["KPI Score", "Fail %"],
        "Value": [kpi_value, remaining]
    })

    fig = px.pie(
        df,
        names="Label",
        values="Value",
        hole=0.55,
        color="Label",
        color_discrete_map={"KPI Score": "green", "Fail %": "red"}
    )
    fig.update_traces(textinfo="label+percent")
    fig.update_layout(title=f"{agent} – KPI Breakdown", height=350)
    return fig


# ============================================================
# FUNCTION: GENERATE KPI
# ============================================================
def generate_kpi(sales, mistake):

    # ---- CLEAN COLUMNS ----
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

    # ---- Detect SO/BILL ----
    so_bill_col = find_column(mistake.columns, "SO", "BILL")
    if so_bill_col is None:
        st.error("❌ Could NOT detect SO/BILL column.")
        st.stop()

    mistake_SO = mistake[
        mistake[so_bill_col].astype(str).str.upper().str.replace(" ", "") == "SO"
    ]

    # ============================================================
    # DAILY KPI
    # ============================================================
    daily_so = sales.groupby(["SO_DATE", "AGENT"]).size().reset_index(name="SO_COUNT")
    daily_mc = (
        mistake_SO.groupby(["MISTAKE_DATE", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    )
    daily_mp = (
        mistake_SO.groupby(["MISTAKE_DATE", "AGENT"])["NO OF MISTAKE"]
        .sum().reset_index(name="TOTAL_MISTAKE_POINTS")
    )

    daily_all = pd.merge(daily_mc, daily_mp,
                         on=["MISTAKE_DATE", "AGENT"], how="outer").fillna(0)

    daily_kpi = pd.merge(
        daily_so, daily_all,
        left_on=["SO_DATE", "AGENT"],
        right_on=["MISTAKE_DATE", "AGENT"],
        how="left"
    ).fillna(0)

    # ---- KPI CALCULATIONS ----
    daily_kpi["KPI_COUNT_OF_MISTAKE_SO"] = (
        1 - (daily_kpi["MISTAKE_COUNT"] / daily_kpi["SO_COUNT"])
    ) * 100

    daily_kpi["KPI_NUMBER_OF_MISTAKES"] = (
        1 - (daily_kpi["TOTAL_MISTAKE_POINTS"] / daily_kpi["SO_COUNT"])
    ) * 100

    # Rename fields
    daily_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    # ---- NEW FAIL COLUMNS ----
    daily_kpi["SO_KPI_SCORE"] = daily_kpi["KPI_COUNT_OF_MISTAKE_SO"]
    daily_kpi["SO_KPI_FAIL"] = 100 - daily_kpi["SO_KPI_SCORE"]

    daily_kpi["MISTAKE_KPI_SCORE"] = daily_kpi["KPI_NUMBER_OF_MISTAKES"]
    daily_kpi["MISTAKE_KPI_FAIL"] = 100 - daily_kpi["MISTAKE_KPI_SCORE"]

    # ---- Column Order ----
    daily_order = [
        "MONTH", "AGENT", "SO_COUNT",
        "COUNT_OF_MISTAKE_SO", "NUMBER_OF_MISTAKES",
        "SO_KPI_SCORE", "SO_KPI_FAIL",
        "MISTAKE_KPI_SCORE", "MISTAKE_KPI_FAIL"
    ]
    daily_kpi = daily_kpi[daily_order]

    # ============================================================
    # MONTHLY KPI
    # ============================================================
    monthly_so = sales.groupby(["MONTH", "AGENT"]).size().reset_index(name="SO_COUNT")
    monthly_mc = mistake_SO.groupby(["MONTH", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    monthly_mp = (
        mistake_SO.groupby(["MONTH", "AGENT"])["NO OF MISTAKE"]
        .sum().reset_index(name="TOTAL_MISTAKE_POINTS")
    )

    monthly_all = pd.merge(monthly_mc, monthly_mp,
                           on=["MONTH", "AGENT"], how="outer").fillna(0)

    monthly_kpi = pd.merge(
        monthly_so, monthly_all,
        on=["MONTH", "AGENT"], how="left"
    ).fillna(0)

    # KPI
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

    # New Fail Columns
    monthly_kpi["SO_KPI_SCORE"] = monthly_kpi["KPI_COUNT_OF_MISTAKE_SO"]
    monthly_kpi["SO_KPI_FAIL"] = 100 - monthly_kpi["SO_KPI_SCORE"]

    monthly_kpi["MISTAKE_KPI_SCORE"] = monthly_kpi["KPI_NUMBER_OF_MISTAKES"]
    monthly_kpi["MISTAKE_KPI_FAIL"] = 100 - monthly_kpi["MISTAKE_KPI_SCORE"]

    # Column Order
    monthly_order = [
        "MONTH", "AGENT", "SO_COUNT",
        "COUNT_OF_MISTAKE_SO", "NUMBER_OF_MISTAKES",
        "SO_KPI_SCORE", "SO_KPI_FAIL",
        "MISTAKE_KPI_SCORE", "MISTAKE_KPI_FAIL"
    ]
    monthly_kpi = monthly_kpi[monthly_order]

    return daily_kpi, monthly_kpi


# ============================================================
# UI START
# ============================================================
st.title("📊 Internal Team KPI Dashboard")

col1, col2 = st.columns(2)
with col1:
    sales_file = st.file_uploader("📥 Upload Sales Order File", type=["xls", "xlsx"])
with col2:
    mistake_file = st.file_uploader("📥 Upload Mistake File", type=["xls", "xlsx"])


# ============================================================
# PROCESS
# ============================================================
if sales_file and mistake_file:

    sales = pd.read_excel(sales_file)
    mistake = pd.read_excel(mistake_file)

    daily_kpi, monthly_kpi = generate_kpi(sales, mistake)

    st.success("✔ Files processed successfully!")

    # Filters
    st.sidebar.header("🔍 Filters")
    view_type = st.sidebar.selectbox("View Mode", ["Daily KPI", "Monthly KPI"])

    agent_filter = st.sidebar.multiselect(
        "Filter by Agent", sorted(daily_kpi["AGENT"].unique())
    )

    if agent_filter:
        daily_kpi = daily_kpi[daily_kpi["AGENT"].isin(agent_filter)]
        monthly_kpi = monthly_kpi[monthly_kpi["AGENT"].isin(agent_filter)]

    # Daily date range
    if view_type == "Daily KPI":
        date_range = st.sidebar.date_input(
            "Select Date Range",
            value=(sales["SO_DATE"].min(), sales["SO_DATE"].max())
        )
        if len(date_range) == 2:
            start, end = map(pd.to_datetime, date_range)
            daily_kpi = daily_kpi[
                (daily_kpi["MONTH"] >= start.strftime("%Y-%m")) &
                (daily_kpi["MONTH"] <= end.strftime("%Y-%m"))
            ]

    # Monthly filter
    if view_type == "Monthly KPI":
        month_filter = st.sidebar.selectbox("Select Month",
                                            sorted(monthly_kpi["MONTH"].unique()))
        monthly_kpi = monthly_kpi[monthly_kpi["MONTH"] == month_filter]

    # ============================================================
    # DISPLAY
    # ============================================================

    if view_type == "Daily KPI":
        st.subheader("📅 Daily KPI Table")
        st.dataframe(daily_kpi, use_container_width=True)

        # BAR CHART
        st.markdown("### 📊 Sales Orders vs Mistake Orders")
        bar = daily_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum().reset_index()

        fig_bar = go.Figure()
        fig_bar.add_bar(name="Sales Orders", x=bar["AGENT"], y=bar["SO_COUNT"],
                        marker_color="lightskyblue", text=bar["SO_COUNT"],
                        textposition="inside")

        fig_bar.add_bar(name="Mistake Orders", x=bar["AGENT"], y=bar["COUNT_OF_MISTAKE_SO"],
                        marker_color="red", text=bar["COUNT_OF_MISTAKE_SO"],
                        textposition="inside")

        fig_bar.update_layout(barmode="group")
        st.plotly_chart(fig_bar, use_container_width=True)

        # TREND
        st.markdown("### 📈 Daily Trend – Number of Mistakes")
        trend = mistake.groupby("MISTAKE_DATE")["NO OF MISTAKE"].sum().reset_index()
        fig_trend = px.line(trend, x="MISTAKE_DATE", y="NO OF MISTAKE", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # PIE CHARTS all agents
        st.subheader("🥧 KPI Pie Charts – All Agents")
        cols = st.columns(3)
        agents = daily_kpi["AGENT"].unique()

        for i, agent in enumerate(agents):
            df_agent = daily_kpi[daily_kpi["AGENT"] == agent]
            val = float(df_agent["SO_KPI_SCORE"].mean())
            with cols[i % 3]:
                st.plotly_chart(create_agent_pie(agent, val))

    # MONTHLY
    else:
        st.subheader("📅 Monthly KPI Table")
        st.dataframe(monthly_kpi, use_container_width=True)

        # BAR
        st.markdown("### 📊 Sales Orders vs Mistake Orders")
        bar = monthly_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum().reset_index()

        fig_bar = go.Figure()
        fig_bar.add_bar(name="Sales Orders", x=bar["AGENT"], y=bar["SO_COUNT"],
                        marker_color="lightskyblue", text=bar["SO_COUNT"],
                        textposition="inside")

        fig_bar.add_bar(name="Mistake Orders", x=bar["AGENT"], y=bar["COUNT_OF_MISTAKE_SO"],
                        marker_color="red", text=bar["COUNT_OF_MISTAKE_SO"],
                        textposition="inside")

        fig_bar.update_layout(barmode="group")
        st.plotly_chart(fig_bar, use_container_width=True)

        # TREND daily inside monthly
        st.markdown("### 📈 Daily Trend – Number of Mistakes")
        trend = mistake.groupby("MISTAKE_DATE")["NO OF MISTAKE"].sum().reset_index()
        fig_trend = px.line(trend, x="MISTAKE_DATE", y="NO OF MISTAKE", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # PIE
        st.subheader("🥧 KPI Pie Charts – All Agents")
        cols = st.columns(3)
        agents = monthly_kpi["AGENT"].unique()

        for i, agent in enumerate(agents):
            df_agent = monthly_kpi[monthly_kpi["AGENT"] == agent]
            val = float(df_agent["SO_KPI_SCORE"].mean())
            with cols[i % 3]:
                st.plotly_chart(create_agent_pie(agent, val))

    # ============================================================
    # DOWNLOAD
    # ============================================================
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        daily_kpi.to_excel(writer, sheet_name="DAILY_KPI", index=False)
        monthly_kpi.to_excel(writer, sheet_name="MONTHLY_KPI", index=False)

    st.download_button(
        "📥 Download KPI Excel",
        buffer.getvalue(),
        file_name="INTERNAL_TEAM_KPI.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("⬆️ Please upload both files to generate KPI report.")
