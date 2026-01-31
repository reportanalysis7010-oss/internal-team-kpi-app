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
# FUNCTION: PIE CHART
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
    fig.update_layout(title=f"{agent}", height=350)

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

    so_bill_col = find_column(mistake.columns, "SO", "BILL")
    if so_bill_col is None:
        st.error("❌ Could NOT detect SO/BILL column.")
        st.stop()

    mistake_SO = mistake[mistake[so_bill_col].astype(str).str.upper().str.replace(" ", "") == "SO"]

    # ================= DAILY =====================
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

    # KPI SCORES
    daily_kpi["SO_KPI_SCORE"] = (1 - (daily_kpi["MISTAKE_COUNT"] / daily_kpi["SO_COUNT"])) * 100
    daily_kpi["MISTAKE_KPI_SCORE"] = (1 - (daily_kpi["TOTAL_MISTAKE_POINTS"] / daily_kpi["SO_COUNT"])) * 100

    daily_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    # FAIL %
    daily_kpi["SO_KPI_FAIL"] = 100 - daily_kpi["SO_KPI_SCORE"]
    daily_kpi["MISTAKE_KPI_FAIL"] = 100 - daily_kpi["MISTAKE_KPI_SCORE"]

    # ADD MONTH for reorder
    daily_kpi["MONTH"] = daily_kpi["SO_DATE"].dt.to_period("M").astype(str)

    # reorder
    daily_order = [
        "MONTH", "AGENT", "SO_COUNT",
        "COUNT_OF_MISTAKE_SO", "NUMBER_OF_MISTAKES",
        "SO_KPI_SCORE", "SO_KPI_FAIL",
        "MISTAKE_KPI_SCORE", "MISTAKE_KPI_FAIL"
    ]

    daily_kpi = daily_kpi.reindex(columns=daily_order)

    # ================= MONTHLY =====================
    monthly_so = sales.groupby(["MONTH", "AGENT"]).size().reset_index(name="SO_COUNT")
    monthly_mc = mistake_SO.groupby(["MONTH", "AGENT"]).size().reset_index(name="MISTAKE_COUNT")
    monthly_mp = mistake_SO.groupby(["MONTH", "AGENT"])["NO OF MISTAKE"].sum().reset_index(name="TOTAL_MISTAKE_POINTS")

    monthly_all = pd.merge(monthly_mc, monthly_mp, on=["MONTH", "AGENT"], how="outer").fillna(0)

    monthly_kpi = pd.merge(monthly_so, monthly_all, on=["MONTH", "AGENT"], how="left").fillna(0)

    monthly_kpi["SO_KPI_SCORE"] = (1 - (monthly_kpi["MISTAKE_COUNT"] / monthly_kpi["SO_COUNT"])) * 100
    monthly_kpi["MISTAKE_KPI_SCORE"] = (1 - (monthly_kpi["TOTAL_MISTAKE_POINTS"] / monthly_kpi["SO_COUNT"])) * 100

    monthly_kpi.rename(columns={
        "MISTAKE_COUNT": "COUNT_OF_MISTAKE_SO",
        "TOTAL_MISTAKE_POINTS": "NUMBER_OF_MISTAKES"
    }, inplace=True)

    monthly_kpi["SO_KPI_FAIL"] = 100 - monthly_kpi["SO_KPI_SCORE"]
    monthly_kpi["MISTAKE_KPI_FAIL"] = 100 - monthly_kpi["MISTAKE_KPI_SCORE"]

    # reorder
    monthly_kpi = monthly_kpi.reindex(columns=daily_order)

    return daily_kpi, monthly_kpi


# ============================================================
# UI SECTION
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
    st.success("✔ Files processed successfully!")

    st.sidebar.header("🔍 FILTERS")
    view_type = st.sidebar.selectbox("View Mode", ["Daily KPI", "Monthly KPI"])

    agent_filter = st.sidebar.multiselect("Filter by Agent", sorted(daily_kpi["AGENT"].unique()))

    if agent_filter:
        daily_kpi = daily_kpi[daily_kpi["AGENT"].isin(agent_filter)]
        monthly_kpi = monthly_kpi[monthly_kpi["AGENT"].isin(agent_filter)]

    # ================= DAILY =====================
    if view_type == "Daily KPI":
        st.subheader("📅 Daily KPI Table")
        st.dataframe(daily_kpi, use_container_width=True)

        # BAR CHART
        st.markdown("### 📊 Sales Orders vs Mistakes (Daily)")
        bar_data = daily_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum().reset_index()

        fig_bar = go.Figure()
        fig_bar.add_bar(x=bar_data["AGENT"], y=bar_data["SO_COUNT"],
                        name="Sales Orders", marker_color="lightskyblue",
                        text=bar_data["SO_COUNT"], textposition="inside")

        fig_bar.add_bar(x=bar_data["AGENT"], y=bar_data["COUNT_OF_MISTAKE_SO"],
                        name="Mistakes", marker_color="red",
                        text=bar_data["COUNT_OF_MISTAKE_SO"], textposition="inside")

        fig_bar.update_layout(barmode="group")
        st.plotly_chart(fig_bar, use_container_width=True)

        # TREND
        st.markdown("### 📈 Daily Trend – Number Of Mistakes")
        trend = daily_kpi.groupby(["MONTH", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()
        fig_trend = px.line(trend, x="MONTH", y="NUMBER_OF_MISTAKES", color="AGENT", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # PIE CHARTS
        st.markdown("### 🥧 KPI Pie Charts (All Agents)")
        agents = daily_kpi["AGENT"].unique()
        cols = st.columns(3)
        for i, agent in enumerate(agents):
            df_a = daily_kpi[daily_kpi["AGENT"] == agent]
            kpi_val = float(df_a["SO_KPI_SCORE"].mean())
            with cols[i % 3]:
                st.plotly_chart(create_agent_pie(agent, kpi_val))


    # ================= MONTHLY =====================
    else:
        st.subheader("📅 Monthly KPI Table")
        st.dataframe(monthly_kpi, use_container_width=True)

        # BAR CHART
        st.markdown("### 📊 Sales Orders vs Mistakes (Monthly)")
        bar_data = monthly_kpi.groupby("AGENT")[["SO_COUNT", "COUNT_OF_MISTAKE_SO"]].sum().reset_index()

        fig_bar = go.Figure()
        fig_bar.add_bar(x=bar_data["AGENT"], y=bar_data["SO_COUNT"],
                        name="Sales Orders", marker_color="lightskyblue",
                        text=bar_data["SO_COUNT"], textposition="inside")

        fig_bar.add_bar(x=bar_data["AGENT"], y=bar_data["COUNT_OF_MISTAKE_SO"],
                        name="Mistakes", marker_color="red",
                        text=bar_data["COUNT_OF_MISTAKE_SO"], textposition="inside")

        fig_bar.update_layout(barmode="group")
        st.plotly_chart(fig_bar, use_container_width=True)

        # Trend
        st.markdown("### 📈 Daily Trend – Number Of Mistakes")
        trend = daily_kpi.groupby(["MONTH", "AGENT"])["NUMBER_OF_MISTAKES"].sum().reset_index()
        fig_trend = px.line(trend, x="MONTH", y="NUMBER_OF_MISTAKES", color="AGENT", markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

        # Pie Charts
        st.markdown("### 🥧 KPI Pie Charts (All Agents)")
        agents = monthly_kpi["AGENT"].unique()
        cols = st.columns(3)
        for i, agent in enumerate(agents):
            df_a = monthly_kpi[monthly_kpi["AGENT"] == agent]
            kpi_val = float(df_a["SO_KPI_SCORE"].mean())
            with cols[i % 3]:
                st.plotly_chart(create_agent_pie(agent, kpi_val))

    # DOWNLOAD EXCEL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        daily_kpi.to_excel(writer, sheet_name="DAILY_SO_KPI", index=False)
        monthly_kpi.to_excel(writer, sheet_name="MONTHLY_KPI", index=False)

    st.download_button(
        "📥 Download KPI Excel Report",
        output.getvalue(),
        file_name="INTERNAL_TEAM_KPI.xlsx"
    )


else:
    st.info("⬆ Upload both files to generate KPI report.")
