# =====================================================
# OANDO PLC — BUSINESS INTELLIGENCE DASHBOARD
# Clean • Safe • Beginner Friendly • No Red Flags
# =====================================================

import os
import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------
st.set_page_config(
    page_title="Oando PLC BI Dashboard",
    layout="wide"
)

# -----------------------------------------------------
# EXECUTIVE MULTI-COLOR THEME
# -----------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(
            135deg,
            rgb(55,16,113),
            rgb(211,0,33),
            rgb(232,110,24),
            rgb(242,179,0),
            rgb(245,212,0)
        );
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# -----------------------------------------------------
# DATA LOADER (ROBUST)
# -----------------------------------------------------
@st.cache_data
def load_data() -> pd.DataFrame:
    """Load Excel safely and clean columns."""
    file_path = "Oando_Oil_Gas_Random_Data.xlsx"

    if not os.path.exists(file_path):
        st.error(f"Excel file not found: {file_path}")
        return pd.DataFrame()

    try:
        data = pd.read_excel(file_path)
    except Exception as exc:
        st.error(f"Failed to read Excel file: {exc}")
        return pd.DataFrame()

    # Clean column names
    data.columns = data.columns.str.strip()

    # Convert Date if present
    if "Date" in data.columns:
        data["Date"] = pd.to_datetime(data["Date"], errors="coerce")

    # Fill key categorical fields safely
    for col in ["Status", "Region", "Customer_Segment"]:
        if col in data.columns:
            data[col] = data[col].fillna("Unknown")

    return data


df = load_data()

# Stop early if no data
if df.empty:
    st.stop()

# -----------------------------------------------------
# HEADER WITH LOGO
# -----------------------------------------------------
logo_col, title_col = st.columns([1, 4])

with logo_col:
    if os.path.exists("oando.svg"):
        st.image("oando.svg", width=120)
    else:
        st.warning("Logo file 'oando.svg' not found.")

with title_col:
    st.title("Oando PLC Operational Intelligence Dashboard")
    st.caption("Executive Energy Analytics")

# -----------------------------------------------------
# SIDEBAR FILTERS
# -----------------------------------------------------
st.sidebar.header(" Filters")

filtered_df = df.copy()

# Status filter
if "Status" in df.columns:
    status_options = sorted(df["Status"].dropna().unique())
    selected_status = st.sidebar.multiselect(
        "Order Status",
        options=status_options,
        default=status_options
    )
    filtered_df = filtered_df[filtered_df["Status"].isin(selected_status)]

# Date filter
    # --- Get min/max safely ---
    min_date_ts = df["Date"].min()
    max_date_ts = df["Date"].max()

    # Convert to python date objects (important)
    min_date = min_date_ts.date() if pd.notnull(min_date_ts) else None
    max_date = max_date_ts.date() if pd.notnull(max_date_ts) else None

    # --- Date input widget ---
    date_range = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date) if min_date and max_date else None
    )
    # --- Apply date filter safely ---
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])

        filtered_df = filtered_df[
            (filtered_df["Date"] >= start_date) &
            (filtered_df["Date"] <= end_date)
            ]

    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])

        filtered_df = filtered_df[
            (filtered_df["Date"] >= start_date) &
            (filtered_df["Date"] <= end_date)
            ]

    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])
        filtered_df = filtered_df[
            (filtered_df["Date"] >= start_date)
            & (filtered_df["Date"] <= end_date)
        ]

# Region filter
if "Region" in df.columns:
    region_options = sorted(df["Region"].dropna().unique())
    selected_region = st.sidebar.multiselect(
        "Region",
        options=region_options,
        default=region_options
    )
    filtered_df = filtered_df[filtered_df["Region"].isin(selected_region)]

# -----------------------------------------------------
# KPI CALCULATIONS (SAFE)
# -----------------------------------------------------
total_orders = len(filtered_df)

fulfilled_orders = (
    filtered_df["Status"].eq("Completed").sum()
    if "Status" in filtered_df.columns else 0
)

pending_orders = (
    filtered_df["Status"].eq("Pending").sum()
    if "Status" in filtered_df.columns else 0
)

cancelled_orders = (
    filtered_df["Status"].eq("Cancelled").sum()
    if "Status" in filtered_df.columns else 0
)

fulfillment_rate = (
    (fulfilled_orders / total_orders) * 100
    if total_orders > 0 else 0.0
)

# -----------------------------------------------------
# KPI DISPLAY
# -----------------------------------------------------
st.subheader(" Key Performance Indicators")

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Total Orders", f"{total_orders:,}")
k2.metric("Fulfilled", f"{fulfilled_orders:,}")
k3.metric("Pending", f"{pending_orders:,}")
k4.metric("Cancelled", f"{cancelled_orders:,}")
k5.metric("Fulfillment Rate", f"{fulfillment_rate:.1f}%")

# -----------------------------------------------------
# VISUALIZATIONS
# -----------------------------------------------------
st.markdown("---")
st.subheader(" Operational Insights")

col1, col2 = st.columns(2)

# Status Distribution
if "Status" in filtered_df.columns:
    status_counts = (
        filtered_df["Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
    )

    fig_status = px.bar(
        status_counts,
        x="Status",
        y="Count",
        color="Status",
        title="Order Status Distribution"
    )
    col1.plotly_chart(fig_status, use_container_width=True)

# Monthly Trend
if "Date" in filtered_df.columns:
    trend_df = (
        filtered_df
        .dropna(subset=["Date"])
        .groupby(filtered_df["Date"].dt.to_period("M"))
        .size()
        .reset_index(name="Orders")
    )

    if not trend_df.empty:
        trend_df["Date"] = trend_df["Date"].astype(str)

        fig_trend = px.area(
            trend_df,
            x="Date",
            y="Orders",
            title="Monthly Order Trend"
        )
        col2.plotly_chart(fig_trend, use_container_width=True)

# Second row
col3, col4 = st.columns(2)

# Fulfillment Pie
if "Status" in filtered_df.columns:
    fig_pie = px.pie(
        filtered_df,
        names="Status",
        hole=0.5,
        title="Fulfillment Status Share"
    )
    col3.plotly_chart(fig_pie, use_container_width=True)

# Delivery Volume
if "Volume_Barrels" in filtered_df.columns:
    fig_hist = px.histogram(
        filtered_df,
        x="Volume_Barrels",
        nbins=30,
        title="Distribution of Delivery Volume"
    )
    col4.plotly_chart(fig_hist, use_container_width=True)

# Revenue by Region
if {"Region", "Revenue_USD"}.issubset(filtered_df.columns):
    st.subheader(" Revenue by Region")

    revenue_region = (
        filtered_df
        .groupby("Region", as_index=False)["Revenue_USD"]
        .sum()
    )

    fig_rev = px.bar(
        revenue_region,
        x="Region",
        y="Revenue_USD",
        color="Region",
        title="Revenue by Region"
    )
    st.plotly_chart(fig_rev, use_container_width=True)

# -----------------------------------------------------

# -----------------------------------------------------
# FULL DATASET VIEWER
# -----------------------------------------------------
st.markdown("---")
st.subheader(" Full Dataset Viewer")

st.caption("Interactive view of the loaded Excel data")

# Show dataframe interactively
st.dataframe(
    filtered_df,
    use_container_width=True,
    height=400
)

# Optional: download button
csv_data = filtered_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="⬇ Download Filtered Data (CSV)",
    data=csv_data,
    file_name="oando_filtered_data.csv",
    mime="text/csv"
)

# EXECUTIVE INSIGHT
# -----------------------------------------------------
st.markdown("###  Executive Insight")

st.info(
    f"""
Current fulfillment rate is **{fulfillment_rate:.1f}%**.  
Monitor cancelled orders and regional revenue concentration
to improve operational efficiency.
"""
)