from __future__ import annotations

from io import BytesIO
from pathlib import Path
import html
import hashlib

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from pdf_export import build_pdf_report
from report_utils import (
    NONE_LABEL,
    build_excel_report,
    format_inr,
    guess_column,
    parse_numeric,
    prepare_analysis,
    read_data_file,
    visible_data,
)

APP_DIR = Path(__file__).parent
SAMPLE_PATH = APP_DIR / "sample_sales.csv"
CURRENCIES = {"INR (₹)": "₹", "USD ($)": "$", "EUR (€)": "€", "GBP (£)": "£"}
CHART_PALETTE = ["#356859", "#b57943", "#507d9a", "#ae6e57", "#8c7188", "#749d91", "#8a9870"]

st.set_page_config(
    page_title="Ledgerly | Business reporting",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root {
        --ink:#1d2924; --muted:#66716b; --line:#dce4de;
        --paper:#f6f7f4; --panel:#ffffff; --panel-2:#eef2ef;
        --accent:#356859; --accent-dark:#24483d; --accent-soft:#e7efe9;
        --warm:#b57943;
      }
      html, body, [class*="css"] { font-family:system-ui,-apple-system,"Segoe UI",sans-serif; color:var(--ink); }
      html { background:var(--paper); color-scheme:light; }
      .stApp { background:var(--paper); color:var(--ink); }
      [data-testid="stHeader"] { background:rgba(246,247,244,.96); }
      [data-testid="stToolbar"] { right:1rem; }
      .block-container { padding-top:1.5rem; padding-bottom:2.5rem; max-width:1440px; }

      [data-testid="stSidebar"] { background:#eef2ef !important; border-right:1px solid var(--line) !important; }
      [data-testid="stSidebar"] > div:first-child { background:transparent !important; }
      [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
      [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
      [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
      [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
      [data-testid="stSidebar"] label,
      [data-testid="stSidebar"] label p,
      [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
      [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
      [data-testid="stSidebar"] legend,
      [data-testid="stSidebar"] [data-testid="stExpander"] summary,
      [data-testid="stSidebar"] [data-testid="stExpander"] summary p { color:#35433c !important; font-weight:500 !important; }
      [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:var(--accent-dark) !important; font-weight:700 !important; }
      [data-testid="stSidebar"] hr { border-color:var(--line) !important; opacity:1; }
      [data-testid="stSidebar"] [data-testid="stRadio"] > label,
      [data-testid="stSidebar"] [data-testid="stSelectbox"] > label,
      [data-testid="stSidebar"] [data-testid="stMultiSelect"] > label,
      [data-testid="stSidebar"] [data-testid="stDateInput"] > label,
      [data-testid="stSidebar"] [data-testid="stFileUploader"] > label { color:#35433c !important; font-weight:650 !important; }
      [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label { background:transparent; border:1px solid transparent; border-radius:8px; padding:7px 9px; margin:3px 0; }
      [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover { background:#e1e9e3; border-color:#cad8cf; }

      input, textarea, [data-baseweb="input"] input { color:var(--ink) !important; background:#fff !important; }
      [data-baseweb="select"] > div, [data-testid="stDateInput"] input, [data-testid="stNumberInput"] input,
      [data-testid="stTextInput"] input, [data-testid="stFileUploaderDropzone"] { background:#fff !important; color:var(--ink) !important; border-color:#cbd6cf !important; }
      [data-baseweb="select"] *, [data-testid="stMultiSelect"] * { color:#27352e; }
      [data-baseweb="popover"] [role="listbox"], [data-baseweb="menu"] { background:#fff !important; border:1px solid var(--line) !important; }
      [data-baseweb="popover"] [role="option"] { color:var(--ink) !important; }
      [data-baseweb="popover"] [role="option"]:hover { background:#eef2ef !important; }
      [data-testid="stFileUploaderDropzone"] { border:1px dashed #aabbb0 !important; border-radius:10px !important; }

      h1,h2,h3,h4 { color:#22342c !important; letter-spacing:-.02em; font-weight:700 !important; }
      h1 { font-size:2rem !important; }
      h2 { font-size:1.45rem !important; }
      h3 { font-size:1.12rem !important; }
      p, label { color:#35433c !important; font-size:14px; } [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] *, .stCaption, .stCaption * { color:#35433c !important; font-size:14px !important; opacity:1 !important; }
      strong,b { color:#24483d !important; font-weight:650 !important; }
      [data-testid="stMarkdownContainer"] a { color:#356859 !important; font-weight:600; }

      [data-testid="stMetric"] { background:#fff !important; padding:16px 17px; border:1px solid #e1e7e2 !important; border-radius:12px; box-shadow:0 1px 2px rgba(27,48,37,.04); }
      [data-testid="stMetricLabel"] { color:#3d4942 !important; font-size:.84rem !important; font-weight:600 !important; overflow:visible !important; }
      [data-testid="stMetricLabel"] * { color:#3d4942 !important; white-space:normal !important; overflow:visible !important; text-overflow:clip !important; line-height:1.25 !important; }
      [data-testid="stMetricValue"] { color:#24483d !important; font-weight:700 !important; letter-spacing:-.025em; white-space:nowrap; }
      [data-testid="stMetricValue"] * { font-size:clamp(.95rem, 1.45vw, 1.55rem) !important; white-space:nowrap !important; }
      [data-testid="stMetricDelta"] { color:#356859 !important; font-weight:600 !important; }
      [data-testid="stVerticalBlockBorderWrapper"] { border-color:#e1e7e2 !important; border-radius:12px !important; background:#fff !important; }
      [data-testid="stDataFrame"] { border:1px solid #dce4de; border-radius:10px; overflow:hidden; }
      [data-testid="stDataFrame"] * { color:#26332d; }

      .hero { padding:26px 28px; background:#24483d; border:1px solid #24483d; border-radius:12px; color:#fff; margin-bottom:22px; box-shadow:none; }
      .hero .eyebrow { font-size:11px; letter-spacing:.11em; text-transform:uppercase; color:#c8ddd1; font-weight:650; margin-bottom:9px; }
      .hero h1 { color:#fff !important; font-size:clamp(24px,2.2vw,32px) !important; margin:0 0 9px 0; letter-spacing:-.03em; }
      .hero p { color:#e8f0eb !important; margin:0; max-width:780px; font-size:15px; line-height:1.6; }
      .hero strong,.hero b { color:#fff !important; font-weight:650 !important; }
      .section-kicker { color:#356859; text-transform:uppercase; letter-spacing:.09em; font-weight:650; font-size:11px; margin-bottom:6px; }
      .muted { color:#4d5b53 !important; font-size:12px; font-weight:500; }
      .insight-card { background:#fff; border:1px solid #e1e7e2; padding:16px 17px; border-radius:12px; min-height:108px; }
      .insight-title { color:#56665d; font-size:11px; font-weight:650; text-transform:uppercase; letter-spacing:.07em; margin-bottom:9px; }
      .insight-value { color:#22342c; font-size:18px; line-height:1.35; font-weight:700; overflow-wrap:anywhere; }
      .insight-detail { color:#4d5b53; font-size:12px; margin-top:6px; font-weight:500; }
      .status-pill { display:inline-block; border:1px solid #c8d9ce; background:#e7efe9; color:#315b4d; border-radius:99px; padding:6px 11px; font-size:11px; font-weight:650; }

      [data-testid="stButton"] button, [data-testid="stDownloadButton"] button, div.stButton > button, div.stDownloadButton > button { border-radius:8px; font-weight:600 !important; min-height:42px; color:#fff !important; background:#356859 !important; border:1px solid #356859 !important; box-shadow:none; }
      [data-testid="stButton"] button *, [data-testid="stDownloadButton"] button * { color:#fff !important; }
      div.stButton > button:hover, div.stDownloadButton > button:hover { background:#284f44 !important; border-color:#284f44 !important; }
      div.stButton > button[kind="primary"] { background:#24483d !important; border-color:#24483d !important; color:#fff !important; }
      hr { border-color:#dce4de !important; opacity:1; }
      footer { visibility:hidden; }
      .stAlert { border-radius:10px; }
      [data-testid="stAlert"] p { color:#27352e !important; }
      [data-testid="stExpander"] { border-color:#dce4de !important; background:#fff !important; }
      ::selection { background:#d8e7dd; color:#1d2924; }
      @media (max-width:900px) { .block-container { padding-left:1rem; padding-right:1rem; } .hero { padding:21px; } }
      @media (prefers-reduced-motion:reduce) { *,*::before,*::after { animation:none !important; scroll-behavior:auto !important; transition:none !important; } }
    </style>
    """,
    unsafe_allow_html=True,
)


def _money(value: float, symbol: str) -> str:
    if pd.isna(value):
        value = 0
    return f"{symbol}{value:,.2f}"


def _safe_text(value: object) -> str:
    return html.escape(str(value))


def _kicker(text: str) -> None:
    st.markdown(f'<div class="section-kicker">{_safe_text(text)}</div>', unsafe_allow_html=True)


def _insight_card(title: str, value: str, detail: str) -> None:
    st.markdown(
        f'<div class="insight-card"><div class="insight-title">{_safe_text(title)}</div>'
        f'<div class="insight-value">{_safe_text(value)}</div><div class="insight-detail">{_safe_text(detail)}</div></div>',
        unsafe_allow_html=True,
    )


def _format_chart(fig, *, height: int = 330):
    """Apply readable typography, accessible contrast and smooth Plotly transitions."""
    fig.update_layout(
        height=height,
        autosize=True,
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        margin=dict(l=44, r=24, t=48, b=48),
        font=dict(family="system-ui, sans-serif", color="#26332d", size=13),
        title=dict(font=dict(size=17, color="#356859", family="system-ui, sans-serif"), x=0.02, xanchor="left"),
        legend=dict(font=dict(size=13, color="#26332d", family="system-ui, sans-serif"), orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=dict(bgcolor="#ffffff", bordercolor="#cbd6cf", font=dict(family="system-ui, sans-serif", color="#26332d", size=13)),
        hovermode="x unified",
        transition=dict(duration=0),
        uirevision="ledgerly-chart",
    )
    fig.update_xaxes(
        showgrid=False,
        linecolor="#d5dfd9",
        zeroline=False,
        automargin=True,
        tickfont=dict(size=13, color="#26332d", family="system-ui, sans-serif"),
        title_font=dict(size=13, color="#356859", family="system-ui, sans-serif"),
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="#e7ece8",
        linecolor="rgba(0,0,0,0)",
        zeroline=False,
        automargin=True,
        tickfont=dict(size=13, color="#26332d", family="system-ui, sans-serif"),
        title_font=dict(size=13, color="#356859", family="system-ui, sans-serif"),
    )
    return fig


# ---------- Sidebar / source selection ----------
with st.sidebar:
    st.markdown("<div style='font-size:12px;color:#356859;font-weight:700;letter-spacing:.12em;margin-bottom:8px'>LEDGERLY</div>", unsafe_allow_html=True)
    st.markdown("## Reporting workspace")
    st.caption("Turn raw sales files into useful business summaries.")
    st.markdown("---")
    source_mode = st.radio("Data source", ["Try sample data", "Upload your own file"], index=0)
    uploaded = None
    if source_mode == "Upload your own file":
        uploaded = st.file_uploader("Choose CSV or Excel file", type=["csv", "xlsx", "xlsm"])
        st.caption("For the public demo, use fictional or anonymised data only.")
    st.markdown("---")
    page = st.radio("WORKSPACE", ["Overview", "Explore data", "Data quality", "Export center"], index=0)
    st.markdown("---")
    currency_label = st.selectbox("Report currency label", list(CURRENCIES.keys()), index=0)
    currency_symbol = CURRENCIES[currency_label]

if source_mode == "Upload your own file" and uploaded is None:
    st.markdown(
        '<div class="hero"><div class="eyebrow">SALES REPORTING</div><h1>Start with your sales data.</h1><p>Upload a CSV or Excel file to review trends, check data quality and export a report.</p></div>',
        unsafe_allow_html=True,
    )
    st.info("Upload a CSV or Excel workbook in the sidebar to start. You can switch to sample data at any time.")
    st.stop()

try:
    if uploaded is not None:
        df_raw = read_data_file(BytesIO(uploaded.getvalue()), uploaded.name)
        source_name = uploaded.name
    else:
        df_raw = read_data_file(SAMPLE_PATH, "sample_sales.csv")
        source_name = "sample_sales.csv · Demo dataset"
except Exception as exc:
    st.error(f"We couldn't read this file. {exc}")
    st.stop()

# ---------- Column mapping + cleaning rules ----------
columns = list(df_raw.columns)
options = [NONE_LABEL] + columns
dataset_signature = hashlib.sha1((source_name + "|" + "|".join(columns)).encode("utf-8")).hexdigest()[:8]
with st.sidebar:
    with st.expander("Data setup", expanded=True):
        st.caption("Map your columns once. Ledgerly uses these mappings across each report view.")
        date_guess = guess_column(columns, ["order date", "transaction date", "sale date", "date", "created at"])
        sales_guess = guess_column(columns, ["total sales", "revenue", "sales amount", "total amount", "amount", "sales", "total"])
        category_guess = guess_column(columns, ["product category", "category", "department", "type"])
        product_guess = guess_column(columns, ["product name", "product", "item name", "item", "sku"])
        order_guess = guess_column(columns, ["order id", "transaction id", "invoice number", "order number", "invoice", "order"])
        qty_guess = guess_column(columns, ["quantity", "qty", "units sold"])
        price_guess = guess_column(columns, ["unit price", "price per unit", "unit cost", "price"])

        date_col = st.selectbox("Date column", options, index=options.index(date_guess), key=f"map_date_{dataset_signature}")
        revenue_mode = st.radio("How should sales be calculated?", ["Use a sales/revenue column", "Calculate quantity × unit price"], index=0 if sales_guess != NONE_LABEL else (1 if qty_guess != NONE_LABEL and price_guess != NONE_LABEL else 0), key=f"revenue_mode_{dataset_signature}")
        sales_col = NONE_LABEL
        quantity_col = NONE_LABEL
        unit_price_col = NONE_LABEL
        if revenue_mode == "Use a sales/revenue column":
            sales_col = st.selectbox("Sales / revenue column", options, index=options.index(sales_guess), key=f"map_sales_{dataset_signature}")
        else:
            quantity_col = st.selectbox("Quantity column", options, index=options.index(qty_guess), key=f"map_qty_{dataset_signature}")
            unit_price_col = st.selectbox("Unit price column", options, index=options.index(price_guess), key=f"map_unit_price_{dataset_signature}")
        category_col = st.selectbox("Category column", options, index=options.index(category_guess), key=f"map_category_{dataset_signature}")
        product_col = st.selectbox("Product column", options, index=options.index(product_guess), key=f"map_product_{dataset_signature}")
        order_col = st.selectbox("Order / invoice ID (optional)", options, index=options.index(order_guess), key=f"map_order_{dataset_signature}")
        remove_duplicates = st.checkbox("Remove exact duplicate rows from analysis", value=True)
        trim_text = st.checkbox("Trim leading/trailing spaces in text", value=True)
        dayfirst_dates = st.checkbox("Interpret dates as DD/MM/YYYY (day first)", value=False, help="Enable this for dates such as 13/10/2026. ISO dates like 2026-10-13 are generally unambiguous.")

try:
    cleaned_source, analysis, quality = prepare_analysis(
        df_raw,
        sales_col=sales_col,
        date_col=date_col,
        category_col=category_col,
        product_col=product_col,
        order_col=order_col,
        quantity_col=quantity_col,
        unit_price_col=unit_price_col,
        revenue_mode="quantity_price" if revenue_mode == "Calculate quantity × unit price" else "column",
        remove_duplicates=remove_duplicates,
        trim_text=trim_text,
        dayfirst_dates=dayfirst_dates,
    )
except ValueError as exc:
    st.markdown(
        '<div class="hero"><div class="eyebrow">One small setup step</div><h1>Map your sales column to continue.</h1><p>Every business spreadsheet is a little different. Choose the column that contains sales/revenue, or select quantity × unit price in the sidebar.</p></div>',
        unsafe_allow_html=True,
    )
    st.warning(str(exc))
    st.stop()

# Data range and categories are global filters shared across every page.
with st.sidebar:
    st.markdown("### Report filters")
    valid_dates = analysis["_report_date"].dropna()
    if date_col != NONE_LABEL and not valid_dates.empty:
        min_day = valid_dates.min().date()
        max_day = valid_dates.max().date()
        range_default = (min_day, max_day)
        date_range = st.date_input("Date range", value=range_default, min_value=min_day, max_value=max_day, key=f"global_date_range_{dataset_signature}")
    else:
        date_range = None
        st.caption("Map a valid date column to enable date filtering.")
    categories = sorted(analysis["_report_category"].dropna().astype(str).unique().tolist())
    selected_categories = st.multiselect("Categories", categories, default=categories, key=f"global_categories_{dataset_signature}")

filtered = analysis.copy()
if date_range is not None and isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_day, end_day = date_range
    mask = filtered["_report_date"].notna() & (filtered["_report_date"].dt.date >= start_day) & (filtered["_report_date"].dt.date <= end_day)
    filtered = filtered.loc[mask].copy()
if selected_categories:
    filtered = filtered[filtered["_report_category"].astype(str).isin(selected_categories)].copy()
else:
    filtered = filtered.iloc[0:0].copy()

valid_sales = filtered["_report_sales"].dropna()
total_sales = float(valid_sales.sum()) if not valid_sales.empty else 0.0
average_sales = float(valid_sales.mean()) if not valid_sales.empty else 0.0
mapped_order_values = filtered.loc[filtered["_report_order"] != "Unspecified", "_report_order"]
unique_orders = int(mapped_order_values.nunique())
valid_rows = int(filtered["_report_sales"].notna().sum())

# Header is persistent across views.
source_display = source_name if len(source_name) < 52 else source_name[:49] + "…"
st.markdown(
    f'<div class="hero"><div class="eyebrow">SALES REPORTING</div><h1>A clear view of your sales.</h1><p>Review sales trends, check data quality and export a report from <b>{_safe_text(source_display)}</b>.</p></div>',
    unsafe_allow_html=True,
)
left, right = st.columns([3, 1])
with left:
    st.markdown(f"### {page}")
    st.caption(f"Showing {len(filtered):,} row(s) after your cleaning rules and current filters.")
with right:
    st.markdown(f'<div style="text-align:right;padding-top:8px"><span class="status-pill">● Data loaded</span></div>', unsafe_allow_html=True)

# ---------- Overview ----------
if page == "Overview":
    _kicker("At a glance")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total sales", _money(total_sales, currency_symbol), help="Sum of rows where sales could be interpreted as a number.")
    k2.metric("Rows analysed", f"{valid_rows:,}", help="Rows with a usable sales value in the filtered view.")
    k3.metric("Average sales / row", _money(average_sales, currency_symbol))
    k4.metric("Unique orders", f"{unique_orders:,}" if unique_orders else "—", help="Shown when an order/invoice column is mapped.")
    st.write("")

    insight_cols = st.columns(3)
    month_frame = filtered.dropna(subset=["_report_date", "_report_sales"]).copy()
    if not month_frame.empty:
        month_frame["_month"] = month_frame["_report_date"].dt.to_period("M").astype(str)
        monthly_series = month_frame.groupby("_month")["_report_sales"].sum().sort_index()
        peak_month = monthly_series.idxmax() if len(monthly_series) else "—"
        peak_sales = float(monthly_series.max()) if len(monthly_series) else 0.0
        with insight_cols[0]:
            _insight_card("Strongest month", peak_month, f"{_money(peak_sales, currency_symbol)} in recorded sales")
    else:
        with insight_cols[0]:
            _insight_card("Strongest month", "Not available", "Map dates and check valid date values")
    category_series = filtered.dropna(subset=["_report_sales"]).groupby("_report_category")["_report_sales"].sum().sort_values(ascending=False)
    with insight_cols[1]:
        if not category_series.empty:
            cat_name = str(category_series.index[0])
            _insight_card("Leading category", cat_name, f"{_money(float(category_series.iloc[0]), currency_symbol)} in recorded sales")
        else:
            _insight_card("Leading category", "No data", "No valid sales in current view")
    product_series = filtered.dropna(subset=["_report_sales"]).groupby("_report_product")["_report_sales"].sum().sort_values(ascending=False)
    with insight_cols[2]:
        if not product_series.empty:
            _insight_card("Top product", str(product_series.index[0]), f"{_money(float(product_series.iloc[0]), currency_symbol)} in recorded sales")
        else:
            _insight_card("Top product", "No data", "Map a product column for product insights")

    st.write("")
    chart_left, chart_right = st.columns([1.45, 1])
    with chart_left, st.container(border=True):
        st.markdown("#### Sales trend")
        if month_frame.empty:
            st.info("No valid date-and-sales rows are available. Check your date mapping and date values.")
        else:
            monthly = month_frame.groupby("_month", as_index=False)["_report_sales"].sum().rename(columns={"_month":"Month", "_report_sales":"Sales"})
            trend_style = st.radio("Trend style", ["Area", "Line", "Bars"], horizontal=True, label_visibility="collapsed", key="trend_style")
            if trend_style == "Area":
                fig = px.area(monthly, x="Month", y="Sales", markers=True, color_discrete_sequence=["#356859"])
                fig.update_traces(fillcolor="rgba(53,104,89,.14)", line=dict(width=3), marker=dict(size=7, color="#356859"),
                                  hovertemplate="<b>%{x}</b><br>Sales: " + currency_symbol + "%{y:,.2f}<extra></extra>")
            elif trend_style == "Line":
                fig = px.line(monthly, x="Month", y="Sales", markers=True, color_discrete_sequence=["#356859"])
                fig.update_traces(line=dict(width=3, shape="spline", smoothing=0.65), marker=dict(size=8, color="#356859", line=dict(width=2, color="#ffffff")),
                                  hovertemplate="<b>%{x}</b><br>Sales: " + currency_symbol + "%{y:,.2f}<extra></extra>")
            else:
                fig = px.bar(monthly, x="Month", y="Sales", color_discrete_sequence=["#356859"])
                fig.update_traces(marker_line_width=0,
                                  hovertemplate="<b>%{x}</b><br>Sales: " + currency_symbol + "%{y:,.2f}<extra></extra>")
            fig = _format_chart(fig, height=365)
            fig.update_layout(title="Monthly sales", yaxis_title="Sales", xaxis_title="", margin=dict(l=58, r=22, t=54, b=52), bargap=0.28)
            fig.update_yaxes(tickprefix=currency_symbol, tickformat="~s", title_text="Sales")
            fig.update_xaxes(tickangle=-20, nticks=9, title_text="")
            st.caption("Hover for exact values · use the chart toolbar to zoom or reset · switch styles above to explore the trend.")
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "responsive": True, "scrollZoom": True, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}, key="sales-trend-chart")
    with chart_right, st.container(border=True):
        st.markdown("#### Sales by category")
        if category_series.empty:
            st.info("Map a category column to compare categories.")
        else:
            category_chart = category_series.head(7).sort_values(ascending=True).rename_axis("Category").reset_index(name="Sales")
            category_chart["Display"] = category_chart["Category"].astype(str).map(lambda x: x if len(x) <= 22 else x[:19] + "…")
            category_chart["Share"] = category_chart["Sales"] / max(float(category_chart["Sales"].sum()), 1e-9) * 100
            fig = px.bar(category_chart, x="Sales", y="Display", orientation="h", text="Sales", custom_data=["Category", "Share"], color_discrete_sequence=["#356859"])
            fig.update_traces(texttemplate=currency_symbol + "%{x:,.0f}", textposition="outside", cliponaxis=False, marker_line_width=0, marker_color=CHART_PALETTE[:len(category_chart)],
                              hovertemplate="<b>%{customdata[0]}</b><br>Sales: " + currency_symbol + "%{x:,.2f}<br>Share of shown categories: %{customdata[1]:.1f}%<extra></extra>")
            fig = _format_chart(fig, height=max(340, 38*len(category_chart)+76))
            max_value = float(category_chart["Sales"].max()) if not category_chart.empty else 0
            fig.update_layout(title="Category performance", xaxis_title="Sales", yaxis_title="", margin=dict(l=104, r=80, t=54, b=48), showlegend=False, hovermode="closest")
            category_min = float(category_chart["Sales"].min()) if not category_chart.empty else 0
            category_floor = category_min * 1.15 if category_min < 0 else 0
            fig.update_xaxes(range=[category_floor, max_value * 1.32 if max_value > 0 else 1], tickprefix=currency_symbol, tickformat="~s", title_text="Sales")
            fig.update_yaxes(showgrid=False, tickfont=dict(size=12, color="#26332d", family="system-ui, sans-serif"))
            st.caption("Hover over a bar for the full category name, sales total and share.")
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "responsive": True, "scrollZoom": True}, key="category-performance-chart")

    bottom_left, bottom_right = st.columns([1.25, 1])
    with bottom_left, st.container(border=True):
        st.markdown("#### Top products by sales")
        if product_series.empty:
            st.info("Map a product column to see product performance.")
        else:
            top_products = product_series.head(8).sort_values().rename_axis("Product").reset_index(name="Sales")
            top_products["Display"] = top_products["Product"].astype(str).map(lambda x: x if len(x) <= 25 else x[:22] + "…")
            fig = px.bar(top_products, x="Sales", y="Display", orientation="h", text="Sales", custom_data=["Product"], color_discrete_sequence=["#356859"])
            fig.update_traces(texttemplate=currency_symbol + "%{x:,.0f}", textposition="outside", cliponaxis=False, marker_line_width=0, marker_color=CHART_PALETTE[:len(top_products)],
                              hovertemplate="<b>%{customdata[0]}</b><br>Sales: " + currency_symbol + "%{x:,.2f}<extra></extra>")
            fig = _format_chart(fig, height=max(310, 38*len(top_products)+85))
            max_product_sales = float(top_products["Sales"].max()) if not top_products.empty else 0
            fig.update_layout(title="Top products", xaxis_title="Sales", yaxis_title="", margin=dict(l=130, r=84, t=54, b=48), showlegend=False, hovermode="closest")
            product_min = float(top_products["Sales"].min()) if not top_products.empty else 0
            product_floor = product_min * 1.15 if product_min < 0 else 0
            fig.update_xaxes(range=[product_floor, max_product_sales * 1.34 if max_product_sales > 0 else 1], tickprefix=currency_symbol, tickformat="~s", title_text="Sales")
            fig.update_yaxes(showgrid=False, tickfont=dict(size=12, color="#26332d", family="system-ui, sans-serif"))
            st.caption("Long product names are shortened on the axis; hover over a bar to see the full name and exact sales.")
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "responsive": True, "scrollZoom": True}, key="top-products-chart")
    with bottom_right, st.container(border=True):
        st.markdown("#### Data health snapshot")
        health_items = [
            ("Exact duplicates found", quality["duplicate_rows_detected"], "Removed from analysis" if remove_duplicates else "Still included"),
            ("Rows with missing values", quality["rows_with_missing_values"], "Source-wide check"),
            ("Invalid sales values", quality["invalid_sales_values"], "Excluded from numeric sales metrics"),
            ("Unreadable dates", quality["invalid_dates"], "Date-based charts only"),
        ]
        for label, number, detail in health_items:
            c1, c2 = st.columns([2, 1])
            c1.markdown(f"**{label}**<div class='muted'>{_safe_text(detail)}</div>", unsafe_allow_html=True)
            c2.markdown(f"<div style='text-align:right;font-size:20px;color:{'#356859' if number == 0 else '#b57943'};font-weight:650'>{number:,}</div>", unsafe_allow_html=True)
            st.markdown("<hr style='margin:8px 0'>", unsafe_allow_html=True)
    st.caption("Insights are descriptive and depend on the available data. Sales are not the same as profit; refunds, taxes, discounts and cost data may need separate treatment.")

# ---------- Explore data ----------
elif page == "Explore data":
    _kicker("Inspect the rows behind the numbers")
    st.markdown("Use search and row limits to locate specific records in your current filtered dataset.")
    e1, e2, e3 = st.columns([2, 1, 1])
    with e1:
        search_term = st.text_input("Search all visible columns", placeholder="Try a product name, order ID, or category…")
    with e2:
        page_size = st.selectbox("Rows per page", [10, 25, 50, 100], index=1)
    visible = visible_data(filtered)
    if search_term.strip():
        search_mask = visible.astype("string").apply(lambda col: col.str.contains(search_term.strip(), case=False, na=False, regex=False)).any(axis=1)
        visible = visible.loc[search_mask]
    total_matches = len(visible)
    pages = max(1, (total_matches + page_size - 1) // page_size)
    with e3:
        page_number = st.number_input("Page", min_value=1, max_value=pages, value=1, step=1, disabled=(total_matches == 0))
    start = (int(page_number) - 1) * page_size
    st.caption(f"{total_matches:,} matching row(s) · page {page_number} of {pages}")
    if visible.empty:
        st.info("No rows match the current search and report filters.")
    else:
        st.dataframe(visible.iloc[start:start+page_size], width="stretch", hide_index=True, height=420)
    st.download_button(
        "Download current filtered rows (CSV)",
        data=visible.to_csv(index=False).encode("utf-8-sig"),
        file_name="filtered_business_data.csv",
        mime="text/csv",
        width="content",
    )

# ---------- Data quality ----------
elif page == "Data quality":
    _kicker("Trust the inputs before trusting the output")
    st.markdown("These checks describe the source dataset after the selected text-trimming and duplicate-removal rules. Missing values are reported, not automatically guessed or filled.")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Original rows", f"{quality['original_rows']:,}")
    q2.metric("Exact duplicates", f"{quality['duplicate_rows_detected']:,}")
    q3.metric("Rows with missing cells", f"{quality['rows_with_missing_values']:,}")
    q4.metric("Missing cells", f"{quality['missing_cells']:,}")

    st.markdown("#### Missing values by column")
    missing = quality["missing_by_column"].rename("Missing cells").to_frame()
    missing["Missing %"] = (missing["Missing cells"] / max(quality["original_rows"], 1) * 100).round(1)
    missing = missing.sort_values(["Missing cells", "Missing %"], ascending=False)
    st.dataframe(missing, width="stretch", height=min(360, 44 + 36*len(missing)))

    checks = pd.DataFrame([
        {"Check": "Exact duplicate rows", "Count": quality["duplicate_rows_detected"], "Interpretation": "Rows matching every source column. Similar orders are not necessarily duplicates."},
        {"Check": "Invalid sales values", "Count": quality["invalid_sales_values"], "Interpretation": "Sales couldn't be parsed as a number; excluded from sales metrics."},
        {"Check": "Invalid dates", "Count": quality["invalid_dates"], "Interpretation": "Date couldn't be interpreted; excluded from date-based charts."},
        {"Check": "Negative sales values", "Count": quality["negative_sales_rows"], "Interpretation": "May represent refunds, credits or adjustments; verify with the business."},
    ])
    st.markdown("#### Validation checks")
    st.dataframe(checks, width="stretch", hide_index=True)
    if quality["duplicate_rows_detected"]:
        st.warning(f"Found {quality['duplicate_rows_detected']:,} exact duplicate row(s). {'They are removed from analysis.' if remove_duplicates else 'They remain in analysis because duplicate removal is switched off.'}")
    if quality["invalid_sales_values"]:
        st.warning(f"{quality['invalid_sales_values']:,} sales value(s) could not be parsed. Inspect the source before using this as a formal report.")
    if quality["negative_sales_rows"]:
        st.info("Negative sales can be legitimate refunds or accounting adjustments. This app flags them but does not delete or reinterpret them.")
    if not (quality["duplicate_rows_detected"] or quality["invalid_sales_values"] or quality["invalid_dates"] or quality["missing_cells"]):
        st.success("No issues were detected by these basic validation checks. This does not guarantee that the source data is correct.")

# ---------- Export center ----------
elif page == "Export center":
    _kicker("Deliverables")
    st.markdown("Create shareable files from the current filtered view. Check the date/category filters in the sidebar before exporting.")
    x1, x2, x3 = st.columns(3)
    with x1, st.container(border=True):
        st.markdown("#### Cleaned data")
        st.caption("CSV containing the source columns for the current filtered rows.")
        st.download_button("Download filtered CSV", data=visible_data(filtered).to_csv(index=False).encode("utf-8-sig"), file_name="filtered_sales_data.csv", mime="text/csv", width="stretch")
    with x2, st.container(border=True):
        st.markdown("#### Excel report")
        st.caption("Multiple worksheets: overview, monthly trend, categories, top products, data quality and filtered rows.")
        try:
            excel_bytes = build_excel_report(filtered, source_name=source_name, quality=quality, currency_symbol=currency_symbol)
            st.download_button("Download Excel workbook", data=excel_bytes, file_name="business_performance_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
        except Exception as exc:
            st.error(f"Excel export could not be generated: {exc}")
    with x3, st.container(border=True):
        st.markdown("#### PDF summary")
        st.caption("A concise printable report with KPIs, observations, category performance and quality checks.")
        if date_range is not None and isinstance(date_range, (tuple, list)) and len(date_range) == 2:
            date_range_label = f"{date_range[0]} to {date_range[1]}"
        else:
            date_range_label = "All available dates / no date mapping"
        categories_label = ", ".join(selected_categories) if selected_categories else "None selected"
        try:
            pdf_bytes = build_pdf_report(filtered, source_name, quality, currency_symbol, date_range_label, categories_label)
            st.download_button("Download PDF report", data=pdf_bytes, file_name="business_performance_report.pdf", mime="application/pdf", width="stretch")
        except Exception as exc:
            st.error(f"PDF export could not be generated: {exc}")

    st.markdown("#### Before you share a report")
    st.info("Confirm the selected filters, verify totals against the source system, and explain any dropped duplicate rows or invalid values. This app provides descriptive analysis and is not a substitute for accounting review.")

st.markdown("---")
st.markdown(
    "<div style='display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;color:#536158;font-weight:700;font-size:11px'>"
    "<span>LEDGERLY · BUSINESS REPORT AUTOMATOR</span><span>Built for transparent, practical reporting · Always verify source figures</span></div>",
    unsafe_allow_html=True,
)
