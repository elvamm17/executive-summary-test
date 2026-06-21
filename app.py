from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src.config import MASTER_DATA_PATH
from src.pipeline import export_outputs, run_pipeline


ROOT_DIR = Path(__file__).resolve().parent
SAMPLE_MASTER_DATA_PATH = ROOT_DIR / "sample_master_data.xlsx"
BRAND_RED = "#d71920"
LIGHT_BG = "#f6f7f9"
TEXT = "#1f2933"
MUTED = "#6b7280"
PLATFORM_ORDER = ["Amazon", "TikTok", "Temu"]

st.set_page_config(page_title="Wei Long NA Management Review", layout="wide")

st.markdown(
    f"""
    <style>
    .stApp {{ background: {LIGHT_BG}; color: {TEXT}; }}
    .block-container {{ padding-top: 1.3rem; padding-bottom: 2rem; max-width: 1420px; }}
    h1, h2, h3 {{ color: {TEXT}; letter-spacing: 0; }}
    div[data-testid="stMetric"] {{
        background: white; border: 1px solid #eceff3; border-radius: 12px;
        padding: 16px 18px; box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }}
    .hero {{
        background: linear-gradient(135deg, #ffffff 0%, #fff5f5 100%);
        border: 1px solid #f0d7d9; border-radius: 18px; padding: 24px;
        box-shadow: 0 12px 30px rgba(215, 25, 32, 0.06);
    }}
    .section-title {{ font-size: 1.35rem; font-weight: 800; margin: 1.8rem 0 0.75rem 0; }}
    .kpi-label {{ color: {MUTED}; font-size: 0.84rem; text-transform: uppercase; letter-spacing: .04em; margin-bottom: .25rem; }}
    .kpi-value {{ font-size: 2.05rem; line-height: 1.1; font-weight: 850; color: {TEXT}; }}
    .kpi-sub {{ color: {MUTED}; font-size: .88rem; margin-top: .3rem; }}
    .card {{
        background: white; border: 1px solid #eceff3; border-radius: 14px; padding: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04); height: 100%;
    }}
    .pillar-title {{ font-size: 1.2rem; font-weight: 850; color: {BRAND_RED}; margin-bottom: .8rem; }}
    .progress-track {{ height: 13px; background: #e5e7eb; border-radius: 999px; overflow: hidden; margin: 8px 0 6px 0; }}
    .progress-fill {{ height: 13px; background: {BRAND_RED}; border-radius: 999px; }}
    .progress-row {{ display: flex; justify-content: space-between; color: {MUTED}; font-size: .9rem; }}
    .rank-card {{
        background: white; border: 1px solid #eceff3; border-radius: 14px; padding: 15px;
        margin-bottom: 10px; box-shadow: 0 8px 22px rgba(15, 23, 42, 0.035);
    }}
    .rank-num {{
        display: inline-block; background: {BRAND_RED}; color: white; border-radius: 999px;
        width: 30px; height: 30px; line-height: 30px; text-align: center; font-weight: 850; margin-right: 8px;
    }}
    .rank-title {{ font-size: 1.05rem; font-weight: 800; color: {TEXT}; }}
    .rank-meta {{ color: {MUTED}; font-size: .88rem; margin-top: 4px; }}
    .balanced-card {{
        min-height: 360px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .overview-number {{
        font-size: 2.35rem;
        line-height: 1.05;
        font-weight: 900;
        color: {BRAND_RED};
        margin: 0.15rem 0 0.75rem 0;
    }}
    .overview-target {{
        font-size: 1.65rem;
        line-height: 1.1;
        font-weight: 850;
        color: {TEXT};
        margin: 0.15rem 0 0 0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_master_data(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        return {}
    xls = pd.ExcelFile(path)
    return {sheet: pd.read_excel(path, sheet_name=sheet) for sheet in xls.sheet_names}


def resolve_master_data_path() -> Path:
    if MASTER_DATA_PATH.exists():
        return MASTER_DATA_PATH
    if SAMPLE_MASTER_DATA_PATH.exists():
        return SAMPLE_MASTER_DATA_PATH
    return MASTER_DATA_PATH


def money(value: float, decimals: int = 0) -> str:
    value = 0 if pd.isna(value) else float(value)
    return f"${value:,.{decimals}f}"


def number(value: float) -> str:
    value = 0 if pd.isna(value) else float(value)
    return f"{value:,.0f}"


def pct(value: float) -> str:
    return "-" if pd.isna(value) else f"{float(value):.1%}"


def ratio(value: float) -> str:
    return "-" if pd.isna(value) else f"{float(value):.2f}x"


def canonical_platform(value: object) -> str:
    text = "" if pd.isna(value) else str(value).strip()
    key = text.lower().replace(" ", "")
    if key == "tiktok":
        return "TikTok"
    if key == "amazon":
        return "Amazon"
    if key == "temu":
        return "Temu"
    return text


def title_case_flavor(value: object) -> str:
    text = "" if pd.isna(value) else str(value).strip()
    return " ".join(part.capitalize() for part in text.replace("_", " ").split())


def add_product_display_names(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    master = out.get("master_sku", pd.Series([""] * len(out), index=out.index)).fillna("").astype(str).str.strip()
    flavor_raw = out.get("flavor", pd.Series([""] * len(out), index=out.index)).fillna("").astype(str).str.strip()
    flavor = flavor_raw.map(title_case_flavor)
    pack = out.get("pack_size", pd.Series([""] * len(out), index=out.index)).fillna("").astype(str).str.strip()
    missing = master.eq("") | master.str.lower().isin(["nan", "none", "unknown", "unmatched"]) | flavor_raw.eq("") | pack.eq("")
    out["sku_title"] = (flavor + " " + pack).str.replace(r"\s+", " ", regex=True).str.strip()
    out["sku_title"] = out["sku_title"].mask(missing | out["sku_title"].str.lower().eq("other"), "Unmapped SKU")
    out["product_display_name"] = out["sku_title"]
    return out


def ensure_date(df: pd.DataFrame, col: str = "order_date") -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df
    out = df.copy()
    out[col] = pd.to_datetime(out[col], errors="coerce")
    if "month" not in out.columns:
        out["month"] = out[col].dt.to_period("M").astype(str)
    return out


def prep_target_table(target_table: pd.DataFrame) -> pd.DataFrame:
    if target_table.empty:
        return pd.DataFrame(columns=["month", "platform", "target_gmv"])
    out = target_table.copy()
    out["platform"] = out["platform"].map(canonical_platform)
    out["month"] = pd.to_datetime(out["month"], errors="coerce").dt.to_period("M").astype(str)
    out["target_gmv"] = pd.to_numeric(out["target_gmv"], errors="coerce").fillna(0)
    return out


def apply_filters(fact_order: pd.DataFrame, fact_product: pd.DataFrame, product_expanded: pd.DataFrame):
    st.sidebar.title("Management Review")
    platforms = [p for p in PLATFORM_ORDER if p in set(fact_order.get("platform", pd.Series(dtype=str)).dropna())]
    selected_platforms = st.sidebar.multiselect("Platform", platforms, default=platforms)

    dates = fact_order["order_date"].dropna() if "order_date" in fact_order.columns else pd.Series(dtype="datetime64[ns]")
    if not dates.empty:
        date_value = st.sidebar.date_input("Date Range", [dates.min().date(), dates.max().date()])
        if isinstance(date_value, tuple) and len(date_value) == 2:
            start_ts = pd.Timestamp(date_value[0])
            end_ts = pd.Timestamp(date_value[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        else:
            start_ts, end_ts = dates.min(), dates.max()
    else:
        start_ts, end_ts = pd.Timestamp.min, pd.Timestamp.max

    def filter_df(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        out = df.copy()
        if selected_platforms and "platform" in out.columns:
            out = out[out["platform"].isin(selected_platforms)]
        if "order_date" in out.columns:
            out = out[(out["order_date"].isna()) | ((out["order_date"] >= start_ts) & (out["order_date"] <= end_ts))]
        return out

    return filter_df(fact_order), filter_df(fact_product), filter_df(product_expanded)


def active_month(fact_order: pd.DataFrame) -> str:
    dates = fact_order["order_date"].dropna() if "order_date" in fact_order.columns else pd.Series(dtype="datetime64[ns]")
    if dates.empty:
        return pd.Timestamp.today().to_period("M").strftime("%Y-%m")
    return dates.max().to_period("M").strftime("%Y-%m")


def progress_bar(label: str, value: float, cap: float = 1.0):
    raw = 0 if pd.isna(value) else float(value)
    width = max(0, min(raw / cap, 1)) * 100
    st.markdown(
        f"""
        <div class="progress-row"><span>{label}</span><strong>{pct(raw)}</strong></div>
        <div class="progress-track"><div class="progress-fill" style="width:{width:.1f}%"></div></div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value: str, sub: str = ""):
    st.markdown(
        f"""
        <div class="card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def platform_summary(fact_order: pd.DataFrame, target_table: pd.DataFrame, month_key: str) -> pd.DataFrame:
    if fact_order.empty:
        return pd.DataFrame(columns=["platform", "gmv", "orders", "aov", "target_gmv", "achievement"])
    data = fact_order.groupby("platform", as_index=False).agg(gmv=("gmv_incl_shipping_tax", "sum"), orders=("order_id", "nunique"))
    data["aov"] = data["gmv"] / data["orders"].replace(0, pd.NA)
    targets = target_table[target_table["month"].eq(month_key)].groupby("platform", as_index=False)["target_gmv"].sum()
    data = data.merge(targets, on="platform", how="left").fillna({"target_gmv": 0})
    data["achievement"] = data["gmv"] / data["target_gmv"].replace(0, pd.NA)
    order = {p: i for i, p in enumerate(PLATFORM_ORDER)}
    data["sort"] = data["platform"].map(order).fillna(99)
    return data.sort_values("sort").drop(columns="sort")


def business_overview(fact_order: pd.DataFrame, target_table: pd.DataFrame):
    st.title("Business Overview")
    st.caption("Summary -> Platform -> Product -> Geo")
    if fact_order.empty:
        st.info("No order data loaded yet.")
        return

    month_key = active_month(fact_order)
    total_gmv = fact_order["gmv_incl_shipping_tax"].sum()
    orders = fact_order["order_id"].nunique()
    monthly_target = target_table[target_table["month"].eq(month_key)]["target_gmv"].sum()
    achievement = total_gmv / monthly_target if monthly_target else pd.NA
    today = pd.Timestamp.today()
    if month_key == today.to_period("M").strftime("%Y-%m"):
        date_progress = today.day / today.days_in_month
    else:
        month_dates = fact_order.loc[fact_order["month"].eq(month_key), "order_date"].dropna()
        date_progress = month_dates.dt.day.max() / month_dates.dt.days_in_month.max() if not month_dates.empty else pd.NA
    platform_mix = fact_order.groupby("platform", as_index=False)["gmv_incl_shipping_tax"].sum()

    c1, c2 = st.columns([0.45, 0.55])
    with c1:
        st.markdown('<div class="card balanced-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="pillar-title">Business Progress</div><div class="kpi-label">GMV Achievement</div><div class="kpi-value" style="color:{BRAND_RED}">{pct(achievement)}</div>', unsafe_allow_html=True)
        progress_bar("Achievement", achievement)
        st.markdown(
            f"""
            <div class="kpi-label" style="margin-top:1.3rem">Total GMV</div>
            <div class="overview-number">{money(total_gmv)}</div>
            <div class="kpi-label">Monthly Target</div>
            <div class="overview-target">{money(monthly_target)}</div>
            <div class="kpi-sub">{month_key} | {orders:,} orders</div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card balanced-card">', unsafe_allow_html=True)
        fig = px.pie(platform_mix, names="platform", values="gmv_incl_shipping_tax", hole=0.55, title="GMV Mix by Platform")
        fig.update_traces(textinfo="percent+label")
        fig.update_layout(
            showlegend=True,
            legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center"),
            margin=dict(l=10, r=10, t=52, b=28),
            height=360,
            paper_bgcolor="white",
            plot_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Platform Pillars</div>', unsafe_allow_html=True)
    pillars = platform_summary(fact_order, target_table, month_key)
    cols = st.columns(3)
    for idx, platform in enumerate(PLATFORM_ORDER):
        item = pillars[pillars["platform"].eq(platform)]
        row = item.iloc[0] if not item.empty else pd.Series({"gmv": 0, "orders": 0, "aov": 0, "target_gmv": 0, "achievement": pd.NA})
        with cols[idx]:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f'<div class="pillar-title">{platform}</div>', unsafe_allow_html=True)
            a, b, c = st.columns(3)
            a.metric("GMV", money(row["gmv"]))
            b.metric("Orders", number(row["orders"]))
            c.metric("AOV", money(row["aov"], 2))
            progress_bar("Achievement", row["achievement"])
            st.caption(f"Target {money(row['target_gmv'])}")
            st.markdown("</div>", unsafe_allow_html=True)


def donut(df: pd.DataFrame, names: str, values: str, title: str):
    fig = px.pie(df, names=names, values=values, hole=0.58, title=title)
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(showlegend=False, margin=dict(l=10, r=10, t=50, b=10), paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)


def mix_data(df: pd.DataFrame, dim: str, value: str = "gmv_product", top_n: int = 8) -> pd.DataFrame:
    if df.empty or dim not in df.columns:
        return pd.DataFrame(columns=[dim, value])
    out = df.copy()
    out[dim] = out[dim].fillna("").replace("", "Unmapped SKU")
    out[dim] = out[dim].mask(out[dim].astype(str).str.lower().isin(["nan", "none", "unknown", "unmatched", "other"]), "Unmapped SKU")
    out = out.groupby(dim, as_index=False)[value].sum().sort_values(value, ascending=False)
    return out.head(top_n)


def leaderboard_card(rank: int, row: pd.Series):
    st.markdown(
        f"""
        <div class="rank-card">
            <span class="rank-num">{rank}</span><span class="rank-title">{row.get('sku_title', row.get('master_sku', 'Unclassified'))}</span>
            <div class="rank-meta">{row.get('platform', '')} | Master SKU {row.get('master_sku', '')}</div>
            <div class="rank-meta">GMV incl. shipping {money(row.get('gmv_incl_shipping_tax', row.get('gmv_product', 0)))} | Product GMV {money(row.get('gmv_product', 0))} | Qty {number(row.get('quantity', 0))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def product_performance(product_expanded: pd.DataFrame):
    st.title("Product Performance")
    if product_expanded.empty:
        st.info("No product data loaded yet.")
        return

    data = add_product_display_names(product_expanded)
    data["flavor_en"] = data["flavor"].fillna("") if "flavor" in data.columns else ""
    qty_col = "analysis_quantity" if "analysis_quantity" in data.columns else "quantity"

    st.markdown('<div class="section-title">All Platform Mix</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        sku_mix = mix_data(data, "sku_title", value="gmv_incl_shipping_tax", top_n=10).sort_values("gmv_incl_shipping_tax", ascending=True)
        fig = px.bar(sku_mix, x="gmv_incl_shipping_tax", y="sku_title", orientation="h", title="SKU Mix - GMV includes shipping", text_auto=".2s")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=430, paper_bgcolor="white", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        donut(mix_data(data, "flavor_en", value="gmv_incl_shipping_tax"), "flavor_en", "gmv_incl_shipping_tax", "Flavor Mix - GMV includes shipping")

    st.markdown('<div class="section-title">By Platform</div>', unsafe_allow_html=True)
    tabs = st.tabs(PLATFORM_ORDER)
    for tab, platform in zip(tabs, PLATFORM_ORDER):
        with tab:
            subset = data[data["platform"].eq(platform)]
            if subset.empty:
                st.info(f"No {platform} product data.")
                continue
            c1, c2, c3 = st.columns(3)
            with c1:
                sku_mix = mix_data(subset, "sku_title", value="gmv_incl_shipping_tax", top_n=8).sort_values("gmv_incl_shipping_tax", ascending=True)
                fig = px.bar(sku_mix, x="gmv_incl_shipping_tax", y="sku_title", orientation="h", title="SKU Mix - GMV includes shipping", text_auto=".2s")
                fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=390, paper_bgcolor="white", yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                donut(mix_data(subset, "flavor_en", value="gmv_incl_shipping_tax"), "flavor_en", "gmv_incl_shipping_tax", "Flavor Mix - GMV includes shipping")
            with c3:
                pack = mix_data(subset, "pack_size", "gmv_incl_shipping_tax", top_n=10)
                fig = px.bar(pack, x="gmv_incl_shipping_tax", y="pack_size", orientation="h", title="Pack Size Mix - GMV includes shipping", text_auto=".2s")
                fig.update_layout(yaxis={"categoryorder": "total ascending"}, margin=dict(l=10, r=10, t=50, b=10), paper_bgcolor="white")
                st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">Top SKU Leaderboard</div>', unsafe_allow_html=True)
    cols = st.columns(3)
    for idx, platform in enumerate(PLATFORM_ORDER):
        subset = data[data["platform"].eq(platform)]
        top = subset.groupby(["platform", "master_sku", "sku_title", "flavor_en", "pack_size"], dropna=False, as_index=False).agg(
            gmv_incl_shipping_tax=("gmv_incl_shipping_tax", "sum"),
            gmv_product=("gmv_product", "sum"),
            quantity=(qty_col, "sum"),
        ).sort_values("gmv_incl_shipping_tax", ascending=False).head(3)
        with cols[idx]:
            st.markdown(f'<div class="pillar-title">{platform}</div>', unsafe_allow_html=True)
            if top.empty:
                st.info("No data.")
            for rank, (_, row) in enumerate(top.iterrows(), start=1):
                leaderboard_card(rank, row)


def geo_analysis(fact_order: pd.DataFrame, product_expanded: pd.DataFrame):
    st.title("Geo Analysis")
    if fact_order.empty or "state" not in fact_order.columns:
        st.info("No state-level data loaded yet.")
        return

    platform = st.radio("Platform", ["All"] + PLATFORM_ORDER, horizontal=True)
    geo = fact_order.copy()
    if platform != "All":
        geo = geo[geo["platform"].eq(platform)]
    geo["state"] = geo["state"].fillna("").astype(str)
    geo = geo[geo["state"].ne("")]
    geo = geo[~((geo["platform"].eq("Temu")) & (geo["state"].eq("MA")))]
    by_state = geo.groupby("state", as_index=False).agg(gmv=("gmv_incl_shipping_tax", "sum"), orders=("order_id", "nunique"))
    by_state["aov"] = by_state["gmv"] / by_state["orders"].replace(0, pd.NA)

    fig = px.choropleth(
        by_state,
        locations="state",
        locationmode="USA-states",
        color="gmv",
        scope="usa",
        color_continuous_scale=["#fff1f2", BRAND_RED],
        hover_data={"gmv": ":,.0f", "orders": True, "aov": ":.2f"},
        title="GMV by State",
    )
    fig.update_layout(margin=dict(l=0, r=0, t=48, b=0), paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">State Insights</div>', unsafe_allow_html=True)
    states = sorted(by_state["state"].dropna().unique())
    if not states:
        st.info("No eligible state data. Temu MA is excluded from map analysis per rule.")
        return
    selected_state = st.selectbox("Select State", states, index=0)
    st.subheader(selected_state)

    product = product_expanded.copy()
    if platform != "All":
        product = product[product["platform"].eq(platform)]
    product["state"] = product["state"].fillna("").astype(str)
    product = product[(product["state"].eq(selected_state)) & ~(product["platform"].eq("Temu") & product["state"].eq("MA"))]
    product["flavor_en"] = product["flavor"].fillna("") if "flavor" in product.columns else ""
    qty_col = "analysis_quantity" if "analysis_quantity" in product.columns else "quantity"
    c1, c2 = st.columns(2)
    top_sku = product.groupby("master_sku", dropna=False, as_index=False).agg(gmv=("gmv_product", "sum"), quantity=(qty_col, "sum")).sort_values("gmv", ascending=False).head(3)
    top_flavor = product.groupby("flavor_en", dropna=False, as_index=False).agg(gmv=("gmv_product", "sum"), quantity=(qty_col, "sum")).sort_values("gmv", ascending=False).head(3)
    with c1:
        st.markdown('<div class="pillar-title">Top SKU</div>', unsafe_allow_html=True)
        for rank, (_, row) in enumerate(top_sku.iterrows(), start=1):
            leaderboard_card(rank, pd.Series({"master_sku": row["master_sku"], "platform": selected_state, "flavor_en": "", "pack_size": "", "gmv_product": row["gmv"], "quantity": row["quantity"]}))
    with c2:
        st.markdown('<div class="pillar-title">Top Flavor</div>', unsafe_allow_html=True)
        for rank, (_, row) in enumerate(top_flavor.iterrows(), start=1):
            leaderboard_card(rank, pd.Series({"master_sku": row["flavor_en"], "platform": selected_state, "flavor_en": "", "pack_size": "", "gmv_product": row["gmv"], "quantity": row["quantity"]}))


def bundle_performance(fact_product: pd.DataFrame, bundle_mapping: pd.DataFrame):
    st.title("Bundle Performance")
    if bundle_mapping.empty:
        st.info("Bundle Mapping 尚未补充，后续用于分析各平台组包表现。")
        return
    data = fact_product.copy()
    data["bundle_platform_sku_id"] = data["platform_sku"].fillna("")
    data["bundle_name"] = data.get("bundle_name", data["bundle_platform_sku_id"]).fillna("")
    bundle = data.groupby(["platform", "bundle_platform_sku_id", "bundle_name"], dropna=False, as_index=False).agg(
        gmv=("gmv_incl_shipping_tax", "sum"),
        orders=("order_id", "nunique"),
        quantity=("quantity", "sum"),
    )
    bundle["aov"] = bundle["gmv"] / bundle["orders"].replace(0, pd.NA)
    fig = px.bar(bundle.sort_values("gmv", ascending=False).head(20), x="bundle_platform_sku_id", y="gmv", color="platform", title="Bundle GMV", text_auto=".2s")
    st.plotly_chart(fig, use_container_width=True)


def product_quadrant(product_expanded: pd.DataFrame, unit_economics: pd.DataFrame | None = None):
    st.title("Product Quadrant")
    if product_expanded.empty:
        st.info("No product data loaded yet.")
        return
    data = product_expanded.copy()
    sku_col = "master_sku" if "master_sku" in data.columns else "erp_sku"
    qty_col = "analysis_quantity" if "analysis_quantity" in data.columns else "quantity"
    quad = data.groupby(sku_col, dropna=False, as_index=False).agg(gmv_product=("gmv_product", "sum"), quantity=(qty_col, "sum"), orders=("order_id", "nunique"))
    use_profit = False
    if unit_economics is not None and not unit_economics.empty and {"master_sku", "profit_margin"}.issubset(unit_economics.columns):
        ue = unit_economics.copy()
        ue["profit_margin"] = pd.to_numeric(ue["profit_margin"], errors="coerce")
        quad = quad.merge(ue[["master_sku", "profit_margin"]].drop_duplicates("master_sku"), left_on=sku_col, right_on="master_sku", how="left")
        use_profit = quad["profit_margin"].notna().any()
    y_col = "profit_margin" if use_profit else "quantity"
    st.caption("Profit Quadrant: x = GMV Product, y = Profit Margin." if use_profit else "Sales Performance Quadrant: x = GMV Product, y = Quantity. 利润四象限待 Unit_Economics / 成本表补齐后启用。")
    fig = px.scatter(quad, x="gmv_product", y=y_col, size="orders", hover_name=sku_col, text=sku_col, title="Product Quadrant")
    fig.add_vline(x=quad["gmv_product"].mean(), line_dash="dash", line_color="gray")
    fig.add_hline(y=quad[y_col].mean(), line_dash="dash", line_color="gray")
    fig.update_traces(textposition="top center")
    st.plotly_chart(fig, use_container_width=True)


st.title("Wei Long North America Management Review")

with st.sidebar:
    if st.button("Run Data Pipeline", type="primary"):
        with st.spinner("Cleaning platform data and exporting master_data.xlsx..."):
            result = run_pipeline()
            export_outputs(result)
            st.cache_data.clear()
        st.success("Pipeline complete.")

data_path = resolve_master_data_path()
data = load_master_data(data_path)
if not data:
    st.warning("No master data workbook was found.")
    st.markdown(
        """
        本地运行请先执行：

        ```bash
        python run_pipeline.py
        ```

        Streamlit Community Cloud 部署时，请上传脱敏后的 `output/master_data.xlsx`，
        或保留仓库根目录的 `sample_master_data.xlsx` 作为 Demo 数据。
        """
    )
    st.stop()

if data_path.name == "sample_master_data.xlsx":
    st.sidebar.info("Using sample_master_data.xlsx for demo deployment.")
else:
    st.sidebar.caption(f"Data source: {data_path}")

fact_order = ensure_date(data.get("Fact_Order", pd.DataFrame()))
fact_product = ensure_date(data.get("Fact_Product", pd.DataFrame()))
product_expanded = ensure_date(data.get("Product_Expanded", pd.DataFrame()))
target_table = prep_target_table(data.get("Target_Table", pd.DataFrame()))
bundle_mapping = data.get("Bundle_Mapping", pd.DataFrame())
unit_economics = data.get("Unit_Economics", pd.DataFrame())

fact_order, fact_product, product_expanded = apply_filters(fact_order, fact_product, product_expanded)

page = st.sidebar.radio("Page", ["Business Overview", "Product Performance", "Geo Analysis", "Bundle Performance", "Product Quadrant"])

if page == "Business Overview":
    business_overview(fact_order, target_table)
elif page == "Product Performance":
    product_performance(product_expanded)
elif page == "Geo Analysis":
    geo_analysis(fact_order, product_expanded)
elif page == "Bundle Performance":
    bundle_performance(fact_product, bundle_mapping)
else:
    product_quadrant(product_expanded, unit_economics)
