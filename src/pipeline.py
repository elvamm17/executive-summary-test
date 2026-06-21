from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .config import (
    AMAZON_PATTERNS,
    CONFIG_DIR,
    MASTER_DATA_PATH,
    OUTPUT_DIR,
    RAW_DIR,
    TEMU_GUCANG_PATTERNS,
    TEMU_ORDER_INFO_PATTERNS,
    TEMU_SETTLED_PATTERNS,
    TEMU_UNSETTLED_PATTERNS,
    TIKTOK_PATTERNS,
    UNMATCHED_SKU_PATH,
)
from .io_utils import (
    date_series,
    detect_header_row,
    find_files_any,
    first_existing,
    money_series,
    normalize_columns,
    read_config_csv,
    read_excel_sheet,
    read_table,
    text_series,
    value_series,
)
from .sku_utils import attach_sku_master, clean_platform_sku, map_platform_sku, split_batch_order_numbers


PRODUCT_COLUMNS = [
    "platform",
    "order_id",
    "order_date",
    "state",
    "platform_sku",
    "clean_platform_sku",
    "erp_sku",
    "master_sku",
    "product_family",
    "category",
    "flavor",
    "pack_size",
    "quantity",
    "gmv_product",
    "shipping_revenue",
    "gmv_incl_shipping_tax",
    "platform_discount",
    "seller_discount",
    "settlement_status",
    "source_file",
]

ORDER_COLUMNS = [
    "platform",
    "order_id",
    "order_date",
    "state",
    "quantity",
    "gmv_product",
    "shipping_revenue",
    "gmv_incl_shipping_tax",
    "aov",
    "settlement_status",
    "source_file",
]


STATE_ALIASES = {
    "AL": "AL", "ALABAMA": "AL", "阿拉巴马": "AL",
    "AK": "AK", "ALASKA": "AK", "阿拉斯加": "AK",
    "AZ": "AZ", "ARIZONA": "AZ", "亚利桑那": "AZ",
    "AR": "AR", "ARKANSAS": "AR", "阿肯色": "AR",
    "CA": "CA", "CALIFORNIA": "CA", "加利福尼亚": "CA", "加州": "CA",
    "CO": "CO", "COLORADO": "CO", "科罗拉多": "CO",
    "CT": "CT", "CONNECTICUT": "CT", "康涅狄格": "CT",
    "DE": "DE", "DELAWARE": "DE", "特拉华": "DE",
    "FL": "FL", "FLORIDA": "FL", "佛罗里达": "FL",
    "GA": "GA", "GEORGIA": "GA", "佐治亚": "GA", "乔治亚": "GA",
    "HI": "HI", "HAWAII": "HI", "夏威夷": "HI",
    "ID": "ID", "IDAHO": "ID", "爱达荷": "ID",
    "IL": "IL", "ILLINOIS": "IL", "伊利诺伊": "IL",
    "IN": "IN", "INDIANA": "IN", "印第安纳": "IN",
    "IA": "IA", "IOWA": "IA", "艾奥瓦": "IA", "爱荷华": "IA",
    "KS": "KS", "KANSAS": "KS", "堪萨斯": "KS",
    "KY": "KY", "KENTUCKY": "KY", "肯塔基": "KY",
    "LA": "LA", "LOUISIANA": "LA", "路易斯安那": "LA",
    "ME": "ME", "MAINE": "ME", "缅因": "ME",
    "MD": "MD", "MARYLAND": "MD", "马里兰": "MD",
    "MA": "MA", "MASSACHUSETTS": "MA", "马萨诸塞": "MA",
    "MI": "MI", "MICHIGAN": "MI", "密歇根": "MI",
    "MN": "MN", "MINNESOTA": "MN", "明尼苏达": "MN",
    "MS": "MS", "MISSISSIPPI": "MS", "密西西比": "MS",
    "MO": "MO", "MISSOURI": "MO", "密苏里": "MO",
    "MT": "MT", "MONTANA": "MT", "蒙大拿": "MT",
    "NE": "NE", "NEBRASKA": "NE", "内布拉斯加": "NE",
    "NV": "NV", "NEVADA": "NV", "内华达": "NV",
    "NH": "NH", "NEW HAMPSHIRE": "NH", "新罕布什尔": "NH",
    "NJ": "NJ", "NEW JERSEY": "NJ", "新泽西": "NJ",
    "NM": "NM", "NEW MEXICO": "NM", "新墨西哥": "NM",
    "NY": "NY", "NEW YORK": "NY", "纽约": "NY",
    "NC": "NC", "NORTH CAROLINA": "NC", "北卡罗来纳": "NC",
    "ND": "ND", "NORTH DAKOTA": "ND", "北达科他": "ND",
    "OH": "OH", "OHIO": "OH", "俄亥俄": "OH",
    "OK": "OK", "OKLAHOMA": "OK", "俄克拉荷马": "OK",
    "OR": "OR", "OREGON": "OR", "俄勒冈": "OR",
    "PA": "PA", "PENNSYLVANIA": "PA", "宾夕法尼亚": "PA",
    "RI": "RI", "RHODE ISLAND": "RI", "罗德岛": "RI",
    "SC": "SC", "SOUTH CAROLINA": "SC", "南卡罗来纳": "SC",
    "SD": "SD", "SOUTH DAKOTA": "SD", "南达科他": "SD",
    "TN": "TN", "TENNESSEE": "TN", "田纳西": "TN",
    "TX": "TX", "TEXAS": "TX", "德克萨斯": "TX", "德州": "TX",
    "UT": "UT", "UTAH": "UT", "犹他": "UT",
    "VT": "VT", "VERMONT": "VT", "佛蒙特": "VT",
    "VA": "VA", "VIRGINIA": "VA", "弗吉尼亚": "VA",
    "WA": "WA", "WASHINGTON": "WA", "华盛顿": "WA",
    "WV": "WV", "WEST VIRGINIA": "WV", "西弗吉尼亚": "WV",
    "WI": "WI", "WISCONSIN": "WI", "威斯康星": "WI",
    "WY": "WY", "WYOMING": "WY", "怀俄明": "WY",
    "DC": "DC", "DISTRICT OF COLUMBIA": "DC", "华盛顿特区": "DC",
}


def normalize_state(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    key = text.upper().replace(".", "").strip()
    return STATE_ALIASES.get(key, STATE_ALIASES.get(text, text if len(text) == 2 and text.isalpha() else ""))


def normalize_state_series(series: pd.Series) -> pd.Series:
    return series.fillna("").map(normalize_state)


def zero_small(series: pd.Series, threshold: float = 0.01) -> pd.Series:
    out = pd.to_numeric(series, errors="coerce").fillna(0.0)
    return out.mask(out.abs().lt(threshold), 0.0)


def safe_join_unique(series: pd.Series, sep: str = "/") -> str:
    values = []
    for value in series:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not text or text.lower() in {"nan", "none", "null"}:
            continue
        values.append(text)
    if not values:
        return ""
    return sep.join(sorted(set(values)))


@dataclass
class PipelineResult:
    fact_order: pd.DataFrame
    fact_product: pd.DataFrame
    sku_master: pd.DataFrame
    platform_sku_mapping: pd.DataFrame
    bundle_mapping: pd.DataFrame
    target_table: pd.DataFrame
    product_expanded: pd.DataFrame
    unmatched_sku: pd.DataFrame
    source_notes: pd.DataFrame


def load_config() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sku_master = _load_sku_master()
    platform_sku_mapping = _load_optional_config(
        ("platform_sku_mapping", "platform sku mapping", "平台sku映射", "sku_mapping"),
        ["platform", "platform_sku", "clean_platform_sku", "erp_sku", "master_sku", "notes"],
    )
    bundle_mapping = _load_optional_config(
        ("bundle_mapping", "bundle mapping", "组包", "bundle"),
        ["platform", "platform_sku", "erp_sku", "bundle_sku", "bundle_name", "component_erp_sku", "component_qty"],
    )
    target_table = _load_target_table()
    target_table["month"] = pd.to_datetime(target_table["month"], errors="coerce").dt.to_period("M").astype(str)
    target_table["target_gmv"] = pd.to_numeric(target_table["target_gmv"], errors="coerce").fillna(0.0)
    return sku_master, platform_sku_mapping, bundle_mapping, target_table


def _find_config_file(patterns: tuple[str, ...]) -> Path | None:
    normalized = [p.lower().replace("_", " ").replace("-", " ") for p in patterns]
    for path in sorted(CONFIG_DIR.glob("*")):
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        name = path.stem.lower().replace("_", " ").replace("-", " ")
        if any(p in name for p in normalized):
            return path
    return None


def _standardize_sku_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def _load_sku_master() -> pd.DataFrame:
    columns = ["erp_sku", "master_sku", "product_family", "category", "flavor", "pack_size", "unit_size", "case_pack", "status"]
    path = _find_config_file(("sku master", "sku_master", "sku"))
    if path is None:
        return pd.DataFrame(columns=columns)
    if path.suffix.lower() == ".csv":
        header = detect_header_row(path, ["erp_sku"])
        df = pd.read_csv(path, header=header, encoding="utf-8-sig")
    else:
        header = detect_header_row(path, ["erp_sku"])
        df = pd.read_excel(path, header=header)
    df = normalize_columns(df)
    rename = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in {"erp_sku", "erp sku", "商品编码"}:
            rename[col] = "erp_sku"
        elif key in {"master_sku", "master sku"}:
            rename[col] = "master_sku"
        elif key in {"product_family", "family"}:
            rename[col] = "product_family"
        elif key == "category":
            rename[col] = "category"
        elif key in {"flavor", "flavor_en", "口味"}:
            rename[col] = "flavor"
        elif key == "pack_size":
            rename[col] = "pack_size"
        elif key in {"规格", "unit_size"}:
            rename[col] = "unit_size"
        elif "箱规" in key or key == "case_pack":
            rename[col] = "case_pack"
        elif "状态" in key or key == "status":
            rename[col] = "status"
    df = df.rename(columns=rename)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns].copy()
    df["erp_sku"] = _standardize_sku_text(df["erp_sku"])
    df = df[df["erp_sku"].ne("")]
    df["master_sku"] = _standardize_sku_text(df["master_sku"]).mask(lambda s: s.eq(""), df["erp_sku"])
    for col in ("product_family", "category"):
        df[col] = df[col].ffill().fillna("")
    for col in ("flavor", "pack_size", "unit_size", "case_pack", "status"):
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df.drop_duplicates("erp_sku")


def _load_optional_config(patterns: tuple[str, ...], columns: list[str]) -> pd.DataFrame:
    path = _find_config_file(patterns)
    if path is None:
        return pd.DataFrame(columns=columns)
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    else:
        df = pd.read_excel(path, dtype=str).fillna("")
    df = normalize_columns(df)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[columns]


def _load_target_table() -> pd.DataFrame:
    columns = ["month", "platform", "target_gmv"]
    path = _find_config_file(("target_table", "target table", "target"))
    if path is None:
        return pd.DataFrame(columns=columns)
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, encoding="utf-8-sig")
    else:
        df = pd.read_excel(path)
    df = normalize_columns(df)
    if "platform" not in df.columns and "channel" in df.columns:
        df = df.rename(columns={"channel": "platform"})
    if "target_gmv" not in df.columns:
        if "target_usd" in df.columns:
            df = df.rename(columns={"target_usd": "target_gmv"})
        elif "target_rmb" in df.columns:
            df = df.rename(columns={"target_rmb": "target_gmv"})
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns].copy()
    df = df[df["platform"].fillna("").astype(str).str.strip().ne("")]
    df["platform"] = df["platform"].astype(str).str.strip().str.title()
    return df


def clean_amazon(paths: list[Path], mapping: pd.DataFrame, sku_master: pd.DataFrame) -> pd.DataFrame:
    frames = []
    for path in paths:
        header = detect_header_row(path, ["order id", "sku", "product sales"])
        df = pd.read_csv(path, header=header, encoding="utf-8-sig") if path.suffix.lower() == ".csv" else pd.read_excel(path, header=header)
        df = normalize_columns(df)
        if df.empty:
            continue
        out = pd.DataFrame(index=df.index)
        out["platform"] = "Amazon"
        out["order_id"] = text_series(df, ["order_id", "amazon-order-id", "Amazon Order Id", "Order ID", "order id"])
        out["order_date"] = date_series(df, ["purchase-date", "Purchase Date", "order_date", "Order Date", "date/time", "Date"])
        out["state"] = normalize_state_series(text_series(df, ["orderstate", "order state", "ship-state", "Ship State", "state", "State"]))
        out["platform_sku"] = text_series(df, ["sku", "SKU", "merchant-sku", "Merchant SKU"])
        out["clean_platform_sku"] = out["platform_sku"].map(clean_platform_sku)
        out["erp_sku"] = out["clean_platform_sku"]
        out["quantity"] = pd.to_numeric(value_series(df, ["quantity-purchased", "Quantity", "quantity", "qty"], 1), errors="coerce").fillna(1)
        product_sales = money_series(df, ["product_sales", "product sales", "Product Sales", "item-price", "Item Price"])
        rebates = money_series(df, ["promotional_rebates", "promotional rebates", "Promotional Rebates", "promotion-ids amount"])
        shipping = money_series(df, ["shipping_credits", "shipping credits", "Shipping Credits", "shipping-price", "Shipping Price"])
        out["gmv_product"] = product_sales + rebates
        out["shipping_revenue"] = shipping
        out["gmv_incl_shipping_tax"] = out["gmv_product"] + out["shipping_revenue"]
        out["platform_discount"] = 0.0
        out["seller_discount"] = 0.0
        out["settlement_status"] = ""
        out["source_file"] = path.name
        out = out[out["order_id"].ne("") & out["order_id"].str.lower().ne("order id")]
        frames.append(out)
    if not frames:
        return pd.DataFrame(columns=PRODUCT_COLUMNS)
    return attach_sku_master(map_platform_sku(pd.concat(frames, ignore_index=True), mapping), sku_master)[PRODUCT_COLUMNS]


def clean_tiktok(paths: list[Path], mapping: pd.DataFrame, sku_master: pd.DataFrame) -> pd.DataFrame:
    frames = []
    for path in paths:
        df = normalize_columns(read_table(path))
        if df.empty:
            continue
        out = pd.DataFrame(index=df.index)
        out["platform"] = "TikTok"
        out["order_id"] = text_series(df, ["order_id", "Order ID", "Order/adjustment ID", "订单号"])
        out["order_date"] = date_series(df, ["Created Time", "Order Created Time", "order_date", "Order Date", "创建时间"])
        out["state"] = normalize_state_series(text_series(df, ["State", "Province", "Recipient State", "收件人省/州"]))
        seller_sku = "Seller sku input by the seller in the product system."
        out["platform_sku"] = text_series(df, [seller_sku, "Seller SKU", "Seller sku", "SKU Seller Input", "商家SKU"])
        out["clean_platform_sku"] = out["platform_sku"].map(clean_platform_sku)
        out["erp_sku"] = out["clean_platform_sku"]
        out["quantity"] = pd.to_numeric(value_series(df, ["Quantity", "SKU Quantity", "Qty", "数量"], 1), errors="coerce").fillna(1)
        out["platform_discount"] = money_series(df, ["Total platform discount in this SKU ID.", "Total platform discount", "Platform Discount"])
        out["seller_discount"] = money_series(df, ["Total seller discount in this SKU ID.", "Total seller discount", "Seller Discount"])
        out["gmv_product"] = money_series(df, ["SKU Subtotal After Discount", "Sku Subtotal After Discount", "SKU subtotal after discount"])
        out["shipping_revenue"] = money_series(df, ["Shipping Fee After Discount", "Shipping fee after discount"])
        out["gmv_incl_shipping_tax"] = out["gmv_product"] + out["shipping_revenue"]
        out["settlement_status"] = ""
        out["source_file"] = path.name
        out = out[
            out["order_id"].ne("")
            & ~out["order_id"].str.lower().str.contains("platform unique|order id", regex=True, na=False)
            & ~out["platform_sku"].str.lower().str.contains("seller sku input by the seller", regex=False, na=False)
        ]
        frames.append(out)
    if not frames:
        return pd.DataFrame(columns=PRODUCT_COLUMNS)
    mapped = map_platform_sku(pd.concat(frames, ignore_index=True), mapping)
    mapped["erp_sku"] = mapped["erp_sku"].fillna("").mask(mapped["erp_sku"].fillna("").eq(""), mapped["clean_platform_sku"])
    return attach_sku_master(mapped, sku_master)[PRODUCT_COLUMNS]


def _load_temu_order_sku_bridge(raw_dir: Path) -> pd.DataFrame:
    order_info_files = find_files_any(raw_dir, TEMU_ORDER_INFO_PATTERNS)
    gucang_files = find_files_any(raw_dir, TEMU_GUCANG_PATTERNS)
    if not order_info_files or not gucang_files:
        return pd.DataFrame(columns=["po_order_id", "batch_order_id", "gucang_order_id", "erp_sku", "state", "order_date"])

    order_info_frames = []
    for path in order_info_files:
        df = normalize_columns(read_table(path))
        batch_col = first_existing(df, ["批次订单号", "batch_order_id", "Batch Order No", "batch order number"])
        po_col = first_existing(df, ["订单号", "PO单号", "po_order_id"])
        state_col = first_existing(df, ["省份", "省/州", "state", "State"])
        date_col = first_existing(df, ["订单创建时间", "order_date", "Order Date"])
        if batch_col is None:
            continue
        tmp = df[[batch_col] + ([po_col] if po_col else []) + ([state_col] if state_col else []) + ([date_col] if date_col else [])].copy()
        tmp.columns = ["raw_batch_order_id"] + (["po_order_id"] if po_col else []) + (["state"] if state_col else []) + (["order_date"] if date_col else [])
        tmp["batch_order_id"] = tmp["raw_batch_order_id"].apply(split_batch_order_numbers)
        tmp = tmp.explode("batch_order_id")
        tmp["state"] = normalize_state_series(tmp.get("state", pd.Series([""] * len(tmp), index=tmp.index)))
        tmp["order_date"] = pd.to_datetime(tmp.get("order_date", pd.NaT), errors="coerce")
        order_info_frames.append(tmp)
    if not order_info_frames:
        return pd.DataFrame(columns=["po_order_id", "batch_order_id", "gucang_order_id", "erp_sku", "state", "order_date"])
    order_info = pd.concat(order_info_frames, ignore_index=True)

    bridge_frames = []
    product_frames = []
    for path in gucang_files:
        xls = pd.ExcelFile(path)
        sheet1 = "Sheet1" if "Sheet1" in xls.sheet_names else next((s for s in xls.sheet_names if "订单" in str(s)), xls.sheet_names[0])
        orders = normalize_columns(read_excel_sheet(path, sheet1))
        ref_col = first_existing(orders, ["订单参考号", "order_reference_no", "Order Reference No"])
        gc_col = first_existing(orders, ["谷仓订单号", "gucang_order_id", "GC Order No", "订单号"])
        if ref_col and gc_col:
            bridge_frames.append(orders[[ref_col, gc_col]].rename(columns={ref_col: "order_reference_no", gc_col: "gucang_order_id"}))

        product_sheet = next((s for s in xls.sheet_names if "商品" in str(s) or "product" in str(s).lower()), None)
        if product_sheet:
            products = normalize_columns(read_excel_sheet(path, product_sheet))
            gc_col = first_existing(products, ["谷仓订单号", "gucang_order_id", "GC Order No", "订单号"])
            sku_col = first_existing(products, ["商品编码", "erp_sku", "SKU", "sku"])
            qty_col = first_existing(products, ["数量", "quantity", "Qty"])
            if gc_col and sku_col:
                tmp = products[[gc_col, sku_col] + ([qty_col] if qty_col else [])].copy()
                tmp.columns = ["gucang_order_id", "erp_sku"] + (["bridge_quantity"] if qty_col else [])
                product_frames.append(tmp)

    if not bridge_frames or not product_frames:
        return pd.DataFrame(columns=["po_order_id", "batch_order_id", "gucang_order_id", "erp_sku", "state", "order_date"])
    bridge = pd.concat(bridge_frames, ignore_index=True).drop_duplicates()
    products = pd.concat(product_frames, ignore_index=True).drop_duplicates()
    out = order_info.merge(bridge, left_on="batch_order_id", right_on="order_reference_no", how="left").merge(products, on="gucang_order_id", how="left")
    out["batch_order_id"] = out["batch_order_id"].fillna("").astype(str).str.strip()
    out["po_order_id"] = out.get("po_order_id", "").fillna("").astype(str).str.strip()
    out["erp_sku"] = _standardize_sku_text(out["erp_sku"])
    out = out.drop_duplicates(["po_order_id", "batch_order_id", "gucang_order_id", "erp_sku"])
    has_sku = out["erp_sku"].ne("")
    out = out[has_sku | ~out["po_order_id"].isin(out.loc[has_sku, "po_order_id"])]
    keep = ["po_order_id", "batch_order_id", "gucang_order_id", "erp_sku", "state", "order_date"]
    if "bridge_quantity" in out.columns:
        keep.append("bridge_quantity")
    return out[keep]


def clean_temu(paths: list[Path], mapping: pd.DataFrame, sku_master: pd.DataFrame) -> pd.DataFrame:
    bridge = _load_temu_order_sku_bridge(RAW_DIR)
    sku_master_set = set(sku_master["erp_sku"].astype(str))
    frames = []
    for path in paths:
        df = pd.read_excel(path, header=[0, 1]) if path.suffix.lower() in {".xlsx", ".xls"} else read_table(path)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(x).strip() for x in col if str(x) != "nan" and not str(x).startswith("Unnamed")).strip()
                for col in df.columns
            ]
        df = normalize_columns(df)
        if df.empty:
            continue

        lines = pd.DataFrame(index=df.index)
        lines["order_id"] = text_series(df, ["PO单号", "po_order_id", "order_id", "订单号"]).replace("", pd.NA).ffill().fillna("")
        lines = lines[lines["order_id"].ne("")]
        if lines.empty:
            continue
        lines["platform_sku"] = text_series(df.loc[lines.index], ["商品信息 * 销售件数 SKU货号", "SKU货号", "SKU ID", "商品信息 * 销售件数 SKU ID", "SKU", "sku", "平台SKU", "商品编码"])
        lines["clean_platform_sku"] = lines["platform_sku"].map(clean_platform_sku)
        qty = pd.to_numeric(value_series(df, ["商品信息 * 销售件数 件数", "件数", "Quantity", "quantity"], 1), errors="coerce").fillna(1)
        lines["line_quantity"] = pd.to_numeric(qty.loc[lines.index], errors="coerce").fillna(1)
        lines["sales_return"] = money_series(df.loc[lines.index], ["销售回款", "sales_return"])
        lines["sales_reversal"] = money_series(df.loc[lines.index], ["销售冲回", "sales_reversal"])
        lines["shipping_return"] = money_series(df.loc[lines.index], ["运费回款", "shipping_return"])
        lines["shipping_reversal"] = money_series(df.loc[lines.index], ["运费冲回", "shipping_reversal"])

        order_financial = lines.groupby("order_id", as_index=False).agg(
            quantity=("line_quantity", "sum"),
            sales_return=("sales_return", "sum"),
            sales_reversal=("sales_reversal", "sum"),
            shipping_return=("shipping_return", "sum"),
            shipping_reversal=("shipping_reversal", "sum"),
            platform_sku=("platform_sku", lambda s: " + ".join(sorted(set(x for x in s.astype(str) if x and x != "nan")))),
            clean_platform_sku=("clean_platform_sku", lambda s: " + ".join(sorted(set(x for x in s.astype(str) if x and x != "nan")))),
        )
        order_financial["gmv_product"] = zero_small((order_financial["sales_return"] - order_financial["sales_reversal"].abs()) * 1.25)
        order_financial["shipping_revenue"] = zero_small(order_financial["shipping_return"] - order_financial["shipping_reversal"].abs())
        order_financial["gmv_incl_shipping_tax"] = zero_small(order_financial["gmv_product"] + order_financial["shipping_revenue"])
        order_financial["settlement_status"] = "Unsettled" if any(p in path.name.lower() for p in ("unsettled", "未结算")) else "Settled"
        order_financial["source_file"] = path.name

        attrs = bridge.groupby("po_order_id", as_index=False).agg(
            state=("state", lambda s: next((x for x in s.astype(str) if x and x != "nan"), "")),
            order_date=("order_date", "min"),
        ).rename(columns={"po_order_id": "order_id"})
        order_financial = order_financial.merge(attrs, on="order_id", how="left")
        order_financial["state"] = normalize_state_series(order_financial.get("state", pd.Series([""] * len(order_financial), index=order_financial.index)))

        bridge_components = bridge[bridge["erp_sku"].fillna("").ne("")].copy()
        if "bridge_quantity" not in bridge_components.columns:
            bridge_components["bridge_quantity"] = 1
        bridge_components["quantity"] = pd.to_numeric(bridge_components["bridge_quantity"], errors="coerce").fillna(1)
        bridge_components = (
            bridge_components.groupby(["po_order_id", "erp_sku"], as_index=False)
            .agg(quantity=("quantity", "sum"))
            .rename(columns={"po_order_id": "order_id"})
        )

        fallback = lines.groupby(["order_id", "platform_sku", "clean_platform_sku"], as_index=False).agg(quantity=("line_quantity", "sum"))
        fallback["erp_sku"] = fallback["clean_platform_sku"].where(fallback["clean_platform_sku"].isin(sku_master_set), "")
        fallback = fallback[["order_id", "platform_sku", "clean_platform_sku", "erp_sku", "quantity"]]

        components = bridge_components.merge(order_financial[["order_id", "platform_sku", "clean_platform_sku"]], on="order_id", how="left")
        orders_with_bridge = set(components["order_id"].dropna())
        components = pd.concat([components, fallback[~fallback["order_id"].isin(orders_with_bridge)]], ignore_index=True)
        components["quantity"] = pd.to_numeric(components["quantity"], errors="coerce").fillna(1)
        components["platform_sku"] = components["platform_sku"].fillna("")
        components["clean_platform_sku"] = components["clean_platform_sku"].fillna(components["platform_sku"].map(clean_platform_sku))
        components = components.groupby(["order_id", "platform_sku", "clean_platform_sku", "erp_sku"], dropna=False, as_index=False).agg(quantity=("quantity", "sum"))

        out = components.merge(order_financial, on=["order_id", "platform_sku", "clean_platform_sku"], how="left", suffixes=("", "_order"))
        total_qty = out.groupby("order_id")["quantity"].transform("sum")
        row_count = out.groupby("order_id")["order_id"].transform("size")
        alloc_weight = out["quantity"] / total_qty.where(total_qty.gt(0), pd.NA)
        alloc_weight = alloc_weight.fillna(1 / row_count)
        for money_col in ("gmv_product", "shipping_revenue", "gmv_incl_shipping_tax"):
            out[money_col] = zero_small(out[money_col] * alloc_weight)
        out["platform"] = "Temu"
        out["platform_discount"] = 0.0
        out["seller_discount"] = 0.0
        frames.append(out)
    if not frames:
        return pd.DataFrame(columns=PRODUCT_COLUMNS)
    mapped = map_platform_sku(pd.concat(frames, ignore_index=True), mapping)
    return attach_sku_master(mapped, sku_master)[PRODUCT_COLUMNS]


def build_product_expanded(fact_product: pd.DataFrame, bundle_mapping: pd.DataFrame, sku_master: pd.DataFrame) -> pd.DataFrame:
    if fact_product.empty:
        return fact_product.copy()
    for col in ["platform", "platform_sku", "erp_sku", "bundle_sku", "bundle_name", "component_erp_sku", "component_qty"]:
        if col not in bundle_mapping.columns:
            bundle_mapping[col] = ""
    bundles = bundle_mapping.copy()
    bundles["component_qty"] = pd.to_numeric(bundles["component_qty"], errors="coerce").fillna(1)
    bundles["clean_platform_sku"] = bundles["platform_sku"].map(clean_platform_sku)
    join_cols = ["platform", "clean_platform_sku", "erp_sku"]
    expanded = fact_product.merge(
        bundles[join_cols + ["bundle_sku", "bundle_name", "component_erp_sku", "component_qty"]],
        on=join_cols,
        how="left",
    )
    expanded["is_bundle"] = expanded["component_erp_sku"].fillna("").ne("")
    expanded["analysis_erp_sku"] = expanded["component_erp_sku"].fillna("").mask(
        expanded["component_erp_sku"].fillna("").eq(""), expanded["erp_sku"]
    )
    expanded["analysis_quantity"] = expanded["quantity"] * expanded["component_qty"].fillna(1)
    attrs = sku_master.rename(columns={"erp_sku": "analysis_erp_sku"})
    keep = ["analysis_erp_sku", "master_sku", "product_family", "category", "flavor", "pack_size", "unit_size", "case_pack"]
    for col in keep:
        if col not in attrs.columns:
            attrs[col] = ""
    expanded = expanded.drop(columns=["master_sku", "product_family", "category", "flavor", "pack_size", "unit_size", "case_pack"], errors="ignore")
    expanded = expanded.merge(attrs[keep].drop_duplicates("analysis_erp_sku"), on="analysis_erp_sku", how="left")
    return expanded


def build_fact_order(fact_product: pd.DataFrame) -> pd.DataFrame:
    if fact_product.empty:
        return pd.DataFrame(columns=ORDER_COLUMNS)
    group_cols = ["platform", "order_id"]
    agg = {
        "order_date": "min",
        "state": lambda s: next((x for x in s.astype(str) if x and x != "nan"), ""),
        "quantity": "sum",
        "gmv_product": "sum",
        "shipping_revenue": "sum",
        "gmv_incl_shipping_tax": "sum",
        "settlement_status": safe_join_unique,
        "source_file": lambda s: safe_join_unique(s, sep="; "),
    }
    out = fact_product.groupby(group_cols, as_index=False).agg(agg)
    for col in ("gmv_product", "shipping_revenue", "gmv_incl_shipping_tax"):
        out[col] = zero_small(out[col])
    out["aov"] = out["gmv_incl_shipping_tax"]
    return out[ORDER_COLUMNS]


def collect_unmatched(fact_product: pd.DataFrame) -> pd.DataFrame:
    if fact_product.empty:
        return pd.DataFrame(columns=["platform", "platform_sku", "clean_platform_sku", "source_file", "row_count"])
    unmatched = fact_product[fact_product["erp_sku"].fillna("").eq("")]
    if unmatched.empty:
        return pd.DataFrame(columns=["platform", "platform_sku", "clean_platform_sku", "source_file", "row_count"])
    return (
        unmatched.groupby(["platform", "platform_sku", "clean_platform_sku", "source_file"], dropna=False)
        .size()
        .reset_index(name="row_count")
        .sort_values(["platform", "row_count"], ascending=[True, False])
    )


def run_pipeline() -> PipelineResult:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    sku_master, platform_sku_mapping, bundle_mapping, target_table = load_config()

    amazon_files = [p for p in find_files_any(RAW_DIR, AMAZON_PATTERNS) if "temu" not in p.name.lower()]
    tiktok_files = find_files_any(RAW_DIR, TIKTOK_PATTERNS)
    temu_files = [
        p
        for p in find_files_any(RAW_DIR, TEMU_SETTLED_PATTERNS + TEMU_UNSETTLED_PATTERNS)
        if not any(x in p.name.lower() for x in ("订单信息", "订单转谷仓", "order_info", "warehouse", "gucang"))
    ]

    amazon = clean_amazon(amazon_files, platform_sku_mapping, sku_master)
    tiktok = clean_tiktok(tiktok_files, platform_sku_mapping, sku_master)
    temu = clean_temu(temu_files, platform_sku_mapping, sku_master)
    fact_product = pd.concat([amazon, tiktok, temu], ignore_index=True)
    if not fact_product.empty:
        fact_product["order_date"] = pd.to_datetime(fact_product["order_date"], errors="coerce")
        fact_product["month"] = fact_product["order_date"].dt.to_period("M").astype(str)
    else:
        fact_product["month"] = pd.Series(dtype=str)
    fact_order = build_fact_order(fact_product)
    if not fact_order.empty:
        fact_order["month"] = pd.to_datetime(fact_order["order_date"], errors="coerce").dt.to_period("M").astype(str)
    product_expanded = build_product_expanded(fact_product, bundle_mapping, sku_master)
    unmatched = collect_unmatched(fact_product)
    source_notes = pd.DataFrame(
        [
            {"area": "Amazon", "files": "; ".join(p.name for p in amazon_files), "rule": "GMV product = product_sales + promotional_rebates; shipping = shipping_credits."},
            {"area": "TikTok", "files": "; ".join(p.name for p in tiktok_files), "rule": "ERP SKU sourced from seller SKU, then optional mapping/master enrichment."},
            {"area": "Temu", "files": "; ".join(p.name for p in temu_files), "rule": "GMV includes settled and unsettled; SKU bridge uses order info -> Gucang order -> product code."},
            {"area": "Config", "files": str(CONFIG_DIR), "rule": "Missing optional platform_sku_mapping/bundle_mapping config files are not invented; empty sheets are exported when absent."},
            {"area": "Unmatched SKU", "files": str(UNMATCHED_SKU_PATH.name), "rule": "Rows without ERP SKU are exported for manual mapping; no guessing applied."},
        ]
    )
    return PipelineResult(
        fact_order=fact_order,
        fact_product=fact_product,
        sku_master=sku_master,
        platform_sku_mapping=platform_sku_mapping,
        bundle_mapping=bundle_mapping,
        target_table=target_table,
        product_expanded=product_expanded,
        unmatched_sku=unmatched,
        source_notes=source_notes,
    )


def export_outputs(result: PipelineResult) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(MASTER_DATA_PATH, engine="openpyxl") as writer:
        result.fact_order.to_excel(writer, sheet_name="Fact_Order", index=False)
        result.fact_product.to_excel(writer, sheet_name="Fact_Product", index=False)
        result.sku_master.to_excel(writer, sheet_name="SKU_Master", index=False)
        result.platform_sku_mapping.to_excel(writer, sheet_name="Platform_SKU_Mapping", index=False)
        result.bundle_mapping.to_excel(writer, sheet_name="Bundle_Mapping", index=False)
        result.target_table.to_excel(writer, sheet_name="Target_Table", index=False)
        result.product_expanded.to_excel(writer, sheet_name="Product_Expanded", index=False)
        result.source_notes.to_excel(writer, sheet_name="Source_Notes", index=False)
    with pd.ExcelWriter(UNMATCHED_SKU_PATH, engine="openpyxl") as writer:
        result.unmatched_sku.to_excel(writer, sheet_name="unmatched_sku", index=False)


def main() -> None:
    result = run_pipeline()
    export_outputs(result)
    print(f"Exported {MASTER_DATA_PATH}")
    print(f"Exported {UNMATCHED_SKU_PATH} ({len(result.unmatched_sku)} unmatched SKU rows)")


if __name__ == "__main__":
    main()
