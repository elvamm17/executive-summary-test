from __future__ import annotations

import re

import pandas as pd


def clean_platform_sku(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"-(FBM|FBA)$", "", text, flags=re.IGNORECASE)
    return text.strip()


def split_batch_order_numbers(value: object) -> list[str]:
    if pd.isna(value):
        return []
    text = str(value)
    text = re.sub(r"[\u4e00-\u9fff]+", " ", text)
    parts = re.split(r"[,，;；/、\s]+", text)
    return [p.strip() for p in parts if p.strip()]


def attach_sku_master(df: pd.DataFrame, sku_master: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    master_cols = [
        "erp_sku",
        "master_sku",
        "product_family",
        "category",
        "flavor",
        "pack_size",
        "unit_size",
        "case_pack",
    ]
    for col in master_cols:
        if col not in sku_master.columns:
            sku_master[col] = ""
    attrs = sku_master[master_cols].drop_duplicates("erp_sku")
    out = df.merge(attrs, on="erp_sku", how="left", suffixes=("", "_master"))
    if "master_sku_master" in out.columns:
        out["master_sku"] = out["master_sku"].fillna("").mask(
            out["master_sku"].fillna("").eq(""), out["master_sku_master"].fillna("")
        )
        out = out.drop(columns=["master_sku_master"])
    out["master_sku"] = out["master_sku"].fillna("").mask(out["master_sku"].fillna("").eq(""), out["erp_sku"])
    for col in ("product_family", "category", "flavor", "pack_size", "unit_size", "case_pack"):
        out[col] = out[col].fillna("")
    return out


def map_platform_sku(df: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    map_cols = ["platform", "platform_sku", "clean_platform_sku", "erp_sku", "master_sku"]
    for col in map_cols:
        if col not in mapping.columns:
            mapping[col] = ""
    m = mapping[map_cols].copy()
    m["platform"] = m["platform"].astype(str).str.strip()
    m["clean_platform_sku"] = m["clean_platform_sku"].fillna("").astype(str).map(clean_platform_sku)
    missing_clean = m["clean_platform_sku"].eq("")
    m.loc[missing_clean, "clean_platform_sku"] = m.loc[missing_clean, "platform_sku"].map(clean_platform_sku)
    m = m.drop_duplicates(["platform", "clean_platform_sku"])
    out = df.merge(
        m[["platform", "clean_platform_sku", "erp_sku", "master_sku"]],
        on=["platform", "clean_platform_sku"],
        how="left",
        suffixes=("", "_mapped"),
    )
    if "erp_sku_mapped" in out.columns:
        out["erp_sku"] = out.get("erp_sku", "").fillna("").mask(
            out.get("erp_sku", "").fillna("").eq(""), out["erp_sku_mapped"].fillna("")
        )
        out = out.drop(columns=["erp_sku_mapped"])
    if "master_sku_mapped" in out.columns:
        out["master_sku"] = out.get("master_sku", "").fillna("").mask(
            out.get("master_sku", "").fillna("").eq(""), out["master_sku_mapped"].fillna("")
        )
        out = out.drop(columns=["master_sku_mapped"])
    out["erp_sku"] = out["erp_sku"].fillna("").astype(str).str.strip()
    out["master_sku"] = out["master_sku"].fillna("").astype(str).str.strip()
    return out
