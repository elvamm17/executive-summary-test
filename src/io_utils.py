from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


def find_files(directory: Path, patterns: Iterable[str]) -> list[Path]:
    patterns_lc = tuple(p.lower() for p in patterns)
    files: list[Path] = []
    for path in directory.glob("**/*"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        name = path.name.lower()
        if path.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        if all(p in name for p in patterns_lc):
            files.append(path)
    return sorted(files)


def find_files_any(directory: Path, patterns: Iterable[str]) -> list[Path]:
    patterns_lc = tuple(p.lower() for p in patterns)
    files: list[Path] = []
    for path in directory.glob("**/*"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        name = path.name.lower()
        if path.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        if any(p in name for p in patterns_lc):
            files.append(path)
    return sorted(files)


def read_table(path: Path, sheet_name=0) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for encoding in ("utf-8-sig", "utf-8", "gb18030", "latin1"):
            try:
                return pd.read_csv(path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path)
    return pd.read_excel(path, sheet_name=sheet_name)


def read_excel_sheet(path: Path, sheet_name: str | int) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def first_existing(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in lookup:
            return lookup[key]
    return None


def value_series(df: pd.DataFrame, candidates: Iterable[str], default=0) -> pd.Series:
    col = first_existing(df, candidates)
    if col is None:
        return pd.Series([default] * len(df), index=df.index)
    values = df[col]
    if isinstance(values, pd.DataFrame):
        values = values.iloc[:, 0]
    return values


def text_series(df: pd.DataFrame, candidates: Iterable[str], default="") -> pd.Series:
    return value_series(df, candidates, default=default).fillna("").astype(str).str.strip()


def money_series(df: pd.DataFrame, candidates: Iterable[str]) -> pd.Series:
    s = value_series(df, candidates, default=0)
    if not isinstance(s, pd.Series):
        s = pd.Series([s] * len(df), index=df.index)
    cleaned = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace("¥", "", regex=False)
        .str.replace("￥", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def date_series(df: pd.DataFrame, candidates: Iterable[str]) -> pd.Series:
    col = first_existing(df, candidates)
    if col is None:
        return pd.Series([pd.NaT] * len(df), index=df.index)
    return pd.to_datetime(df[col], errors="coerce")


def read_config_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns)
    df = pd.read_csv(path, dtype=str).fillna("")
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[columns]


def detect_header_row(path: Path, required: Iterable[str], sheet_name=0, max_rows: int = 20) -> int:
    if path.suffix.lower() == ".csv":
        preview = read_table(path, sheet_name=sheet_name).head(0)
        if any(str(c).strip().lower() in {r.lower() for r in required} for c in preview.columns):
            return 0
        raw = pd.read_csv(path, header=None, nrows=max_rows, encoding="utf-8-sig")
    else:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_rows)
    required_lc = {r.lower() for r in required}
    for idx, row in raw.iterrows():
        values = {str(v).strip().lower() for v in row.tolist() if pd.notna(v)}
        if required_lc.intersection(values):
            return int(idx)
    return 0
