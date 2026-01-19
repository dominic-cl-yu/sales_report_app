# process.py
# -*- coding: utf-8 -*-
"""销售透视报表处理核心（无 UI）

功能：
- 读取订单状态 Excel（支持多 sheet，自动识别表头）
- 清洗/规范化列名与 Team
- 按 Factory / Team / Order Type 生成按月透视表（Order Qty / SAH / Sales）
- 输出为一个 Excel 文件（9 个 sheet：3 个 Factory * 3 个指标）

说明：
- 本模块不依赖 streamlit；可被 Streamlit/Tkinter/CLI 等界面调用。
- Customer 字段固定写入 "ALL"（可在常量 CUSTOMER_LABEL_DEFAULT 修改）。

作者：基于用户提供的逻辑重构
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import time as time_mod
from datetime import datetime, date, time
from typing import Dict, List, Tuple, Optional, Any

import pandas as pd

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.filters import AutoFilter, FilterColumn


# ======================
# 自动默认值（不暴露任何 UI 选项）
# ======================
SCAN_ROWS_DEFAULT = 50
CUSTOMER_LABEL_DEFAULT = "ALL"  # Pivot 报表中 Customer 字段固定写入


# ======================
# 配置
# ======================
TARGET_COLS = [
    "Factory",
    "Team",
    "Order Type",
    "Order Qty",
    "SAH",
    "Sales (USD)",
    "GP (USD)",
    "Product Type",
    "Request Garment Delivery (DeadLine ex-fty)",
    "Customer Delivery Date",
]

ORDER_TYPE_ALLOWED = {"AO", "FR"}

ORDER_TYPE_SORT_ORDER = {
    "SO": 0,
    "AO": 1,
    "Forecast-FR": 2,
    "Total": 999,
}

VALID_TEAMS = {
    "Sports",
    "Fancy",
    "SW",
    "Cotton Panty",
    "Brands-COS",
    "Brands-Stories",
}

# 指标配置：每个 Factory 输出 3 张表（Order Qty / SAH / Sales (USD)）
METRIC_SPECS: List[Tuple[str, str]] = [
    ("Order Qty", "Order Qty"),
    ("SAH", "SAH"),
    ("Sales (USD)", "Sales (USD)"),
]


# ======================
# Team 映射（规则）
# ======================
def apply_known_mapping(team: str) -> Tuple[str, bool]:
    """将 team 依据已知规则映射到标准 Team。

    返回： (映射后的 team, 是否命中规则)
    """
    if pd.isna(team) or not str(team).strip():
        return team, False

    team_str = str(team).strip()
    lower = team_str.lower()

    if "sw" in lower:
        return "SW", True
    if "sports" in lower:
        return "Sports", True
    if "fancy" in lower:
        return "Fancy", True
    if "cotton panty reservation" in lower:
        return "Cotton Panty", True

    cleaned = re.sub(r"\s*[Rr]eservation.*$", "", team_str).strip()
    if cleaned != team_str and cleaned in VALID_TEAMS:
        return cleaned, True

    return team_str, False


def categorize_team(team: str, cache: Dict[str, str]) -> str:
    """带缓存的 Team 规范化。"""
    team_str = str(team).strip() if team else ""
    if team_str in cache:
        return cache[team_str]

    mapped, was_rule_mapped = apply_known_mapping(team_str)
    result = mapped if was_rule_mapped else team_str

    cache[team_str] = result
    return result


# ======================
# 规范化工具
# ======================
def _norm(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("\u00A0", " ")
    s = s.lower()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("–", "-").replace("—", "-")
    return s


def _norm_key(s: Any) -> str:
    s = _norm(s)
    s = re.sub(r"[^a-z0-9()%-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ======================
# 列别名
# ======================
def _build_header_aliases() -> Dict[str, List[str]]:
    return {
        "Factory": ["factory"],
        "Team": ["team"],
        "Order Type": ["order type", "order_type", "ordertype"],
        "Order Qty": ["order qty", "order quantity", "qty"],
        "SAH": ["sah"],
        "Sales (USD)": ["sales (usd)", "sales usd", "sales"],
        "GP (USD)": ["gp (usd)", "gp usd", "gp"],
        "Product Type": ["product type"],
        "Request Garment Delivery (DeadLine ex-fty)": [
            "request garment delivery (deadline ex-fty)",
            "requested garment delivery",
            "deadline ex-fty",
            "ex-fty",
        ],
        "Customer Delivery Date": ["customer delivery date", "delivery date"],
    }


ALIASES = _build_header_aliases()


# ======================
# 表头检测（pd.ExcelFile 版本）
# ======================
def _score_row_as_header(row_values: List[object]) -> Tuple[int, Dict[str, int]]:
    normed = [_norm_key(v) for v in row_values]
    matched: Dict[str, int] = {}

    for target, alias_list in ALIASES.items():
        for idx, cell in enumerate(normed):
            if not cell:
                continue
            for a in alias_list:
                a_norm = _norm_key(a)
                if cell == a_norm or a_norm in cell or cell in a_norm:
                    matched[target] = idx
                    break
            if target in matched:
                break

    return len(matched), matched


def detect_header_row(
    xls: pd.ExcelFile,
    sheet_name: str,
    scan_rows: int = SCAN_ROWS_DEFAULT,
) -> Tuple[int, int, Dict[str, int]]:
    """在工作表前 scan_rows 行中自动找表头所在行。"""
    preview = pd.read_excel(
        xls,
        sheet_name=sheet_name,
        header=None,
        nrows=scan_rows,
    )
    best_row = 0
    best_score = -1
    best_matched: Dict[str, int] = {}

    for r in range(len(preview)):
        score, matched = _score_row_as_header(preview.iloc[r].tolist())
        if score > best_score:
            best_score = score
            best_row = r
            best_matched = matched

    return best_row, best_score, best_matched


# ======================
# Order Type 列选择
# ======================
def _pick_order_type_column(df: pd.DataFrame) -> Optional[str]:
    candidates = [
        c
        for c in df.columns
        if _norm_key(c) in {"order type", "ordertype", "order_type"}
    ]
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    def score(col: str) -> Tuple[int, int]:
        s = df[col].dropna().astype(str).str.strip().str.upper()
        s = s[s != ""]
        return (s.isin(ORDER_TYPE_ALLOWED)).sum(), len(s)

    ranked = [(c, *score(c)) for c in candidates]
    ranked.sort(key=lambda x: (x[1], x[2]), reverse=True)
    return ranked[0][0] if ranked and ranked[0][1] > 0 else None


# ======================
# 列重命名/选择
# ======================
def _rename_and_select(df: pd.DataFrame) -> pd.DataFrame:
    rename_map: Dict[str, str] = {}
    used_targets = set()

    for col in df.columns:
        col_norm = _norm_key(col)
        for target, alias_list in ALIASES.items():
            if target in used_targets:
                continue
            for a in alias_list:
                a_norm = _norm_key(a)
                if col_norm == a_norm or a_norm in col_norm or col_norm in a_norm:
                    rename_map[col] = target
                    used_targets.add(target)
                    break
            if col in rename_map:
                break

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()]

    for c in TARGET_COLS:
        if c not in df.columns:
            df[c] = pd.NA

    return df[TARGET_COLS].copy()


# ======================
# 底部/无效行清理
# ======================
def _drop_footer_note_rows(df: pd.DataFrame) -> pd.DataFrame:
    """去掉类似页脚的“备注行”——通常是一些编号说明文本。"""
    key_cols = ["Order Qty", "SAH", "Sales (USD)", "GP (USD)"]
    text_cols = ["Factory", "Team", "Product Type"]
    date_cols = [
        "Request Garment Delivery (DeadLine ex-fty)",
        "Customer Delivery Date",
    ]

    existing_key = [c for c in key_cols if c in df.columns]
    existing_text = [c for c in text_cols if c in df.columns]
    existing_date = [c for c in date_cols if c in df.columns]

    has_key_signal = pd.Series(False, index=df.index)
    for c in existing_key:
        numeric = pd.to_numeric(df[c], errors="coerce")
        has_key_signal = has_key_signal | numeric.notna()

    has_date_signal = pd.Series(False, index=df.index)
    for c in existing_date:
        dt = pd.to_datetime(df[c], errors="coerce")
        has_date_signal = has_date_signal | dt.notna()

    has_text_signal = pd.Series(False, index=df.index)
    for c in existing_text:
        s = df[c].astype(str).str.strip()
        has_text_signal = has_text_signal | (s.ne("") & s.ne("nan") & s.ne("none"))

    non_empty_count = df.apply(
        lambda r: sum(
            (str(v).strip() not in ("", "nan", "None", "none")) and (pd.notna(v))
            for v in r.values
        ),
        axis=1,
    )

    note_like = pd.Series(False, index=df.index)
    scan_cols = list(df.columns[: min(10, len(df.columns))])
    pattern = re.compile(r"^\s*\d+(\.\d+)?\s+\S+.*$", re.IGNORECASE)
    for c in scan_cols:
        s = df[c].astype(str)
        note_like = note_like | s.str.match(pattern, na=False)

    keep = has_key_signal | has_date_signal | has_text_signal
    drop = (~keep) & (non_empty_count <= 1) & note_like
    return df.loc[~drop].copy()


def _remove_footer_rows(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """更激进的尾部清理：去掉 Summary/Benchmark 行，及明显无效行。"""
    if len(df) == 0:
        return df

    df_str = df.astype(str)

    summary_pattern = re.compile(r"\bsummary\b", re.IGNORECASE)
    has_summary = pd.Series(False, index=df.index)
    for col in df_str.columns:
        exact_match = df_str[col].str.strip().str.lower() == "summary"
        word_match = df_str[col].str.contains(summary_pattern, na=False, regex=True)
        has_summary = has_summary | exact_match | word_match

    benchmark_pattern = re.compile(r"^\s*\d+(\.\d+)?\s+benchmark\s*$", re.IGNORECASE)
    has_benchmark = pd.Series(False, index=df.index)
    for col in df_str.columns:
        has_benchmark = has_benchmark | df_str[col].str.match(benchmark_pattern, na=False)

    non_empty_mask = df.astype(str).apply(
        lambda row: row.str.strip().notna()
        & (row.str.strip() != "")
        & (row.str.strip().str.lower() != "nan"),
        axis=1,
    )
    non_empty_count = non_empty_mask.sum(axis=1)

    numeric_cols = ["Order Qty", "SAH", "Sales (USD)", "GP (USD)"]
    has_numeric = pd.Series(False, index=df.index)
    for col in numeric_cols:
        if col in df.columns:
            numeric_vals = pd.to_numeric(df[col], errors="coerce")
            has_numeric = has_numeric | numeric_vals.notna()

    date_cols = [
        "Request Garment Delivery (DeadLine ex-fty)",
        "Customer Delivery Date",
    ]
    has_date = pd.Series(False, index=df.index)
    for col in date_cols:
        if col in df.columns:
            date_vals = pd.to_datetime(df[col], errors="coerce")
            has_date = has_date | date_vals.notna()

    factory_valid = pd.Series(True, index=df.index)
    team_valid = pd.Series(True, index=df.index)
    if "Factory" in df.columns:
        factory_str = df["Factory"].astype(str).str.strip()
        factory_valid = (factory_str != "") & (factory_str.str.lower() != "nan") & factory_str.notna()
    if "Team" in df.columns:
        team_str = df["Team"].astype(str).str.strip()
        team_valid = (team_str != "") & (team_str.str.lower() != "nan") & team_str.notna()

    is_valid = factory_valid & team_valid & (has_numeric | has_date)
    drop_mask = has_summary | has_benchmark | ((non_empty_count < 2) & ~has_numeric & ~has_date & ~is_valid)

    return df.loc[~drop_mask].copy()


# ======================
# 校验（ExcelFile 版本）
# ======================
def sanity_check(
    xls: pd.ExcelFile,
    *,
    scan_rows: int = SCAN_ROWS_DEFAULT,
    min_matched: int = 6,
    min_rows_per_sheet: int = 1,
) -> Tuple[bool, str]:
    if not xls.sheet_names:
        return False, "Excel 文件中没有任何工作表。"

    issues: List[str] = []
    valid_sheets = 0

    for sheet in xls.sheet_names:
        best_row, best_score, matched = detect_header_row(xls, sheet, scan_rows)
        if best_score < min_matched:
            issues.append(
                f"工作表「{sheet}」：匹配列不足（{best_score}/{len(ALIASES)}），至少需要 {min_matched}。已匹配：{', '.join(matched.keys()) or '无'}"
            )
            continue

        df = pd.read_excel(xls, sheet_name=sheet, header=best_row).dropna(how="all")
        if len(df) < min_rows_per_sheet:
            issues.append(f"工作表「{sheet}」：数据行不足（{len(df)} < {min_rows_per_sheet}）")
            continue

        valid_sheets += 1

    if valid_sheets == 0:
        return False, "文件不满足基本要求：\n" + "\n".join(issues)

    expl = "" if not issues else "以下工作表已跳过/警告：\n" + "\n".join(issues) + "\n将继续处理有效工作表。"
    return True, expl


# ======================
# 读取并清洗（合并为一张 DataFrame）
# ======================
def process_excel(
    xls: pd.ExcelFile,
    *,
    scan_rows: int = SCAN_ROWS_DEFAULT,
    min_matched: int = 6,
    footer_scan_rows: int = 300,
) -> pd.DataFrame:
    """读取并清洗（合并为一张 DataFrame）。

    优化点：
    - 自动跳过 Summary/Benchmark 等非数据 sheet（依据表头匹配分数）
    - 对于超大 sheet，只在末尾 N 行做“footer/note”清理，避免整表逐行 apply 造成耗时
    """
    team_cache: Dict[str, str] = {}
    all_dfs: List[pd.DataFrame] = []

    for sheet in xls.sheet_names:
        header_row, score, _matched = detect_header_row(xls, sheet, scan_rows)
        if score < min_matched:
            continue

        df = pd.read_excel(xls, sheet_name=sheet, header=header_row).dropna(how="all")

        # footer/note 清理（仅扫末尾，提高性能）
        if len(df) > footer_scan_rows:
            cut = max(0, len(df) - footer_scan_rows)
            head = df.iloc[:cut]
            tail = _drop_footer_note_rows(df.iloc[cut:])
            df = pd.concat([head, tail], axis=0)
        else:
            df = _drop_footer_note_rows(df)

        sheet_upper = str(sheet).upper()
        if sheet_upper == "SO":
            df["Order Type"] = "SO"
        elif "AO" in sheet_upper or "FR" in sheet_upper:
            picked = _pick_order_type_column(df)
            if picked:
                df["Order Type"] = df[picked].astype(str).str.strip().str.upper()
            else:
                df["Order Type"] = "AO" if "AO" in sheet_upper else "FR"
            df = df[df["Order Type"].isin(ORDER_TYPE_ALLOWED)]

        cleaned = _rename_and_select(df)

        if "Product Type" in cleaned.columns:
            cleaned["Product Type"] = cleaned["Product Type"].apply(
                lambda x: str(x).strip().upper() if pd.notna(x) else x
            )

        if "Team" in cleaned.columns:
            unique_teams = cleaned["Team"].dropna().unique()
            team_mapping = {t: categorize_team(t, cache=team_cache) for t in unique_teams}
            cleaned["Team"] = cleaned["Team"].map(team_mapping)

        cleaned.insert(0, "_Sheet", sheet)
        cleaned = cleaned.loc[:, ~cleaned.columns.duplicated()]

        # footer 清理（仅扫末尾，提高性能）
        if len(cleaned) > footer_scan_rows:
            cut = max(0, len(cleaned) - footer_scan_rows)
            head = cleaned.iloc[:cut]
            tail = _remove_footer_rows(cleaned.iloc[cut:], sheet)
            cleaned = pd.concat([head, tail], axis=0)
        else:
            cleaned = _remove_footer_rows(cleaned, sheet)

        if len(cleaned) > 0 and not cleaned.isna().all().all():
            all_dfs.append(cleaned)

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame(columns=["_Sheet"] + TARGET_COLS)


# ======================
# Pivot 生成
# ======================


def _coerce_numeric_series(s: pd.Series) -> pd.Series:
    """将混合类型序列尽可能转为数值（用于 Pivot 计算）。

    背景：部分 OSR 导出的 Excel 中，`Order Qty` 列会被错误套用“日期格式”，
    pandas 读取后会变成 `datetime`（例如 1900-01-03）。
    若直接 `pd.to_numeric` 会变成 NaN，导致透视表该月份全部为 0。

    处理策略：
    - datetime / Timestamp：转换回 Excel serial number（openpyxl 的 to_excel）
    - time：换算为一天的比例（Excel 对时间的存储方式）
    - 字符串：去掉逗号、空格、括号负数等常见格式后再 to_numeric
    """
    if s is None:
        return pd.Series(dtype='float64')

    # already numeric
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors='coerce')

    ser = s.copy()

    # datetime64 dtype (rare in object-mixed columns)
    if pd.api.types.is_datetime64_any_dtype(ser):
        out = ser.map(lambda x: to_excel(pd.Timestamp(x).to_pydatetime()) if pd.notna(x) else None)
        return pd.to_numeric(out, errors='coerce')

    # object dtype: handle datetime/date/time and string
    # datetime / Timestamp
    mask_dt = ser.map(lambda x: isinstance(x, (datetime, pd.Timestamp)))
    if mask_dt.any():
        ser.loc[mask_dt] = ser.loc[mask_dt].map(
            lambda x: to_excel(x.to_pydatetime()) if isinstance(x, pd.Timestamp) else to_excel(x)
        )

    # date (without time)
    mask_date = ser.map(lambda x: isinstance(x, date) and not isinstance(x, datetime))
    if mask_date.any():
        ser.loc[mask_date] = ser.loc[mask_date].map(
            lambda d: to_excel(datetime.combine(d, datetime.min.time()))
        )

    # time -> fraction of day
    mask_time = ser.map(lambda x: isinstance(x, time))
    if mask_time.any():
        ser.loc[mask_time] = ser.loc[mask_time].map(
            lambda t: (t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6) / 86400
        )

    # strings cleaning
    mask_str = ser.map(lambda x: isinstance(x, str))
    if mask_str.any():
        cleaned = ser.loc[mask_str].astype(str).str.strip()
        cleaned = cleaned.str.replace(r'^\((.*)\)$', r'-\1', regex=True)
        cleaned = cleaned.str.replace(',', '', regex=False)
        cleaned = cleaned.str.replace(r'[^0-9eE\.\+\-]', '', regex=True)
        ser.loc[mask_str] = cleaned

    return pd.to_numeric(ser, errors='coerce')

def generate_pivot_tables(
    df: pd.DataFrame,
    *,
    value_col: str = "Order Qty",
    report_date: str,
    customer: str,
) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    if len(df) == 0:
        return {}, []

    df = df.copy()
    if value_col not in df.columns:
        df[value_col] = 0

    df[value_col] = _coerce_numeric_series(df[value_col]).fillna(0)
    df["Customer Delivery Date"] = pd.to_datetime(df["Customer Delivery Date"], errors="coerce")
    df["MonthLabel"] = df["Customer Delivery Date"].dt.strftime("%b-%y")

    df["Order Type"] = df["Order Type"].replace({"FR": "Forecast-FR"})

    min_date = df["Customer Delivery Date"].min()
    max_date = df["Customer Delivery Date"].max()
    if pd.notna(min_date) and pd.notna(max_date):
        all_months = pd.date_range(min_date, max_date, freq="MS").strftime("%b-%y").tolist()
    else:
        all_months = []

    product_types = sorted(df["Product Type"].dropna().unique())
    result: Dict[str, pd.DataFrame] = {}

    # NOTE: 输出列顺序要求：Order Type 紧挨在 Team 的左侧（即：... Order Type, Team ...）。
    # 这样在每个 Excel Table 中，筛选下拉的顺序也会按期望显示。
    base_cols = ["Factory", "Order Type", "Team", "Customer", "Product Type", "Date"]

    for pt in product_types:
        sub_df = df[df["Product Type"] == pt].copy()
        if len(sub_df) == 0:
            continue

        pivot = pd.pivot_table(
            sub_df,
            values=value_col,
            index=["Factory", "Team", "Order Type"],
            columns="MonthLabel",
            aggfunc="sum",
            fill_value=0,
        )
        pivot = pivot.reindex(columns=all_months, fill_value=0).reset_index()

        # 调整三大维度列顺序：Factory, Order Type, Team
        # （保留后续 sort 的逻辑：仍然按 Factory -> Team -> Order Type 排序更易读）
        pivot = pivot.reindex(columns=["Factory", "Order Type", "Team"] + [c for c in pivot.columns if c not in {"Factory", "Team", "Order Type"}])

        pivot["Customer"] = customer
        pivot["Product Type"] = pt
        pivot["Date"] = report_date

        month_2025 = [c for c in all_months if str(c).endswith("-25")]
        month_2026 = [c for c in all_months if str(c).endswith("-26")]
        pivot["2025 Ttl"] = pivot[month_2025].sum(axis=1) if month_2025 else 0
        pivot["2026 Ttl"] = pivot[month_2026].sum(axis=1) if month_2026 else 0
        pivot["Ttl"] = pivot["2025 Ttl"] + pivot["2026 Ttl"]

        pivot["__sort_key"] = pivot["Order Type"].map(ORDER_TYPE_SORT_ORDER).fillna(999)
        pivot = pivot.sort_values(by=["Factory", "Team", "__sort_key"], ascending=[True, True, True])
        pivot = pivot.drop(columns=["__sort_key"])

        final_cols = base_cols + all_months + ["2025 Ttl", "2026 Ttl", "Ttl"]
        pivot = pivot[[c for c in final_cols if c in pivot.columns]]
        result[pt] = pivot

    # 修剪：去掉全为 0 的前后月份
    if all_months and result:
        zero_months = {
            m
            for m in all_months
            if all(p.get(m, pd.Series([0])).sum() == 0 for p in result.values())
        }
        first_idx = 0
        while first_idx < len(all_months) and all_months[first_idx] in zero_months:
            first_idx += 1
        last_idx = len(all_months) - 1
        while last_idx >= first_idx and all_months[last_idx] in zero_months:
            last_idx -= 1
        trimmed_months = all_months[first_idx : last_idx + 1] if first_idx <= last_idx else []
    else:
        trimmed_months = []

    for pt, pivot in list(result.items()):
        cols_to_drop = set(all_months) - set(trimmed_months)
        pivot = pivot.drop(columns=cols_to_drop, errors="ignore")

        current_2025 = [c for c in trimmed_months if str(c).endswith("-25")]
        current_2026 = [c for c in trimmed_months if str(c).endswith("-26")]
        pivot["2025 Ttl"] = pivot[current_2025].sum(axis=1) if current_2025 else 0
        pivot["2026 Ttl"] = pivot[current_2026].sum(axis=1) if current_2026 else 0
        pivot["Ttl"] = pivot["2025 Ttl"] + pivot["2026 Ttl"]

        result[pt] = pivot

    return result, trimmed_months


# ======================
# Excel 输出
# ======================
def _format_headers_in_range(ws, header_row: int, num_cols: int):
    orange_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    light_blue_fill = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")
    for col_idx in range(1, num_cols + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        col_name = cell.value
        cell.fill = orange_fill if col_name == "Factory" else light_blue_fill


def _sanitize_table_name(name: str) -> str:
    """Sanitize table name to Excel rules."""
    sanitized = re.sub(r"[^a-zA-Z0-9_]", "_", str(name))
    if sanitized and sanitized[0].isdigit():
        sanitized = "T_" + sanitized
    if len(sanitized) > 255:
        sanitized = sanitized[:255]
    if not sanitized:
        sanitized = "Table1"
    return sanitized


def _get_unique_table_name(ws, base_name: str) -> str:
    base_sanitized = _sanitize_table_name(base_name)
    existing_names = set()

    if hasattr(ws, "tables") and ws.tables:
        existing_names.update(ws.tables.keys())

    if hasattr(ws, "parent") and ws.parent:
        for sheet in ws.parent.worksheets:
            if hasattr(sheet, "tables") and sheet.tables:
                existing_names.update(sheet.tables.keys())

    candidate = base_sanitized
    counter = 1
    while candidate in existing_names:
        suffix = f"_{counter}"
        max_base_len = 255 - len(suffix)
        candidate = base_sanitized[:max_base_len] + suffix
        counter += 1

    return candidate


def _write_table_block(
    ws,
    df: pd.DataFrame,
    title_text: str,
    start_row: int,
    report_date: str,
    factory_name: str,
    table_name: str,
    total_team_label: str = "Total",
    total_product_type_label: str = "ALL",
    allowed_filter_columns: Optional[List[str]] = None,
) -> int:
    """写入一个表块（标题 + 表头 + 数据 + Total 行），返回下一可用行。"""
    if df.empty:
        return start_row

    current_row = start_row
    num_cols = len(df.columns)
    end_col_letter = get_column_letter(num_cols)

    # Title row
    ws.merge_cells(f"A{current_row}:{end_col_letter}{current_row}")
    title_cell = ws.cell(row=current_row, column=1, value=title_text)
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="DDEBF7")
    ws.row_dimensions[current_row].height = 30
    current_row += 1

    # Header row
    header_row = current_row
    ws.append(list(df.columns))
    for c_idx in range(1, num_cols + 1):
        ws.cell(row=header_row, column=c_idx).font = Font(bold=True)
    _format_headers_in_range(ws, header_row, num_cols)
    current_row += 1

    # Data rows
    data_start_row = current_row
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    last_data_row = data_start_row + len(df) - 1
    current_row = last_data_row + 1

    # Create Excel Table for header + data rows only (excludes Total row)
    table_ref = f"A{header_row}:{end_col_letter}{last_data_row}"
    unique_table_name = _get_unique_table_name(ws, table_name)

    table = Table(displayName=unique_table_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False,
    )

    # Configure AutoFilter: show filter dropdowns only on allowed columns
    table.autoFilter = AutoFilter(ref=table_ref)

    headers = list(df.columns)
    allowed = set(allowed_filter_columns or ["Order Type"])
    allowed_colIds = {headers.index(c) for c in allowed if c in headers}

    # Hide filter buttons for all columns EXCEPT allowed columns
    table.autoFilter.filterColumn = [
        FilterColumn(colId=colId, hiddenButton=True)
        for colId in range(len(headers))
        if colId not in allowed_colIds
    ]

    ws.add_table(table)

    # Total row (pre-calculated values)
    # NOTE:
    # - openpyxl 不会计算公式，也不会写入公式的缓存结果；
    #   因此若使用 SUBTOTAL 等公式，部分预览器/未自动重算的 Excel 可能显示为空，
    #   造成“Total 行/Total 数值丢失”的观感。
    # - 这里改为：在 Python 端直接汇总并写入数值，确保打开即有结果。
    total_row_num = current_row

    # 输出列顺序：Factory, Order Type, Team, Customer, Product Type, Date
    base_cols = ["Factory", "Order Type", "Team", "Customer", "Product Type", "Date"]
    actual_month_cols = [
        c
        for c in df.columns
        if c not in base_cols and c not in ["2025 Ttl", "2026 Ttl", "Ttl"]
    ]

    def _sum_numeric(col: str) -> float:
        """安全求和：优先走数值 dtype；否则尝试转数值后求和。"""
        if col not in df.columns:
            return 0.0
        ser = df[col]
        if pd.api.types.is_numeric_dtype(ser):
            return float(ser.fillna(0).sum())
        s = pd.to_numeric(ser, errors="coerce").fillna(0)
        return float(s.sum())

    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=total_row_num, column=c_idx)
        cell.font = Font(bold=True)

        if col_name in base_cols:
            if col_name == "Factory":
                cell.value = factory_name
            elif col_name == "Customer":
                cell.value = (
                    df["Customer"].iloc[0]
                    if "Customer" in df.columns and len(df) > 0
                    else CUSTOMER_LABEL_DEFAULT
                )
            elif col_name == "Team":
                cell.value = total_team_label
            elif col_name == "Product Type":
                cell.value = total_product_type_label
            elif col_name == "Date":
                cell.value = report_date
            elif col_name == "Order Type":
                cell.value = "Total"

        elif col_name in actual_month_cols or col_name in ("2025 Ttl", "2026 Ttl", "Ttl"):
            total_val = _sum_numeric(col_name)
            # 如果是整数（如 Order Qty），尽量写成 int，避免出现 1234.0
            if abs(total_val - round(total_val)) < 1e-9:
                cell.value = int(round(total_val))
            else:
                cell.value = total_val

        else:
            cell.value = 0

    current_row += 1
    return current_row


def _write_factory_sheet_no_table(
    ws,
    factory_name: str,
    pivot_tables: Dict[str, pd.DataFrame],
    report_date: str,
    metric_label: str,
) -> bool:
    """每个 Factory sheet：

    结构：
    1) 顶部：ALL DATA 汇总表（All Teams + Product Types）
    2) 下方：按 **Team -> Product Type** 拆分多个小表（满足“同 Team 聚在一起”的查看需求）
    """

    # Build consolidated DataFrame for this factory
    parts: List[pd.DataFrame] = []
    for product_type in sorted(pivot_tables.keys()):
        pivot_df = pivot_tables[product_type]
        if pivot_df is None or pivot_df.empty:
            continue
        sub = pivot_df[pivot_df["Factory"] == factory_name].copy()
        if not sub.empty:
            parts.append(sub)

    if not parts:
        return False

    factory_all_df = pd.concat(parts, ignore_index=True)

    # Sort for readability (Team -> Product Type -> Order Type)
    if "Order Type" in factory_all_df.columns:
        factory_all_df["__sort_key"] = factory_all_df["Order Type"].map(ORDER_TYPE_SORT_ORDER).fillna(999)

        sort_cols: List[str] = []
        for c in ("Team", "Product Type", "__sort_key"):
            if c in factory_all_df.columns:
                sort_cols.append(c)

        if sort_cols:
            factory_all_df = factory_all_df.sort_values(by=sort_cols, ascending=True)

        factory_all_df = factory_all_df.drop(columns=["__sort_key"], errors="ignore").reset_index(drop=True)

    current_row = 1
    table_counter = 1
    safe_factory = _sanitize_table_name(factory_name)

    # ----------------------
    # Consolidated ALL DATA
    # ----------------------
    if not factory_all_df.empty:
        title_text = f"{factory_name} — ALL --- {metric_label} (All Teams + Product Types)"
        table_name = f"T_{safe_factory}_ALL_{table_counter}"
        table_counter += 1

        current_row = _write_table_block(
            ws,
            factory_all_df,
            title_text,
            current_row,
            report_date,
            factory_name,
            table_name,
            total_team_label="Total",
            total_product_type_label="ALL",
            allowed_filter_columns=["Team", "Order Type", "Product Type"],
        )

        # spacing
        current_row += 3

    # ----------------------
    # Split tables (Team -> Product Type)
    # ----------------------
    team_to_parts: Dict[str, List[Tuple[str, pd.DataFrame]]] = {}

    for product_type in sorted(pivot_tables.keys()):
        pivot_df = pivot_tables[product_type]
        if pivot_df is None or pivot_df.empty:
            continue

        factory_pt_df = pivot_df[pivot_df["Factory"] == factory_name].copy()
        if factory_pt_df.empty:
            continue

        team_series = factory_pt_df["Team"].fillna("").astype(str)
        for team in pd.unique(team_series):
            team_df = factory_pt_df.loc[team_series == team].copy()
            if team_df.empty:
                continue
            team_to_parts.setdefault(team, []).append((product_type, team_df))

    team_order = {
        "Sports": 0,
        "Fancy": 1,
        "SW": 2,
        "Cotton Panty": 3,
        "Brands-COS": 4,
        "Brands-Stories": 5,
    }

    def _team_sort_key(t: str):
        t_clean = str(t).strip()
        if not t_clean:
            return (2, 999, "")
        if t_clean in team_order:
            return (0, team_order[t_clean], t_clean)
        return (1, 999, t_clean)

    for team in sorted(team_to_parts.keys(), key=_team_sort_key):
        parts_list = team_to_parts[team]
        # within same team, sort by Product Type
        parts_list_sorted = sorted(parts_list, key=lambda x: str(x[0]))

        for product_type, team_df in parts_list_sorted:
            team_title = team if str(team).strip() else "（空）"
            title_text = f"{factory_name} — {team_title} --- {product_type} ({metric_label})"

            safe_team = _sanitize_table_name(str(team))
            safe_pt = _sanitize_table_name(str(product_type))
            table_name = f"T_{safe_factory}_{safe_team}_{safe_pt}_{table_counter}"
            table_counter += 1

            current_row = _write_table_block(
                ws,
                team_df,
                title_text,
                current_row,
                report_date,
                factory_name,
                table_name,
                total_team_label="Total",
                total_product_type_label=product_type,
            )

            current_row += 3

    # ----------------------
    # Column widths
    # ----------------------
    max_col = max(ws.max_column, 1)
    # 输出列顺序已调整为：Factory, Order Type, Team, Customer, Product Type, Date, ...
    for col_idx in range(1, max_col + 1):
        if col_idx == 1:  # Factory
            ws.column_dimensions[get_column_letter(col_idx)].width = 12
        elif col_idx == 2:  # Order Type
            ws.column_dimensions[get_column_letter(col_idx)].width = 14
        elif col_idx == 3:  # Team
            ws.column_dimensions[get_column_letter(col_idx)].width = 22
        elif col_idx in (4, 5, 6):  # Customer / Product Type / Date
            ws.column_dimensions[get_column_letter(col_idx)].width = 14
        else:
            ws.column_dimensions[get_column_letter(col_idx)].width = 11

    return True


# ======================
# 工作表名称安全处理
# ======================
_SHEET_INVALID_CHARS_RE = re.compile(r"[:\\/?*\[\]]")


def _safe_sheet_name(wb, desired: str) -> str:
    name = _SHEET_INVALID_CHARS_RE.sub("_", str(desired)).strip()
    name = name[:31].rstrip()
    if not name:
        name = "Sheet1"

    base = name
    i = 2
    while name in wb.sheetnames:
        suffix = f" ({i})"
        max_len = 31 - len(suffix)
        name = (base[:max_len].rstrip()) + suffix
        i += 1
    return name


def _write_workbook_no_table(writer, pivot_tables: Dict[str, pd.DataFrame], report_date: str, metric_label: str):
    all_factories = set()
    for pt_df in pivot_tables.values():
        all_factories.update(pt_df["Factory"].dropna().unique())

    for factory in sorted(all_factories):
        factory_str = str(factory)
        desired_name = f"{factory_str} - {metric_label}"
        safe_name = _safe_sheet_name(writer.book, desired_name)
        ws = writer.book.create_sheet(safe_name)
        has_data = _write_factory_sheet_no_table(ws, factory_str, pivot_tables, report_date, metric_label)
        if not has_data:
            writer.book.remove(ws)


def _safe_remove(path: str, retries: int = 12, delay_s: float = 0.15) -> None:
    for _ in range(retries):
        try:
            os.remove(path)
            return
        except FileNotFoundError:
            return
        except PermissionError:
            time_mod.sleep(delay_s)
    try:
        os.remove(path)
    except Exception:
        pass


def pivot_report_multi_to_xlsx_bytes_no_table(
    pivot_tables_by_metric: Dict[str, Dict[str, pd.DataFrame]],
    report_date: str,
) -> bytes:
    """生成一个 Excel（包含多个指标的多个 sheet），返回 bytes。"""
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            for metric_label, pivot_tables in pivot_tables_by_metric.items():
                _write_workbook_no_table(writer, pivot_tables, report_date, metric_label)

            wb = writer.book
            try:
                wb.calculation.calcMode = "auto"
                wb.calculation.fullCalcOnLoad = True
                wb.calculation.forceFullCalc = True
            except Exception:
                pass

            if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
                ws0 = wb["Sheet"]
                if ws0.max_row == 1 and ws0.max_column == 1 and ws0["A1"].value is None:
                    wb.remove(ws0)

        with open(temp_path, "rb") as f:
            return f.read()

    finally:
        _safe_remove(temp_path)


# ======================
# 高层封装：从上传文件 bytes 生成报表
# ======================
class ReportError(RuntimeError):
    """用于 UI 友好提示的异常。"""


def generate_pivot_report_from_upload(
    excel_bytes: bytes,
    *,
    filename: str,
    scan_rows: int = SCAN_ROWS_DEFAULT,
    customer_label: str = CUSTOMER_LABEL_DEFAULT,
    report_date: Optional[str] = None,
) -> Tuple[bytes, Dict[str, Any]]:
    """从上传的 Excel bytes 生成 pivot 报表。

    Returns:
        (pivot_xlsx_bytes, stats)

    stats:
        rows_used: int
        factories: List[str]
        product_types: List[str]
        months: Dict[str, List[str]]  # metric -> months
        report_date: str
    """
    if report_date is None:
        report_date = datetime.now().strftime("%b-%d")

    if not filename:
        raise ReportError("缺少文件名。")

    ext = os.path.splitext(filename)[1].lower()

    warnings_text: str = ""

    try:
        if ext != ".xls":
            excel_io = io.BytesIO(excel_bytes)
            with pd.ExcelFile(excel_io, engine="openpyxl") as xls:
                ok, expl = sanity_check(xls, scan_rows=scan_rows)
                if not ok:
                    raise ReportError(expl)

                warnings_text = expl or ""

                combined_df = process_excel(xls, scan_rows=scan_rows)

        else:
            # .xls：需要 xlrd
            try:
                import xlrd  # noqa: F401
            except Exception:
                raise ReportError("当前环境未安装 xlrd，无法读取 .xls。请安装 xlrd==2.0.1 或将文件另存为 .xlsx。")

            fd, tmp_xls = tempfile.mkstemp(suffix=".xls")
            os.close(fd)
            try:
                with open(tmp_xls, "wb") as f:
                    f.write(excel_bytes)
                with pd.ExcelFile(tmp_xls, engine="xlrd") as xls:
                    ok, expl = sanity_check(xls, scan_rows=scan_rows)
                    if not ok:
                        raise ReportError(expl)

                    warnings_text = expl or ""
                    combined_df = process_excel(xls, scan_rows=scan_rows)
            finally:
                _safe_remove(tmp_xls)

    except ReportError:
        raise
    except Exception as e:
        raise ReportError(f"无法读取 Excel 文件：{e}")

    if len(combined_df) == 0:
        raise ReportError("清洗后没有找到可用数据行，请检查 Excel 格式是否符合要求。")

    # Generate pivot tables for each metric
    pivot_tables_by_metric: Dict[str, Dict[str, pd.DataFrame]] = {}
    month_cols_by_metric: Dict[str, List[str]] = {}

    for metric_label, metric_col in METRIC_SPECS:
        pivot_tables, month_cols = generate_pivot_tables(
            combined_df,
            value_col=metric_col,
            report_date=report_date,
            customer=customer_label,
        )
        if not pivot_tables:
            raise ReportError(f"未能生成透视表（{metric_label}），请检查该列是否存在有效值。")
        pivot_tables_by_metric[metric_label] = pivot_tables
        month_cols_by_metric[metric_label] = month_cols

    pivot_bytes = pivot_report_multi_to_xlsx_bytes_no_table(pivot_tables_by_metric, report_date)

    stats = {
        "report_date": report_date,
        "warnings": warnings_text,
        "rows_used": int(len(combined_df)),
        "factories": sorted(combined_df["Factory"].dropna().astype(str).unique().tolist()),
        "product_types": sorted(combined_df["Product Type"].dropna().astype(str).unique().tolist()),
        "months": month_cols_by_metric,
    }

    return pivot_bytes, stats
