# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime
import re
from typing import Any, Iterable, Optional

import numpy as np
import pandas as pd

def _time_str_from_ts(ts_start, ts_end=None, default_interval_min=15):
    """统一生成 'HH:MM~HH:MM' 时间段字符串（仅用于展示/选择，不参与计算口径）。"""
    try:
        if ts_start is None or pd.isna(ts_start):
            return "无时间信息"
        ts_start = pd.to_datetime(ts_start, errors='coerce')
        if pd.isna(ts_start):
            return "无时间信息"
        if ts_end is None or pd.isna(ts_end):
            ts_end = ts_start + pd.Timedelta(minutes=default_interval_min)
        else:
            ts_end = pd.to_datetime(ts_end, errors='coerce')
            if pd.isna(ts_end):
                ts_end = ts_start + pd.Timedelta(minutes=default_interval_min)
        return f"{ts_start.strftime('%H:%M')}~{ts_end.strftime('%H:%M')}"
    except Exception:
        return "无时间信息"

def _strip_internal_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    """移除内部辅助列，避免影响既有导出/对账内容。"""
    if df is None or df.empty:
        return df
    drop_cols = ['_ts', '_cell_key', 'SEG__time_str', 'SEG__ts_start']
    for c in drop_cols:
        if c in df.columns:
            try:
                del df[c]
            except Exception:
                pass
    return df

def _safe_sheet_name(name: str, used: set) -> str:
    """Excel sheet 名称最长 31 字符，且不能包含 : \\ / ? * [ ]"""
    if name is None:
        name = ""
    s = str(name)
    s = re.sub(r'[:\\/\?\*\[\]]', '_', s).strip()
    if not s:
        s = "Sheet"
    s = s[:31]
    base = s
    i = 1
    while s in used:
        suffix = f"_{i}"
        s = (base[:31 - len(suffix)] + suffix)[:31]
        i += 1
    used.add(s)
    return s

def format_time_range(row):
    try:
        ts = row.get('STD__time_start'); te = row.get('STD__time_end')
        if pd.isna(ts): return ""
        return f"{ts.strftime('%H:%M')}~{te.strftime('%H:%M')}" if pd.notna(te) else ts.strftime('%H:%M')
    except Exception: return ""
