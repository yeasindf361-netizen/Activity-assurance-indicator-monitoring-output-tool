# -*- coding: utf-8 -*-
from __future__ import annotations

import copy
import re
from typing import Iterable, List, Optional

import numpy as np
import pandas as pd

def _robust_parse_time_part(x):
    import kpi_tool.core.time_handler as TH
    return TH._robust_parse_time_part(x)

def _robust_time_any_to_start_end(x):
    import kpi_tool.core.time_handler as TH
    return TH._robust_time_any_to_start_end(x)


from kpi_tool.config import constants as C

def clean_header_name(header):
    # 说明：原版本会把括号内容整体删除，导致 (QCI=1)/(5QI1) 等关键识别信息丢失，
    # 从而使 kpi_connect 误匹配到 QCI=1 字段，进而造成“无线接通率/掉线率”等指标异常。
    # 新逻辑：保留括号内文本，仅移除括号符号，以便后续匹配可以正确区分 QCI/VoLTE/VoNR 等字段。
    if not isinstance(header, str):
        return str(header)
    h = str(header)
    h = re.sub(r'_\d{10,}', '', h)
    h = h.replace('（', '(').replace('）', ')')
    # 仅移除括号符号，不删除括号内文本
    h = re.sub(r'[()]', '', h)
    # 去空白
    h = re.sub(r'\s+', '', h)
    return h.strip().upper()

def auto_match_column_safe(columns, keywords, exclusions=None):
    if isinstance(keywords, str): keywords = [keywords]
    if isinstance(exclusions, str): exclusions = [exclusions]
    col_map_cleaned = {clean_header_name(c): c for c in columns}
    for kw in keywords:
        kw_upper = clean_header_name(str(kw))
        if kw_upper in col_map_cleaned:
            raw_col = col_map_cleaned[kw_upper]
            if exclusions and any(ex.upper() in str(raw_col).upper() for ex in exclusions): pass
            else: return raw_col
        for col_clean, original_col in col_map_cleaned.items():
            if kw_upper in col_clean:
                if exclusions and (any(ex.upper() in col_clean for ex in exclusions) or any(ex.upper() in str(original_col).upper() for ex in exclusions)): continue
                return original_col
    return None

class KeyNormalizer:
    @staticmethod
    def normalize_id(series: pd.Series) -> pd.Series:
        if series is None or series.empty: return series
        return series.astype(str).str.replace(r"\s+", "", regex=True)\
                .str.replace(r"\.0$", "", regex=True)\
                .str.replace(r"\D", "", regex=True)

    @staticmethod
    def normalize_name(series: pd.Series) -> pd.Series:
        if series is None or series.empty: return series
        return series.astype(str).str.replace(r"\s+", "", regex=True)\
                .str.replace(r"\.0$", "", regex=True)

class FieldStandardizer:
    def __init__(self):
        self.global_candidates = copy.deepcopy(C.GLOBAL_CANDIDATES)

    def standardize_df(self, df, rat):
        if df is None or df.empty: return df
        df_std = df.copy()
        for std, cfg in self.global_candidates.items():
            if isinstance(cfg, tuple): cands, exclusions = cfg
            else: cands, exclusions = cfg, None
            matched_col = auto_match_column_safe(df.columns, cands, exclusions)
            std_col = f"STD__{std}"
            if matched_col:
                df_std[f"SRC__{std}"] = matched_col
                df_std[f"RAW__{std}"] = df[matched_col]
                # 交通量单位兼容：部分厂家输出为 MB，需要换算为 GB（保持 STD__kpi_traffic_gb 语义不变）
                if std == 'kpi_traffic_gb' and isinstance(matched_col, str) and ('MB' in matched_col.upper() or 'MB' in matched_col):
                    df_std[std_col] = pd.to_numeric(df[matched_col], errors='coerce') / 1024.0
                # 华为部分 5G 指标 “5QI话务量(erl)_t” 口径偏大：按历史对账口径/10
                elif std == 'kpi_vonr_traffic_erl' and isinstance(matched_col, str) and ('5QI话务量' in matched_col):
                    df_std[std_col] = pd.to_numeric(df[matched_col], errors='coerce') / 10.0
                else:
                    df_std[std_col] = df[matched_col]
            else:
                df_std[std_col] = np.nan
                df_std[f"SRC__{std}"] = ""
                df_std[f"RAW__{std}"] = np.nan
                # --- 时间字段强健解析（关键修复：避免诺基亚DAY被当作纳秒时间 => 1970） ---
        if 'STD__time_start' in df_std.columns:
            st_raw = df_std['STD__time_start']
            # 支持 "start;end"
            if st_raw.astype(str).str.contains(';', na=False).any():
                _st, _ed = _robust_time_any_to_start_end(st_raw)
                df_std['STD__time_start'] = _st
                df_std['STD__time_end'] = _ed
            else:
                df_std['STD__time_start'] = _robust_parse_time_part(st_raw)
                # 结束时间若存在则解析
                if 'STD__time_end' in df_std.columns and not df_std['STD__time_end'].isna().all():
                    df_std['STD__time_end'] = _robust_parse_time_part(df_std['STD__time_end'])
                # 若结束时间全空，默认 +15min
                if 'STD__time_end' in df_std.columns and df_std['STD__time_end'].isna().all():
                    df_std['STD__time_end'] = df_std['STD__time_start'] + pd.to_timedelta(15, unit='m')

        # 若 time_start 仍全 NaT，再用 time_any 补
        if df_std['STD__time_start'].isna().all() and not df_std['STD__time_any'].isna().all():
            _st, _ed = _robust_time_any_to_start_end(df_std['STD__time_any'])
            df_std['STD__time_start'] = _st
            df_std['STD__time_end'] = _ed
        return df_std

# 全局标准化器（与旧版保持一致：单例）
standardizer = FieldStandardizer()

