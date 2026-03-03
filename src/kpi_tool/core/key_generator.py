# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Optional

import pandas as pd
import numpy as np

from kpi_tool.config import constants as C
from kpi_tool.core.standardizer import auto_match_column_safe, KeyNormalizer

def _first_non_empty_series(df: pd.DataFrame, cols: list) -> pd.Series:
    """按列优先级逐行取第一个非空值（空定义：NaN / '' / 'nan'）"""
    cols = [c for c in cols if c in df.columns]
    if not cols:
        return pd.Series([""] * len(df), index=df.index)
    s = df[cols[0]]
    for c in cols[1:]:
        cur = df[c]
        # 判空
        s_str = s.astype(str).str.strip().str.lower()
        s_empty = s.isna() | (s_str.eq("")) | (s_str.eq("nan")) | (s_str.eq("none"))
        if s_empty.any():
            s = s.where(~s_empty, cur)
    return s

def generate_list_keys(df):
    if df is None or df.empty: return df
    ecgi_col = auto_match_column_safe(df.columns, ['CGI/ECGI', 'ECGI', 'CGI', 'CI'])
    df['list_key_id'] = KeyNormalizer.normalize_id(df[ecgi_col]) if ecgi_col else ""
    name_col = auto_match_column_safe(df.columns, ['小区中文名', '小区名称'])
    if name_col: df['list_key_name'] = KeyNormalizer.normalize_name(df[name_col])
    else: df['list_key_name'] = ""
    return df

def generate_raw_keys(df, tech_type):
    if df is None or df.empty: 
        return df
    df['raw_key_name'] = ""
    df['raw_key_id'] = ""
    df['raw_key_fuzzy_base'] = ""

    # —— 关键修复：用 STD__cell_name（各厂家已标准化）做主 name key，逐行兜底 —— #
    name_cols = [
        'STD__cell_name',
        'DU物理小区名称', 'DU物理小区名',
        '小区中文名', '小区名称',
        'CEL_NAME', 'CellName'
    ]
    name_s = _first_non_empty_series(df, name_cols)
    df['raw_key_name'] = KeyNormalizer.normalize_name(name_s)
    df['raw_key_fuzzy_base'] = df['raw_key_name'].apply(
        lambda x: str(x)[C.PREFIX_STRIP_LEN:] if len(str(x)) > C.PREFIX_STRIP_LEN else str(x)
    )

    # —— ID key —— #
    if tech_type == '5G':
        # 5G：优先 masterOperatorId（中兴），其它厂家一般走 NAME_key
        id_cols = ['masterOperatorId', 'Global Cell ID', 'ECGI', 'CGI', 'ENB_CELL']
        id_s = _first_non_empty_series(df, id_cols)
        df['raw_key_id'] = KeyNormalizer.normalize_id(id_s)
    else:
        # 4G：优先 ENB_CELL/CGI，兜底 46000 + ENB_ID + CELL_ID / eNodeBId+cellId
        cgi_cols = ['ENB_CELL', 'CGI']
        cgi_s = _first_non_empty_series(df, cgi_cols)
        cgi_norm = KeyNormalizer.normalize_id(cgi_s)

        enb_cols = ['eNodeBId', 'eNBID', 'ENBID', 'EN_ID', 'ENB_ID']
        cid_cols = ['cellId', 'CellId', 'CI', 'CELL_ID', '本地小区标识']
        enb_s = _first_non_empty_series(df, enb_cols)
        cid_s = _first_non_empty_series(df, cid_cols)
        enb_norm = KeyNormalizer.normalize_id(enb_s)
        cid_norm = KeyNormalizer.normalize_id(cid_s)

        join = pd.Series([""] * len(df), index=df.index)
        mask_valid = (enb_norm != "") & (cid_norm != "")
        join.loc[mask_valid] = "46000" + enb_norm[mask_valid] + cid_norm[mask_valid]

        # 优先用 CGI/ENB_CELL（如果有值）
        df['raw_key_id'] = np.where(cgi_norm != "", cgi_norm, join)

    return df
