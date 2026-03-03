# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime
import glob
import os
import re
from typing import Any, Dict, List, Tuple

import numpy as np
import pandas as pd

from kpi_tool.io.data_loader import load_data_frame

from kpi_tool.config.logging_config import log_print
from kpi_tool.core.key_generator import generate_list_keys, generate_raw_keys
from kpi_tool.core.standardizer import auto_match_column_safe
from kpi_tool.utils.helpers import _time_str_from_ts, _strip_internal_time_cols

# ================= 时间段缓存（保持旧版的全局字典结构） =================

_ALL_TIME_SEGMENT_DFS: Dict[str, Dict[str, pd.DataFrame]] = {'4G': {}, '5G': {}}
_ALL_TIME_SEGMENT_META: Dict[str, Dict[str, Any]] = {'4G': {}, '5G': {}}
_ALL_TIME_SEGMENTS: Dict[str, Dict[str, List[Tuple[datetime.datetime, datetime.datetime]]]] = {'4G': {}, '5G': {}}
_SELECTED_TIME_SEGMENTS: Dict[str, Tuple[datetime.datetime, datetime.datetime]] = {}
# Backwards-compatibility alias used by older modules
_ALL_RAW_FULL_WITH_SEG: Dict[str, pd.DataFrame] = {'4G': None, '5G': None}
_TIME_STR_TO_TS: Dict[str, Dict[str, List[datetime.datetime]]] = {'4G': {}, '5G': {}}

class TimeSegmentCache:
    """封装旧版时间段全局缓存（单例）。"""
    _instance: 'TimeSegmentCache | None' = None

    def __new__(cls) -> 'TimeSegmentCache':
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def set_segments(self, tech: str, dfs: Dict[str, pd.DataFrame], meta: Dict[str, Any], ts_map: Dict[str, List[Tuple[datetime.datetime, datetime.datetime]]]) -> None:
        _ALL_TIME_SEGMENT_DFS[tech] = dfs
        _ALL_TIME_SEGMENT_META[tech] = meta
        _ALL_TIME_SEGMENTS[tech] = ts_map

    def get_full(self, tech: str) -> Dict[str, pd.DataFrame]:
        return _ALL_TIME_SEGMENT_DFS.get(tech, {})

    def get_meta(self, tech: str) -> Dict[str, Any]:
        return _ALL_TIME_SEGMENT_META.get(tech, {})

    def get_segments(self, tech: str) -> Dict[str, List[Tuple[datetime.datetime, datetime.datetime]]]:
        return _ALL_TIME_SEGMENTS.get(tech, {})

    def set_selected(self, selected: Dict[str, Tuple[datetime.datetime, datetime.datetime]]) -> None:
        _SELECTED_TIME_SEGMENTS.clear()
        _SELECTED_TIME_SEGMENTS.update(selected)

    def get_selected(self) -> Dict[str, Tuple[datetime.datetime, datetime.datetime]]:
        return _SELECTED_TIME_SEGMENTS

CACHE = TimeSegmentCache()

def _robust_parse_time_part(part: pd.Series) -> pd.Series:
    if part is None or part.empty:
        return part
    s = part.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=s.index)
    # 12位：YYYYMMDDHHMM
    m12 = s.str.fullmatch(r"\d{12}")
    if m12.any():
        out.loc[m12] = pd.to_datetime(s.loc[m12], format="%Y%m%d%H%M", errors="coerce")
    # 14位：YYYYMMDDHHMMSS
    m14 = s.str.fullmatch(r"\d{14}")
    if m14.any():
        out.loc[m14] = pd.to_datetime(s.loc[m14], format="%Y%m%d%H%M%S", errors="coerce")
    # 8位：YYYYMMDD
    m8 = s.str.fullmatch(r"\d{8}")
    if m8.any():
        out.loc[m8] = pd.to_datetime(s.loc[m8], format="%Y%m%d", errors="coerce")
    # 其它：让 pandas 自行解析（例如 2026-01-07 15:00）
    rest = ~(m12 | m14 | m8)
    if rest.any():
        out.loc[rest] = pd.to_datetime(s.loc[rest], errors="coerce")
    return out

def _robust_time_any_to_start_end(time_any: pd.Series):
    if time_any is None or time_any.empty:
        return time_any, time_any
    s = time_any.astype(str).str.strip()
    # 支持 "start;end"
    has_sep = s.str.contains(";", na=False)
    start_part = s.copy()
    end_part = pd.Series("", index=s.index)
    if has_sep.any():
        spl = s.loc[has_sep].str.split(";", n=1, expand=True)
        start_part.loc[has_sep] = spl[0].astype(str)
        end_part.loc[has_sep] = spl[1].astype(str)
    start_ts = _robust_parse_time_part(start_part)
    end_ts = _robust_parse_time_part(end_part)
    # 没有显式 end 或 end 解析失败：默认 +15min
    end_ts = end_ts.where(end_ts.notna(), start_ts + pd.to_timedelta(15, unit="m"))
    return start_ts, end_ts

def _select_latest_per_cell(full_df: pd.DataFrame, tech_type: str) -> pd.DataFrame:
    """
    默认模式：每个小区各自取最新可用时间段（最大化覆盖率）。
    - 对每个小区保留其最新时间段下的全部记录（同一时间段多条保留）。
    """
    if full_df is None or full_df.empty:
        return full_df

    if '_ts' not in full_df.columns:
        return full_df

    df = full_df.copy()
    df = generate_raw_keys(df, tech_type)

    # 构造小区唯一键：优先 ID，其次名称（兼容不同厂家字段缺失）
    id_part = df['raw_key_id'].astype(str) if 'raw_key_id' in df.columns else ""
    name_part = df['raw_key_name'].astype(str) if 'raw_key_name' in df.columns else ""
    df['_cell_key'] = (id_part.fillna("").str.strip() + "||" + name_part.fillna("").str.strip()).astype(str)

    # 若 key 仍为空，则兜底用 STD__cell_name/原始小区名
    if (df['_cell_key'] == "||").all():
        cell_col = auto_match_column_safe(df.columns, ['STD__cell_name', '小区名称', '小区中文名', 'CELL_NAME', 'cell_name'])
        if cell_col:
            df['_cell_key'] = df[cell_col].astype(str).fillna("").str.strip()

    # 计算每个小区的最新时间（忽略 NaT，若全为 NaT 则保留 NaT 记录）
    max_ts = df.groupby('_cell_key')['_ts'].transform('max')
    keep_mask = (df['_ts'] == max_ts) | (df['_ts'].isna() & max_ts.isna())
    out = df.loc[keep_mask].copy()
    return out

def _ensure_list_keys_for_time_select(df_list: pd.DataFrame) -> pd.DataFrame:
    """与 waterfall_merge 保持一致的清单 key 生成（仅用于时间段覆盖统计/选择）。"""
    if df_list is None or df_list.empty:
        return df_list

    df = df_list.copy()
    if '活动名称' not in df.columns:
        act_col = auto_match_column_safe(df.columns, ['活动名称', '保障活动', '活动'])
        if act_col and act_col != '活动名称':
            df.rename(columns={act_col: '活动名称'}, inplace=True)
        elif '活动名称' not in df.columns:
            df['活动名称'] = "未知活动"

    if 'list_key_name' not in df.columns:
        name_col = auto_match_column_safe(df.columns, ['小区中文名', '小区名称', 'CELL_NAME', '小区名'])
        df['list_key_name'] = df[name_col].astype(str).str.strip() if name_col else ""

    if 'list_key_id' not in df.columns:
        id_col = auto_match_column_safe(df.columns, ['ECGI', 'CGI', 'ENB_CELL', '小区ID', 'NCI'])
        df['list_key_id'] = (df[id_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                             if id_col else "")

    if '_list_idx' not in df.columns:
        df['_list_idx'] = np.arange(len(df), dtype=int)

    return df

def get_available_time_segments(list_file_path: str, raw_dir_4g: str, raw_dir_5g: str) -> dict:
    """
    返回每个活动在网管数据中实际覆盖的时间段及小区覆盖情况。
    注意：该函数只做“展示/选择用统计”，不改变任何计算口径。
    返回示例：
    {
        "活动A": [
            {"time_str":"14:00~14:15","start":datetime,"end":datetime,"cell_count":12,"vendor_breakdown":{"华为":8,"诺基亚":4}},
            {"time_str":"15:00~15:15","start":datetime,"end":datetime,"cell_count":10,"vendor_breakdown":{"华为":10}},
        ],
        "活动B": [...]
    }
    """
    # 保障清单
    df_list = load_data_frame(list_file_path)
    if df_list is None or df_list.empty:
        return {}

    # 与 main() 一致的列归一（避免 GUI/CLI 不一致）
    region_col_raw = auto_match_column_safe(df_list.columns, ['区域', '场景'])
    if region_col_raw and region_col_raw != '区域':
        df_list.rename(columns={region_col_raw: '区域'}, inplace=True)
    if '区域' not in df_list.columns:
        df_list['区域'] = '整体'

    act_col = auto_match_column_safe(df_list.columns, ['活动名称', '保障活动', '活动'])
    if act_col and act_col != '活动名称':
        df_list.rename(columns={act_col: '活动名称'}, inplace=True)
    if '活动名称' not in df_list.columns:
        df_list['活动名称'] = "未知活动"
    df_list['_list_idx'] = np.arange(len(df_list), dtype=int)

    df_list = _ensure_list_keys_for_time_select(df_list)

    # 若尚未缓存全量时间段数据，则先加载（不会影响后续 main 的运行）
    if _ALL_RAW_FULL_WITH_SEG.get('4G') is None and os.path.exists(raw_dir_4g):
        load_and_distribute(raw_dir_4g, '4G')
    if _ALL_RAW_FULL_WITH_SEG.get('5G') is None and os.path.exists(raw_dir_5g):
        load_and_distribute(raw_dir_5g, '5G')

    # 汇总 4G/5G 覆盖（按活动+time_str）
    result = {}
    activities = df_list['活动名称'].astype(str).fillna("未知活动").unique().tolist()

    # 为避免 args 未初始化导致 waterfall_merge 崩溃，此处不依赖 getattr(args, 'only_activity', None) 过滤
    for act in activities:
        result[act] = {}

    for tech in ['4G', '5G']:
        raw_full = _ALL_RAW_FULL_WITH_SEG.get(tech)
        if raw_full is None or raw_full.empty:
            continue

        try:
            merged = waterfall_merge(raw_full, df_list, tech)
        except Exception:
            continue

        if merged is None or merged.empty:
            continue

        # 缺少 time_str 时兜底
        if 'SEG__time_str' not in merged.columns:
            merged['SEG__time_str'] = "无时间信息"

        # cell_count：以清单行(_list_idx)去重计数
        grp = merged.dropna(subset=['_list_idx']).groupby(['活动名称', 'SEG__time_str'])['_list_idx'].nunique()
        # vendor_breakdown
        gvend = merged.dropna(subset=['_list_idx']).groupby(['活动名称', 'SEG__time_str', '厂家'])['_list_idx'].nunique()

        for (act, tstr), cnt in grp.items():
            d = result.setdefault(str(act), {})
            entry = d.get(tstr, {"cell_count": 0, "vendor_breakdown": {}})
            entry["cell_count"] = int(entry.get("cell_count", 0) + int(cnt))
            d[tstr] = entry

        for (act, tstr, vendor), cnt in gvend.items():
            d = result.setdefault(str(act), {})
            entry = d.get(tstr, {"cell_count": 0, "vendor_breakdown": {}})
            vb = entry.get("vendor_breakdown", {})
            vname = str(vendor) if vendor is not None else "Unknown"
            vb[vname] = int(vb.get(vname, 0) + int(cnt))
            entry["vendor_breakdown"] = vb
            d[tstr] = entry

    # 补充 start/end，并输出为列表（按时间排序）
    final = {}
    for act, seg_map in result.items():
        seg_list = []
        for tstr, d in seg_map.items():
            # start/end 推断：取 4G/5G 中同 time_str 的最新 start（若存在）
            start_dt = None
            end_dt = None
            candidates = []
            for tech in ['4G', '5G']:
                for ts in _TIME_STR_TO_TS.get(tech, {}).get(tstr, []):
                    if ts is not None and pd.notna(ts):
                        candidates.append(pd.to_datetime(ts))
            if candidates:
                start_dt = max(candidates)
                # end：优先从 meta 中取
                end_candidates = []
                for tech in ['4G', '5G']:
                    for ts, meta in _ALL_TIME_SEGMENT_META.get(tech, {}).items():
                        if meta.get("time_str") == tstr:
                            e = meta.get("end")
                            if e is not None and pd.notna(e):
                                end_candidates.append(pd.to_datetime(e))
                end_dt = max(end_candidates) if end_candidates else (start_dt + pd.Timedelta(minutes=15))

            seg_list.append({
                "time_str": tstr,
                "start": start_dt.to_pydatetime() if start_dt is not None else None,
                "end": end_dt.to_pydatetime() if end_dt is not None else None,
                "cell_count": int(d.get("cell_count", 0)),
                "vendor_breakdown": d.get("vendor_breakdown", {})
            })

        # 按 start 排序（无 start 的放最后）
        seg_list.sort(key=lambda x: (x["start"] is None, x["start"] or datetime.datetime.min))
        final[act] = seg_list

    return final

def set_selected_time_segments(selected):
    """
    用户选择后调用（供 GUI 使用）：
      selected = {"活动A": "15:00~15:15", "活动B": "14:00~14:15"}
    未指定的活动仍使用默认“每小区最新时间段”逻辑。
    """
    global _SELECTED_TIME_SEGMENTS
    try:
        if not selected:
            _SELECTED_TIME_SEGMENTS = {}
            log_print("时间段选择已清空：全部活动将使用默认“每小区最新时间段”模式", "INFO")
            return

        if not isinstance(selected, dict):
            log_print("set_selected_time_segments 参数必须为 dict[str,str]，已忽略。", "WARN")
            return

        cleaned = {}
        for k, v in selected.items():
            kk = str(k).strip()
            vv = str(v).strip()
            if kk and vv:
                cleaned[kk] = vv
        _SELECTED_TIME_SEGMENTS = cleaned
        log_print(f"已设置活动时间段选择：{len(cleaned)} 个活动将使用用户指定时间段", "INFO")
    except Exception:
        _SELECTED_TIME_SEGMENTS = {}

def _apply_selected_time_segments_to_raw(raw_default_df: pd.DataFrame, df_list: pd.DataFrame, tech_type: str) -> pd.DataFrame:
    """
    若用户通过 set_selected_time_segments 为某些活动指定了时间段：
      - 对指定活动：仅使用该时间段下的数据（统一时间段，可能导致部分小区无数据）
      - 对未指定活动：仍使用默认“每小区最新时间段”
    说明：该函数只调整 waterfall_merge 的输入 raw_df，不改动匹配/计算口径。
    """
    if raw_default_df is None or isinstance(raw_default_df, list):
        return raw_default_df

    if not _SELECTED_TIME_SEGMENTS:
        # 未指定：直接使用默认 df（每小区最新）
        return raw_default_df

    raw_full = _ALL_RAW_FULL_WITH_SEG.get(tech_type)
    if raw_full is None or raw_full.empty:
        return raw_default_df

    try:
        df_list2 = _ensure_list_keys_for_time_select(df_list)
        df_full = raw_full.copy()

        # 确保有 _ts
        if '_ts' not in df_full.columns and 'STD__time_start' in df_full.columns:
            df_full['_ts'] = pd.to_datetime(df_full['STD__time_start'], errors='coerce')

        df_full = generate_raw_keys(df_full, tech_type)

        id_part = df_full['raw_key_id'].astype(str) if 'raw_key_id' in df_full.columns else ""
        name_part = df_full['raw_key_name'].astype(str) if 'raw_key_name' in df_full.columns else ""
        df_full['_cell_key'] = (id_part.fillna("").str.strip() + "||" + name_part.fillna("").str.strip()).astype(str)

        # default df 也做同样 cell_key（用于剔除被选择活动覆盖的小区）
        df_def = raw_default_df.copy()
        if '_ts' not in df_def.columns and 'STD__time_start' in df_def.columns:
            df_def['_ts'] = pd.to_datetime(df_def['STD__time_start'], errors='coerce')
        df_def = generate_raw_keys(df_def, tech_type)
        id_part2 = df_def['raw_key_id'].astype(str) if 'raw_key_id' in df_def.columns else ""
        name_part2 = df_def['raw_key_name'].astype(str) if 'raw_key_name' in df_def.columns else ""
        df_def['_cell_key'] = (id_part2.fillna("").str.strip() + "||" + name_part2.fillna("").str.strip()).astype(str)

        keep_mask = pd.Series(True, index=df_def.index)
        selected_parts = []

        for act, tstr in (_SELECTED_TIME_SEGMENTS or {}).items():
            act = str(act).strip()
            tstr = str(tstr).strip()
            if not act or not tstr:
                continue

            list_sub = df_list2[df_list2['活动名称'].astype(str) == act]
            if list_sub.empty:
                continue

            id_set = set(list_sub.get('list_key_id', pd.Series(dtype=str)).astype(str).fillna("").str.strip().tolist())
            name_set = set(list_sub.get('list_key_name', pd.Series(dtype=str)).astype(str).fillna("").str.strip().tolist())

            # 识别该活动涉及的小区（跨所有时间段）
            match_full = pd.Series(False, index=df_full.index)
            if 'raw_key_id' in df_full.columns:
                match_full |= df_full['raw_key_id'].astype(str).isin(id_set)
            if 'raw_key_name' in df_full.columns:
                match_full |= df_full['raw_key_name'].astype(str).isin(name_set)

            affected_cells = set(df_full.loc[match_full, '_cell_key'].astype(str).tolist())
            if affected_cells:
                keep_mask &= ~df_def['_cell_key'].astype(str).isin(affected_cells)

            # 解析用户选择的时间段对应的 ts_start（取该 time_str 在本制式中的最新 ts）
            ts_list = _TIME_STR_TO_TS.get(tech_type, {}).get(tstr, []) or []
            ts_list = [pd.to_datetime(x, errors='coerce') for x in ts_list if x is not None]
            ts_list = [x for x in ts_list if pd.notna(x)]
            chosen_ts = max(ts_list) if ts_list else None

            # 若映射表中不存在该 time_str，尝试从数据中反推
            if chosen_ts is None and 'SEG__time_str' in df_full.columns and 'SEG__ts_start' in df_full.columns:
                cand = df_full.loc[df_full['SEG__time_str'].astype(str) == tstr, 'SEG__ts_start']
                cand = pd.to_datetime(cand, errors='coerce').dropna().unique()
                if len(cand) > 0:
                    chosen_ts = max(cand)

            # 选择该时间段下的原始记录
            if chosen_ts is not None and 'SEG__ts_start' in df_full.columns:
                seg_mask = (pd.to_datetime(df_full['SEG__ts_start'], errors='coerce') == pd.to_datetime(chosen_ts))
            else:
                seg_mask = pd.Series(False, index=df_full.index)

            sel_rows = df_full.loc[match_full & seg_mask].copy()

            if sel_rows is not None and not sel_rows.empty:
                selected_parts.append(sel_rows)
                log_print(f"活动[{act}] {tech_type} 使用用户指定时间段: {tstr}", "INFO")
            else:
                log_print(f"活动[{act}] {tech_type} 指定时间段 {tstr} 下无数据（该活动此制式将按指定时间段入统，可能无输出）", "WARN")

        out = df_def.loc[keep_mask].copy()
        if selected_parts:
            out = pd.concat([out] + selected_parts, ignore_index=True)

        # 清理内部辅助列，避免影响既有导出
        out = _strip_internal_time_cols(out)
        for c in ['raw_key_id', 'raw_key_name', 'raw_key_fuzzy_base', '_cell_key']:
            if c in out.columns:
                try:
                    del out[c]
                except Exception:
                    pass

        return out

    except Exception:
        return raw_default_df