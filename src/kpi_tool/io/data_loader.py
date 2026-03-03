# -*- coding: utf-8 -*-
from __future__ import annotations
from concurrent.futures import ThreadPoolExecutor, as_completed

import glob
import os
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd

from kpi_tool.config.logging_config import log_print
from kpi_tool.core.key_generator import generate_raw_keys
from kpi_tool.core.standardizer import auto_match_column_safe, standardizer

def get_activity_list(list_path: str):
    """供 GUI 调用：从保障小区清单中读取“活动名称”列表。

    说明：
    - 自动识别活动列（表头包含“活动”关键词优先）。
    - 去空、去重（保留首次出现顺序）。
    - 仅做读取与识别，不影响任何业务口径/计算逻辑。
    """
    if not list_path or (not os.path.exists(list_path)):
        return []

    # 优先读取第 1 个工作表；失败则读取全部并取首个非空表
    df = None
    try:
        df = pd.read_excel(list_path, sheet_name=0, dtype=str)
    except Exception:
        try:
            sheets = pd.read_excel(list_path, sheet_name=None, dtype=str)
            for _name, _df in sheets.items():
                if _df is not None and len(_df) > 0:
                    df = _df
                    break
        except Exception:
            df = None

    if df is None or df.empty:
        return []

    # 识别“活动名称”列
    act_col = auto_match_column_safe(
        df.columns,
        keywords=["活动名称", "保障活动", "活动", "活动名", "活动名称(必填)", "活动名称（必填）"],
        exclusions=["活动类型", "活动级别"]
    )
    if act_col is None:
        # 兜底：表头中包含“活动”的列
        for c in df.columns:
            if isinstance(c, str) and ("活动" in c):
                act_col = c
                break

    if act_col is None:
        raise ValueError("未识别到活动名称列（请检查保障清单表头是否包含‘活动名称’等字段）")

    s = df[act_col].fillna("").astype(str).str.strip()
    acts = []
    seen = set()
    for a in s.tolist():
        if not a:
            continue
        if a not in seen:
            seen.add(a)
            acts.append(a)
    return acts

def load_data_frame(file_path):
    try:
        if str(file_path).lower().endswith('.csv'): return pd.read_csv(file_path, encoding='gbk', on_bad_lines='skip')
        try: return pd.read_excel(file_path)
        except Exception: return pd.read_csv(file_path, encoding='gbk', on_bad_lines='skip')
    except Exception:
        try: return pd.read_csv(file_path, encoding='utf-8', on_bad_lines='skip')
        except Exception: return None

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

def load_and_distribute(directory, tech_type):
    import kpi_tool.core.time_handler as TH
    """
    兼容旧接口（返回3元组不变）：
      - 返回默认用于统计的 raw_df（已升级为“每小区最新时间段”模式）
      - prev_df 仍保留（全局上一时间段，若存在）
      - time_window 返回全局最新起始时间（用于展示）

    同时：内部会缓存“全部时间段”数据，供 GUI 时间段查看/选择使用。

    额外增强（不改变任何统计/匹配/输出口径，仅提升可观测性与稳健性）：
      - 目录为空 / 文件均读取失败时，打印明确 WARN，便于快速定位“0 匹配”的根因。
    """
    if not os.path.exists(directory):
        try:
            log_print(f"{tech_type} 网管指标目录不存在: {directory}", "WARN")
        except Exception:
            pass
        return None, None, []

    # 仅扫描当前目录（保持旧行为），不递归子目录
    try:
        all_entries = [os.path.join(directory, x) for x in os.listdir(directory)]
        sub_dirs = [x for x in all_entries if os.path.isdir(x)]
        if sub_dirs:
            # 提示但不改变行为
            try:
                log_print(f"{tech_type} 网管指标目录包含子目录（工具默认不递归读取）：{directory}", "WARN")
            except Exception:
                pass
    except Exception:
        pass

    files = glob.glob(os.path.join(directory, "*.*"))
    # 过滤掉目录/临时文件
    try:
        files = [f for f in files if os.path.isfile(f) and not os.path.basename(f).startswith('~$')]
    except Exception:
        pass

    if not files:
        try:
            log_print(f"{tech_type} 网管指标目录为空（未发现可读取文件）：{directory}", "WARN")
        except Exception:
            pass
        return None, None, []

    dfs = []
    errors = []  # (file, err)

    def _worker(f):
        try:
            df = load_data_frame(f)
            if df is not None and not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                df['_source_file'] = os.path.basename(f)
                df = standardizer.standardize_df(df, tech_type)
                return df, None, f
            return None, None, f
        except Exception as e:
            return None, f"{type(e).__name__}: {e}", f

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(_worker, f) for f in files]
        for future in as_completed(futures):
            try:
                df, err, fp = future.result()
                if df is not None and not df.empty:
                    dfs.append(df)
                else:
                    if err:
                        errors.append((os.path.basename(fp), err))
            except Exception as e:
                # 理论上极少出现：future.result() 自身异常
                try:
                    errors.append(("future.result()", f"{type(e).__name__}: {e}"))
                except Exception:
                    pass

    if not dfs:
        # 目录有文件但全部未能读入：给出可操作提示（仅日志，不改变行为）
        try:
            exts = {}
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                exts[ext] = exts.get(ext, 0) + 1
            ext_str = ", ".join([f"{k}:{v}" for k, v in sorted(exts.items(), key=lambda x: (-x[1], x[0]))])
            log_print(f"{tech_type} 未加载到任何可用数据（文件数 {len(files)} | 扩展名分布 {ext_str}）。本次将导致匹配为 0。", "WARN")
            # 输出前几条错误，避免刷屏
            if errors:
                for fn, msg in errors[:5]:
                    log_print(f"{tech_type} 读取失败示例: {fn} | {msg}", "WARN")
            else:
                log_print(f"{tech_type} 所有文件读取结果为空（可能是文件为空/表头不识别/格式不支持）。", "WARN")
            log_print(f"{tech_type} 请确认：1) 文件是否放在该目录根下（不在子目录）；2) 文件是否为 .xlsx/.xls/.xlsm；3) 文件是否可正常打开。", "WARN")
        except Exception:
            pass
        return None, None, []

    full = pd.concat(dfs, ignore_index=True)

    # ---- 时间段识别：按 STD__time_start 分段（与旧逻辑一致）----
    t_col = 'STD__time_start'
    if t_col in full.columns:
        full['_ts'] = pd.to_datetime(full[t_col], errors='coerce')
        times = sorted(full['_ts'].dropna().unique())
        if not times:
            # 有 time 列但解析失败：仍返回全量，且缓存为“无时间信息”
            full['SEG__ts_start'] = pd.NaT
            full['SEG__time_str'] = "无时间信息"
            TH._ALL_RAW_FULL_WITH_SEG[tech_type] = full
            TH._ALL_TIME_SEGMENT_DFS[tech_type] = {"无时间信息": full}
            TH._ALL_TIME_SEGMENT_META[tech_type] = {"无时间信息": {"time_str": "无时间信息", "start": None, "end": None}}
            TH._TIME_STR_TO_TS[tech_type] = {"无时间信息": [None]}
            try:
                log_print(f"{tech_type} 时间字段存在但解析失败（STD__time_start 全部无效）。默认按全量入统，时间窗显示为“未知时间”。", "WARN")
            except Exception:
                pass
            return _strip_internal_time_cols(full.copy()), None, []

        # 构造每行的展示时间段（优先用 STD__time_end，缺失则默认 +15min）
        if 'STD__time_end' in full.columns:
            full['SEG__time_str'] = [
                _time_str_from_ts(s, e) for s, e in zip(full['_ts'], full['STD__time_end'])
            ]
        else:
            full['SEG__time_str'] = [_time_str_from_ts(s, None) for s in full['_ts']]
        full['SEG__ts_start'] = full['_ts']

        # 缓存：全量+分段 dict
        seg_dfs = {}
        seg_meta = {}
        ts_map = {}
        for ts in times:
            seg = full[full['_ts'] == ts].copy()
            # 该段的 end：尽量从 STD__time_end 推断
            end_ts = None
            if 'STD__time_end' in seg.columns:
                cand = pd.to_datetime(seg['STD__time_end'], errors='coerce').dropna()
                if not cand.empty:
                    end_ts = cand.max()
            time_str = _time_str_from_ts(ts, end_ts)
            seg_dfs[ts] = seg
            seg_meta[ts] = {
                "time_str": time_str,
                "start": pd.to_datetime(ts),
                "end": pd.to_datetime(end_ts) if end_ts is not None else pd.to_datetime(ts) + pd.Timedelta(minutes=15)
            }
            ts_map.setdefault(time_str, []).append(pd.to_datetime(ts))

        TH._ALL_RAW_FULL_WITH_SEG[tech_type] = full
        TH._ALL_TIME_SEGMENT_DFS[tech_type] = seg_dfs
        TH._ALL_TIME_SEGMENT_META[tech_type] = seg_meta
        TH._TIME_STR_TO_TS[tech_type] = ts_map

        # 默认 df：每小区最新时间段
        df_latest = _select_latest_per_cell(full, tech_type)
        df_latest = _strip_internal_time_cols(df_latest)

        prev = None
        if len(times) >= 2:
            prev = _strip_internal_time_cols(seg_dfs[times[-2]].copy())

        try:
            log_print(f"{tech_type} 已加载 {len(files)} 份文件 | 时间段 {len(times)} 段 | 默认：每小区最新时间段", "INFO")
        except Exception:
            pass

        return df_latest, prev, times[-1]

    # 无时间字段：仍缓存为单段
    full['_ts'] = pd.NaT
    full['SEG__ts_start'] = pd.NaT
    full['SEG__time_str'] = "无时间信息"
    TH._ALL_RAW_FULL_WITH_SEG[tech_type] = full
    TH._ALL_TIME_SEGMENT_DFS[tech_type] = {"无时间信息": full}
    TH._ALL_TIME_SEGMENT_META[tech_type] = {"无时间信息": {"time_str": "无时间信息", "start": None, "end": None}}
    TH._TIME_STR_TO_TS[tech_type] = {"无时间信息": [None]}
    try:
        log_print(f"{tech_type} 已加载 {len(files)} 份文件 | 未识别到时间字段 | 默认：全量入统", "WARN")
    except Exception:
        pass
    return _strip_internal_time_cols(full.copy()), None, None


