# -*- coding: utf-8 -*-
from __future__ import annotations

import difflib
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from kpi_tool.config import constants as C
from kpi_tool.config.constants import FUZZY_THRESHOLD_4G, FUZZY_THRESHOLD_5G, SUBSTRING_MATCH_MIN_RATIO
from kpi_tool.config.logging_config import log_print
from kpi_tool.core.key_generator import generate_list_keys, generate_raw_keys
from kpi_tool.core.standardizer import standardizer, auto_match_column_safe, KeyNormalizer

# Backwards-compatibility: some modules reference a module-level 'args' populated by main
try:
    from main import args  # type: ignore
except Exception:
    args = None

# 快速模糊匹配：优先 rapidfuzz（不存在时回退 difflib）
try:
    from rapidfuzz import process, fuzz  # type: ignore
    HAS_RAPIDFUZZ = True
except Exception:
    process = None  # type: ignore
    fuzz = None     # type: ignore
    HAS_RAPIDFUZZ = False

def _core_fuzzy_match(raw_df, list_df, tech_type):
    if raw_df.empty or list_df.empty: return None
    threshold = FUZZY_THRESHOLD_5G if tech_type == '5G' else FUZZY_THRESHOLD_4G
    queries = raw_df['raw_key_fuzzy_base'].unique().tolist()
    choices = list_df['list_key_name'].unique().tolist()
    choice_map = list_df.set_index('list_key_name')
    q_to_c_map = {}
    for q in queries:
        if not q or len(q) < 2: continue
        best_match = None; best_score = 0
        
        # 优化的子串匹配：要求双向子串且长度>=4，避免短词误匹配（如"宁都"匹配"宁都小布"）
        # 增强：要求子串长度至少占较短字符串的70%，避免仅凭通用后缀误匹配
        for c in choices:
            if len(q) >= 4:
                min_len = min(len(q), len(c))
                # 双向子串检查
                if q in c:
                    # q是c的子串，检查q的长度是否足够（至少占c的70%）
                    if len(q) >= min_len * SUBSTRING_MATCH_MIN_RATIO:
                        best_match = c
                        best_score = 100
                        break
                elif c in q:
                    # c是q的子串，检查c的长度是否足够（至少占q的70%）
                    if len(c) >= min_len * SUBSTRING_MATCH_MIN_RATIO:
                        best_match = c
                        best_score = 100
                        break
        
        if best_score < 100:
             if HAS_RAPIDFUZZ:
                 res = process.extractOne(q, choices, scorer=fuzz.ratio, score_cutoff=threshold*100)
                 if res and res[1] > best_score: best_match, best_score = res[0], res[1]
             else:
                 for c in choices:
                     s = difflib.SequenceMatcher(None, q, c).ratio() * 100
                     if s >= threshold*100 and s > best_score: best_score, best_match = s, c
        if best_match and best_score >= threshold*100:
            q_to_c_map[q] = (best_match, best_score)
    if not q_to_c_map: return None
    map_data = []
    for q, (choice, score) in q_to_c_map.items():
        if choice in choice_map.index:
            row_dict = choice_map.loc[choice].iloc[0].to_dict() if isinstance(choice_map.loc[choice], pd.DataFrame) else choice_map.loc[choice].to_dict()
            map_data.append({'_query_key': q, '_fuzzy_score': score, **row_dict})
    df_map = pd.DataFrame(map_data)
    merged = pd.merge(raw_df, df_map, left_on='raw_key_fuzzy_base', right_on='_query_key', how='inner')
    merged['match_method'] = 'FUZZY_STRIP_PREFIX'
    return merged

def waterfall_merge(df_raw, df_list, tech_type):
    '''
    三段匹配（名称 > ID > 模糊），但按“清单小区维度”严格去重：
    - 先 EXACT_NAME；
    - 若某活动 EXACT_NAME 已覆盖该活动清单全部 list_key_name，则该活动仅保留 EXACT_NAME 结果（跳过 ID/模糊，避免重复统计）；
    - 若未覆盖，则只对“剩余未匹配清单小区”补充 EXACT_ID 和 FUZZY_STRIP_PREFIX。
    '''
    if df_raw is None or df_raw.empty or df_list is None or df_list.empty:
        return pd.DataFrame()

    # 仅用同制式清单参与匹配（避免 4G/5G 同名导致串匹配、计数偏大）
    freq_col = auto_match_column_safe(df_list.columns, ['频段', '制式', '网络制式'])
    if freq_col:
        s = df_list[freq_col].astype(str).str.upper()
        if tech_type == '4G':
            df_list = df_list[s.str.contains('4G') | s.str.contains('LTE')].copy()
        else:
            df_list = df_list[s.str.contains('5G') | s.str.contains('NR')].copy()
        if df_list.empty:
            return pd.DataFrame()

    df_raw = df_raw.copy()
    df_raw['_tracker_idx'] = np.arange(len(df_raw))
    df_raw = generate_raw_keys(df_raw, tech_type)

    # 保底生成清单 Key
    if 'list_key_name' not in df_list.columns:
        name_col = auto_match_column_safe(df_list.columns, ['小区中文名', '小区名称', 'CELL_NAME', '小区名'])
        df_list = df_list.copy()
        df_list['list_key_name'] = df_list[name_col].astype(str).str.strip() if name_col else ""

    if 'list_key_id' not in df_list.columns:
        id_col = auto_match_column_safe(df_list.columns, ['ECGI', 'CGI', 'ENB_CELL', '小区ID', 'NCI'])
        df_list = df_list.copy()
        df_list['list_key_id'] = KeyNormalizer.normalize_id(df_list[id_col]) if id_col else ""

    # 清单缺列兜底
    if '活动名称' not in df_list.columns:
        df_list = df_list.copy()
        df_list['活动名称'] = "未知活动"
    if '_list_idx' not in df_list.columns:
        df_list = df_list.copy()
        df_list['_list_idx'] = np.arange(len(df_list), dtype=int)

    # ========= 智能名称修正：当 list_key_name 在网管中找不到但 list_key_id 能找到时，自动用网管名称替换 =========
    # 目的：解决保障清单 B 列（小区中文名）人工填写不规范导致匹配失败的问题
    id_to_name_map = df_raw.dropna(subset=['raw_key_id']).drop_duplicates(subset=['raw_key_id']).set_index('raw_key_id')['raw_key_name'].to_dict()
    raw_names_set = set(df_raw['raw_key_name'].dropna().unique())

    _name_fix_log = []  # 记录名称修正日志

    # 向量化名称修正（替代 apply）
    df_list = df_list.copy()
    lid_series = df_list['list_key_id'].astype(str).str.strip()
    lname_series = df_list['list_key_name'].astype(str).str.strip()
    name_in_raw = lname_series.isin(raw_names_set)
    id_in_map = lid_series.isin(id_to_name_map)
    need_fix = lname_series.ne('') & (~name_in_raw) & id_in_map
    new_names = lid_series.map(id_to_name_map)
    if need_fix.any():
        fix_rows = df_list[need_fix]
        for _, r in fix_rows.iterrows():
            lid = str(r.get('list_key_id', '')).strip()
            old = str(r.get('list_key_name', '')).strip()
            _name_fix_log.append((lid, old, id_to_name_map.get(lid, '')))
        df_list['list_key_name'] = np.where(need_fix, new_names, lname_series)

    # 打印名称修正日志
    if _name_fix_log:
        log_print(f"[智能名称修正] {tech_type} 共修正 {len(_name_fix_log)} 个小区名称：", "INFO")
        for lid, old_name, new_name in _name_fix_log[:5]:  # 只打印前5条
            log_print(f"  ID={lid}: '{old_name}' → '{new_name}'", "INFO")
        if len(_name_fix_log) > 5:
            log_print(f"  ... 还有 {len(_name_fix_log) - 5} 条修正记录", "INFO")


    if getattr(args, 'only_activity', None):
        df_list = df_list[df_list['活动名称'].astype(str) == str(getattr(args, 'only_activity', None))].copy()
        if df_list.empty:
            log_print(f"❌ 指定活动未在保障清单中找到: {getattr(args, 'only_activity', None)}", "WARN")
            return

    matched_dfs = []
    activities = df_list['活动名称'].astype(str).fillna("未知活动").unique().tolist()

    for act in activities:
        list_sub = df_list[df_list['活动名称'].astype(str) == act].copy()
        if list_sub.empty:
            continue

        target_names = set(list_sub['list_key_name'].astype(str).fillna("").tolist())
        target_names.discard("")
        can_cover_check = len(target_names) > 0

        # 1) EXACT_NAME
        m1 = pd.DataFrame()
        if 'raw_key_name' in df_raw.columns:
            subset = df_raw[df_raw['raw_key_name'] != ""]
            if not subset.empty:
                m1 = pd.merge(subset, list_sub, left_on='raw_key_name', right_on='list_key_name', how='inner')
                if not m1.empty:
                    m1['match_method'] = 'EXACT_NAME'

        matched_names = set()
        if not m1.empty and 'list_key_name' in m1.columns:
            matched_names = set(m1['list_key_name'].astype(str).fillna("").tolist())
            matched_names.discard("")

        # 覆盖判定：以“清单行（_list_idx）覆盖”为准，避免同名/重复名导致误判
        if (not m1.empty) and ('_list_idx' in m1.columns) and (m1['_list_idx'].nunique() == list_sub['_list_idx'].nunique()):
            matched_dfs.append(m1)
            continue

        used_list_idx = set()
        if not m1.empty and '_list_idx' in m1.columns:
            used_list_idx = set(pd.to_numeric(m1['_list_idx'], errors='coerce').dropna().astype(int).tolist())

        list_rem = list_sub[~list_sub['_list_idx'].isin(used_list_idx)].copy()

        # 2) EXACT_ID（只对剩余清单）
        m2 = pd.DataFrame()
        if not list_rem.empty and 'raw_key_id' in df_raw.columns and 'list_key_id' in list_rem.columns:
            subset = df_raw[df_raw['raw_key_id'] != ""]
            if not subset.empty:
                m2 = pd.merge(subset, list_rem, left_on='raw_key_id', right_on='list_key_id', how='inner')
                if not m2.empty:
                    m2['match_method'] = 'EXACT_ID'
                    used2 = set(pd.to_numeric(m2['_list_idx'], errors='coerce').dropna().astype(int).tolist())
                    used_list_idx |= used2

        list_rem2 = list_sub[~list_sub['_list_idx'].isin(used_list_idx)].copy()

        # 3) FUZZY（只对剩余清单）
        m3 = None
        if not list_rem2.empty and 'raw_key_fuzzy_base' in df_raw.columns:
            subset = df_raw[df_raw['raw_key_fuzzy_base'] != ""]
            if not subset.empty:
                m3 = _core_fuzzy_match(subset, list_rem2, tech_type)

        if not m1.empty:
            matched_dfs.append(m1)
        if not m2.empty:
            matched_dfs.append(m2)
        if m3 is not None and not m3.empty:
            matched_dfs.append(m3)

    if not matched_dfs:
        return pd.DataFrame()

    final_df = pd.concat(matched_dfs, ignore_index=True)

    # ========= 一对多去重：确保每个清单行只对应唯一的网管数据 =========
    # 防止同一清单小区匹配到多个网管小区（不同频点/扇区）导致统计翻倍
    if '_list_idx' in final_df.columns and not final_df.empty:
        before_cnt = len(final_df)
        
        # 排序：按流量和用户数降序（确保保留业务量最大的记录）
        sort_cols = []
        if 'STD__kpi_traffic_gb' in final_df.columns:
            final_df['_sort_traffic'] = pd.to_numeric(final_df['STD__kpi_traffic_gb'], errors='coerce').fillna(0)
            sort_cols.append('_sort_traffic')
        if 'STD__kpi_rrc_users_max' in final_df.columns:
            final_df['_sort_users'] = pd.to_numeric(final_df['STD__kpi_rrc_users_max'], errors='coerce').fillna(0)
            sort_cols.append('_sort_users')
        
        if sort_cols:
            final_df = final_df.sort_values(sort_cols, ascending=False)
        
        # 按 _list_idx 去重，保留第一条（即排序后业务量最大的）
        final_df = final_df.drop_duplicates(subset=['_list_idx'], keep='first')
        
        # 清理临时排序列
        for c in ['_sort_traffic', '_sort_users']:
            if c in final_df.columns:
                del final_df[c]
        
        after_cnt = len(final_df)
        if before_cnt > after_cnt:
            log_print(f"[一对多去重] {tech_type} 去重前 {before_cnt} 行 → 去重后 {after_cnt} 行（移除 {before_cnt - after_cnt} 条重复匹配）", "INFO")

    # ========= 多对一去重：确保每个网管小区在每个活动中只出现一次 =========
    # 防止保障清单中同一小区被重复录入（不同 _list_idx）导致统计翻倍
    if '活动名称' in final_df.columns and 'raw_key_id' in final_df.columns and not final_df.empty:
        before_cnt2 = len(final_df)
        # 匹配方法优先级：EXACT_NAME(最优) > EXACT_ID > FUZZY_STRIP_PREFIX
        method_priority = {'EXACT_NAME': 0, 'EXACT_ID': 1, 'FUZZY_STRIP_PREFIX': 2}
        if 'match_method' in final_df.columns:
            final_df['_match_pri'] = final_df['match_method'].map(method_priority).fillna(9)
            final_df = final_df.sort_values('_match_pri', ascending=True)
        # 只对 raw_key_id 非空的行去重（避免空 ID 被误合并）
        mask_has_id = final_df['raw_key_id'].astype(str).str.strip() != ''
        df_has_id = final_df[mask_has_id].drop_duplicates(subset=['活动名称', 'raw_key_id'], keep='first')
        df_no_id = final_df[~mask_has_id]
        final_df = pd.concat([df_has_id, df_no_id], ignore_index=True)
        if '_match_pri' in final_df.columns:
            del final_df['_match_pri']
        after_cnt2 = len(final_df)
        if before_cnt2 > after_cnt2:
            log_print(f"[多对一去重] {tech_type} 去重前 {before_cnt2} 行 → 去重后 {after_cnt2} 行（移除 {before_cnt2 - after_cnt2} 条清单重复录入）", "INFO")

    # 仅删除内部辅助列（保留 list_key_id/list_key_name/raw_key_name/raw_key_id/_list_idx 便于对账输出）
    # 注意：raw_key_id 必须保留，用于 calculator.py 中的 _safe_get_cell_name 函数通过 ID 查找小区名称
    drop_cols = ['_tracker_idx', 'raw_key_fuzzy_base', '_query_key', '_fuzzy_score']
    for c in drop_cols:
        if c in final_df.columns:
            del final_df[c]
    return final_df
