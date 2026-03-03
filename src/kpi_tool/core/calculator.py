# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from kpi_tool.config import project_config as PC
from kpi_tool.config.project_config import get_output_cols, get_percent_rows
from kpi_tool.config.logging_config import log_print
from kpi_tool.core.standardizer import standardizer
from kpi_tool.utils.helpers import format_time_range


def _vec_pick_val(df, candidates):
    """向量化版 _pick_val：从多个候选列中取第一个非空值"""
    result = pd.Series('', index=df.index, dtype=str)
    for col in reversed(candidates):
        if col in df.columns:
            s = df[col].astype(str).str.strip()
            valid = s.ne('') & s.str.lower().ne('nan') & s.str.lower().ne('none')
            result = result.where(~valid, s)
    return result


def _poor_quality_fallback_vectorized(df, tech, _prefer_raw_drop_func):
    """向量化的质差判定回退逻辑（无配置门限时使用）"""
    if df is None or df.empty:
        return pd.DataFrame()

    limit_util = 90 if tech == '5G' else 85

    # 向量化提取数值列
    connect = pd.to_numeric(df.get('STD__kpi_connect', pd.Series(0, index=df.index)), errors='coerce').fillna(0)
    drop_val = pd.to_numeric(df.get('STD__kpi_drop', pd.Series(0, index=df.index)), errors='coerce').fillna(0)
    interf = pd.to_numeric(df.get('STD__kpi_ul_interf_dbm', pd.Series(-120, index=df.index)), errors='coerce').fillna(-120)
    util = pd.to_numeric(df.get('KPI_UTIL', pd.Series(0, index=df.index)), errors='coerce').fillna(0)
    users = pd.to_numeric(df.get('STD__kpi_rrc_users_max', pd.Series(0, index=df.index)), errors='coerce').fillna(0)

    # 布尔掩码
    m_low_connect = (connect > 0) & (connect <= 90)
    m_high_drop = drop_val >= 3
    m_high_interf = (interf >= -105) & (interf != 0)
    m_high_load = (util >= limit_util) & (users >= 100)

    any_poor = m_low_connect | m_high_drop | m_high_interf | m_high_load

    if tech == '5G':
        vonr_connect = pd.to_numeric(df.get('STD__kpi_vonr_connect', pd.Series(0, index=df.index)), errors='coerce').fillna(0)
        vonr_drop = pd.to_numeric(df.get('STD__kpi_vonr_drop_5qi1', pd.Series(0, index=df.index)), errors='coerce').fillna(0)
        m_vonr_low = (vonr_connect > 0) & (vonr_connect <= 90)
        m_vonr_high_drop = vonr_drop >= 3
        any_poor = any_poor | m_vonr_low | m_vonr_high_drop

    if not any_poor.any():
        return pd.DataFrame()

    # 只处理质差行
    idx = df.index[any_poor]
    n = len(idx)

    # 构建 reasons 和 details（用列表拼接，比逐行 iterrows 快得多）
    reason_list = [''] * n
    detail_list = [''] * n

    m_lc = m_low_connect.reindex(idx).values
    m_hd = m_high_drop.reindex(idx).values
    m_hi = m_high_interf.reindex(idx).values
    m_hl = m_high_load.reindex(idx).values
    c_vals = connect.reindex(idx).values
    d_vals = drop_val.reindex(idx).values
    i_vals = interf.reindex(idx).values
    u_vals = util.reindex(idx).values
    usr_vals = users.reindex(idx).values

    # 掉线率展示值：优先 STD，回退 RAW
    v_std = pd.to_numeric(df.get('STD__kpi_drop', np.nan), errors='coerce').reindex(idx).values
    v_raw = pd.to_numeric(df.get('RAW__kpi_drop', np.nan), errors='coerce').reindex(idx).values
    drop_show = np.where(np.isfinite(v_std) & ~np.isnan(v_std), v_std, v_raw)

    if tech == '5G':
        m_vl = m_vonr_low.reindex(idx).values
        m_vd = m_vonr_high_drop.reindex(idx).values

    for i in range(n):
        r = []
        d = []
        if m_lc[i]:
            r.append('低接通')
            d.append(f'接通率:{c_vals[i]:.2f}%')
        if m_hd[i]:
            r.append('高掉线')
            ds = drop_show[i]
            if np.isfinite(ds):
                d.append(f'掉线率:{ds:.2f}%')
        if m_hi[i]:
            r.append('高干扰')
            d.append(f'干扰:{i_vals[i]:.1f}dBm')
        if m_hl[i]:
            r.append('高负荷')
            d.append(f'利用率:{u_vals[i]:.1f}%,用户:{int(usr_vals[i])}')
        if tech == '5G':
            if m_vl[i]:
                r.append('VoNR低接通')
            if m_vd[i]:
                r.append('VoNR高掉线')
        reason_list[i] = ';'.join(r)
        detail_list[i] = ' '.join(d)

    poor_df = df.loc[idx]

    # 向量化提取小区名称和ID
    name_cols = ['小区名称', '小区中文名', 'CELL_NAME', 'cell_name', 'STD__cell_name', 'list_key_name']
    id_cols = ['CGI/ECGI', 'ECGI', 'CGI', 'ENB_CELL', 'NCI', 'list_key_id', 'join_key']

    result_df = pd.DataFrame({
        '活动名称': poor_df.get('活动名称', '').values if '活动名称' in poor_df.columns else '',
        '区域': poor_df.get('区域', '').values if '区域' in poor_df.columns else '',
        '小区名称': _vec_pick_val(poor_df, name_cols).values,
        'CGI/ECGI': _vec_pick_val(poor_df, id_cols).values,
        '制式': tech,
        '厂家': poor_df.get('厂家', '').values if '厂家' in poor_df.columns else '',
        '质差类型': reason_list,
        '备注': detail_list,
    })

    if not result_df.empty:
        result_df = result_df.drop_duplicates(subset=['活动名称', '小区名称', 'CGI/ECGI', '质差类型'], keep='first')
    return result_df

def get_poor_quality(merged_df: pd.DataFrame, tech_type: str) -> pd.DataFrame:
    return _get_poor_quality_impl(merged_df, tech_type)


def _get_poor_quality_impl(df, tech_type):

    def _pick_val(row, keys):
        """逐个字段取第一个非空值（用于补全小区名称/ID等展示列）"""
        for k in keys:
            try:
                v = row.get(k, "")
            except Exception:
                v = ""
            if v is None:
                continue
            s = str(v).strip()
            if s == "" or s.lower() == "nan":
                continue
            return v
        return ""

    def _prefer_raw_drop(row):
        """质差明细展示用：优先取已缩放的 STD__kpi_drop（与门限判定一致），避免 RAW 未缩放导致显示值偏小"""
        v_std = pd.to_numeric(row.get("STD__kpi_drop", None), errors="coerce")
        if pd.notna(v_std):
            return float(v_std)
        try:
            v_raw = row.get("RAW__kpi_drop", None)
        except Exception:
            v_raw = None
        v_raw = pd.to_numeric(v_raw, errors="coerce")
        if pd.notna(v_raw):
            return float(v_raw)
        return np.nan

    """质差判定（可配置）。

    优先使用项目配置 Excel 的 Thresholds sheet（enabled=1 的规则）。
    - 支持 op: <, <=, >, >=, BETWEEN
    - 支持复合条件：and_kpi_id/and_op/and_th1/and_th2（例如高负荷：利用率>=阈值 且 用户数>=100）
    若未提供 Thresholds，则回退到脚本内置兜底规则（保持历史行为）。
    """
    if df is None or df.empty:
        return pd.DataFrame()
    rows = []

    tech = str(tech_type).upper()
    rules_by_tech = {}
    try:
        rules_by_tech = PC.PROJECT_CFG.get("threshold_rules", {}) if isinstance(PC.PROJECT_CFG, dict) else {}
    except Exception:
        rules_by_tech = {}

    def _to_num(x):
        try:
            if pd.isna(x):
                return None
            return float(x)
        except Exception:
            return None

    def _eval_op(v, op, th1, th2=None):
        if v is None:
            return False
        op = (op or "").upper()
        t1 = _to_num(th1)
        t2 = _to_num(th2)
        if t1 is None and op not in ("ISNULL", "NOTNULL"):
            return False
        try:
            if op in ("<", "LT"):
                return v < t1
            if op in ("<=", "LE"):
                return v <= t1
            if op in (">", "GT"):
                return v > t1
            if op in (">=", "GE"):
                return v >= t1
            if op in ("BETWEEN", "RANGE"):
                if t2 is None:
                    return False
                lo, hi = (t1, t2) if t1 <= t2 else (t2, t1)
                return lo <= v <= hi
            if op in ("!=", "NE"):
                return v != t1
            if op in ("=", "==", "EQ"):
                return v == t1
        except Exception:
            return False
        return False

    def _rule_to_col(rule_kpi_id, rule_std_field):
        # 优先使用 std_field；否则用 kpi_id 映射到 std_field
        sf = (rule_std_field or "").strip()
        kid = (rule_kpi_id or "").strip()
        if not sf and kid:
            try:
                sf = PC.PROJECT_CFG.get("kpi_id_to_std", {}).get(kid, "")
            except Exception:
                sf = ""
        if not sf and kid:
            # 兜底：按 STD__{kpi_id}
            sf = kid if kid.startswith(("STD__", "RAW__", "SRC__", "KPI_")) else f"STD__{kid}"
        return sf

    def _value_from_row(row, col):
        if not col:
            return None
        # 兼容 KPI_UTIL 等派生列
        if col in row:
            return _to_num(row.get(col))
        # 兼容未带前缀：默认取 STD__
        if ("STD__" + col) in row:
            return _to_num(row.get("STD__" + col))
        return None

    # 取规则：本制式 + BOTH
    tech_rules = []
    if isinstance(rules_by_tech, dict):
        tech_rules += list(rules_by_tech.get(tech, []) or [])
        tech_rules += list(rules_by_tech.get("BOTH", []) or [])

    # 如果没有配置门限，则回退向量化逻辑
    if not tech_rules:
        return _poor_quality_fallback_vectorized(df, tech, _prefer_raw_drop)

    # ===== 配置化门限判定 =====
    # 将 rule 中 and_kpi_id 映射到 std_field（一次性）
    try:
        kmap = PC.PROJECT_CFG.get("kpi_id_to_std", {})
    except Exception:
        kmap = {}
    for r in tech_rules:
        if r.get("and_kpi_id") and not r.get("and_std_field"):
            r["and_std_field"] = kmap.get(r["and_kpi_id"], "")

    for _, row in df.iterrows():
        reasons = []
        details = []
        for rule in sorted(tech_rules, key=lambda x: int(x.get("priority", 100))):
            col = _rule_to_col(rule.get("kpi_id", ""), rule.get("std_field", ""))
            v = _value_from_row(row, col)
            if v is None:
                continue
            ok = _eval_op(v, rule.get("op", ""), rule.get("th1", np.nan), rule.get("th2", np.nan))
            if not ok:
                continue
            # 复合条件（可选）
            if rule.get("and_kpi_id") or rule.get("and_std_field"):
                col2 = _rule_to_col(rule.get("and_kpi_id", ""), rule.get("and_std_field", ""))
                v2 = _value_from_row(row, col2)
                ok2 = _eval_op(v2, rule.get("and_op", ""), rule.get("and_th1", np.nan), rule.get("and_th2", np.nan))
                if not ok2:
                    continue

            pt = rule.get("poor_type", "") or "质差"
            if pt not in reasons:
                reasons.append(pt)
                # 详情展示
                dname = None
                try:
                    dname = PC.PROJECT_CFG.get("kpi_id_to_display", {}).get(rule.get("kpi_id", ""), None)
                except Exception:
                    dname = None
                if not dname:
                    dname = rule.get("kpi_id", "") or col
                # 百分比与数值统一保留两位
                details.append(f"{dname}:{v:.2f}")

        if reasons:
            cell_name = _pick_val(row, ['小区名称','小区中文名','CELL_NAME','cell_name','STD__cell_name','list_key_name'])
            cell_id = _pick_val(row, ['CGI/ECGI','ECGI','CGI','ENB_CELL','NCI','list_key_id','join_key'])
            rows.append({
                '活动名称': row.get('活动名称', ''),
                '区域': row.get('区域', ''),
                '小区名称': cell_name,
                'CGI/ECGI': cell_id,
                '制式': tech,
                '厂家': row.get('厂家', ''),
                '质差类型': ";".join(reasons),
                '备注': " ".join(details)
            })
    result_df = pd.DataFrame(rows)
    if not result_df.empty:
        result_df = result_df.drop_duplicates(subset=['活动名称', '小区名称', 'CGI/ECGI', '质差类型'], keep='first')
    return result_df

def _apply_catalog_generic_kpis(s: dict, grp: pd.DataFrame, tech_type: str):
    tech = str(tech_type).upper()
    rows = PC.PROJECT_CFG.get("kpi_by_tech", {}).get(tech, []) if isinstance(PC.PROJECT_CFG, dict) else []
    if not rows:
        return

    def _series(field):
        if field in grp.columns:
            return pd.to_numeric(grp[field], errors='coerce')
        return None

    for r in rows:
        name = r.get('display_name','')
        if not name or name in s:
            continue
        field = r.get('std_field', None)
        if field is None or (isinstance(field, float) and pd.isna(field)) or str(field).strip().upper()=='NAN':
            continue
        field = str(field).strip()
        ser = _series(field)
        if ser is None:
            continue

        # 过滤 NaN / 0
        if r.get('dropna',1):
            ser = ser[ser.notna()]
        if r.get('drop0',1):
            ser = ser[ser != 0]

        if ser.empty:
            s[name] = '指标项缺失'
            continue

        # 掉线率 RAW 可能为小数比值（0.002）或百分比（0.2），做自适配
        if r.get('unit') == 'PERCENT' and field.startswith('RAW__'):
            mx = ser.max()
            if mx <= 2:  # 经验阈值：<=2 认为是百分比值或小数比值
                # 若多数值小于 0.5，按小数比值转百分比
                if (ser.abs() < 0.5).mean() > 0.6:
                    ser = ser * 100

        agg = r.get('agg','MEAN')
        if agg == 'SUM':
            val = float(ser.fillna(0).sum())
        elif agg == 'MAX':
            val = float(ser.max())
        elif agg == 'MIN':
            val = float(ser.min())
        else:  # MEAN
            val = float(ser.mean())

        dec = int(r.get('decimals',2) or 2)
        try:
            s[name] = round(val, dec)
        except Exception:
            s[name] = val

def calculate_kpis(merged_df: pd.DataFrame, tech_type: str) -> pd.DataFrame:
    return _calculate_kpis_impl(merged_df, tech_type)


def _calculate_kpis_impl(merged_df, tech_type):
    # ====== 智能名称映射（在计算前应用）======
    # 用网管的完整名称替换清单的简写名称，确保后续显示正确
    try:
        from kpi_tool.core.smart_name_mapper import build_cgi_to_name_mapping, apply_smart_name_mapping
        
        # 构建CGI->完整名称映射表
        cgi_to_name = build_cgi_to_name_mapping(merged_df, tech_type)
        
        # 应用映射：用网管完整名称替换清单简写
        if cgi_to_name:
            merged_df = apply_smart_name_mapping(merged_df, cgi_to_name)
    except Exception:
        # 映射失败不影响主流程
        pass
    # ====== 智能名称映射结束 ======
    
    if merged_df is None or merged_df.empty:
        return []
    if '厂家' not in merged_df.columns:
        merged_df['厂家'] = "Unknown"

    # ---- 1) 数值化：不要把 NaN 强行填 0（否则成功率均值会被“缺失=0”拉低）----
    col_map = standardizer.global_candidates
    non_numeric_keys = {'time_start', 'time_end', 'time_any', 'cell_name'}
    sum_keys = {'kpi_rrc_users_max', 'kpi_traffic_gb', 'kpi_volte_traffic_erl', 'kpi_vonr_traffic_erl', 'kpi_connect_num', 'kpi_connect_den', 'kpi_volte_connect_num', 'kpi_volte_connect_den'}

    for k in col_map.keys():
        std_col = f"STD__{k}"
        if std_col not in merged_df.columns:
            merged_df[std_col] = np.nan
            continue
        if k in non_numeric_keys:
            continue

        s = pd.to_numeric(merged_df[std_col], errors='coerce')
        if k in sum_keys:
            merged_df[std_col] = s.fillna(0)
        else:
            merged_df[std_col] = s  # 保留 NaN（用于均值排除缺失）

    # ---- 2) 统一单位：0~1 小数 => 0~100 百分数（按厂家&分位数判断，兼容少量脏值）----
    pct_keys_high = [
        'kpi_connect', 'kpi_ho_intra',
        'kpi_vonr_connect', 'kpi_nr2lte_ho',
        'kpi_volte_connect', 'kpi_volte_ho',
        'kpi_util_max', 'kpi_prb_ul_util', 'kpi_prb_dl_util'
    ]
    drop_keys = ['kpi_drop', 'kpi_vonr_drop', 'kpi_volte_drop']

    def _scale_frac_to_pct(idx, std_col, q=0.95, thresh=1.5):
        try:
            s = pd.to_numeric(merged_df.loc[idx, std_col], errors='coerce')
            s2 = s.dropna()
            if s2.empty:
                return
            qv = float(s2.quantile(q))
            # 用分位数而不是 max：避免少量“100”脏值影响判断
            if qv <= thresh:
                merged_df.loc[idx, std_col] = s * 100.0
        except Exception:
            pass

    # success/util：按厂家分别判断是否为 0~1（向量化 transform）
    vendor_groups = merged_df.groupby('厂家')
    for k in pct_keys_high:
        std_col = f"STD__{k}"
        if std_col not in merged_df.columns:
            continue
        s = pd.to_numeric(merged_df[std_col], errors='coerce')
        q95 = s.groupby(merged_df['厂家']).transform(
            lambda x: x.dropna().quantile(0.95) if len(x.dropna()) > 0 else np.nan)
        need_scale = q95.notna() & (q95 <= 1.5)
        merged_df[std_col] = np.where(need_scale & s.notna(), s * 100.0, s)

    # drop：更保守——只有“几乎全部 <=1.0”才当 0~1 小数（避免把 0.01% 错当成 1%）
    def _scale_drop_if_frac(idx, std_col):
        try:
            s = pd.to_numeric(merged_df.loc[idx, std_col], errors='coerce')
            s2 = s.dropna()
            if s2.empty:
                return
            qv = float(s2.quantile(0.95))
            mx = float(s2.max())
            # 更严格：仅当掉线率几乎全部为极小小数（疑似 0~1 比值）才 *100
            if qv <= 0.01 and mx <= 0.1:
                merged_df.loc[idx, std_col] = s * 100.0
        except Exception:
            pass

    for k in drop_keys:
        std_col = f"STD__{k}"
        if std_col not in merged_df.columns:
            continue
        s = pd.to_numeric(merged_df[std_col], errors='coerce')
        q95 = s.groupby(merged_df['厂家']).transform(
            lambda x: x.dropna().quantile(0.95) if len(x.dropna()) > 0 else np.nan)
        mx = s.groupby(merged_df['厂家']).transform(
            lambda x: x.dropna().max() if len(x.dropna()) > 0 else np.nan)
        need_scale = q95.notna() & (q95 <= 0.01) & mx.notna() & (mx <= 0.1)
        merged_df[std_col] = np.where(need_scale & s.notna(), s * 100.0, s)

    # ---- 3) 时间窗字符串 & 利用率综合（向量化）----
    ts = pd.to_datetime(merged_df.get('STD__time_start'), errors='coerce')
    te = pd.to_datetime(merged_df.get('STD__time_end'), errors='coerce')
    ts_str = ts.dt.strftime('%H:%M').fillna('')
    te_str = te.dt.strftime('%H:%M').fillna('')
    merged_df['time_str'] = np.where(
        ts.notna() & te.notna(), ts_str + '~' + te_str,
        np.where(ts.notna(), ts_str, '')
    )
    # KPI_UTIL: 取三者最大（已有统一单位的前提下）
    merged_df['KPI_UTIL'] = merged_df[['STD__kpi_util_max', 'STD__kpi_prb_ul_util', 'STD__kpi_prb_dl_util']].max(axis=1)

    # ---- 4) 统计函数：成功率类排除 NaN/0（0 常是缺失/无效），掉线/掉话保留 0 ----
    def safe_mean(s, drop_zero=True):
        """统一口径：先排除 NaN 和 0，再计算平均值（MEAN(排NaN/0)）。
        drop_zero 参数保留仅为兼容旧调用，但不再改变行为。
        """
        s = pd.to_numeric(s, errors='coerce')
        s = s.where(s != 0)
        s = s.dropna()
        return float(s.mean()) if len(s) else 0.0

    def safe_sum(s):
        s = pd.to_numeric(s, errors='coerce').fillna(0)
        return float(s.sum()) if len(s) else 0.0

    def safe_max(s):
        s = pd.to_numeric(s, errors='coerce').dropna()
        return float(s.max()) if len(s) else 0.0

    def calc_rate(group, std_num_name, std_den_name, std_pct_name, drop_zero_pct=True):
        """优先用 分子/分母 做加权平均；若不存在，则回落到百分比列的简单平均。"""
        if std_num_name in group.columns and std_den_name in group.columns:
            num = pd.to_numeric(group[std_num_name], errors='coerce')
            den = pd.to_numeric(group[std_den_name], errors='coerce')
            mask = den > 0
            if mask.any():
                total_num = num[mask].clip(lower=0).sum()
                total_den = den[mask].sum()
                if total_den > 0:
                    return float(total_num / total_den * 100.0)
        if std_pct_name in group.columns:
            s = pd.to_numeric(group[std_pct_name], errors='coerce')
            if drop_zero_pct:
                s = s.where(s != 0)
            s = s.dropna()
            if len(s):
                return float(s.mean())
        return 0.0

    def calc_rate_mean(group, std_num_name, std_den_name, std_pct_name, drop_zero_pct=True):
        """统一口径：先得到“行级百分比”，再做 MEAN(排NaN/0)。

        行级百分比来源优先级：
        1）若 std_pct_name 列存在且该行值非 NaN 且非 0，直接使用；
        2）否则若 num/den 同时存在且 den>0，则用 num/den*100 计算补齐；
        3）最后对结果排除 NaN 和 0，取平均；若无有效值返回 np.nan。
        """
        # 1) 百分比列
        if std_pct_name in group.columns:
            pct = pd.to_numeric(group[std_pct_name], errors='coerce')
        else:
            pct = pd.Series(np.nan, index=group.index)

        # 2) num/den 兜底补齐
        if (std_num_name in group.columns) and (std_den_name in group.columns):
            num = pd.to_numeric(group[std_num_name], errors='coerce')
            den = pd.to_numeric(group[std_den_name], errors='coerce')
            computed = (num / den) * 100.0
            computed = computed.where(den > 0)
            pct = pct.where(pct.notna() & (pct != 0), computed)

        pct = pd.to_numeric(pct, errors='coerce')
        if drop_zero_pct:
            pct = pct.where(pct != 0)
        pct = pct.dropna()
        return float(pct.mean()) if len(pct) else np.nan



    results = []
    if '活动名称' not in merged_df.columns:
        return []

    # 区域规范化：去空格/NaN/None，确保“场内/场外/整体”分组不丢失
    has_region = '区域' in merged_df.columns
    if has_region:
        def _norm_region(x):
            try:
                if pd.isna(x):
                    return ''
            except Exception:
                pass
            s = str(x).strip()
            if s.lower() in ('nan', 'none'):
                return ''
            return s
        merged_df['区域'] = merged_df['区域'].apply(_norm_region)

    # ===== 构建 ID→名称 查找表（用于 _safe_get_cell_name 兜底查找）=====
    _id_to_raw_name = {}   # list_key_id / raw_key_id → 原始网管小区名称
    _id_to_list_name = {}  # list_key_id → 清单小区名称(B列)
    def _is_valid(v):
        if v is None: return False
        s = str(v).strip()
        return s and s.lower() not in ('nan', 'none', '')

    # 从merged_df构建：raw_key_id → raw_key_name（原始网管名称，最优先）
    if 'raw_key_id' in merged_df.columns and 'raw_key_name' in merged_df.columns:
        tmp = merged_df[['raw_key_id', 'raw_key_name']].drop_duplicates()
        rid = tmp['raw_key_id'].astype(str).str.strip()
        rname = tmp['raw_key_name'].astype(str).str.strip()
        valid = rid.ne('') & rid.str.lower().ne('nan') & rname.ne('') & rname.str.lower().ne('nan')
        _id_to_raw_name = dict(zip(rid[valid], rname[valid]))
    # 也用 list_key_id → raw_key_name（因为匹配后同一行的list_key_id对应raw_key_name）
    if 'list_key_id' in merged_df.columns and 'raw_key_name' in merged_df.columns:
        tmp = merged_df[['list_key_id', 'raw_key_name']].drop_duplicates()
        lid = tmp['list_key_id'].astype(str).str.strip()
        rname = tmp['raw_key_name'].astype(str).str.strip()
        valid = lid.ne('') & lid.str.lower().ne('nan') & rname.ne('') & rname.str.lower().ne('nan')
        for k, v in zip(lid[valid], rname[valid]):
            _id_to_raw_name.setdefault(k, v)
    # 从merged_df构建：list_key_id → list_key_name（清单名称，次优先）
    if 'list_key_id' in merged_df.columns and 'list_key_name' in merged_df.columns:
        tmp = merged_df[['list_key_id', 'list_key_name']].drop_duplicates()
        lid = tmp['list_key_id'].astype(str).str.strip()
        lname = tmp['list_key_name'].astype(str).str.strip()
        valid = lid.ne('') & lid.str.lower().ne('nan') & lname.ne('') & lname.str.lower().ne('nan')
        _id_to_list_name = dict(zip(lid[valid], lname[valid]))

    def _safe_get_cell_name(row):
        """安全获取小区名称
        
        优先级：
        1. 当前行的名称字段（raw_key_name等）
        2. 用list_key_id查找原始网管中的小区名称
        3. 用list_key_id查找保障清单中的小区中文名(B列)
        4. 用ID字段(ECGI/CGI)作为显示
        5. 返回'-'
        """
        # 第1优先：当前行直接取名称字段
        name_candidates = ['raw_key_name', 'list_key_name', 'STD__cell_name', '小区名称', '小区中文名', 'CELL_NAME']
        for field in name_candidates:
            try:
                val = row.get(field)
                if val is not None:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in ('nan', 'none', ''):
                        return val_str
            except Exception:
                continue
        
        # 第2优先：用list_key_id查找原始网管小区名称
        # 第3优先：用list_key_id查找清单小区名称
        for id_field in ['list_key_id', 'raw_key_id']:
            try:
                kid = str(row.get(id_field, '')).strip()
                if kid and kid.lower() not in ('nan', 'none', ''):
                    if kid in _id_to_raw_name:
                        return _id_to_raw_name[kid]
                    if kid in _id_to_list_name:
                        return _id_to_list_name[kid]
            except Exception:
                continue
        
        # 第4优先：直接显示ID
        for field in ['list_key_id', 'raw_key_id', 'CGI/ECGI', 'ECGI', 'CGI']:
            try:
                val = row.get(field)
                if val is not None:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in ('nan', 'none', ''):
                        return val_str
            except Exception:
                continue
        
        return '-'

    def _calc_one(act, region_label, grp):
        """计算单个(活动, 区域)的指标结果。region_label 传入期望显示的区域名称（如 整体/场内/场外）。"""
        s = {
            '指标时间': grp['time_str'].iloc[0] if 'time_str' in grp.columns else "",
            '活动名称': act,
            '区域': region_label if region_label else "整体",
            '厂家': "/".join(sorted(grp['厂家'].astype(str).unique()))
        }

        # SUM 类指标需按网管小区去重，防止清单重复条目导致同一小区被重复累加
        grp_unique = grp
        dedup_col = None
        for c in ['raw_key_name', 'raw_key_id', 'STD__cell_name']:
            if c in grp.columns and grp[c].notna().any() and (grp[c].astype(str).str.strip() != '').any():
                dedup_col = c
                break
        if dedup_col is not None:
            grp_unique = grp.drop_duplicates(subset=[dedup_col], keep='first')

        s['总用户数'] = int(safe_sum(grp_unique['STD__kpi_rrc_users_max']))
        s['总流量(GB)'] = round(safe_sum(grp_unique['STD__kpi_traffic_gb']), 2)

        # 成功率类：排除 0/NaN
        # 成功率类：统一按 MEAN(排NaN/0) 口径
        # - 若已标准化出 STD__kpi_connect，则直接对该列取均值
        # - 若仅存在 num/den，则按( num/den*100 )逐行计算后再取均值（不是加权）
        _conn = pd.to_numeric(grp.get('STD__kpi_connect', pd.Series([], dtype=float)), errors='coerce')
        if _conn.notna().any():
            s['无线接通率(%)'] = round(safe_mean(_conn, drop_zero=True), 2)
        else:
            _num = pd.to_numeric(grp.get('STD__kpi_connect_num', pd.Series([], dtype=float)), errors='coerce')
            _den = pd.to_numeric(grp.get('STD__kpi_connect_den', pd.Series([], dtype=float)), errors='coerce')
            _ratio = (_num / _den * 100).replace([float('inf'), -float('inf')], pd.NA)
            s['无线接通率(%)'] = round(safe_mean(_ratio, drop_zero=True), 2) if _ratio.notna().any() else '指标项缺失'
        s['系统内切换出成功率(%)'] = round(safe_mean(grp['STD__kpi_ho_intra'], drop_zero=True), 2)

        # 掉线率：优先使用 RAW__kpi_drop 计算均值（仅排NaN，0值视为有效）
        # 说明：部分厂家 RAW__kpi_drop 已是“百分比值”（如 0.2584 表示 0.2584%）；
        #       部分厂家是“小数比值”（如 0.00155 表示 0.155%）。此处做一次轻量口径自适应。
        _raw_drop = pd.to_numeric(grp.get('RAW__kpi_drop', pd.Series([], dtype=float)), errors='coerce')
        if _raw_drop.notna().any():
            _d = _raw_drop.dropna().copy()
            _nz = _d[_d != 0]
            # 若非零值整体非常小（<=0.1），更像是小数比值，转换成百分比（*100）
            if len(_nz) > 0 and (_nz.quantile(0.95) <= 0.01) and (_nz.max() <= 0.1):
                _d = _d * 100
            s['无线掉线率(%)'] = round(float(_d.mean()), 3)
        else:
            # fallback：若无 RAW__kpi_drop，则尝试 STD__kpi_drop，但不再用“>=2 全部判缺失”这一硬规则
            _std_drop = pd.to_numeric(grp.get('STD__kpi_drop', pd.Series([], dtype=float)), errors='coerce')
            if _std_drop.notna().any():
                _d = _std_drop.dropna()
                s['无线掉线率(%)'] = round(float(_d.mean()), 3)
            else:
                s['无线掉线率(%)'] = '指标项缺失'

        s['平均干扰(dBm)'] = round(safe_mean(grp.get('STD__kpi_ul_interf_dbm', pd.Series([], dtype=float)), drop_zero=True), 2)

        if tech_type == '5G':
            s['VoNR无线接通率(%)'] = round(safe_mean(grp['STD__kpi_vonr_connect'], drop_zero=True), 2)
            s['VoNR到VoLTE切换成功率(%)'] = round(safe_mean(grp['STD__kpi_nr2lte_ho'], drop_zero=True), 2)
            s['VoNR掉线率(5QI1)(%)'] = round(safe_mean(grp['STD__kpi_vonr_drop'], drop_zero=False), 2)
            s['VoNR话务量(Erl)'] = round(safe_sum(grp['STD__kpi_vonr_traffic_erl']), 2)
            s['5G利用率最大值(%)'] = round(safe_max(grp['KPI_UTIL']), 2)
        else:
            # VoLTE无线接通率：统一按 MEAN(排NaN/0) 口径（不是加权）
            _v = pd.to_numeric(grp.get('STD__kpi_volte_connect', pd.Series([], dtype=float)), errors='coerce')
            if _v.notna().any():
                s['VoLTE无线接通率(%)'] = round(safe_mean(_v, drop_zero=True), 2)
            else:
                _n = pd.to_numeric(grp.get('STD__kpi_volte_connect_num', pd.Series([], dtype=float)), errors='coerce')
                _d = pd.to_numeric(grp.get('STD__kpi_volte_connect_den', pd.Series([], dtype=float)), errors='coerce')
                _r = (_n / _d * 100).replace([float('inf'), -float('inf')], pd.NA)
                s['VoLTE无线接通率(%)'] = round(safe_mean(_r, drop_zero=True), 2) if _r.notna().any() else '指标项缺失'
            if 'STD__kpi_volte_ho' in grp.columns and pd.to_numeric(grp['STD__kpi_volte_ho'], errors='coerce').notna().any():
                s['VoLTE切换成功率(%)'] = round(safe_mean(grp['STD__kpi_volte_ho'], drop_zero=True), 2)
            else:
                s['VoLTE切换成功率(%)'] = '指标项缺失'
            s['E-RAB掉话率(QCI=1)(%)'] = round(safe_mean(grp['STD__kpi_volte_drop'], drop_zero=False), 2)
            s['VoLTE话务量(Erl)'] = round(safe_sum(grp['STD__kpi_volte_traffic_erl']), 2)
            s['4G利用率最大值(%)'] = round(safe_max(grp['KPI_UTIL']), 2)

        # 最忙小区：避免 KPI_UTIL 全 NaN 导致 idxmax 崩溃
        if grp['KPI_UTIL'].notna().any():
            max_idx = grp['KPI_UTIL'].idxmax()
            busy_cell = grp.loc[max_idx]
        else:
            busy_cell = grp.iloc[0]

        c_name = _safe_get_cell_name(busy_cell)
        s['最大利用率小区'] = c_name
        s['最大利用率小区的利用率'] = round(float(busy_cell.get('KPI_UTIL', 0) or 0), 2)
        s['最大利用率小区的用户数'] = int(pd.to_numeric(busy_cell.get('STD__kpi_rrc_users_max', 0), errors='coerce') or 0)

        limit = 90 if tech_type == '5G' else 85
        high_load = grp[(grp['KPI_UTIL'] >= limit) & (grp['STD__kpi_rrc_users_max'] >= 100)]
        s['高负荷小区数'] = len(high_load)

        # ========= 质差小区数统计（用于微信简报状态判定）========
        # 条件：接通率<95 或 掉线率>1 或 干扰>-100
        _pq_connect = pd.to_numeric(grp.get('STD__kpi_connect', pd.Series([], dtype=float)), errors='coerce')
        _pq_drop = pd.to_numeric(grp.get('STD__kpi_drop', pd.Series([], dtype=float)), errors='coerce')
        _pq_interf = pd.to_numeric(grp.get('STD__kpi_ul_interf_dbm', pd.Series([], dtype=float)), errors='coerce')
        poor_quality_mask = (_pq_connect < 95) | (_pq_drop > 1) | (_pq_interf > -100)
        poor_quality_cnt = int(poor_quality_mask.sum())
        s['质差小区数'] = poor_quality_cnt

        _apply_catalog_generic_kpis(s, grp, tech_type)

        cols = get_output_cols(tech_type)
        # 若配置要求输出但本轮未生成，填充为‘指标项缺失’（保持可对账）
        for c in cols:
            if c not in s:
                s[c] = '指标项缺失'

        return (pd.DataFrame([s], columns=cols), high_load, poor_quality_cnt)

    if not has_region:
        for act, grp in merged_df.groupby('活动名称'):
            results.append(_calc_one(act, "整体", grp))
    else:
        act_to_regs = {}
        # 先按(活动, 区域)分别计算
        for (act, reg0), grp in merged_df.groupby(['活动名称', '区域'], dropna=False):
            reg = reg0 if str(reg0).strip() != '' else "整体"
            results.append(_calc_one(act, reg, grp))
            act_to_regs.setdefault(act, set()).add(reg)

        # 对“存在区域划分”的活动，补充“整体”（全量汇总）——保证微信简报按区域完整输出
        for act, regs in act_to_regs.items():
            if regs == {'整体'}:
                continue
            if '整体' in regs:
                continue
            grp_all = merged_df[merged_df['活动名称'] == act]
            if grp_all is None or grp_all.empty:
                continue
            results.append(_calc_one(act, "整体", grp_all))
    return results
# ----------------------------------------------------------------------
# Backwards-compatibility shim
# Some callers expect: import kpi_tool.core.calculator as CALC; CALC.CALC.calculate_kpis(...)
# Keep behavior identical by delegating to the existing module-level functions.

class _CalcShim:
    @staticmethod
    def calculate_kpis(tech_type, df):
        return calculate_kpis(df, tech_type)

    @staticmethod
    def get_poor_quality(tech_type, df):
        return get_poor_quality(df, tech_type)

CALC = _CalcShim


