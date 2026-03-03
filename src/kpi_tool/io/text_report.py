# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime
import os
from typing import Any, Dict, List, Optional

import pandas as pd
import numpy as np

from kpi_tool.config.logging_config import log_print
from kpi_tool.utils.helpers import format_time_range

def generate_text_report(res_4g_list, res_5g_list, has_region_col: bool = False) -> str:
    return _generate_text_report_impl(res_4g_list, res_5g_list, has_region_col=has_region_col)


def _generate_text_report_impl(res_4g_list, res_5g_list, has_region_col=False):
    """仅用于生成微信简报(txt)的排版输出。

    严格遵守：不改变任何计算逻辑/统计口径/数据处理流程，只组织输出文本。

    本次调整焦点：当同一个活动涉及多个区域（整体/场内/场外/其它...）时：
      - 标题行/厂家行/起始分隔线：每个活动仅输出一次；
      - 按固定顺序输出“场景块”：整体 → 场内 → 场外 → 其它（名称排序）；
      - 场景块之间仅使用“━━━━━━━━━━━━━━”分隔（末尾不重复）；
      - 每个场景块内部：先输出5G（若有），再输出4G（若有）；

    其他说明：仅调整微信简报排版，其它逻辑保持不变。
    """

    def _is_blank(x):
        if x is None:
            return True
        s = str(x).strip()
        return (s == '' or s.lower() == 'nan' or s.lower() == 'none')

    def _norm_region(x):
        """区域空值统一归为“整体”（仅用于TXT排版，不影响任何数据/计算）。"""
        return '整体' if _is_blank(x) else str(x).strip()

    def _fmt_time_window(t):
        """将各种时间表示统一为 HH:MM–HH:MM（默认 +15min）。"""
        if _is_blank(t):
            return '未知时间'
        s = str(t).strip()
        # 已包含“–”或“~”的区间：尽量抽取 HH:MM
        if '–' in s:
            parts = s.split('–', 1)
            if len(parts) == 2:
                return f"{parts[0].strip()}–{parts[1].strip()}"
        if '~' in s:
            a, b = s.split('~', 1)
            return f"{a.strip()}–{b.strip()}"
        if '至' in s:
            a, b = s.split('至', 1)
            return f"{a.strip()}–{b.strip()}"
        # 尝试解析单点时间
        try:
            dt = pd.to_datetime(s, errors='coerce')
            if pd.notna(dt):
                dt2 = dt + pd.to_timedelta(15, unit='m')
                return f"{dt.strftime('%H:%M')}–{dt2.strftime('%H:%M')}"
        except Exception:
            pass
        # 最后兜底：直接返回原串
        return s

    def _to_float(x):
        if x is None:
            return None
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        s = str(x).strip()
        if s == '' or s.lower() in ('nan', 'none'):
            return None
        s = s.replace('%', '').replace('％', '')
        try:
            return float(s)
        except Exception:
            return None

    def _fmt_pct(x):
        """百分比统一 xx.xx%（若为文本如“指标项缺失”则原样返回）。"""
        if _is_blank(x):
            return '指标项缺失'
        fv = _to_float(x)
        if fv is None:
            return str(x)
        # 可能出现 0.0123 表示 1.23% 的情况：不擅自改变口径，只按当前数值显示
        return f"{fv:.2f}%"

    def _fmt_gb(x):
        if _is_blank(x):
            return '0.00GB'
        fv = _to_float(x)
        if fv is None:
            s = str(x).strip()
            if s.upper().endswith('GB'):
                return s
            return s + 'GB'
        return f"{fv:.2f}GB"

    def _fmt_int(x):
        fv = _to_float(x)
        if fv is None:
            return 0
        try:
            return int(round(fv))
        except Exception:
            return 0

    def _status_by_hl(hl_cnt, pq_cnt=0):
        """
        微信简报状态判定（仅用于展示，不改变任何计算逻辑）
        新规则：只有当 (高负荷小区数 > 5) 且 (质差小区数 > 5) 同时满足时，返回 '⚠️异常'
        其他情况均返回 '✅稳定'
        """
        try:
            return '⚠️异常' if (int(hl_cnt) > 5 and int(pq_cnt) > 5) else '✅稳定'
        except Exception:
            return '✅稳定'

    def _safe_str(x, fallback='指标项缺失'):
        return fallback if _is_blank(x) else str(x)

    def _is_missing_text(v):
        if _is_blank(v):
            return True
        ss = str(v).strip()
        return (ss == '' or ss.lower() == 'nan' or ss in ('指标项缺失', '--', '-'))

    def _has_data(summary: dict, tech: str) -> bool:
        """用于微信简报：当活动/场景只涉及单网时，另一网不输出空段。"""
        if not isinstance(summary, dict) or not summary:
            return False
        u = _to_float(summary.get('总用户数', None))
        t = _to_float(summary.get('总流量(GB)', None))
        if (u is not None and u > 0) or (t is not None and t > 0):
            return True
        keys = (['无线接通率(%)', '无线掉线率(%)', '系统内切换出成功率(%)', 'VoLTE无线接通率(%)', 'VoLTE切换成功率(%)', 'E-RAB掉话率(QCI=1)(%)']
                if tech == '4G' else
                ['无线接通率(%)', '无线掉线率(%)', '系统内切换出成功率(%)', 'VoNR无线接通率(%)', 'VoNR到VoLTE切换成功率(%)', 'VoNR掉线率(5QI1)(%)'])
        for k in keys:
            vv = summary.get(k, None)
            # 允许 0（如掉线 0%）视为有数据
            if not _is_missing_text(vv):
                return True
        return False

    def _build_block_4g(s, hl_cnt, pq_cnt=0):
        status = _status_by_hl(hl_cnt, pq_cnt)
        l1 = f"🟦 4G｜状态 {status}｜高负荷 {int(hl_cnt) if str(hl_cnt).isdigit() else hl_cnt}｜整体用户 {_fmt_int(s.get('总用户数'))}｜整体流量 {_fmt_gb(s.get('总流量(GB)'))}"
        l2 = f"🔥 最忙小区：{_safe_str(s.get('最大利用率小区'))}｜利用率 {_fmt_pct(s.get('最大利用率小区的利用率'))}｜用户 {_fmt_int(s.get('最大利用率小区的用户数'))}"
        l3 = f"数据：接通 {_fmt_pct(s.get('无线接通率(%)'))} ｜掉线 {_fmt_pct(s.get('无线掉线率(%)'))} ｜切换 {_fmt_pct(s.get('系统内切换出成功率(%)'))}"
        l4 = f"语音：VoLTE接通 {_fmt_pct(s.get('VoLTE无线接通率(%)'))} ｜切换 {_fmt_pct(s.get('VoLTE切换成功率(%)'))} ｜ERAB掉话 {_fmt_pct(s.get('E-RAB掉话率(QCI=1)(%)'))}"
        return [l1, l2, l3, l4]

    def _build_block_5g(s, hl_cnt, pq_cnt=0):
        status = _status_by_hl(hl_cnt, pq_cnt)
        l1 = f"🟩 5G｜状态 {status}｜高负荷 {int(hl_cnt) if str(hl_cnt).isdigit() else hl_cnt}｜整体用户 {_fmt_int(s.get('总用户数'))}｜整体流量 {_fmt_gb(s.get('总流量(GB)'))}"
        l2 = f"🔥 最忙小区：{_safe_str(s.get('最大利用率小区'))}｜利用率 {_fmt_pct(s.get('最大利用率小区的利用率'))}｜用户 {_fmt_int(s.get('最大利用率小区的用户数'))}"
        l3 = f"数据：接通 {_fmt_pct(s.get('无线接通率(%)'))} ｜掉线 {_fmt_pct(s.get('无线掉线率(%)'))} ｜切换 {_fmt_pct(s.get('系统内切换出成功率(%)'))}"
        l4 = f"语音：VoNR接通 {_fmt_pct(s.get('VoNR无线接通率(%)'))} ｜切换 {_fmt_pct(s.get('VoNR到VoLTE切换成功率(%)'))} ｜5QI1掉线 {_fmt_pct(s.get('VoNR掉线率(5QI1)(%)'))}"
        return [l1, l2, l3, l4]

    def _split_vendors(v):
        if _is_blank(v):
            return []
        ss = str(v).strip()
        for d in ['、', ';', '；', ',', '，', '|', '｜']:
            ss = ss.replace(d, '/')
        return [p.strip() for p in ss.split('/') if p and p.strip() and p.strip().lower() not in ('nan', 'none')]

    # 1) 组装 (活动, 区域) -> (summary_row_dict, highload_count, poor_quality_count)
    map_4g = {}
    for r in res_4g_list or []:
        if r and (not r[0].empty):
            row = r[0].iloc[0].to_dict()
            act = row.get('活动名称', '未知活动')
            reg = _norm_region(row.get('区域', '整体'))
            hl_cnt = len(r[1]) if r[1] is not None else 0
            pq_cnt = r[2] if len(r) > 2 else 0  # 质差小区数（兼容旧返回格式）
            map_4g[(act, reg)] = (row, hl_cnt, pq_cnt)

    map_5g = {}
    for r in res_5g_list or []:
        if r and (not r[0].empty):
            row = r[0].iloc[0].to_dict()
            act = row.get('活动名称', '未知活动')
            reg = _norm_region(row.get('区域', '整体'))
            hl_cnt = len(r[1]) if r[1] is not None else 0
            pq_cnt = r[2] if len(r) > 2 else 0  # 质差小区数（兼容旧返回格式）
            map_5g[(act, reg)] = (row, hl_cnt, pq_cnt)

    all_keys = sorted(set(map_4g.keys()) | set(map_5g.keys()))

    # 活动分组
    from collections import defaultdict
    act_groups = defaultdict(set)
    for act, reg in all_keys:
        act_groups[act].add(_norm_region(reg))

    # 分割线：严格使用“--------------------------------”
    divider = '-' * 32
    scene_divider = '━━━━━━━━━━━━━━'

    def _reg_sort(x):
        s = _norm_region(x)
        if s == '整体':
            return (0, '')
        if s == '场内':
            return (1, '')
        if s == '场外':
            return (2, '')
        return (3, s)

    blocks = []

    for act in sorted(act_groups.keys()):
        regions_all = sorted(act_groups[act], key=_reg_sort)

        # 判断是否为“同一活动涉及多个区域”的场景：存在除“整体”以外的区域
        multi_region = bool(has_region_col and any(r != '整体' for r in regions_all))

        # 预扫描：仅纳入有数据的区域（避免空块、避免厂家统计被空块影响）
        regions_with_data = []
        vset = set()
        time_candidates = []

        for reg in regions_all:
            s4, hl4, pq4 = map_4g.get((act, reg), ({}, 0, 0))
            s5, hl5, pq5 = map_5g.get((act, reg), ({}, 0, 0))
            has4 = _has_data(s4, '4G')
            has5 = _has_data(s5, '5G')
            if not (has4 or has5):
                continue
            regions_with_data.append(reg)
            if has4:
                vset.update(_split_vendors(s4.get('厂家', None)))
                tv = s4.get('指标时间', None)
                if not _is_blank(tv):
                    time_candidates.append((reg, tv))
            if has5:
                vset.update(_split_vendors(s5.get('厂家', None)))
                tv = s5.get('指标时间', None)
                if not _is_blank(tv):
                    time_candidates.append((reg, tv))

        if not regions_with_data:
            # 该活动无任何可输出数据，跳过
            continue

        # 标题时间：优先取“整体”，否则取第一个有数据的候选
        tval = None
        for reg, tv in time_candidates:
            if _norm_region(reg) == '整体':
                tval = tv
                break
        if _is_blank(tval) and time_candidates:
            tval = time_candidates[0][1]
        tstr = _fmt_time_window(tval)

        vendor_str = '/'.join(sorted(vset)) if vset else '未知厂家'

        # 标题行 / 厂家行 / 起始分隔线：每个活动只输出一次
        blocks.append(f"📡 {act}（{tstr}）")
        blocks.append(f"{vendor_str} 指标监控通报")
        blocks.append(divider)

        if multi_region:
            # 场景块：整体 → 场内 → 场外 → 其它（名称排序）
            regions_sorted = sorted(set(regions_with_data), key=_reg_sort)

            printed = 0
            for reg in regions_sorted:
                s4, hl4, pq4 = map_4g.get((act, reg), ({}, 0, 0))
                s5, hl5, pq5 = map_5g.get((act, reg), ({}, 0, 0))
                has4 = _has_data(s4, '4G')
                has5 = _has_data(s5, '5G')
                if not (has4 or has5):
                    continue

                if printed > 0:
                    blocks.append(scene_divider)

                # 场景标题：用于分块识别（仅排版，不改任何计算/取值）
                blocks.append(f"※场景：{_norm_region(reg)}")

                # 每个场景块内部：先 5G，后 4G（均为“有数据才输出”）
                if has5:
                    blocks.extend(_build_block_5g(s5, hl5, pq5))
                if has5 and has4:
                    blocks.append('')
                if has4:
                    blocks.extend(_build_block_4g(s4, hl4, pq4))

                printed += 1

            # 简报末尾：保留结束分隔线（每个活动一条）
            blocks.append(divider)
        else:
            # 非多区域场景：保持原来“单块输出”的信息结构（仅输出有数据的网络）
            reg = _norm_region(regions_with_data[0])
            s4, hl4, pq4 = map_4g.get((act, reg), ({}, 0, 0))
            s5, hl5, pq5 = map_5g.get((act, reg), ({}, 0, 0))
            has4 = _has_data(s4, '4G')
            has5 = _has_data(s5, '5G')

            # 维持原有顺序：4G 在前，5G 在后
            if has4:
                blocks.extend(_build_block_4g(s4, hl4, pq4))
            if has4 and has5:
                blocks.append('')
            if has5:
                blocks.extend(_build_block_5g(s5, hl5, pq5))

            blocks.append(divider)

        # 活动间空行（最后会统一裁剪尾部空行，保证全文末尾为 divider）
        blocks.append('')

    # 去掉末尾空行，确保全文以 divider 结尾
    while blocks and blocks[-1] == '':
        blocks.pop()

    return "\n".join(blocks)
