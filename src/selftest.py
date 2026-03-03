# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any, Dict

import numpy as np
import pandas as pd

from kpi_tool.config.logging_config import log_print

def _coerce_float(x):
    try:
        if isinstance(x, str):
            # 去掉可能的百分号/空格
            x = x.replace("%","").strip()
        return float(x)
    except Exception:
        return None

def get_default_selftest_expected():
    # 你明确提供的“手动计算正确值”，作为默认自检基准。
    # 若现场数据（小区数/时间窗/提取口径）不同，可自行修改为外部 expected.json / expected.xlsx。
    return {
        ("4G","华为4G测试"): {
            "无线接通率(%)": 99.83,
            "无线掉线率(%)": 0.015,
            "系统内切换出成功率(%)": 99.35,
            "VoLTE无线接通率(%)": 99.90,
            "VoLTE话务量(Erl)": 9.68,
            "平均干扰(dBm)": -116.50,
        },
        ("4G","诺基亚4G测试"): {
            "无线接通率(%)": 99.94,
            "无线掉线率(%)": 0.014,
            "系统内切换出成功率(%)": 99.62,
            "VoLTE无线接通率(%)": 99.96,
        },
        ("5G","中兴5G测试"): {
            "总用户数": 4428,
            "无线接通率(%)": 99.58,
            "无线掉线率(%)": 0.04,
            "VoNR无线接通率(%)": 99.67,
            "VoNR话务量(Erl)": 12.30,
            "平均干扰(dBm)": -114.55,
            "最大利用率小区的用户数": 215,
        },
        ("5G","华为5G测试"): {
            "VoNR话务量(Erl)": 25.83,
        }
    }

def run_selftest(res_4g_list, res_5g_list, cfg):
    expected = get_default_selftest_expected()

    got = {}
    for r in res_4g_list:
        if not r[0].empty:
            d = r[0].iloc[0].to_dict()
            got[("4G", str(d.get("活动名称","")))] = d
    for r in res_5g_list:
        if not r[0].empty:
            d = r[0].iloc[0].to_dict()
            got[("5G", str(d.get("活动名称","")))] = d

    def tol_for(k):
        if k in ("无线掉线率(%)",):
            return cfg["selftest_tol_drop"]
        if "话务量(Erl)" in k:
            return cfg["selftest_tol_erl"]
        if "干扰(dBm)" in k:
            return cfg["selftest_tol_dbm"]
        if k in ("总用户数","最大利用率小区的用户数"):
            return cfg["selftest_tol_users"]
        return cfg["selftest_tol_pct"]

    lines = []
    ok_all = True
    lines.append("============================================================")
    lines.append("✅ 自检报告（SELFTEST）")
    lines.append("============================================================")

    for key, exp_map in expected.items():
        tech, act = key
        if key not in got:
            ok_all = False
            lines.append(f"[FAIL] {tech}｜{act}：未在输出结果中找到该活动（请确认 only_activity/清单活动名/输入文件）")
            continue
        d = got[key]
        lines.append(f"--- {tech}｜{act} ---")
        for metric, exp in exp_map.items():
            got_val = d.get(metric, None)
            gv = _coerce_float(got_val)
            tol = tol_for(metric)
            if gv is None:
                ok_all = False
                lines.append(f"  [FAIL] {metric}: 期望 {exp}，实际无法解析（{got_val}）")
                continue
            ok = abs(gv - float(exp)) <= tol
            ok_all = ok_all and ok
            lines.append(("  [PASS] " if ok else "  [FAIL] ") + f"{metric}: 期望 {exp}，实际 {gv:.6g}，容差 ±{tol}")
    lines.append("------------------------------------------------------------")
    lines.append("总体结果: " + ("PASS" if ok_all else "FAIL"))
    return ok_all, "\n".join(lines)
