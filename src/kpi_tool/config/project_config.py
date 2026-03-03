# -*- coding: utf-8 -*-
from __future__ import annotations

import configparser
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from kpi_tool.config import constants as C
from kpi_tool.config.logging_config import log_print

# ================= 项目配置（全局） =================

PROJECT_CFG: Dict[str, Any] = {
    "project_name": "DEFAULT",
    "enabled_4g": True,
    "enabled_5g": True,
    "output": {},
    "percent_rows": [],
    "catalog": {},
    "kpi_poor_quality": {},
    "kpi_high_load": {},
    "vendor_map": {},
}

CFG: Dict[str, Any] = {
    "enable_template_mode": False,
    "enable_fuzzy_match": True,
    "enable_verbose_log": False,
}

# 保持旧版默认值（如需调整请在外部传入）
PROJECT_DEFAULT_CONFIG_DIR = r"E:\Tool_Build\配置文件"

def _pick_project_config_path(args, config_ini_path):
    """确定项目配置Excel路径：
    优先级：命令行 --project_config > config.ini [project].project_config > config_dir/项目配置.xlsx > C.BASE_DIR/项目配置_模板_30指标.xlsx
    """
    # 1) CLI
    if getattr(args, "project_config", ""):
        return getattr(args, 'project_config', None)

    # 2) config.ini
    try:
        cp = configparser.ConfigParser()
        if config_ini_path and os.path.exists(config_ini_path):
            _read_ini_safely(cp, config_ini_path)
        if "project" in cp and cp["project"].get("project_config", "").strip():
            return cp["project"].get("project_config").strip()
    except Exception:
        pass

    # 3) config_dir/项目配置.xlsx
    cfg_dir = ""
    if getattr(args, "config_dir", ""):
        cfg_dir = getattr(args, 'config_dir', None)
    else:
        # config.ini 也可提供 config_dir
        try:
            cp = configparser.ConfigParser()
            if config_ini_path and os.path.exists(config_ini_path):
                _read_ini_safely(cp, config_ini_path)
            if "project" in cp and cp["project"].get("config_dir", "").strip():
                cfg_dir = cp["project"].get("config_dir").strip()
        except Exception:
            pass
    if not cfg_dir:
        cfg_dir = PROJECT_DEFAULT_CONFIG_DIR
    cand = os.path.join(cfg_dir, "项目配置.xlsx")
    if os.path.exists(cand):
        return cand

    # 4) fallback template in C.BASE_DIR
    cand2 = os.path.join(C.BASE_DIR, "项目配置_模板_30指标.xlsx")
    if os.path.exists(cand2):
        return cand2

    # 最后兜底：不启用配置化
    return ""

def load_project_config_excel(config_path: str | None = None) -> dict:
    """加载项目配置（兼容旧版 API）。"""
    return _load_project_config_excel_impl(config_path)


def _load_project_config_excel_impl(project_config_path: str):
    """读取项目配置Excel。

    要求：至少包含 KPI_Catalog sheet（列：display_name/tech/std_field/agg/dropna/drop0/unit/decimals/enabled）。
    若缺失/读取失败，则返回空配置（保持脚本原有硬编码口径与输出）。

    输出：
      - kpi_by_tech: {"4G": [dict,...], "5G": [dict,...]} 仅 enabled=1
      - output_cols_by_tech: 以 display_name 顺序拼接（前置：指标时间/活动名称/区域/厂家）
      - percent_rows_by_tech: unit==PERCENT 的 display_name 列表（用于转置表按“行”设置百分比格式）
    """
    cfg = {
        "project_config": project_config_path,
        "kpi_catalog": None,
        "kpi_by_tech": {"4G": [], "5G": []},
        "output_cols_by_tech": {"4G": None, "5G": None},
        "percent_rows_by_tech": {"4G": [], "5G": []},
    }
    if not project_config_path or not os.path.exists(project_config_path):
        return cfg
    try:
        df = pd.read_excel(project_config_path, sheet_name="KPI_Catalog")
    except Exception:
        return cfg

    if df is None or df.empty or "display_name" not in df.columns or "tech" not in df.columns:
        return cfg

    # 统一字段
    df = df.copy()
    for c in ["enabled", "dropna", "drop0", "decimals"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "enabled" in df.columns:
        df = df[df["enabled"].fillna(1) == 1]

    # 顺序：优先 order（若存在）
    if "order" in df.columns:
        df["order"] = pd.to_numeric(df["order"], errors="coerce")
        df = df.sort_values(["tech", "order"], kind="mergesort")

    cfg["kpi_catalog"] = df

    def _rows(tech):
        sub = df[df["tech"].astype(str).str.upper() == tech]
        out = []
        for _, r in sub.iterrows():
            out.append({
                "kpi_id": str(r.get("kpi_id", "") or r.get("id", "") or "").strip(),
                "display_name": str(r.get("display_name", "") or "").strip(),
                "std_field": r.get("std_field", None),
                "agg": str(r.get("agg", "") or "").strip().upper(),
                "dropna": int(r.get("dropna", 1) if pd.notna(r.get("dropna", 1)) else 1),
                "drop0": int(r.get("drop0", 1) if pd.notna(r.get("drop0", 1)) else 1),
                "unit": str(r.get("unit", "") or "").strip().upper(),
                "decimals": int(r.get("decimals", 2) if pd.notna(r.get("decimals", 2)) else 2),
            })
        # 去除空 display_name
        out = [x for x in out if x["display_name"]]
        return out

    cfg["kpi_by_tech"]["4G"] = _rows("4G")
    cfg["kpi_by_tech"]["5G"] = _rows("5G")

    base_cols = ["指标时间", "活动名称", "区域", "厂家"]
    for tech in ["4G", "5G"]:
        kpi_cols = [r["display_name"] for r in cfg["kpi_by_tech"][tech]]
        # 避免重复
        kpi_cols = [c for i, c in enumerate(kpi_cols) if c and c not in kpi_cols[:i] and c not in base_cols]
        cfg["output_cols_by_tech"][tech] = base_cols + kpi_cols
        cfg["percent_rows_by_tech"][tech] = [r["display_name"] for r in cfg["kpi_by_tech"][tech] if r.get("unit") == "PERCENT"]

    

    # ================= 读取 Thresholds（质差门限规则） =================
    # 支持字段：poor_type, kpi_id, std_field, op, th1, th2, tech, enabled, priority
    # 可选复合条件：and_kpi_id, and_std_field, and_op, and_th1, and_th2
    try:
        df_th = pd.read_excel(project_config_path, sheet_name="Thresholds")
    except Exception:
        df_th = None

    rules = {"4G": [], "5G": [], "BOTH": []}
    if df_th is not None and not df_th.empty:
        df_th = df_th.copy()
        df_th.columns = [str(c).strip() for c in df_th.columns]
        # 兼容中文列名
        rename_map = {
            "质差类型": "poor_type",
            "指标ID": "kpi_id",
            "指标": "kpi_id",
            "比较符": "op",
            "阈值1": "th1",
            "阈值2": "th2",
            "制式": "tech",
            "启用": "enabled",
            "优先级": "priority",
            "且指标ID": "and_kpi_id",
            "且标准字段": "and_std_field",
            "且比较符": "and_op",
            "且阈值1": "and_th1",
            "且阈值2": "and_th2",
            "标准字段": "std_field",
            "STD字段": "std_field",
        }
        for k, v in rename_map.items():
            if k in df_th.columns and v not in df_th.columns:
                df_th.rename(columns={k: v}, inplace=True)

        if "enabled" in df_th.columns:
            df_th["enabled"] = pd.to_numeric(df_th["enabled"], errors="coerce").fillna(0).astype(int)
            df_th = df_th[df_th["enabled"] == 1]

        if not df_th.empty:
            if "priority" in df_th.columns:
                df_th["priority"] = pd.to_numeric(df_th["priority"], errors="coerce").fillna(100).astype(int)
            else:
                df_th["priority"] = 100

            if "tech" not in df_th.columns:
                df_th["tech"] = "BOTH"
            df_th["tech"] = df_th["tech"].astype(str).str.upper().replace({"LTE": "4G", "NR": "5G"}).fillna("BOTH")

            for _, r in df_th.sort_values(["tech", "priority"]).iterrows():
                rule = {
                    "poor_type": str(r.get("poor_type", "") or "").strip(),
                    "kpi_id": str(r.get("kpi_id", "") or "").strip(),
                    "std_field": str(r.get("std_field", "") or "").strip(),
                    "op": str(r.get("op", "") or "").strip().upper(),
                    "th1": r.get("th1", np.nan),
                    "th2": r.get("th2", np.nan),
                    "priority": int(r.get("priority", 100)),
                    "and_kpi_id": str(r.get("and_kpi_id", "") or "").strip(),
                    "and_std_field": str(r.get("and_std_field", "") or "").strip(),
                    "and_op": str(r.get("and_op", "") or "").strip().upper(),
                    "and_th1": r.get("and_th1", np.nan),
                    "and_th2": r.get("and_th2", np.nan),
                }
                t = str(r.get("tech", "BOTH") or "BOTH").upper()
                if t not in ("4G", "5G"):
                    t = "BOTH"
                if rule["poor_type"] and (rule["kpi_id"] or rule["std_field"]):
                    rules[t].append(rule)
    cfg["threshold_rules"] = rules

    # ================= 读取 VendorMap（字段映射候选列名） =================
    try:
        df_vm = pd.read_excel(project_config_path, sheet_name="VendorMap")
    except Exception:
        df_vm = None

    vendor_map = {}
    if df_vm is not None and not df_vm.empty:
        df_vm = df_vm.copy()
        df_vm.columns = [str(c).strip() for c in df_vm.columns]
        rename_map2 = {
            "标准字段": "std_field",
            "STD字段": "std_field",
            "候选列名": "src_candidates",
            "候选列": "src_candidates",
            "排除列名": "exclude_candidates",
            "排除列": "exclude_candidates",
            "启用": "enabled",
        }
        for k, v in rename_map2.items():
            if k in df_vm.columns and v not in df_vm.columns:
                df_vm.rename(columns={k: v}, inplace=True)

        if "enabled" in df_vm.columns:
            df_vm["enabled"] = pd.to_numeric(df_vm["enabled"], errors="coerce").fillna(1).astype(int)
            df_vm = df_vm[df_vm["enabled"] == 1]

        for _, r in df_vm.iterrows():
            std_field = str(r.get("std_field", "") or "").strip()
            if not std_field:
                continue
            cands = str(r.get("src_candidates", "") or "").strip()
            excls = str(r.get("exclude_candidates", "") or "").strip()
            cand_list = [x.strip() for x in cands.split(";") if x.strip()] if cands else []
            excl_list = [x.strip() for x in excls.split(";") if x.strip()] if excls else []
            vendor_map.setdefault(std_field, {"cands": [], "excls": []})
            for x in cand_list:
                if x not in vendor_map[std_field]["cands"]:
                    vendor_map[std_field]["cands"].append(x)
            for x in excl_list:
                if x not in vendor_map[std_field]["excls"]:
                    vendor_map[std_field]["excls"].append(x)
    cfg["vendor_map"] = vendor_map

    # ================= OutputLayout（可选） =================
    try:
        df_ol = pd.read_excel(project_config_path, sheet_name="OutputLayout")
    except Exception:
        df_ol = None
    cfg["output_layout"] = df_ol if (df_ol is not None and not df_ol.empty) else None

    # ================= 构建 kpi_id ↔ std_field/display_name 映射（供门限/对账使用） =================
    kpi_id_to_std = {}
    kpi_id_to_display = {}
    display_to_kpi_id = {}
    for tech in ["4G", "5G"]:
        for r in cfg.get("kpi_by_tech", {}).get(tech, []):
            kid = str(r.get("kpi_id", "") or "").strip()
            dname = str(r.get("display_name", "") or "").strip()
            sf = str(r.get("std_field", "") or "").strip()
            if kid:
                if sf and kid not in kpi_id_to_std:
                    kpi_id_to_std[kid] = sf
                if dname and kid not in kpi_id_to_display:
                    kpi_id_to_display[kid] = dname
                if dname and dname not in display_to_kpi_id:
                    display_to_kpi_id[dname] = kid
    cfg["kpi_id_to_std"] = kpi_id_to_std
    cfg["kpi_id_to_display"] = kpi_id_to_display
    cfg["display_to_kpi_id"] = display_to_kpi_id


    return cfg

def apply_vendor_map_to_global_candidates():
    """将项目配置 VendorMap 中的候选列名注入到 C.GLOBAL_CANDIDATES。

    说明：
    - C.GLOBAL_CANDIDATES 的 key 形如 'kpi_connect'（不含 STD__/RAW__/SRC__ 前缀）
    - VendorMap 的 std_field 可填 'STD__kpi_connect' 或 'kpi_connect'
    - 注入后会去重追加，保证原有兜底候选仍有效
    """
    try:
        vm = PROJECT_CFG.get("vendor_map", {}) if isinstance(PROJECT_CFG, dict) else {}
    except Exception:
        vm = {}
    if not vm:
        return
    for std_field, cfgv in vm.items():
        if not std_field:
            continue
        core = str(std_field).strip()
        if core.startswith("STD__"):
            core = core.replace("STD__", "", 1)
        if core.startswith("RAW__"):
            core = core.replace("RAW__", "", 1)
        if core.startswith("SRC__"):
            core = core.replace("SRC__", "", 1)
        cands = cfgv.get("cands", []) if isinstance(cfgv, dict) else []
        excls = cfgv.get("excls", []) if isinstance(cfgv, dict) else []
        if core not in C.GLOBAL_CANDIDATES:
            C.GLOBAL_CANDIDATES[core] = []
        # 保持原有候选优先级：新候选追加到末尾
        for x in cands:
            if x and x not in C.GLOBAL_CANDIDATES[core]:
                C.GLOBAL_CANDIDATES[core].append(x)
        # exclusions 目前仅用于 auto_match_column_safe 的第二参数形式（(cands, exclusions)）
        # 为不破坏既有结构，这里若 exclusions 有值，追加到末尾的 exclusions 列表中（如 C.GLOBAL_CANDIDATES[core] 已为 tuple 则合并）
        if excls:
            # 若原本是 (cands, exclusions)
            if isinstance(C.GLOBAL_CANDIDATES.get(core), tuple) and len(C.GLOBAL_CANDIDATES[core]) == 2:
                old_c, old_e = C.GLOBAL_CANDIDATES[core]
                merged_e = list(old_e) if isinstance(old_e, (list, tuple)) else []
                for e in excls:
                    if e and e not in merged_e:
                        merged_e.append(e)
                C.GLOBAL_CANDIDATES[core] = (list(old_c), merged_e)
            else:
                # 仅记录在 PROJECT_CFG 中，匹配时由 auto_match_column_safe 读取 vendor_map 更精确（后续可升级）
                pass

def get_output_cols(tech_type: str) -> List[str]:
    """封装 PROJECT_CFG 访问：获取输出列（兼容旧版逻辑）。"""
    return PROJECT_CFG.get("output", {}).get(tech_type, []) or (C.C.COLS_STD_4G if tech_type == "4G" else C.C.COLS_STD_5G)

def get_percent_rows() -> List[str]:
    """封装 PROJECT_CFG 访问：获取百分比行列表（兼容旧版逻辑）。"""
    return PROJECT_CFG.get("percent_rows", []) or []

def _read_ini_safely(cp, ini_path: str) -> bool:
    # Read INI with UTF-8 BOM tolerance.
    # PowerShell 5.x `Set-Content -Encoding utf8` writes UTF-8 with BOM by default, which can
    # break configparser section header parsing (line starts with [section] (UTF-8 BOM removed)).
    if not ini_path:
        return False
    try:
        if not os.path.exists(ini_path):
            return False
    except Exception:
        return False

    # First try: utf-8-sig (will strip BOM if present)
    try:
        cp.read(ini_path, encoding="utf-8-sig")
        return True
    except Exception:
        pass

    # Fallback: read bytes, strip UTF-8 BOM manually, then parse from string
    try:
        with open(ini_path, "rb") as f:
            raw = f.read()
        if raw.startswith(b"\xef\xbb\xbf"):
            raw = raw[3:]
        txt = raw.decode("utf-8", errors="replace")
        cp.read_string(txt)
        return True
    except Exception:
        return False

def load_config(config_path: str):
    cfg = {
        # 自检容差（可按项目口径调整）
        "selftest_tol_pct": 0.05,     # 百分比类容差（±0.05）
        "selftest_tol_drop": 0.005,   # 掉线率容差（±0.005）
        "selftest_tol_erl": 0.05,     # Erl容差（±0.05）
        "selftest_tol_dbm": 0.2,      # dBm容差（±0.2）
        "selftest_tol_users": 2.0,    # 用户数容差（±2）
    }
    cp = configparser.ConfigParser()
    if not config_path:
        config_path = os.path.join(C.BASE_DIR, "config.ini")
    if os.path.exists(config_path):
        _read_ini_safely(cp, config_path)
        sec = cp["selftest"] if "selftest" in cp else {}
        for k in list(cfg.keys()):
            if k in sec:
                try:
                    cfg[k] = float(sec.get(k))
                except Exception:
                    pass
    return cfg, config_path
