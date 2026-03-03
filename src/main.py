# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import datetime
import glob
import os
import sys
import time
import traceback
import warnings
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
from kpi_tool.core.standardizer import auto_match_column_safe

from kpi_tool.config import constants as C
from kpi_tool.config import project_config as PC
from kpi_tool.config.logging_config import log_print
from kpi_tool.core import time_handler as TH
from kpi_tool.io import data_loader as DL
from kpi_tool.core import matching as MATCH
from kpi_tool.core import calculator as CALC
from kpi_tool.io import excel_writer as EX
from kpi_tool.io import text_report as TR

import selftest as ST
import archive as ARCH

warnings.filterwarnings('ignore')  # 与旧版保持一致：屏蔽警告

args = None  # 与旧版一致（parse_args 会写入）

def parse_args():
    parser = argparse.ArgumentParser(description="活动保障指标监控工具（支持自检/按活动/按制式运行）")
    parser.add_argument("--only_activity", "--activity", dest="only_activity", default="", help="仅运行指定活动名称（与保障清单“活动名称”完全一致）")
    parser.add_argument("--only_tech", "--tech", dest="only_tech", default="", choices=["", "4G", "5G"], help="仅运行指定制式：4G 或 5G")
    parser.add_argument("--selftest", action="store_true", help="执行自检（用于快速对账），自检时默认仍会产出输出文件，便于排查")
    parser.add_argument("--config", default="", help="配置文件路径（默认读取脚本目录下 config.ini）")
    parser.add_argument("--config_dir", default="", help="商用配置化：配置文件目录（默认 E:\\Tool_Build\\配置文件）")
    parser.add_argument("--project_config", default="", help="商用配置化：项目配置Excel路径（默认在配置目录下找 项目配置.xlsx；否则回退脚本目录的 项目配置_模板_30指标.xlsx）")
    a = parser.parse_args()
    global args
    args = a
    return a

def main(target_activities: list[str] | None = None, log_info_cb=None, progress_cb=None, base_dir: str | None = None,
         list_file_path: str | None = None, raw_dir_4g: str | None = None, raw_dir_5g: str | None = None,
         output_dir: str | None = None, log_dir: str | None = None):
    return _main_impl(target_activities=target_activities, log_info_cb=log_info_cb, progress_cb=progress_cb,
                     base_dir=base_dir, list_file_path=list_file_path, raw_dir_4g=raw_dir_4g, raw_dir_5g=raw_dir_5g,
                     output_dir=output_dir, log_dir=log_dir)


def _main_impl(target_activities=None, log_info_cb=None, progress_cb=None, base_dir: str | None = None, list_file_path: str | None = None, raw_dir_4g: str | None = None, raw_dir_5g: str | None = None, output_dir: str | None = None, log_dir: str | None = None):
    # 路径覆盖（用于 GUI/外部调用；不改变旧版默认值行为）
    if base_dir:
        C.init_base_dir(base_dir)
    if output_dir:
        C.OUTPUT_DIR = output_dir
    if log_dir:
        C.LOG_DIR = log_dir
    if list_file_path:
        C.LIST_FILE_PATH = list_file_path
    if raw_dir_4g:
        C.RAW_DIR_4G = raw_dir_4g
    if raw_dir_5g:
        C.RAW_DIR_5G = raw_dir_5g

    args = parse_args()
    cfg, cfg_path = PC.load_config(getattr(args, 'config', None))
    PC.CFG = cfg

    # ===== 商用配置化：加载项目配置（指标清单/门限/字段映射/输出布局）=====
    try:
        # 解析配置目录（CLI 优先）
        cfg_dir = getattr(args, "config_dir", "") or PC.PROJECT_DEFAULT_CONFIG_DIR
        PC.PROJECT_CFG["config_dir"] = cfg_dir
        proj_path = PC._pick_project_config_path(args, cfg_path)
        if proj_path:
            PC.PROJECT_CFG = PC.load_project_config_excel(proj_path)
            PC.PROJECT_CFG["config_dir"] = cfg_dir
            log_print(f"项目配置: {proj_path}", "INFO")
        else:
            log_print("项目配置: 未启用（将使用脚本内置口径）", "INFO")
    except Exception as e:
        log_print(f"项目配置加载失败，回退内置口径: {e}", "WARN")

    # 注入 VendorMap 候选列名（在读取原始KPI前执行）
    try:
        PC.apply_vendor_map_to_global_candidates()
    except Exception:
        pass

    log_print(f"配置文件: {cfg_path}", "INFO")
    os.system('cls' if os.name == 'nt' else 'clear')
    log_print("活动保障指标监控工具 V16.2 (商用配置化：指标/门限/映射)", "HEADER")
    log_print(f"输出目录: {C.OUTPUT_DIR}", "SUB")
    
    log_print("正在读取保障小区清单...", "SUB")
    df_list = DL.load_data_frame(C.LIST_FILE_PATH)
    if df_list is None:
        log_print(f"❌ 错误：找不到清单文件 {C.LIST_FILE_PATH}", "WARN")
        return

    
    # —— 微信简报格式判定：保障清单是否存在‘区域’列（仅用于TXT排版，不影响计算）——
    region_col_raw = auto_match_column_safe(df_list.columns, ['区域', '场景'])
    has_region_col = region_col_raw is not None
    if region_col_raw and region_col_raw != '区域':
        df_list.rename(columns={region_col_raw: '区域'}, inplace=True)
    if not has_region_col and '区域' not in df_list.columns:
        df_list['区域'] = '整体'

    act_col = auto_match_column_safe(df_list.columns, ['活动名称', '保障活动', '活动'])
    if act_col:
        df_list.rename(columns={act_col: '活动名称'}, inplace=True)
    else:
        df_list['活动名称'] = "未知活动"

    # 清单原始列（用于导出“识别到的小区明细”，保持与《保障小区清单》一致）
    list_export_cols = list(df_list.columns)
    # 为对账输出保留清单行索引（后续匹配成功后可回溯到清单明细）
    df_list['_list_idx'] = np.arange(len(df_list), dtype=int)

    activities = df_list['活动名称'].dropna().unique().tolist()
    log_print(f"识别到 {len(activities)} 个保障活动: {', '.join(activities[:3])}...", "SUCCESS")

    log_print("正在加载 4G/5G 网管原始指标 (多线程)...", "SUB")
    raw_4g, _, time_4g = ([], None, None)
    raw_5g, _, time_5g = ([], None, None)
    if getattr(args, 'only_tech', '') in ('', '4G'):
        raw_4g, _, time_4g = DL.load_and_distribute(C.RAW_DIR_4G, '4G')
    if getattr(args, 'only_tech', '') in ('', '5G'):
        raw_5g, _, time_5g = DL.load_and_distribute(C.RAW_DIR_5G, '5G')
    # 若 4G/5G 均未加载到网管数据，则本次匹配将为 0（仅提示，不中断）
    try:
        if (raw_4g is None or (hasattr(raw_4g, "empty") and raw_4g.empty)) and (raw_5g is None or (hasattr(raw_5g, "empty") and raw_5g.empty)):
            log_print("⚠️ 未加载到任何网管指标数据（4G/5G 均为空）。请确认运行前网管目录已放入新文件；上次运行结束后目录会被自动清空。", "WARN")
    except Exception:
        pass

    t_window = time_5g if time_5g else (time_4g if time_4g else "未知时间")
    log_print(f"指标时间窗: {t_window}", "SUCCESS")

    log_print("执行三级火箭匹配 (名称 > ID > 模糊)...", "SUB")

    # 若 GUI 为部分活动指定了时间段，则在匹配前应用选择（未指定活动仍使用默认“每小区最新时间段”）
    if TH._SELECTED_TIME_SEGMENTS:
        raw_4g = _apply_selected_time_segments_to_raw(raw_4g, df_list, '4G')
        raw_5g = _apply_selected_time_segments_to_raw(raw_5g, df_list, '5G')
    else:
        log_print("未手动选择时间段：使用默认“每小区最新时间段”模式（最大化覆盖）", "INFO")

    merged_4g = MATCH.waterfall_merge(raw_4g, df_list, '4G')
    merged_5g = MATCH.waterfall_merge(raw_5g, df_list, '5G')
    
    print(f"\n{'='*60}\n  📊 匹配结果统计 (时间: {t_window})\n{'='*60}")
    print(f"{'活动名称':<20} | {'4G匹配':<8} | {'5G匹配':<8} | {'高负荷':<6} | {'质差':<6}")
    print("-" * 60)
    
    res_4g_list = CALC.CALC.calculate_kpis('4G', merged_4g)
    res_5g_list = CALC.CALC.calculate_kpis('5G', merged_5g)
    poor_4g = CALC.CALC.get_poor_quality('4G', merged_4g)
    poor_5g = CALC.CALC.get_poor_quality('5G', merged_5g)
    
    summary_map = {act: {'4g':0, '5g':0, 'load':0, 'poor':0} for act in activities}
    
    if not merged_4g.empty:
        for act, c in merged_4g.groupby('活动名称')['_list_idx'].nunique().items(): 
            if act in summary_map: summary_map[act]['4g'] = c
    if not merged_5g.empty:
        for act, c in merged_5g.groupby('活动名称')['_list_idx'].nunique().items(): 
            if act in summary_map: summary_map[act]['5g'] = c
            
    for res in res_4g_list + res_5g_list:
        if not res[1].empty:
            act = res[0]['活动名称'].iloc[0]
            if act in summary_map: summary_map[act]['load'] += len(res[1])
            
    if not poor_4g.empty:

            
        if '_list_idx' in poor_4g.columns:

            
            ser = poor_4g.dropna(subset=['_list_idx']).groupby('活动名称')['_list_idx'].nunique()

            
        else:

            
            ser = poor_4g.groupby('活动名称').size()

            
        for act, c in ser.items():

            
            if act in summary_map: summary_map[act]['poor'] += int(c)

            
    if not poor_5g.empty:

            
        if '_list_idx' in poor_5g.columns:

            
            ser = poor_5g.dropna(subset=['_list_idx']).groupby('活动名称')['_list_idx'].nunique()

            
        else:

            
            ser = poor_5g.groupby('活动名称').size()

            
        for act, c in ser.items():

            
            if act in summary_map: summary_map[act]['poor'] += int(c)
            
    for act in activities:
        d = summary_map.get(act, {})
        print(f"{act[:18]:<20} | {d.get('4g',0):<8} | {d.get('5g',0):<8} | {d.get('load',0):<6} | {d.get('poor',0):<6}")
    print("-" * 60)
    
    log_print("正在生成最终报表...", "SUB")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(C.OUTPUT_DIR, f"通报结果_{ts}.xlsx")
    txt_path = os.path.join(C.OUTPUT_DIR, f"微信通报简报_{ts}.txt")
    
    # 确保输出目录存在（打包为 onefile/Release 后首次运行也能写入）

    
    os.makedirs(C.OUTPUT_DIR, exist_ok=True)

    
    

    
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        if res_4g_list:
            df = pd.concat([r[0] for r in res_4g_list], ignore_index=True)
            EX.safe_write_excel(writer, df, "4G指标监控")
        if res_5g_list:
            df = pd.concat([r[0] for r in res_5g_list], ignore_index=True)
            EX.safe_write_excel(writer, df, "5G指标监控")
        poor_all = pd.concat([poor_4g, poor_5g], ignore_index=True)
        if not poor_all.empty and '_list_idx' in poor_all.columns:
            poor_all = poor_all.drop(columns=['_list_idx'])
        if not poor_all.empty and '匹配方式' in poor_all.columns:
            poor_all = poor_all.drop(columns=['匹配方式'])
        poor_all.to_excel(writer, "质差小区明细", index=False)
        # 导出每个活动“实际识别到的小区明细”（对账核对用）
        EX.export_recognized_cell_details(writer, df_list, merged_4g, merged_5g, activities, list_export_cols)
        # 导出指标计算过程（逐项对账用：公式+逐小区原始值+中间量）
        EX.export_kpi_calc_formulas(writer)
        EX.export_kpi_calc_details(writer, "4G", merged_4g)
        EX.export_kpi_calc_details(writer, "5G", merged_5g)

    
    EX.style_excel(out_path)
    
    txt_content = TR.generate_text_report(res_4g_list, res_5g_list, has_region_col=has_region_col)
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(txt_content)


    # 自检：将当前输出结果与预期值做快速对账（如需关闭，运行时不加 --selftest 即可）
    if getattr(args, 'selftest', False):
        ok, report = ST.run_selftest(res_4g_list, res_5g_list, cfg)
        report_path = os.path.join(C.OUTPUT_DIR, f"运行自检报告_{ts}.txt")
        with open(report_path, 'w', encoding='utf-8') as ff:
            ff.write(report)
        log_print('自检完成: ' + ('PASS' if ok else 'FAIL'), 'SUCCESS' if ok else 'WARN')
        log_print(f'自检报告: {report_path}', 'INFO')
        
    log_print(f"✅ 处理完成！", "SUCCESS")
    log_print(f"   📄 Excel: {out_path}", "INFO")
    log_print(f"   📝 简报: {txt_path}", "INFO")


    # === 历史数据归档（调整后逻辑：归档上一次输出 + 本次网管原始数据；输出目录保留本次最新） ===
    try:
        _keep_files = [out_path, txt_path]
        if 'report_path' in locals() and report_path:
            _keep_files.append(report_path)

        arch = ARCH.archive_run_data(ts, _keep_files, base_dir=C.BASE_DIR)

        # 归档明细（仅日志，不影响任何既有统计/输出）
        _cnt = {}
        try:
            if isinstance(arch, dict):
                _cnt = arch.get("counts", {}) or {}
        except Exception:
            _cnt = {}

        log_print("📦 已归档：上一次输出文件 + 本次网管原始数据；“输出结果”仅保留本次最新 Excel/TXT", "INFO")
        log_print(f"   📄 Excel(最新): {out_path}", "INFO")
        log_print(f"   📝 简报(最新): {txt_path}", "INFO")
        if 'report_path' in locals() and report_path:
            log_print(f"   🧪 自检报告(最新): {report_path}", "INFO")

        if _cnt:
            try:
                log_print(
                    f"   📦 归档明细: 上次输出移动 {_cnt.get('moved_prev_output', 0)} | 本次输出复制 {_cnt.get('copied_current_output', 0)} | 4G原始移动 {_cnt.get('moved_raw_4g', 0)} | 5G原始移动 {_cnt.get('moved_raw_5g', 0)}",
                    "INFO"
                )
                if _cnt.get('moved_raw_4g', 0) == 0 and _cnt.get('moved_raw_5g', 0) == 0:
                    log_print(
                        "   ⚠️ 本次未归档到任何网管原始文件（4G/5G均为0）。若本次匹配为0，请确认运行前已将新网管指标文件放入网管目录（上次运行结束后目录会被自动清空）。",
                        "WARN"
                    )
            except Exception:
                pass
    except Exception as e:
        log_print(f"历史数据归档失败（已忽略，不影响本次结果生成）：{e}", "WARN")

    # 运行结束后自动打开输出结果与微信简报（仅新增功能，不影响计算逻辑）
    try:
        if os.name == 'nt':
            os.startfile(out_path)
            os.startfile(txt_path)
    except Exception:
        pass

    os.system('pause' if os.name == 'nt' else 'true')

if __name__ == '__main__':
    # 兼容旧版：直接运行 main.py 时自动推导 BASE_DIR，并初始化常量路径
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    if not C.BASE_DIR:
        C.init_base_dir(base_dir)
    # 配置日志（如未在外部初始化）
    try:
        from kpi_tool.config.logging_config import configure_logging
        configure_logging(C.LOG_DIR)
    except Exception:
        pass
    main()


