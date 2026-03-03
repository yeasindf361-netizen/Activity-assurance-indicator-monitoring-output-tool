# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime
import glob
import os
import shutil
from typing import List

from kpi_tool.config import constants as C
from kpi_tool.config.logging_config import log_print

def archive_run_data(run_ts: str, output_files: list, base_dir: str = None) -> dict:
    """
    运行结束后自动归档历史数据（调整后逻辑）：

    目标：
      1) “输出结果”目录始终只保留本次最新生成的 Excel/TXT（output_files），便于快速查看最新结果；
         目录中的其它历史输出文件（如上一次运行遗留的 Excel/TXT/报告等）将被移动归档。
      2) 网管指标目录（4G/5G）在本次运行成功后会被移动归档并清空，便于下次放入新数据。
      3) 历史数据目录完整保存每次运行的完整数据（本次输出 + 本次网管原始文件）：
         - 本次输出文件：复制到历史目录（不影响“输出结果”目录保留最新文件）
         - 本次网管原始文件：移动到历史目录（确保输入目录清空）

    失败仅告警，不中断程序。

    返回：
      {
        "moved_map": {src: dst, ...},   # 仅包含被“移动”的文件
        "counts": {
            "moved_prev_output": int,   # 输出目录被移动归档的历史输出数
            "copied_current_output": int,# 本次最新输出复制到历史目录数
            "moved_raw_4g": int,        # 本次归档移动的 4G 原始文件数
            "moved_raw_5g": int         # 本次归档移动的 5G 原始文件数
        },
        "hist_root": "历史数据/<run_ts>"  # 本次归档根目录
      }
    """
    moved_map = {}
    counts = {
        "moved_prev_output": 0,
        "copied_current_output": 0,
        "moved_raw_4g": 0,
        "moved_raw_5g": 0
    }
    hist_root_cur = None

    try:
        if not run_ts:
            return {"moved_map": moved_map, "counts": counts, "hist_root": hist_root_cur}

        base_dir = base_dir or C.BASE_DIR
        hist_base = os.path.join(base_dir, "历史数据")
        try:
            os.makedirs(hist_base, exist_ok=True)
        except Exception:
            pass

        # 本次运行归档目录
        hist_root_cur = os.path.join(hist_base, run_ts)
        dst_out_cur = os.path.join(hist_root_cur, "输出结果")
        dst_4g = os.path.join(hist_root_cur, "4G网管指标")
        dst_5g = os.path.join(hist_root_cur, "5G网管指标")
        for d in [dst_out_cur, dst_4g, dst_5g]:
            try:
                os.makedirs(d, exist_ok=True)
            except Exception:
                pass

        def _abs(p):
            try:
                return os.path.abspath(p)
            except Exception:
                return p

        keep_set = set()
        for fp in (output_files or []):
            if fp:
                keep_set.add(_abs(fp))

        # 尝试从文件名中提取时间戳（优先 YYYYMMDD_HHMMSS，其次 YYYYMMDDHHMMSS）
        def _extract_ts_from_name(name: str) -> str:
            try:
                if not name:
                    return ""
                m = re.search(r"(\d{8}_\d{6})", name)
                if m:
                    return m.group(1)
                m = re.search(r"(\d{14})", name)
                if m:
                    s = m.group(1)
                    return f"{s[:8]}_{s[8:]}"
                return ""
            except Exception:
                return ""

        # 1) 归档“上一次输出文件”：移动 OUTPUT_DIR 中除本次 Excel/TXT 外的历史输出文件
        try:
            if os.path.exists(OUTPUT_DIR):
                for fn in os.listdir(OUTPUT_DIR):
                    try:
                        src_fp = os.path.join(OUTPUT_DIR, fn)
                        if not os.path.isfile(src_fp):
                            continue
                        ext = os.path.splitext(fn)[1].lower()
                        if ext not in [".xlsx", ".xls", ".xlsm", ".txt"]:
                            continue

                        abs_src = _abs(src_fp)
                        if abs_src in keep_set:
                            continue  # 本次最新输出，保留在 OUTPUT_DIR

                        # 归档目录：优先按文件名自带时间戳归档；否则归到“本次运行/输出结果/上一次输出”
                        ts_guess = _extract_ts_from_name(fn)
                        if ts_guess and ts_guess != run_ts:
                            dst_out_prev = os.path.join(hist_base, ts_guess, "输出结果")
                        else:
                            dst_out_prev = os.path.join(dst_out_cur, "上一次输出")
                        try:
                            os.makedirs(dst_out_prev, exist_ok=True)
                        except Exception:
                            pass

                        dst_fp = os.path.join(dst_out_prev, fn)
                        shutil.move(src_fp, dst_fp)
                        moved_map[src_fp] = dst_fp
                        counts["moved_prev_output"] += 1
                    except Exception as e:
                        try:
                            log_print(f"归档上一次输出文件失败: {fn} | {e}", "WARN")
                        except Exception:
                            pass
        except Exception as e:
            try:
                log_print(f"扫描输出目录归档失败: {e}", "WARN")
            except Exception:
                pass

        # 2) 将“本次最新输出文件”复制到本次运行历史目录（不移动，保持 OUTPUT_DIR 仍保留最新）
        for fp in (output_files or []):
            try:
                if not fp:
                    continue
                if os.path.exists(fp):
                    dst_fp = os.path.join(dst_out_cur, os.path.basename(fp))
                    shutil.copy2(fp, dst_fp)
                    counts["copied_current_output"] += 1
            except Exception as e:
                try:
                    log_print(f"复制本次输出到历史目录失败: {fp} | {e}", "WARN")
                except Exception:
                    pass

        # 3) 移动网管原始数据（目录下所有文件）到本次运行历史目录（确保输入目录清空）
        def _move_all_files(src_dir, dst_dir, counter_key):
            try:
                if not src_dir or not os.path.exists(src_dir):
                    return
                for fn in os.listdir(src_dir):
                    src_fp = os.path.join(src_dir, fn)
                    if os.path.isfile(src_fp):
                        try:
                            dst_fp = os.path.join(dst_dir, fn)
                            shutil.move(src_fp, dst_fp)
                            moved_map[src_fp] = dst_fp
                            counts[counter_key] += 1
                        except Exception as ee:
                            try:
                                log_print(f"归档网管文件失败: {src_fp} | {ee}", "WARN")
                            except Exception:
                                pass
            except Exception as e:
                try:
                    log_print(f"归档目录失败: {src_dir} | {e}", "WARN")
                except Exception:
                    pass

        _move_all_files(C.RAW_DIR_4G, dst_4g, "moved_raw_4g")
        _move_all_files(C.RAW_DIR_5G, dst_5g, "moved_raw_5g")

        try:
            log_print(f"历史数据归档完成: {hist_root_cur} | 上次输出移动 {counts['moved_prev_output']} | 本次输出复制 {counts['copied_current_output']} | 4G原始移动 {counts['moved_raw_4g']} | 5G原始移动 {counts['moved_raw_5g']}", "INFO")
        except Exception:
            pass

    except Exception as e:
        try:
            log_print(f"历史数据归档异常: {e}", "WARN")
        except Exception:
            pass

    return {"moved_map": moved_map, "counts": counts, "hist_root": hist_root_cur}
