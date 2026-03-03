# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from typing import Any, Dict

# ================= 版本信息 =================

APP_VERSION = "V16.2"  # 应用版本号（集中管理）

# ================= 路径配置（由 init_base_dir() 初始化） =================

BASE_DIR: str = ""
LIST_FILE_PATH: str = ""
RAW_DIR_4G: str = ""
RAW_DIR_5G: str = ""
OUTPUT_DIR: str = ""
LOG_DIR: str = ""

def init_base_dir(base_dir: str) -> None:
    """初始化 BASE_DIR 及其派生路径（与旧版 backend_logic.py 的默认值保持一致）。"""
    global BASE_DIR, LIST_FILE_PATH, RAW_DIR_4G, RAW_DIR_5G, OUTPUT_DIR, LOG_DIR
    BASE_DIR = base_dir
    LIST_FILE_PATH = os.path.join(BASE_DIR, "保障小区清单", "保障小区清单.xlsx")
    RAW_DIR_4G = os.path.join(BASE_DIR, "网管指标", "4G指标")
    RAW_DIR_5G = os.path.join(BASE_DIR, "网管指标", "5G指标")
    OUTPUT_DIR = os.path.join(BASE_DIR, "输出结果")
    LOG_DIR = os.path.join(BASE_DIR, "logs")

# ================= 业务常量 =================

FUZZY_THRESHOLD_NAME = 0.85
FUZZY_THRESHOLD_FALLBACK = 0.80
FUZZY_THRESHOLD_4G = 0.85
FUZZY_THRESHOLD_5G = 0.85
PREFIX_STRIP_LEN = 6
SUBSTRING_MATCH_MIN_RATIO = 0.7  # 子串匹配最小长度比例（要求子串长度至少占较短字符串的70%）

PROJECT_DEFAULT_CONFIG_DIR = r"E:\Tool_Build\配置文件"

COLS_STD_4G = [
    '指标时间', '活动名称', '区域', '厂家', 
    '总用户数', '总流量(GB)',
    '无线接通率(%)', '无线掉线率(%)', '系统内切换出成功率(%)',
    'VoLTE无线接通率(%)', 'VoLTE切换成功率(%)', 'E-RAB掉话率(QCI=1)(%)', 'VoLTE话务量(Erl)',
    '平均干扰(dBm)', '4G利用率最大值(%)',
    '最大利用率小区', '最大利用率小区的用户数', '最大利用率小区的利用率', '高负荷小区数', '质差小区数'
]

COLS_STD_5G = [
    '指标时间', '活动名称', '区域', '厂家', 
    '总用户数', '总流量(GB)',
    '无线接通率(%)', '无线掉线率(%)', '系统内切换出成功率(%)',
    'VoNR无线接通率(%)', 'VoNR到VoLTE切换成功率(%)', 'VoNR掉线率(5QI1)(%)', 'VoNR话务量(Erl)',
    '平均干扰(dBm)', '5G利用率最大值(%)',
    '最大利用率小区', '最大利用率小区的用户数', '最大利用率小区的利用率', '高负荷小区数', '质差小区数'
]

# FieldStandardizer 使用的候选列名字典（保持原样）
GLOBAL_CANDIDATES: Dict[str, Any] = {
            'time_start': ['开始时间', 'Start Time', 'TIME_START', '时间', 'DAY'],
            'time_end': ['结束时间', 'End Time', 'TIME_END'],
            'time_any': ['DAY', '时间', '日期', '统计时间', '指标时间'],
            'cell_name': ['小区名称', '小区中文名', 'CellName', 'NRCellName', 'DU物理小区名称', 'CEL_NAME', 'DU物理小区名'],
            'kpi_connect': (['无线接通率', '无线接通率YC', '小区无线接通率', 'RRC连接成功率'], ['VoNR', 'VoLTE', 'QCI']),
            'kpi_drop': (['无线掉线率', '掉话率', '无线掉线率（小区）', '无线掉线率(%)'], ['VoNR', 'VoLTE', 'QCI']),
            'kpi_ho_intra': (['系统内切换出成功率', '切换成功率', '切换成功率gk', 'eNB内切换成功率',
                  '切换成功率-新', '切换成功率LTE', '切换成功率QQ', '切换成功率ZB'],
                 ['VoNR', 'VoLTE']),
            'kpi_traffic_gb': ['总流量', '总流量(GB)', '5G总流量(GB)-XJ', '总流量（GB）',
                   '上下行吞吐量MB', '总吞吐量MB', '上下行吞吐量(MB)'],
            'kpi_rrc_users_max': ['RRC最大连接数', 'RRC连接建立最大用户数', '最大激活用户数', '小区内处于RRC连接态的最大用户数', '小区内的最大用户数', 'RRC连接最大连接用户数', 'VOLTE语音最大用户数'],
            'kpi_util_max': ['4G利用率最大值', '5G利用率最大值', '无线利用率', '无线资源利用率'],
            'kpi_prb_ul_util': ['上行PRB平均利用率', 'UL PRB Util', '上行共享信道PRB利用率', '上行PRB利用率'],
            'kpi_prb_dl_util': ['下行PRB平均利用率', 'DL PRB Util', '下行共享信道PRB利用率', '下行PRB利用率'],
            'kpi_ul_interf_dbm': ['小区上行平均干扰电平(dBm)', '上行底噪', '上行每PRB的接收干扰噪声平均值', '上行每PRB的接收干扰噪声平均值(dBm)', '系统上行每个PRB上检测到的干扰噪声的平均值(dBm)', '载波平均噪声干扰'],
            'kpi_vonr_connect': ['VoNR无线接通率', 'VONR接通率', 'VoNR无线接通率(5QI1)'],
            'kpi_vonr_drop': ['VoNR掉线率', 'VONR掉话率', '掉线率(5QI1)(小区级)'],
            'kpi_vonr_traffic_erl': ['VoNR话务量', 'vonr话务量', '5QI话务量(erl)_t'],
            'kpi_nr2lte_ho': ['VoNR到VoLTE切换成功率', 'NR到LTE的系统间切换出成功率', '系统间切换成功率（NG-RAN->EUTRAN）（5QI1）'],
            'kpi_volte_connect': ['VoLTE无线接通率', 'VOLTE接通率', '无线接通率(QCI=1)'],
            'kpi_volte_drop': ['E-RAB掉话率(QCI=1)', 'QCI1掉线率小区级', 'VOLTE掉话率'],
            'kpi_volte_traffic_erl': ['VoLTE话务量', 'QCI1的平均E-RAB数(话务量)', 'volte话务量(Erlang)', 'VoLTE语音话务量'],
            'kpi_volte_ho': ['VoLTE切换成功率', 'VOLTE切换成功率'],
            # 新增：成功率分子/分母（用于加权平均）
            'kpi_connect_num': ['无线接通率分子', 'RRC连接建立完成次数', 'RRC连接成功次数', 'RRC连接建立成功次数'],
            'kpi_connect_den': ['无线接通率分母', 'RRC连接请求次数（包括重发）', 'RRC连接建立请求次数'],
            'kpi_volte_connect_num': ['QCI为1的业务E-RAB建立成功次数', 'QCI1建立成功次数'],
            'kpi_volte_connect_den': ['QCI为1的业务E-RAB建立尝试次数', 'QCI1建立申请次数']
}
# ----------------------------------------------------------------------
# Backwards-compatibility shim: keep historical attribute access constants.C.*
class _CompatC:
    pass

C = _CompatC()

# If these symbols exist in this module, expose them through C as well.
try:
    C.COLS_STD_4G = COLS_STD_4G
except Exception:
    pass

try:
    C.COLS_STD_5G = COLS_STD_5G
except Exception:
    pass
