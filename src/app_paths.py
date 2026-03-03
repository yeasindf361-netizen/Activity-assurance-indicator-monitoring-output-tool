"""统一路径管理模块

所有路径解析逻辑集中在此，消除 backend_logic.py / main_gui.py 中分散的 frozen 检测。

用法：
    from app_paths import APP_DIR, INTERNAL_DIR, DATA_DIR, ...
"""
import os
import sys

# ============================================================
# 核心路径（全局唯一定义）
# ============================================================

def _resolve_app_dir() -> str:
    """应用根目录（exe 所在目录 或 源码目录）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _resolve_internal_dir() -> str:
    """_internal 目录（打包后存放 Python 运行时和捆绑数据）"""
    if getattr(sys, 'frozen', False):
        return os.path.join(APP_DIR, "_internal")
    return APP_DIR  # 开发模式下等同于 APP_DIR


APP_DIR = _resolve_app_dir()
INTERNAL_DIR = _resolve_internal_dir()
IS_FROZEN = getattr(sys, 'frozen', False)

# ============================================================
# 业务路径
# ============================================================

# 数据文件（保障小区清单、配置文件）在 exe 同级
LIST_FILE_PATH = os.path.join(APP_DIR, "保障小区清单", "保障小区清单.xlsx")
CONFIG_DIR = os.path.join(APP_DIR, "配置文件")

# 网管指标目录
RAW_DIR_4G = os.path.join(APP_DIR, "网管指标", "4G指标")
RAW_DIR_5G = os.path.join(APP_DIR, "网管指标", "5G指标")

# 输出目录
OUTPUT_DIR = os.path.join(APP_DIR, "输出结果")
LOG_DIR = os.path.join(APP_DIR, "logs")

# GUI 配置和图标（在 _internal 中，保持 exe 同级整洁）
GUI_CONFIG_FILE = os.path.join(INTERNAL_DIR, "gui_config.json")
ICON_PATH = os.path.join(INTERNAL_DIR, "logo.ico") if IS_FROZEN else os.path.join(APP_DIR, "logo.ico")


def ensure_runtime_dirs():
    """确保运行时目录存在（首次运行时自动创建）"""
    for d in [RAW_DIR_4G, RAW_DIR_5G, OUTPUT_DIR, LOG_DIR]:
        os.makedirs(d, exist_ok=True)
