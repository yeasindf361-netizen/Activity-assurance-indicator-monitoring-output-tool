# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime
import logging
import os
import sys
from typing import Optional

# Ensure stdout/stderr are configured for UTF-8 to avoid GBK console encoding errors
# This does not change any printed text; it only sets the encoding if the runtime supports it.
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

_LOGGING_CONFIGURED: bool = False
LOG_FILENAME: Optional[str] = None
LOG_FILE_PATH: Optional[str] = None

def configure_logging(log_dir: str) -> None:
    """按旧版 backend_logic.py 方式配置日志输出（仅首次生效）。"""
    global _LOGGING_CONFIGURED, LOG_FILENAME, LOG_FILE_PATH
    if _LOGGING_CONFIGURED:
        return
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    LOG_FILENAME = f"run_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    LOG_FILE_PATH = os.path.join(log_dir, LOG_FILENAME)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.FileHandler(LOG_FILE_PATH, encoding='utf-8')]
    )
    _LOGGING_CONFIGURED = True

def log_print(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    if level == "HEADER":
        print(f"\n{'='*60}\n  {msg}\n{'='*60}")
        logging.info(f"=== {msg} ===")
    elif level == "SUB":
        print(f"[{ts}] 🔹 {msg}")
        logging.info(msg)
    elif level == "WARN":
        print(f"[{ts}] ⚠️ {msg}")
        logging.warning(msg)
    elif level == "SUCCESS":
        print(f"[{ts}] ✅ {msg}")
        logging.info(msg)
    else:
        print(f"[{ts}] {msg}")
        logging.info(msg)

def safe_log(msg: str, level: str = "WARN") -> None:
    """统一的安全日志输出：log_print 失败时静默。"""
    try:
        log_print(msg, level=level)
    except Exception:
        pass
