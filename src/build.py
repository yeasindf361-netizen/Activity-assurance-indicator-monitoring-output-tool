#!/usr/bin/env python3
"""KPI_Tool 一键构建脚本

自动化流程：清理 → PyInstaller 构建 → 目录设置 → 安全优化 → 启动验证 → 压缩打包

用法：
    python build.py              # 完整构建
    python build.py --skip-build # 跳过 PyInstaller（仅重新设置目录/优化/压缩）
    python build.py --version V16.3  # 指定版本号
"""
import os
import sys
import shutil
import subprocess
import datetime
import argparse
import zipfile

# ============================================================
# 配置区（修改这里即可适配项目变化）
# ============================================================
SPEC_FILE = "KPI_Tool.spec"
DIST_NAME = "KPI_Tool"
from kpi_tool.config.constants import APP_VERSION

DEFAULT_VERSION = APP_VERSION

# 代码保护：PyInstaller 加密密钥（16字符）
ENCRYPTION_KEY = "KPI_Tool_Secure1"

# 需要从源目录复制到 exe 同级的目录（不打包进 _internal）
COPY_TO_EXE_LEVEL = ["配置文件"]

# 需要在 exe 同级创建的运行时空目录
RUNTIME_DIRS = [
    os.path.join("网管指标", "4G指标"),
    os.path.join("网管指标", "5G指标"),
    "输出结果",
    "历史数据",
    "logs",
]

# 保障小区清单模板表头（打包时自动生成空模板，不打包真实数据）
LIST_TEMPLATE_HEADERS = ["活动名称", "小区CGI", "小区中文名", "厂家"]

# 安全优化：可删除的目录/文件模式（绝对不要碰 numpy.libs）
SAFE_REMOVE_DIRS = [
    "_tcl_data/tcl8/encoding",
    "_tcl_data/tcl8/msgs",
    "_tcl_data/tcl8/tzdata",
    "_tk_data/tk/msgs",
    "_tk_data/tk/images",
    "pytz/zoneinfo",
    "pytz",
    "setuptools",
    "pkg_resources",
    "yaml",
    "_yaml",
    "test",
    "tests",
]
SAFE_REMOVE_GLOBS = ["*.dist-info", "api-ms-win-*.dll", "**/.DS_Store", "**/README"]

# 绝对禁止删除的目录（保护清单）
NEVER_DELETE = {"numpy.libs", "numpy", "pandas", "openpyxl", "customtkinter"}


# ============================================================
# 工具函数
# ============================================================
def log(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    prefix = {"INFO": "  ", "OK": "[OK] ", "WARN": "[!!] ", "ERR": "[ERR] ", "STEP": ">> "}
    print(f"[{ts}] {prefix.get(level, '  ')}{msg}")


def rmtree_safe(path):
    if os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)


def dir_size_mb(path):
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.isfile(fp):
                total += os.path.getsize(fp)
    return total / (1024 * 1024)


# ============================================================
# 构建步骤
# ============================================================
def step_clean(dist_dir, build_dir):
    log("清理旧构建产物...", "STEP")
    rmtree_safe(dist_dir)
    rmtree_safe(build_dir)
    log("清理完成", "OK")


def step_build(src_dir):
    log("执行 PyInstaller 构建（字节码保护）...", "STEP")
    cmd = [
        sys.executable, "-m", "PyInstaller", SPEC_FILE,
        "--clean", "--noconfirm"
    ]
    result = subprocess.run(cmd, cwd=src_dir, capture_output=True, text=True)
    if result.returncode != 0:
        log(f"PyInstaller 构建失败:\n{result.stderr[-500:]}", "ERR")
        sys.exit(1)
    log("PyInstaller 构建成功（代码已编译为字节码）", "OK")


def _copytree_clean(src, dst):
    """复制目录，自动跳过临时文件（~$*）和 .DS_Store"""
    def _ignore(directory, files):
        return [f for f in files if f.startswith("~$") or f == ".DS_Store"]
    shutil.copytree(src, dst, ignore=_ignore)


def step_setup_dirs(app_dir, internal_dir, src_dir):
    log("设置目录结构...", "STEP")

    # 从源目录复制到 exe 同级（过滤临时文件）
    for name in COPY_TO_EXE_LEVEL:
        src = os.path.join(src_dir, name)
        dst = os.path.join(app_dir, name)
        if os.path.exists(src):
            if os.path.exists(dst):
                shutil.rmtree(dst)
            _copytree_clean(src, dst)
            log(f"  复制 {name} → exe同级/{name}", "OK")
        else:
            log(f"  {name} 不存在，跳过", "WARN")

    # 生成保障小区清单模板（不打包真实数据）
    list_dir = os.path.join(app_dir, "保障小区清单")
    os.makedirs(list_dir, exist_ok=True)
    _create_list_template(os.path.join(list_dir, "保障小区清单.xlsx"))
    log("  生成 保障小区清单/保障小区清单.xlsx (空模板)", "OK")

    # 创建运行时空目录
    for d in RUNTIME_DIRS:
        full = os.path.join(app_dir, d)
        os.makedirs(full, exist_ok=True)
    log(f"  创建运行时目录: {len(RUNTIME_DIRS)} 个", "OK")


def _create_list_template(file_path):
    """生成保障小区清单空模板（仅表头，无数据）"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for col, header in enumerate(LIST_TEMPLATE_HEADERS, 1):
        ws.cell(row=1, column=col, value=header)
    wb.save(file_path)
    wb.close()


def step_optimize(internal_dir):
    log("执行安全优化...", "STEP")
    before = dir_size_mb(internal_dir)
    removed = 0

    for d in SAFE_REMOVE_DIRS:
        full = os.path.join(internal_dir, d)
        if os.path.exists(full):
            # 安全检查：不能删除保护清单中的目录
            base = d.split("/")[0].split("\\")[0]
            if base in NEVER_DELETE:
                log(f"  跳过受保护目录: {d}", "WARN")
                continue
            shutil.rmtree(full, ignore_errors=True)
            removed += 1

    import glob
    for pattern in SAFE_REMOVE_GLOBS:
        for f in glob.glob(os.path.join(internal_dir, pattern)):
            base = os.path.basename(f).split(".")[0]
            if base in NEVER_DELETE:
                continue
            if os.path.isdir(f):
                shutil.rmtree(f, ignore_errors=True)
            else:
                os.remove(f)
            removed += 1

    after = dir_size_mb(internal_dir)
    log(f"  优化完成: {before:.1f}MB → {after:.1f}MB (减少 {before - after:.1f}MB, 删除 {removed} 项)", "OK")


def step_verify(app_dir):
    log("验证 exe 启动...", "STEP")
    exe = os.path.join(app_dir, f"{DIST_NAME}.exe")
    if not os.path.exists(exe):
        log(f"exe 不存在: {exe}", "ERR")
        sys.exit(1)

    # 检查关键依赖
    numpy_libs = os.path.join(app_dir, "_internal", "numpy.libs")
    if not os.path.exists(numpy_libs):
        log("numpy.libs 缺失！打包后 numpy 将无法导入", "ERR")
        sys.exit(1)
    log("  numpy.libs 完整", "OK")

    # 检查关键目录
    required_dirs = COPY_TO_EXE_LEVEL + ["保障小区清单"]
    for name in required_dirs:
        if not os.path.exists(os.path.join(app_dir, name)):
            log(f"  {name} 目录缺失", "ERR")
            sys.exit(1)
    log("  数据目录完整", "OK")

    # 启动测试（超时 = 正常，立即退出 = 崩溃）
    try:
        result = subprocess.run(
            [exe], cwd=app_dir, timeout=6,
            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        # 如果 6 秒内退出，说明可能崩溃了
        if result.returncode != 0:
            log(f"exe 启动后异常退出 (code={result.returncode})", "ERR")
            stderr = result.stderr.decode("utf-8", errors="replace")[-300:]
            if stderr.strip():
                log(f"  stderr: {stderr}", "ERR")
            sys.exit(1)
        # returncode=0 且 <6s 退出也可能正常（无数据时快速完成）
        log("  exe 启动正常 (快速退出)", "OK")
    except subprocess.TimeoutExpired:
        # 超时说明 GUI 正常运行中
        log("  exe 启动正常 (GUI 运行中，已超时终止)", "OK")


def step_compress(dist_dir, app_dir, version):
    log("压缩打包...", "STEP")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"KPI_Tool_{version}_{ts}.zip"
    zip_path = os.path.join(dist_dir, zip_name)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(app_dir):
            # 将空目录也写入 zip（确保解压后目录结构完整）
            if not files and not dirs:
                dir_entry = os.path.join(DIST_NAME, os.path.relpath(root, app_dir)) + "/"
                zf.mkdir(dir_entry)
            for f in files:
                # 跳过临时文件
                if f.startswith("~$") or f == ".DS_Store":
                    continue
                fp = os.path.join(root, f)
                arcname = os.path.join(DIST_NAME, os.path.relpath(fp, app_dir))
                zf.write(fp, arcname)

    size_mb = os.path.getsize(zip_path) / (1024 * 1024)
    log(f"  {zip_name} ({size_mb:.1f}MB)", "OK")
    return zip_path


# ============================================================
# 主流程
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="KPI_Tool 一键构建")
    parser.add_argument("--skip-build", action="store_true", help="跳过 PyInstaller 构建")
    parser.add_argument("--version", default=DEFAULT_VERSION, help=f"版本号 (默认 {DEFAULT_VERSION})")
    args = parser.parse_args()

    src_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(src_dir, "dist")
    app_dir = os.path.join(dist_dir, DIST_NAME)
    build_dir = os.path.join(src_dir, "build", DIST_NAME)
    internal_dir = os.path.join(app_dir, "_internal")

    print(f"\n{'='*50}")
    print(f"  KPI_Tool 构建脚本 ({args.version})")
    print(f"{'='*50}\n")

    if not args.skip_build:
        step_clean(app_dir, build_dir)
        step_build(src_dir)

    if not os.path.exists(app_dir):
        log(f"dist/{DIST_NAME} 不存在，请先执行完整构建", "ERR")
        sys.exit(1)

    step_setup_dirs(app_dir, internal_dir, src_dir)
    step_optimize(internal_dir)
    step_verify(app_dir)

    # 清理验证测试产生的临时文件
    import glob as _glob
    for f in _glob.glob(os.path.join(app_dir, "logs", "*")):
        os.remove(f)

    zip_path = step_compress(dist_dir, app_dir, args.version)

    print(f"\n{'='*50}")
    log(f"构建完成: {zip_path}", "OK")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
