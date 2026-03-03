# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

ROOT = Path(SPECPATH).resolve().parents[1]
SRC_DIR = ROOT / 'src'
ASSETS_DIR = ROOT / 'assets'
CONFIGS_DIR = ROOT / 'configs'
TEMPLATES_DIR = ASSETS_DIR / 'templates'
ICON_FILE = ASSETS_DIR / 'logo.ico'

a = Analysis(
    [str(SRC_DIR / 'main_gui.py')],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=[
        (str(CONFIGS_DIR), 'configs'),
        (str(ICON_FILE), '.'),
        (str(TEMPLATES_DIR), 'assets/templates'),
    ],
    hiddenimports=['customtkinter', 'CTkMessagebox', 'rapidfuzz', 'openpyxl', 'pandas', 'numpy'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'sqlite3', '_sqlite3',
        'scipy', 'scipy.*',
        'matplotlib', 'matplotlib.*',
        'IPython', 'IPython.*',
        'jupyter', 'jupyter.*', 'notebook',
        'pytest', 'unittest', 'sphinx', 'docutils',
        'email', 'html.parser', 'pydoc',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='KPI_Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(ICON_FILE)],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='KPI_Tool',
)
