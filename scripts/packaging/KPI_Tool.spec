# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/main_gui.py'],
    pathex=['src'],
    binaries=[],
    datas=[('configs', '配置文件'), ('assets/logo.ico', '.'), ('assets/templates', '保障小区清单')],
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
    icon=['assets/logo.ico'],
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
