# -*- mode: python ; coding: utf-8 -*-
# 在 Windows 上执行：pyinstaller sap_cbs_gui.spec
# 生成单文件、无控制台窗口的 GUI（适合发给同事双击运行）。

import pathlib

PROJECT = pathlib.Path(SPEC).parent.resolve()
ENTRY = PROJECT / "gui_app.py"

block_cipher = None

a = Analysis(
    [str(ENTRY)],
    pathex=[str(PROJECT)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "openpyxl",
        "openpyxl.cell._writer",
        "pandas",
        "numpy",
        "backend_reconcile_sap_cbs",
        "reconcile_sap_cbs",
        "total_reconcile_service",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "flask",
        "werkzeug",
        "jinja2",
        "matplotlib",
        "PIL",
        "IPython",
        "notebook",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="SAP_CBS_对账助手",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
