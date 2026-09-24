# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import importlib.util

ocr_package = Path(importlib.util.find_spec('ddddocr').origin).parent
ocr_models = ('common.onnx', 'common_det.onnx', 'common_old.onnx')

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[(str(ocr_package / model), 'ddddocr') for model in ocr_models],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='智慧工卡管理平台',
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
    icon=['ico_image\\app.ico'],
)
