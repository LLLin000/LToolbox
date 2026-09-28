# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 单文件打包规格：src/engine.py -> dist/LToolbox.exe

无控制台窗口（按钮调用，中途不闪黑框）。启动时必须能在已运行的 PowerPoint 上
做 COM 附着（GetActiveObject），所以 pywin32 的动态模块要显式收集。
"""
import os
from PyInstaller.utils.hooks import collect_submodules

SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))
SRC = os.path.normpath(os.path.join(SPEC_DIR, os.pardir, "src"))

hiddenimports = collect_submodules("win32com") + [
    "pythoncom",
    "pywintypes",
    "win32api",
    "win32timezone",
    "win32com.client",
    "PIL.Image",
    "PIL.ImageOps",
    "pptx",
]

a = Analysis(
    [os.path.join(SRC, "engine.py")],
    pathex=[SRC],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "scipy", "pandas"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="LToolbox",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
)
