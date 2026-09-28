#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 src/engine.py 打成单文件 dist/LToolbox.exe（用户机器不需要装 Python）。

用法:
  python tools/build_exe.py
"""
import os, subprocess, sys

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC = os.path.join(ROOT, "tools", "LToolbox.spec")
DIST = os.path.join(ROOT, "dist")
WORK = os.path.join(ROOT, "build", "pyinstaller")


def main():
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--distpath", DIST, "--workpath", WORK, SPEC]
    print(" ".join(cmd))
    rc = subprocess.call(cmd)
    if rc != 0:
        sys.exit("PyInstaller 失败，返回码 %d" % rc)
    exe = os.path.join(DIST, "LToolbox.exe")
    if not os.path.exists(exe):
        sys.exit("没有生成 " + exe)
    print("\n产物: %s (%.1f MB)" % (exe, os.path.getsize(exe) / 1048576))


if __name__ == "__main__":
    main()
