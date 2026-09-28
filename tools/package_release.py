#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把发布产物打成 zip：exe + ppam + 安装/卸载脚本 + 说明。

ppam 必须在有 PowerPoint 的机器上构建（见 docs/PACKAGING.md），因此它是
随源码提交的；exe 由 CI 从源码构建。

用法:
  python tools/package_release.py                 # 用 dist/LToolbox.exe + release/*.ppam
  python tools/package_release.py --exe e --ppam p
"""
import argparse, glob, os, re, sys, zipfile

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST = os.path.join(ROOT, "dist")
RELEASE = os.path.join(ROOT, "release")
ENGINE = os.path.join(ROOT, "src", "engine.py")
INSTALLER = os.path.join(ROOT, "installer")


def read_version():
    m = re.search(r'^VERSION = "([^"]+)"', open(ENGINE, encoding="utf-8").read(), re.M)
    return m.group(1) if m else "0.0.0"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=None)
    ap.add_argument("--ppam", default=None)
    a = ap.parse_args()
    version = read_version()
    exe = a.exe or os.path.join(DIST, "LToolbox.exe")
    ppam = a.ppam or os.path.join(RELEASE, "LToolbox.ppam")
    if not os.path.exists(ppam):
        cands = sorted(glob.glob(os.path.join(RELEASE, "*.ppam")))
        if len(cands) != 1:
            sys.exit("找不到 release/LToolbox.ppam（先跑 tools/build_ppam.py --mode exe）")
        ppam = cands[0]
    for p in (exe, ppam):
        if not os.path.exists(p):
            sys.exit("缺文件: " + p)

    out = os.path.join(DIST, "LToolbox-%s-win64.zip" % version)
    os.makedirs(DIST, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(exe, "LToolbox.exe")
        z.write(ppam, "LToolbox.ppam")
        for name in ("install.ps1", "uninstall.ps1"):
            p = os.path.join(INSTALLER, name)
            if os.path.exists(p):
                z.write(p, name)
        z.writestr("README.txt",
                   "L的工具箱 v%s\n\n"
                   "安装：右键 install.ps1 -> 使用 PowerShell 运行\n"
                   "或：powershell -ExecutionPolicy Bypass -File install.ps1 -FromDir .\n"
                   "卸载：uninstall.ps1\n" % version)
    print("产物: %s (%.1f MB)" % (out, os.path.getsize(out) / 1048576))


if __name__ == "__main__":
    main()
