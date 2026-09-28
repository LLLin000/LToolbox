#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""冻结 exe 冒烟：能在没有 PowerPoint 的机器（含 CI）上跑。

检查两件事：
  1. 单文件 exe 能启动并写出日志（未崩溃、依赖已打包）；
  2. exe 报出的版本与 src/engine.py 的 VERSION 一致（防"忘了重新构建就发版"）。

用法:
  python tools/smoke_exe.py [dist/LToolbox.exe]
"""
import os, pathlib, re, subprocess, sys, tempfile

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    exe = os.path.abspath(sys.argv[1] if len(sys.argv) > 1
                          else os.path.join(ROOT, "dist", "LToolbox.exe"))
    if not os.path.exists(exe):
        sys.exit("找不到 exe: " + exe)
    version = re.search(r'^VERSION = "([^"]+)"',
                        open(os.path.join(ROOT, "src", "engine.py"),
                             encoding="utf-8").read(), re.M).group(1)
    log = pathlib.Path(tempfile.gettempdir()) / "LToolbox_engine.log"
    log.unlink(missing_ok=True)

    rc = subprocess.run([exe, "version"], timeout=300).returncode
    if rc != 0:
        sys.exit("exe 退出码 %d" % rc)
    if not log.exists():
        sys.exit("exe 没有写日志，可能启动即崩溃")
    text = log.read_text(encoding="utf-8", errors="replace")
    if ("L的工具箱 v%s" % version) not in text:
        sys.exit("exe 版本与 engine.py (%s) 不一致，日志尾部：\n%s"
                 % (version, text[-300:]))
    print("冒烟通过：%s 报告版本 %s" % (os.path.basename(exe), version))


if __name__ == "__main__":
    main()
