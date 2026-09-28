#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
安装 / 卸载 LToolbox（当前用户，无需管理员）。

做三件事：
  1. 复制 dist/LToolbox.exe  ->  %LOCALAPPDATA%\LToolbox\LToolbox.exe   （ppam 里的启动器指向这里）
  2. 复制 dist/LToolbox-<v>.ppam ->  %APPDATA%\Microsoft\AddIns\LToolbox.ppam
  3. 写 HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox（Path/Title/Description/AutoLoad）

不改宏安全设置（AccessVBOM 只有构建时才需要）。

用法:
  python tools/install.py                 # 自动找 dist/ 里的产物
  python tools/install.py --ppam a.ppam --exe b.exe
  python tools/install.py --uninstall
"""
import argparse, glob, os, re, shutil, subprocess, sys, winreg

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST = os.path.join(ROOT, "dist")
ENGINE = os.path.join(ROOT, "src", "engine.py")

LOCAL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "LToolbox")
EXE_DST = os.path.join(LOCAL_DIR, "LToolbox.exe")
ADDINS_DIR = os.path.join(os.environ["APPDATA"], "Microsoft", "AddIns")
PPAM_DST = os.path.join(ADDINS_DIR, "LToolbox.ppam")

ADDIN_KEY = r"Software\Microsoft\Office\16.0\PowerPoint\AddIns"
ADDIN_NAME = "LToolbox"
# 早期失败安装留下的中文名坏包
OLD_NAMES = ["L的工具箱.ppam"]


def read_version():
    m = re.search(r'^VERSION = "([^"]+)"',
                  open(ENGINE, encoding="utf-8").read(), re.M)
    return m.group(1) if m else "0.0.0"


def find_assets(ppam_arg, exe_arg):
    version = read_version()
    ppam = ppam_arg or os.path.join(DIST, "LToolbox-%s.ppam" % version)
    exe = exe_arg or os.path.join(DIST, "LToolbox.exe")
    if not os.path.exists(ppam):
        others = sorted(glob.glob(os.path.join(DIST, "*.ppam")))
        if len(others) == 1:
            ppam = others[0]
        else:
            sys.exit("找不到 PPAM（%s）。先运行 tools/build_ppam.py 或指定 --ppam。" % ppam)
    if not os.path.exists(exe):
        sys.exit("找不到 EXE（%s）。先运行 tools/build_exe.py 或指定 --exe。" % exe)
    return ppam, exe


def purge_legacy():
    for name in OLD_NAMES:
        p = os.path.join(ADDINS_DIR, name)
        if os.path.exists(p):
            try:
                os.remove(p)
                print("   删除旧包:", p)
            except OSError as e:
                print("   删除失败:", p, e)
    for name in [ADDIN_NAME] + OLD_NAMES:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, ADDIN_KEY + "\\" + name)
            print("   注销遗留注册项:", name)
        except FileNotFoundError:
            pass


def register():
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, ADDIN_KEY + "\\" + ADDIN_NAME)
    winreg.SetValueEx(k, "Path", 0, winreg.REG_SZ, PPAM_DST)
    winreg.SetValueEx(k, "Title", 0, winreg.REG_SZ, "L的工具箱")
    winreg.SetValueEx(k, "Description", 0, winreg.REG_SZ,
                      "图片标定 / 比例尺 / 原图像素裁取 · v" + read_version())
    winreg.SetValueEx(k, "AutoLoad", 0, winreg.REG_DWORD, 0xFFFFFFFF)
    winreg.CloseKey(k)


def powerpoint_running():
    """PowerPoint 运行时会锁住已加载的 .ppam，复制必失败；先查再动手。

    注意：tasklist 在中文 Windows 上输出 GBK，不能按 UTF-8 解码（会拿到 None），
    所以这里只比较字节。
    """
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
            capture_output=True).stdout or b""
        return b"POWERPNT.EXE" in out
    except OSError:
        return False


def install(ppam, exe):
    if powerpoint_running():
        sys.exit("PowerPoint 正在运行，加载项文件被锁定。\n"
                 "请完全退出 PowerPoint 后重新运行（本脚本不会强制关闭它）。")
    os.makedirs(LOCAL_DIR, exist_ok=True)
    os.makedirs(ADDINS_DIR, exist_ok=True)
    purge_legacy()
    try:
        shutil.copyfile(exe, EXE_DST)
        print("[1/3] exe   ->", EXE_DST)
    except OSError as e:
        sys.exit("复制 exe 失败（磁盘/权限）：%s" % e)
    try:
        shutil.copyfile(ppam, PPAM_DST)
        print("[2/3] ppam  ->", PPAM_DST)
    except OSError as e:
        sys.exit("复制 ppam 失败（PPT 是否正开着？关掉 PowerPoint 再试）：%s" % e)
    register()
    print("[3/3] 已注册自动加载（HKCU，无需管理员）")
    print("\n完成。完全关闭并重启 PowerPoint 后，出现“L的工具箱”选项卡。"
          "\n（若出现安全提示：文件 > 选项 > 信任中心 > 信任中心设置 > 加载项 勾选允许）")


def uninstall():
    purge_legacy()
    for p in (PPAM_DST,):
        if os.path.exists(p):
            try:
                os.remove(p)
                print("已删除", p)
            except OSError as e:
                print("删除失败:", p, e)
    if os.path.isdir(LOCAL_DIR):
        try:
            shutil.rmtree(LOCAL_DIR)
            print("已删除", LOCAL_DIR)
        except OSError as e:
            print("删除失败:", LOCAL_DIR, e)
    print("卸载完成，请重启 PowerPoint。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ppam", default=None)
    ap.add_argument("--exe", default=None)
    ap.add_argument("--uninstall", action="store_true")
    a = ap.parse_args()
    if a.uninstall:
        uninstall()
        return
    ppam, exe = find_assets(a.ppam, a.exe)
    install(ppam, exe)


if __name__ == "__main__":
    main()
