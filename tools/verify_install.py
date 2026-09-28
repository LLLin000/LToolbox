#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""安装自检：文件 / 注册表 / PowerPoint 是否加载 / 功能区宏能否真的跑通引擎。

用来回答"点了安装但没看到选项卡"这类问题。需要桌面版 PowerPoint：会新开一个实例，
建一个临时演示文稿做一次真实调用，然后关闭不保存、退出。

用法:
  python tools/verify_install.py
"""
import base64, os, subprocess, sys, tempfile, time, zipfile

LOCAL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "LToolbox")
EXE = os.path.join(LOCAL_DIR, "LToolbox.exe")
PPAM = os.path.join(os.environ["APPDATA"], "Microsoft", "AddIns", "LToolbox.ppam")
REG = r"Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox"

# 自检用最小图片（64×64 PNG），避免为了自检再依赖 Pillow
FIXTURE_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAfElEQVR4nNXOQREAIADDsFL/nocIHlyjIGcbZRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncRIncf4OvLpyqgN9ZSiDcwAAAABJRU5ErkJggg==")


def check_files():
    ok = True
    for p in (EXE, PPAM):
        if os.path.exists(p):
            print("  ✓ 文件存在 %s (%.0f KB)" % (p, os.path.getsize(p) / 1024))
        else:
            print("  ✗ 缺文件 %s" % p)
            ok = False
    return ok


def check_registry():
    import winreg
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG)
    except FileNotFoundError:
        print("  ✗ 注册表项不存在: HKCU\\%s" % REG)
        return False
    path, _ = winreg.QueryValueEx(k, "Path")
    autoload, _ = winreg.QueryValueEx(k, "AutoLoad")
    winreg.CloseKey(k)
    good = os.path.normcase(path) == os.path.normcase(PPAM)
    print("  %s Path=%s" % ("✓" if good else "✗", path))
    print("  %s AutoLoad=0x%X" % ("✓" if autoload else "✗", autoload))
    return good and bool(autoload)


def check_ppam_payload():
    with zipfile.ZipFile(PPAM) as z:
        names = z.namelist()
        has_vba = any("vbaProject" in n for n in names)
        has_ui = "customUI/customUI14.xml" in names
        icons = [n for n in names if n.startswith("customUI/images/")]
    print("  %s vbaProject.bin" % ("✓" if has_vba else "✗"))
    print("  %s customUI14.xml" % ("✓" if has_ui else "✗"))
    print("  %s %d 个图标" % ("✓" if icons else "✗", len(icons)))
    return has_vba and has_ui and bool(icons)


def check_macro_roundtrip():
    """功能区宏 -> VBA -> 引擎 exe（安装路径）-> PowerPoint 形状，全链路一次调用。"""
    import win32com.client as wc
    pp = wc.DispatchEx("PowerPoint.Application")
    tmp = tempfile.mkdtemp(prefix="LToolbox_verify_")
    png = os.path.join(tmp, "fixture.png")
    with open(png, "wb") as f:
        f.write(FIXTURE_PNG)
    pres = None
    try:
        addin = None
        for i in range(1, pp.AddIns.Count + 1):
            item = pp.AddIns.Item(i)
            if "LToolbox" in item.Name or "工具箱" in item.Name:
                addin = item
        if addin is None:
            print("  ✗ PowerPoint 的加载项列表里没有 LToolbox（重启 PowerPoint 后再试）")
            return False
        print("  ✓ 加载项已注册: %s（Loaded=%s）" % (addin.Name, addin.Loaded))

        pres = pp.Presentations.Add()
        slide = pres.Slides.Add(1, 12)  # ppLayoutBlank
        pic = slide.Shapes.AddPicture(png, 0, -1, 100, 100, 200, 200)
        box = slide.Shapes.AddShape(1, 150, 150, 100, 100)  # 位于图片内
        pp.ActiveWindow.View.GotoSlide(1)
        slide.Shapes.Range([pic.Name, box.Name]).Select()

        print("  · 调用功能区宏 LT_BatchCrop（等价于点按钮）…")
        pp.Run(addin.Name + "!LT_BatchCrop")
        crop = None
        for _ in range(40):  # VBA 是异步启动 exe 的，最多等 20 s
            time.sleep(0.5)
            for i in range(1, slide.Shapes.Count + 1):
                s = slide.Shapes.Item(i)
                if s.Name.startswith("LT_crop"):
                    crop = s
            if crop is not None:
                break
        if crop is None:
            log = os.path.join(tempfile.gettempdir(), "LToolbox_engine.log")
            print("  ✗ 宏没有产出裁图；引擎日志尾部：")
            if os.path.exists(log):
                print("    " + open(log, encoding="utf-8", errors="replace").read()[-300:].replace("\n", "\n    "))
            return False
        print("  ✓ 宏 -> 引擎 -> 裁图成功: %s" % crop.Name)
        return True
    finally:
        try:
            if pres is not None:
                pres.Saved = True
                pres.Close()
        except Exception:
            pass
        try:
            pp.Quit()
        except Exception:
            pass


def main():
    print("1) 文件")
    ok = check_files()
    print("2) 注册表")
    ok &= check_registry()
    print("3) ppam 内容")
    ok &= check_ppam_payload()
    print("4) 全链路（功能区宏 -> 引擎）")
    ok &= check_macro_roundtrip()
    print()
    print("安装完好。" if ok else "有项目未通过，见上面 ✗；重装：python tools/install.py")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
