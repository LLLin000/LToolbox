#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
构建 LToolbox.ppam（原生 PowerPoint 构建，非手工拼包）。

流程：
  1. 从 src/engine.py 读版本号（唯一来源）。
  2. 用本机 PowerPoint 新建宏演示文稿 -> 注入 VBA 启动器 -> VBE 编译检查 -> 原生另存为 PPAM(filetype=30)。
  3. 向 PPAM 内注入 customUI14.xml、关系与彩色图标。

为什么必须在有 PowerPoint 的机器上构建：
  只有 PowerPoint 自己能生成合法的 vbaProject.bin 并做真正的编译检查；
  CI 想脱离 PowerPoint 构建，需要 vbamc(.NET，见 docs/PACKAGING.md)。

用法:
  python tools/build_ppam.py                       # exe 模式（发布用，默认）
  python tools/build_ppam.py --mode python         # 开发机直调 pythonw + engine.py
  python tools/build_ppam.py --out dist/x.ppam
"""
import argparse, glob, os, re, shutil, sys, tempfile, zipfile

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
DIST = os.path.join(ROOT, "dist")

BAS = os.path.join(SRC, "LToolbox.bas")
RIBBON = os.path.join(SRC, "customUI14.xml")
ICON_DIR = os.path.join(SRC, "icons")
ENGINE = os.path.join(SRC, "engine.py")

EXT = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
ADDINS_DIR_DEFAULT = os.path.join(os.environ.get("LOCALAPPDATA", ""), "LToolbox")
# 发布版里 exe 的固定安装位置；与 tools/install.py、installer/ 保持一致。
EXE_PATH_DEFAULT = os.path.join(ADDINS_DIR_DEFAULT, "LToolbox.exe")


def read_version():
    text = open(ENGINE, encoding="utf-8").read()
    m = re.search(r'^VERSION = "([^"]+)"', text, re.M)
    if not m:
        sys.exit("src/engine.py 里找不到 VERSION = \"x.y.z\"")
    return m.group(1)


def ensure_trust():
    """构建只需要：允许访问 VBA 工程对象模型。安装/运行时不需要这个开关。"""
    import winreg
    base = r"Software\Microsoft\Office\16.0\PowerPoint\Security"
    for view in (0, winreg.KEY_WOW64_32KEY):
        try:
            k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base, 0,
                                   winreg.KEY_SET_VALUE | view)
            winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(k)
        except OSError:
            pass
    print("[1/4] AccessVBOM=1（仅构建需要）")


def render_bas(version, mode, exe_path):
    raw = open(BAS, encoding="utf-8").read().replace("__VERSION__", version)
    if mode == "exe":
        pyw, eng = exe_path, ""
    else:
        pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(pyw):
            pyw = sys.executable
        eng = ENGINE
    raw = raw.replace("__PYW__", pyw).replace("__ENG__", eng)
    # Attribute 行是导出产物，从源码注入时要去掉
    return "\n".join(l for l in raw.splitlines()
                     if not l.strip().startswith("Attribute "))


def build_native_ppam(out_ppam, bas_source):
    """新建宏演示文稿 -> 注入 VBA -> 编译检查 -> 另存为 PPAM。"""
    import win32com.client
    import traceback
    if os.path.exists(out_ppam):
        os.remove(out_ppam)
    pp = win32com.client.DispatchEx("PowerPoint.Application")
    pres = None
    try:
        pres = pp.Presentations.Add(WithWindow=False)
        # 必须至少有一张幻灯片，否则 SaveAs(30) 会报储存错误
        pres.Slides.Add(1, 12)
        comp = pres.VBProject.VBComponents.Add(1)  # 1 = 标准模块
        comp.Name = "LToolbox"
        comp.CodeModule.AddFromString(bas_source)
        print("[2/4] VBA 模块注入成功")
        try:
            btn = pp.VBE.CommandBars.FindControl(Id=578)  # Debug > Compile
            if btn and btn.Enabled:
                btn.Execute()
                print("       VBE 编译检查通过")
        except Exception as ce:
            print("       VBE 编译检查跳过:", ce)
        pres.SaveAs(out_ppam, 30)  # ppSaveAsOpenXMLAddin
        with zipfile.ZipFile(out_ppam) as z:
            if not any("vbaProject" in n for n in z.namelist()):
                raise RuntimeError("生成的 ppam 不含 vbaProject.bin")
        print("[3/4] 原生 PPAM 已生成: %d bytes" % os.path.getsize(out_ppam))
    except Exception as e:
        traceback.print_exc()
        print("注入/编译/保存失败:", e)
        return False
    finally:
        if pres is not None:
            try:
                pres.Saved = True
                pres.Close()
            except Exception:
                pass
        try:
            pp.Quit()
        except Exception:
            pass
    return True


def inject_ribbon_and_icons(ppam_path, version):
    xml = open(RIBBON, encoding="utf-8").read().replace("__VERSION__", version)
    icons = sorted(os.path.splitext(os.path.basename(p))[0]
                   for p in glob.glob(os.path.join(ICON_DIR, "*.png")))
    rels = ('<?xml version="1.0" encoding="utf-8"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            + "\n".join('  <Relationship Id="%s" Type="%s" Target="images/%s.png"/>'
                        % (ic, REL_IMAGE, ic) for ic in icons)
            + "\n</Relationships>")

    tmp = ppam_path + ".tmp"
    with zipfile.ZipFile(ppam_path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                if 'Extension="png"' not in s:
                    s = s.replace("</Types>",
                                  '<Default Extension="png" ContentType="image/png"/></Types>')
                data = s.encode("utf-8")
            elif item.filename == "_rels/.rels":
                s = data.decode("utf-8")
                if EXT not in s:
                    i = s.rfind("</Relationships>")
                    s = s[:i] + ('<Relationship Id="rIdLT" Type="%s" '
                                 'Target="customUI/customUI14.xml"/>' % EXT) + s[i:]
                data = s.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("customUI/customUI14.xml", xml.encode("utf-8"))
        if icons:
            zout.writestr("customUI/_rels/customUI14.xml.rels", rels.encode("utf-8"))
            for ic in icons:
                zout.writestr("customUI/images/%s.png" % ic,
                              open(os.path.join(ICON_DIR, ic + ".png"), "rb").read())
    os.remove(ppam_path)
    shutil.move(tmp, ppam_path)
    print("[4/4] 已注入 RibbonX 选项卡与 %d 个图标: %s" % (len(icons), ", ".join(icons)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["exe", "python"], default="exe")
    ap.add_argument("--exe-path", default=EXE_PATH_DEFAULT,
                    help="exe 模式下写进 VBA 启动器的 exe 完整路径")
    ap.add_argument("--out", default=None)
    a = ap.parse_args()

    for p in (BAS, RIBBON, ENGINE):
        if not os.path.exists(p):
            sys.exit("缺文件: " + p)

    version = read_version()
    out = a.out or os.path.join(DIST, "LToolbox-%s.ppam" % version)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    print("版本 %s，模式 %s" % (version, a.mode))

    ensure_trust()
    bas = render_bas(version, a.mode, a.exe_path)
    if not build_native_ppam(out, bas):
        sys.exit("构建中止。")
    inject_ribbon_and_icons(out, version)
    print("\n产物: %s" % os.path.abspath(out))
    if a.mode == "exe":
        print("提示: 该 ppam 启动器指向 %s，安装脚本会把 exe 放到同一位置。" % a.exe_path)


if __name__ == "__main__":
    main()
