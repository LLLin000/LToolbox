#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""开发自检：用真实 PowerPoint 跑一遍完整链路（标定 -> 取景框 -> 比例尺 -> 原图裁取）。

这是本项目唯一能覆盖"COM + 内嵌原图快照 + 几何换算"的检查，改引擎后必须跑。
需要：桌面版 PowerPoint；会在临时演示文稿里操作，结束时关闭且不保存。

用法:
  python tools/dev_smoke.py                       # 用 dist/LToolbox.exe
  python tools/dev_smoke.py --exe "%LOCALAPPDATA%\\LToolbox\\LToolbox.exe"
"""
import argparse, os, subprocess, sys, tempfile
from PIL import Image, ImageDraw

if hasattr(sys.stdout, "reconfigure"):
    # 英文/CI 控制台（cp1252）无法编码中文，只降级不崩溃
    sys.stdout.reconfigure(errors="replace")

SLIDE_W, SLIDE_H = 600.0, 600.0     # 图片在幻灯片上的尺寸(pt)
PNG_W, PNG_H = 1200, 1200           # 内嵌原图像素
BAR_PX = (100, 700)                 # 测试图上"比例尺"横跨的原图像素
BAR_UM = 500.0                      # 它代表的真实长度
TARGET_UM = 400.0                   # 取景框目标边长


def make_fixture(path):
    im = Image.new("RGB", (PNG_W, PNG_H), "white")
    d = ImageDraw.Draw(im)
    d.rectangle([BAR_PX[0], 1100, BAR_PX[1], 1120], fill="black")
    d.ellipse([400, 300, 800, 700], outline="black", width=8)
    im.save(path)
    return path


def run_engine(exe, *args):
    """按 VBA 的方式调用引擎；不传 --gui，靠日志/形状断言结果。"""
    rc = subprocess.call([exe, *args], timeout=120)
    log = os.path.join(tempfile.gettempdir(), "LToolbox_engine.log")
    tail = ""
    if os.path.exists(log):
        tail = open(log, encoding="utf-8", errors="replace").read()[-400:]
    return rc, tail


def tag(shape, name):
    try:
        return shape.Tags(name)
    except Exception:
        return None


def fail(msg, log=""):
    print("FAIL:", msg)
    if log:
        print(log)
    sys.exit(1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=os.path.join("dist", "LToolbox.exe"))
    a = ap.parse_args()
    exe = os.path.abspath(a.exe)
    if not os.path.exists(exe):
        sys.exit("找不到引擎: " + exe + "（先 python tools/build_exe.py）")

    import win32com.client as wc
    pp = wc.DispatchEx("PowerPoint.Application")
    tmp = tempfile.mkdtemp(prefix="LToolbox_smoke_")
    png = make_fixture(os.path.join(tmp, "fixture.png"))
    pres = None
    checks = []
    try:
        pres = pp.Presentations.Add()
        slide = pres.Slides.Add(1, 12)  # ppLayoutBlank
        # LinkToFile=0(嵌入) / SaveWithDocument=-1：链接式图片不嵌入原图，引擎无法裁取
        pic = slide.Shapes.AddPicture(png, 0, -1, 100, 100, SLIDE_W, SLIDE_H)
        pic.Name = "SmokePicture"
        pp.ActiveWindow.View.GotoSlide(1)
        px_per_pt = PNG_W / SLIDE_W
        start_pt = 100 + BAR_PX[0] / px_per_pt
        end_pt = 100 + BAR_PX[1] / px_per_pt

        # --- ① 标定 -------------------------------------------------------
        line = slide.Shapes.AddLine(start_pt, 650, end_pt, 650)
        line.Select()
        rc, log = run_engine(exe, "calibrate", str(int(BAR_UM)))
        um_per_px = tag(pic, "LTOOLBOX_UM_PER_PX")
        expect = BAR_UM / (BAR_PX[1] - BAR_PX[0])
        if not um_per_px or abs(float(um_per_px) - expect) > 1e-6:
            fail("标定换算率错误: %s，应为 %.6f" % (um_per_px, expect), log)
        checks.append("① 标定 %.6f µm/px 正确" % expect)

        # --- ② 取景框 -----------------------------------------------------
        pic.Select()
        rc, log = run_engine(exe, "makebox", str(int(TARGET_UM)))
        box = None
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes.Item(i)
            if s.Name.startswith("LT_box"):
                box = s
        if box is None:
            fail("没有生成取景框", log)
        want_pt = TARGET_UM / (expect * px_per_pt)
        if abs(box.Width - want_pt) > 2 or abs(box.Height - want_pt) > 2:
            fail("取景框尺寸 %.2f×%.2f pt，应为 %.2f pt" % (box.Width, box.Height, want_pt), log)
        checks.append("② 取景框 %.0f µm = %.1f pt 正确" % (TARGET_UM, box.Width))

        # --- ③ 比例尺 -----------------------------------------------------
        box.Select()
        rc, log = run_engine(exe, "scalebar")
        bar = label = None
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes.Item(i)
            if s.Name.startswith("LT_scalebar"):
                bar = s
            if s.Name.startswith("LT_scalelabel"):
                label = s
        if bar is None or label is None:
            fail("没有生成比例尺黑条/文字", log)
        want_um = TARGET_UM * 0.25      # 上限 25%
        bar_um = bar.Width * expect * px_per_pt
        if bar_um > want_um + 1e-6:
            fail("比例尺 %.1f µm 超过框宽 25%% (%.1f µm)" % (bar_um, want_um), log)
        checks.append("③ 比例尺 %s，%.0f µm（框宽 %.0f µm 的 %.0f%%）"
                      % (label.TextFrame.TextRange.Text, bar_um, TARGET_UM,
                         100 * bar_um / TARGET_UM))

        # --- ④ 原图裁取 ---------------------------------------------------
        slide.Shapes.Range([pic.Name, box.Name]).Select()
        rc, log = run_engine(exe, "crop")
        crop = None
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes.Item(i)
            if s.Name.startswith("LT_crop"):
                crop = s
        if crop is None:
            fail("没有生成裁图", log)
        rect = tag(crop, "LTOOLBOX_SOURCE_PIXEL_RECT")
        x1, y1, x2, y2 = [int(v) for v in rect.split(",")]
        got_px = x2 - x1, y2 - y1
        want_px = round(TARGET_UM / expect)
        if abs(got_px[0] - want_px) > 2 or abs(got_px[1] - want_px) > 2:
            fail("裁图像素 %s，应为约 %d×%d 原图像素" % (got_px, want_px, want_px), log)
        if tag(pic, "LTOOLBOX_UM_PER_PX") != um_per_px:
            fail("裁取后原图标定被改写", log)
        checks.append("④ 裁图 %d×%d 原图像素（原图与标定未改动）" % got_px)

        print("\n".join(checks))
        print("PASS: 引擎端到端通过（%s）" % os.path.basename(exe))
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


if __name__ == "__main__":
    main()
