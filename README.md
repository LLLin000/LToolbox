# L的工具箱 · LToolbox

> PowerPoint 显微图工具：**图片绑定标定 → 规范比例尺 → 原图像素批量裁取**。
>
> 面向把显微镜原图直接丢进 PPT 做组图的人：不再手算比例尺、不再靠截图缩图丢像素。

[![build](https://github.com/LLLin000/LToolbox/actions/workflows/build.yml/badge.svg)](https://github.com/LLLin000/LToolbox/actions/workflows/build.yml)
[![release](https://img.shields.io/github/v/release/LLLin000/LToolbox)](https://github.com/LLLin000/LToolbox/releases/latest)
[![license](https://img.shields.io/github/license/LLLin000/LToolbox)](LICENSE)

Windows + PowerPoint（桌面版，2016/2019/2021/365，x64 推荐）。安装后无需 Python、无需管理员。

---

## 功能

| 功能 | 做什么 | 关键点 |
|---|---|---|
| **① 标定 / 状态** | 在图片自带比例尺上画一条等长直线，填真实长度（µm），标定写入**这张图片自己**（shape Tags） | 图片之后缩放、移动、复制都继承；点①可随时看"已标定 500 µm / 未标定"角标 |
| **② 生成取景框** | 输入想要的真实视野边长（µm），生成红色取景框 | 旧红框保留、新框自动避让；超出现有视野会告诉你"最大可用边长" |
| **③ 处理** | 在当前框右下角生成规范比例尺 | 自动取 1/2/5×10ⁿ 值，约占框宽 20%、不超过 25%；黑条和文字都可单独编辑 |
| **批量原图裁取** | 选中图片 + 一个或多个框，一次裁出小图放回各框位置 | 按 **PPT 内嵌原图分辨率**裁切（不是屏幕截图），原图与红框都不动；不要求标定 |

---

## 安装

### 方式 A：一键（推荐）

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

不带参数时自动下载 [最新 Release](https://github.com/LLLin000/LToolbox/releases/latest) 的发布包；已解压到本地则用 `-FromDir .`。

### 方式 B：安装包

从 Release 下载 `LToolbox-<版本>-setup.exe`，双击安装（当前用户，无需管理员）。

### 方式 C：手动

1. 把 `LToolbox.exe` 放到 `%LOCALAPPDATA%\LToolbox\`；
2. 把 `LToolbox.ppam` 放进 `%APPDATA%\Microsoft\AddIns\`；
3. PowerPoint → 文件 → 选项 → 加载项 → 管理「PowerPoint 加载项」→ 转到 → 添加 `LToolbox.ppam`。

安装后**完全关闭并重启 PowerPoint**，功能区出现「L的工具箱」选项卡。

### 卸载

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1
```

或运行安装包自带的卸载程序。

---

## 快速上手

```
① 选中图片自带比例尺上的那条线  →  点【① 标定 / 状态】→ 填 500（µm）
② 选中图片                      →  点【② 生成取景框】→ 填 700（µm）
③ 选中那个红框                  →  点【③ 处理】→ 框内右下角出现 200 µm 比例尺
④ 选中图片 + 若干红框            →  点【批量原图裁取】→ 每个框位置出现原像素小图
```

---

## 工作原理（决定升级体验）

```text
LToolbox.ppam（启动器外壳）  --调用-->  LToolbox.exe（全部逻辑）
   ├ RibbonX：选项卡 / 5 个按钮      ├ 标定换算（µm ⇄ 原像素）
   ├ 图标（32×32 PNG）               ├ 取景框几何 / 避让
   └ VBA：一行命令，转发按钮点击      ├ 比例尺取值与绘制
                                     └ 原图像素裁取（Pillow）
```

- **标定存在图片的 shape Tags 里**（`LTOOLBOX_UM_PER_PX` 等），随 `.pptx` 保存，跨会话、跨文件复制都在；
- `ppam` 只做外壳，**升级通常只需要换 exe**：把新版 `LToolbox.exe` 覆盖到 `%LOCALAPPDATA%\LToolbox\`，重启 PowerPoint 即可，不用重新安装加载项；
- 只有改动功能区按钮、图标或启动方式时才需要重新构建 `ppam`。

---

## 从源码构建（开发者）

```bash
pip install pywin32 pillow python-pptx pyinstaller

python tools/build_exe.py                        # -> dist/LToolbox.exe
python tools/build_ppam.py --mode exe            # -> dist/LToolbox-<ver>.ppam（需要本机装 PowerPoint）
python tools/install.py                          # 装到当前用户
python tools/install.py --uninstall

python tools/package_release.py                  # -> dist/LToolbox-<ver>-win64.zip

python tools/dev_smoke.py                        # 端到端自检（需 PowerPoint）：标定→取景框→比例尺→裁取
python tools/verify_install.py                   # 安装自检：文件/注册表/加载项/功能区宏全链路
```

- **构建 ppam 必须有桌面版 PowerPoint**：只有 PowerPoint 能生成合法的 `vbaProject.bin` 并做真正的编译检查（`tools/build_ppam.py` 会调用 VBE 的 Compile）；构建时脚本会临时打开 `AccessVBOM`，安装/运行时不需要；
- 开发调试也可以用 `python tools/build_ppam.py --mode python`：ppam 直接调 `pythonw.exe src/engine.py`，改完代码不用重新构建 ppam；
- 发布包内容见 `tools/package_release.py`；`release/LToolbox.ppam` 是**提交进仓库的构建产物**（因为 CI 上没有 PowerPoint，见 `docs/PACKAGING.md`）。

---

## 目录结构

```
src/          engine.py（逻辑主体）· LToolbox.bas（VBA 启动器模板）· customUI14.xml · icons/
tools/        build_exe.py · build_ppam.py · install.py · package_release.py · LToolbox.spec
installer/    install.ps1 · uninstall.ps1 · LToolbox.iss（Inno Setup 模板）
release/      LToolbox.ppam（提交的构建产物）
docs/         DESIGN.md（内部设计规范）· PACKAGING.md（打包/分发调研）· ROADMAP.md
```

---

## 常见问题

**重启 PowerPoint 没看到选项卡？**
文件 → 选项 → 加载项 → 底部「管理」选 PowerPoint 加载项 → 转到，确认 `LToolbox.ppam` 在列表里并勾选；再不行看信任中心 → 加载项是否禁止。

**被宏安全拦截？**
`.ppam` 是宏加载项。若组织策略要求签名宏，需要给 VBA 工程做代码签名（见 `docs/PACKAGING.md` → 签名与信任）。自用可把 `%APPDATA%\Microsoft\AddIns` 加入受信任位置。

**点按钮没反应？**
引擎日志在 `%TEMP%\LToolbox_engine.log`，最后一次调用的入参和输出都在里面；报错同时会以红框提示画在当前幻灯片顶部（`LT_notice`）。

**图片必须先标定吗？**
只有比例尺三步需要；**批量原图裁取不需要标定**。

---

## English

LToolbox is a Windows PowerPoint add-in (VBA launcher + bundled Python engine) for microscopy figures:
calibrate a scale bar **onto the picture itself** (stored in shape tags, so it survives resizing/moving/copying),
generate a physically-sized (µm) field-of-view box, draw a journal-style scale bar inside it, and batch-crop
regions at the **embedded original image resolution** (never a screen grab). Installer is per-user, no admin,
no Python needed. Build with `tools/build_exe.py` + `tools/build_ppam.py`; see `docs/PACKAGING.md` for how
ppam packaging, installers, CI and macro signing work.

## License

MIT — see [LICENSE](LICENSE).
