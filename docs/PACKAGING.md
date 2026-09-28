# 打包与分发调研（PACKAGING）

问题：PowerPoint 插件（VBA 加载项）到底怎么打包？能不能做成安装包？怎么持续更新？
结论与依据如下，末尾是**本项目采用的做法**。

---

## 1. `.ppam` 本体是什么

`.ppam` 就是一个 **OPC（zip）包**，和 `.pptx` 同一套容器规范：

```
LToolbox.ppam
├── [Content_Types].xml
├── _rels/.rels
│     └── Relationship Type=".../2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"
├── ppt/presentation.xml
├── ppt/vbaProject.bin          ← VBA 代码（二进制，编译后）
└── customUI/
    ├── customUI14.xml          ← RibbonX 功能区定义
    ├── _rels/customUI14.xml.rels
    └── images/*.png            ← 按钮图标
```

**推论**：插件本身没有"安装程序"的概念，它就是"一个文件 + 一个注册项"。
所以"能不能做成安装包"= 能不能自动完成那两步（放文件 + 写注册表）。**能**。

官方依据：RibbonX 定制文档、`Presentation.SaveAs` 的 `ppSaveAsOpenXMLAddin = 30`
（[PpSaveAsFileType](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype)）、
[用 Open XML 文件定制功能区](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file)。

---

## 2. 安装/注册的三条路

| 方式 | 做法 | 适用 |
|---|---|---|
| 手动 | PowerPoint → 文件 → 选项 → 加载项 → 添加 | 自己用 |
| 注册表 | `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox`，值 `Path`（ppam 全路径）、`Title`、`Description`、`AutoLoad=DWORD 0xFFFFFFFF` | 脚本/安装包 |
| 企业 | 组策略 / Intune 推同一注册表键 + 文件 | 批量 |

社区实践一致：Inno Setup（或 MSI）把 `.ppam` 复制到 `%APPDATA%\Microsoft\AddIns` 并写上面的 HKCU 键，
即可实现"无人值守自动加载"（[Experts Exchange 实操帖](https://www.experts-exchange.com/questions/29230234/Automatically-load-PowerPoint-Addin-ppam-without-user-intervention-when-PowerPoint-starts.html)、
[SO: Best ways to distribute a .ppam file](https://stackoverflow.com/questions/78603379/best-ways-to-distribute-a-ppam-file)）。

`16.0` 覆盖 Office 2016/2019/2021/365（都是 16.0 主版本），这是当前通用写法。

---

## 3. 构建：三条路线

| 路线 | 依赖 | 优点 | 缺点 |
|---|---|---|---|
| **PowerPoint COM + VBE 注入** | 本机桌面版 PowerPoint | 由 PowerPoint 自己生成合法 `vbaProject.bin`，且能**真正编译检查**（VBE 的 Compile 命令） | 必须有 PowerPoint，无法在 Linux/无 Office 的 CI 上跑 |
| **vbamc**（NetOfficeFw，.NET 全局工具） | .NET 8/10，跨平台 | 可在 GitHub Actions（甚至 Linux）上从 `.vb` 源码编出 `.ppam` | 没有 PowerPoint 的编译检查；需要把 `.bas` 整理成 `.vb` 模块 |
| **手工改 zip** | 无 | 简单 | 不能生成 `vbaProject.bin`，只能改已有 ppam 的外围（Ribbon/图标） |

参考：[NetOfficeFw/vbamc](https://github.com/NetOfficeFw/vbamc)（v2.0.3，跨平台，支持 ppam）、
[DZemens/Build_Script](http://dzmns.co/visual-basic-and-source-control-made-easier/)（Python 把 VBComponents
导进 pptm 再转 ppam、Ribbon 打包注入的完整范例）。

**本项目：路线 1 做本地/发布构建，路线 3 只用于注入 Ribbon 与图标，路线 2 列为可选项。**
理由见第 6 节——ppam 是"外壳 + 启动器"，改动频率极低，没必要为它引入 .NET 工具链。

---

## 4. 安装包形态

| 形态 | 代表 | 说明 |
|---|---|---|
| PowerShell 脚本 | 本仓库 `installer/install.ps1` | 零依赖，双模式：从解压目录装 / 联网下最新 Release 装；能卸载 |
| Inno Setup | 本仓库 `installer/LToolbox.iss` | 生成 `Setup.exe`，用户双击即装；模板可参考 [bovender/VstoAddinInstaller](https://github.com/bovender/VstoAddinInstaller)（ISPP、安装时关闭目标 Office 程序） |
| MSI / Intune | 企业场景 | 需要打包工具（WiX），个人项目不必 |

注意：Office 加载项是**每用户**资源（`HKCU` + `%APPDATA%`），所以安装包应
`PrivilegesRequired=lowest`，不需要管理员——这既省事又少踩 UAC。

---

## 5. 签名与信任（分发绕不开）

| 事实 | 依据 |
|---|---|
| `.ppam` 属于宏启用文件，受信任中心宏策略约束 | [Office 方案开发者安全说明](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/Library-Reference/Concepts/security-notes-for-microsoft-office-solution-developers.md) |
| 从网络下载的文件带 Mark-of-the-Web，Office 默认可能直接拦截 | 同上 |
| 给 VBA 工程做代码签名 + 把证书装到「受信任发布者」，可消除每次提示 | [Trusted publishers for Office files](https://learn.microsoft.com/en-us/microsoft-365-apps/security/trusted-publisher) |
| 自签名只能用 `SelfCert.exe`，仅供自用/测试；对外分发应买代码签名证书 | [Digitally sign your VBA macro project](https://support.microsoft.com/en-us/office/vba/digitally-sign-your-vba-macro-project) |
| 签名建议升级到 V3 签名，防止签名后篡改 | [KB5000676](https://support.microsoft.com/en-us/topic/upgrade-signed-office-vba-macro-projects-to-v3-signature-kb5000676-2b8b3cae-ad64-4b4b-aa85-c4a98ca6da87) |

本项目额外一层：加载项会**启动一个外部 exe**（引擎）。因此分发时
exe 也应做代码签名，否则 SmartScreen 会对首次运行弹"未知发布者"。
自用/实验室用户可在信任中心把 `%APPDATA%\Microsoft\AddIns` 设为受信任位置。

---

## 6. 持续更新怎么设计（本项目的关键决策）

**.ppam 只做外壳，逻辑全在被调用的 exe 里。**

```
LToolbox.ppam  ──WScript.Shell 启动──▶  %LOCALAPPDATA%\LToolbox\LToolbox.exe <action>
   RibbonX + 图标 + VBA(一行转发)          engine.py 打包产物（全部业务逻辑）
```

由此得到：

- **升级 = 换 exe**：覆盖 `%LOCALAPPDATA%\LToolbox\LToolbox.exe`，重启 PowerPoint，结束。
  不用重装加载项，不用重新打开每个 pptx，用户文档里的标定（shape Tags）不受影响；
- **exe 可以由 CI 从源码构建**（Python 有了 PyInstaller）；ppam 变化极慢；
- 反过来，若把逻辑放在 VBA 里，每次修 bug 都要重装 ppam + 重开 PowerPoint，不可持续。

**因此发布流程是**：

```
本地（有 PowerPoint）:
  python tools/build_exe.py                    # 可跑，CI 也跑
  python tools/build_ppam.py --mode exe        # 只在 Ribbon/图标/启动方式变化时才需要
  copy dist\LToolbox-<ver>.ppam release\LToolbox.ppam   # 提交进仓库（CI 上没有 PowerPoint）
打 tag:
  git tag v0.2.0 && git push origin v0.2.0
CI:
  windows job: 构建 exe + 冒烟（dist/LToolbox.exe version）
  release job: 用提交的 ppam + CI 的 exe 打 zip → GitHub Release 附件
```

为什么不把 exe 也提交进仓库：exe 能由源码可复现地构建，放进 git 会让每次提交都膨胀几十 MB；
ppam 无法在 CI 上构建，所以只能提交（约 100 KB，可接受）。

---

## 7. exe 打包（PyInstaller）注意事项

- `python-pptx` 与 `Pillow` 是普通依赖；`win32com` 是**动态导入**，必须显式
  `hiddenimports`（本仓库 `tools/LToolbox.spec` 用 `collect_submodules("win32com")` + `pythoncom`/`pywintypes`）；
- 冻结后 `pythoncom`/`pywintypes` 缺失是已知坑，PyInstaller 自带 pywin32 运行钩子，
  但环境异常时用 `--debug=imports` 定位（[PyInstaller 文档](https://pyinstaller.org/en/stable/man/pyinstaller.html)）；
- 引擎要附着**已经在运行的** PowerPoint（`GetActiveObject`），所以必须是
  `console=False` 的普通 exe（不是服务、不新建实例）；
- `console=False` 下 `sys.stdout is None`，日志文件 `%TEMP%\LToolbox_engine.log` 才是输出通道。

---

## 8. 结论

| 问题 | 答案 |
|---|---|
| ppt 插件怎么打包？ | `.ppam` 是 zip 容器（vbaProject.bin + customUI + 图标），本身即"安装单元" |
| 能做成安装包吗？ | 能。安装包要做的只有两件事：放文件到 `%APPDATA%\Microsoft\AddIns`、写 HKCU 注册表键。已提供 PowerShell 脚本与 Inno Setup 模板 |
| 需要管理员吗？ | 不需要，每用户安装 |
| 可以不装 Python 吗？ | 可以。引擎已用 PyInstaller 打成单文件 `LToolbox.exe` |
| 怎么持续更新？ | 逻辑在 exe，ppam 是外壳；更新只替换 exe，GitHub Release + tag 驱动 |
| 有什么坑？ | 宏安全/签名（对外分发建议买证书）、ppam 无法在 CI 构建（故提交进仓库）、PowerPoint 必须正在运行 |
