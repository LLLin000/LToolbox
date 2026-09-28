# 内部设计规范（DESIGN）

本文只描述**实现约定**：数据放哪、名字怎么起、几何怎么算。改代码前先看这里，
避免出现"同一个字段两套写法""同一个框两个名字"。

对应官方规范：[RibbonX 定制](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file)、
[图标尺寸](https://nolongerset.com/icon-dimensions-in-the-office-ribbon/)、
[RibbonX Image FAQ](https://learn.microsoft.com/en-us/archive/blogs/jensenh/ribbonx-image-faq)。

---

## 1. 单位与坐标

| 量 | 单位 | 来源 |
|---|---|---|
| PowerPoint 形状位置/尺寸 | pt（EMU 的 COM 表示，`Shape.Left/.Width`） | PowerPoint |
| 真实长度 | µm | 用户输入 / 标定线 |
| 标定换算 | µm / 原图像素（`um_per_px`） | `calibrate` |
| 屏幕 pt → 原图像素 | `px_per_pt = 可见像素宽 / Shape.Width` | `_picture_info` |

**关键点：所有物理换算都经过"原图像素"这一层，不经过屏幕像素。**
`px_per_pt` 用内嵌原图的**可见**像素（扣除 PowerPoint 自身的 `crop_left/right/top/bottom`）除以
形状显示尺寸，所以图片被裁剪过也不会算错。

```python
um_per_pt_x = um_per_px * px_per_pt_x      # 横轴
um_per_pt_y = um_per_px * px_per_pt_y      # 纵轴（非等比缩放时与 x 轴不同）
```

---

## 2. 持久化：Tag 键名表（shape Tags）

标定绑定在**图片 shape** 上，不是幻灯片、不是外部文件。Tag 随 `.pptx` 保存，
复制图片一起复制，因此跨会话无需任何"恢复"动作。

| Tag | 挂在谁身上 | 含义 |
|---|---|---|
| `LTOOLBOX_UM_PER_PX` | 图片 | 换算率（唯一的"有没有标定"判据） |
| `LTOOLBOX_REAL_UM` | 图片 | 标定时填的真实长度 |
| `LTOOLBOX_SPAN_PX` | 图片 | 标定线对应的原图像素跨度 |
| `LTOOLBOX_PICTURE_ID` | 框 / 状态角标 / 标定线 | 归属哪张图片（`Shape.Id`） |
| `LTOOLBOX_SOURCE_PICTURE_ID` | 裁出的小图 | 来源图片 |
| `LTOOLBOX_SOURCE_BOX_ID` | 裁出的小图 / 比例尺 / 标签 | 来源取景框 |
| `LTOOLBOX_SOURCE_PIXEL_RECT` | 裁出的小图 | `x1,y1,x2,y2`（原图像素，左闭右开→floor/ceil） |

`get_calibration_record()` 只要 `UM_PER_PX` 缺失或 ≤0 就返回 `None` → 报"未标定"。

新增字段必须遵守：**只增不改名**；旧键作废时在 `calibrate()` 的 legacy 清理列表里登记。

---

## 3. 形状命名表

名字前缀是其它功能的定位依据，改名等于改协议。

| 名称 | 类型 | 命名规则 | 备注 |
|---|---|---|---|
| `LT_box_NN` | 矩形 | 两位序号，同页唯一 | 取景框；用户自己画的矩形不叫这个名 |
| `LT_scalebar_<suffix>` | 矩形 | suffix = 框序号或框 Id | 比例尺黑条 |
| `LT_scalelabel_<suffix>` | 文本框 | 同上 | 比例尺文字，与黑条成对 |
| `LT_crop_<boxId>` | 图片 | 框 Id | 裁出的小图；重裁会删同名旧图 |
| `LT_calibration_line_<picId>` | 直线 | 图片 Id | 标定线（变绿、加粗） |
| `LT_calibration_status_<picId>` | 文本框 | 图片 Id | ✓/✕ 角标 |
| `LT_notice` | 文本框 | 固定 | 顶部红框错误提示，每次调用先清 |

**幂等性约定**：比例尺、标签、裁图在生成前先按名字删旧件，所以重复点击不会叠加。

---

## 4. 几何规则

| 规则 | 取值 | 位置 |
|---|---|---|
| 取景框居中 | 图片正中，然后避让 | `make_box` |
| 避让判据 | `IoU > 0.01` 视为重叠 | `_avoid_overlap` |
| 避让步长 | 每次右移 `1.05×w`，越界回到左边界并下移 `1.05×h` | 同上 |
| 避让上限 | 128 步；仍冲突则报错要求手动挪框 | 同上 |
| 比例尺取值 | `1/2/5 × 10ⁿ` 中 ≤ 框宽 25% 且最接近 20% 者 | `_nice_scale_um` |
| 比例尺内边距 | 框宽的 5.5% | `add_scalebar_to_box` |
| 黑条厚度 | `clamp(框宽×1.8%, 2, 5)` pt | 同上 |
| 文字字号 | `clamp(框宽×5.5%, 8, 14)` pt | 同上 |
| 裁取像素矩形 | `floor(x1), floor(y1), ceil(x2), ceil(y2)`，再 clamp 到原图 | `_crop_pixel_rect` |
| 最小裁取 | 任一边 < 2 px 报错 | `crop_selected_regions` |

**z 序**：`Shape.ZOrder(0)` 在 PPT COM 里是 *bring to front*。比例尺、标签、框、裁图
都调 `ZOrder(0)`，因此后生成的压在前面的上面；每次裁取后调 `_bring_box_overlays_front`
保证框和它的比例尺仍然可见。

---

## 5. 取数：为什么用 `SaveCopyAs` 快照

`_picture_info()` 先 `pres.SaveCopyAs(tmp, 24)` 再用 `python-pptx` 读内嵌原图 blob。

- 不读屏幕上看到的东西，避免缩放/裁剪/DPI 造成误差；
- 快照走临时目录，**不写用户的文件**（`pres.Saved` 状态不变）；
- 用 `python-pptx` 而不是 COM `Shape.Export`，因为要的是**原始字节**，不是渲染结果。

---

## 6. 引擎 CLI 契约

VBA 只负责转发，全部逻辑在 `engine.py`：

```
pythonw.exe engine.py <action> [value] [--gui]
  status        无参          显示/刷新所选图片的标定状态角标
  calibrate     500 [--unit um]  用选中的直线标定它所在的图片
  makebox       700           为已标定图片生成 700 µm 取景框
  scalebar      无参          为当前框生成比例尺
  crop          无参          按选中图片+框批量原像素裁取
  version       无参          打印版本
```

- 退出码非 0 + 顶部 `LT_notice` + `%TEMP%\LToolbox_engine.log` 三重可诊断；
- 单文件 exe 打包后 stdout 不可见，**日志文件才是唯一输出通道**，不要删；
- VBA 启动串允许带一个空参数（exe 模式占位），`main()` 已过滤空 argv。

---

## 7. 功能区与图标

| 项 | 现状 | 官方规范 | 结论 |
|---|---|---|---|
| 选项卡 | 自建 `ltTab`，`insertBeforeMso="TabHome"` | 推荐放在开始旁边 | ✓ |
| 分组 | 2 组：比例尺三步 / 原图分辨率裁取 | 语义分组、按钮 ≤ 6 | ✓ |
| 按钮 | 5 个 `size="large"` | large 32×32，small 16×16 | ✓（全是 large，因此只需要 32 px） |
| 图标 | 32×32 RGBA PNG，透明底，`customUI/images/*.png` | 32 bpp PNG + alpha，32×32 @96 dpi | ✓ |
| 高 DPI | 由 Office 缩放同一张 32 px | 建议补 16/20/32 多尺寸 | 见 ROADMAP |
| 文案 | 每个按钮都有 `screentip` + `supertip` | 必填项 | ✓ |
| 图标生成 | `src/gen_icons.py`（Pillow 代码画图，可复现） | — | ✓ |

图标**必须**通过 `image=` 引用包内资源，不能借用 `imageMso` 内置图标（避免与 Office 自身语义混淆，
详见 [Office 图标指南](https://github.com/OfficeDev/office-js-docs-pr/blob/main/docs/design/add-in-icons.md)）。

---

## 8. 安全边界

| 项 | 约定 |
|---|---|
| `AccessVBOM` | **只有构建 ppam 时**临时开启；安装/运行不需要，安装脚本不得写宏安全设置 |
| 注册表 | 只写 `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\LToolbox` 一个键 |
| 文件 | 只写 `%LOCALAPPDATA%\LToolbox\`（exe）与 `%APPDATA%\Microsoft\AddIns\LToolbox.ppam` |
| 用户文档 | 引擎只用 `SaveCopyAs` 读快照，不改用户 pptx；裁图作为新形状插入当前幻灯片 |
| 旋转图片 | 裁取前要求 `Rotation = 0`，否则拒绝（旋转会让像素映射非线性） |
