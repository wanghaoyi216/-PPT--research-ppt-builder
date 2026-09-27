# 图表技法

这份文件回答两个问题：**这张图该不该画**，以及**画的话怎么画才不翻车**。

## 0. 先决定要不要画图

**结论能用一句话说清，就不要画图。** 「图谱补的是字面上不存在的那一跳」这句话，
画成四根柱子的收益是让人多花五秒钟读坐标轴，代价是版面被占掉三分之一。

真正值得画的三种情况：

| 情况 | 为什么图比字强 |
|---|---|
| 要看**关系 / 流程** | 文字描述拓扑要写五句话，图里三条线就说完了 |
| 要看**量级差** | 「快 12 倍」这句话读者记不住，两根等比柱子的长度差记得住 |
| 要看**趋势 / 拐点** | 随时间变化的量，列表格会让人自己心算斜率 |

反过来，下面这些一律别画：

- 三个数以内的对比 → 用 `kit.table()` 或一行文字带过；
- 结论是一个阈值判断（够不够好） → 写结论，别画图让人替你判断；
- 数据只有一次采样 → 单点没有趋势，硬画成折线是欺骗。

---

## 1. 分工：原生图表还是 matplotlib

两条路都能出图，选错了会付出不同的代价。

|  | **原生 PowerPoint 图表**<br>`kit.column_chart` / `kit.hbar_chart` | **matplotlib 出图**<br>`scripts/figstyle.py` |
|---|---|---|
| 可编辑 | 是。双击进 Excel 改数 | 否。图就是一张 PNG |
| 体积 | 20–60 KB | 200 KB – 3 MB（PNG @200dpi） |
| 样式控制 | 有限：颜色、字号、数据标签能控，网格线/轴脊/刻度格式基本控不了 | 完全自由 |
| 比例 | 天然正确（矢量） | **必须自己管**：`tight=False` 出图 + `kit.fig_fit()` 等比放置 |
| 闸门 | 不触发图片拉伸那两道 | 闸门一按实测 ar 核对，闸门二按 PDF 落位核对 |
| 适合 | 柱状图、条形图、饼图、折线图（单系列） | 多子图、曲线族、箱线/热力、对数轴、误差棒、特殊标注 |

**默认选原生图表。** 图能用原生图表就别用图片：可编辑、体积小几十倍、
改一个数不用重跑出图管线，而且天然没有比例问题。
只有当原生图表画不出你要的样子（多子图、对数轴自定义刻度、缺失值标注、误差棒）时才退回 matplotlib。

```python
# 原生：一句话
kit.column_chart(s, x, y, w, h, ["仅词面", "＋语义"], (0.41, 0.63),
                 [kit.ACC, BAND], series_name="MRR")

# 原生：类别名长 → 换横排
kit.hbar_chart(s, x, y, w, h, ["跨模态预训练", "结构化倒排"], (128, 640),
               [kit.ACC, BAND], num_fmt="0")

# matplotlib：只有原生画不出来时才走这条
import figstyle
figstyle.use_preset("default")
fig, ax = figstyle.plt.subplots(figsize=(6.0, 3.2))   # 比例写死在 figsize 上
bars = ax.bar(names, vals, width=0.55, color=figstyle.SERIES[0],
              edgecolor="white", linewidth=0.6)
figstyle.label_bars(ax, bars, values=vals)
figstyle.despine(ax); figstyle.grid(ax, "y"); figstyle.headroom(ax, max(vals))
figstyle.finish(fig, out_png)      # tight 默认 False
```

---

## 2. 原生柱状图的样式（`kit.style_chart`）

这套样式是固定的，理由和可调项都在参数名里：

| 做法 | 参数 | 为什么 |
|---|---|---|
| 不画标题、不画图例 | `has_title=False` / `has_legend=False` | 幻灯片上标题是页标题，图表再自带一个标题就是重复；单系列图例是纯噪音 |
| 隐藏数值轴与网格 | `_hide_axis(va)` | 数字都标在柱端了，轴和网格是第二份同样的信息，还占地方 |
| 类间距 62% | `GAP_WIDTH` | 柱宽约为槽宽 1.6 倍：柱子不粘连，也不散成孤岛 |
| 逐点填色 + 去边框 | `pt.format.fill` / `line.fill.background()` | 逐点填色才能让「基线 / 本版 / 对照」用不同颜色；描边会在投影上糊成一片 |
| 数据标签外端、Times New Roman 粗体 | `value_labels(..., OUTSIDE_END)` | 标在柱端，读者不必去对照坐标轴；粗体让它在投影上仍然可读 |
| 顶部留 8% | `HEADROOM = 0.08` | 标签画在数据之上，轴上限贴着柱顶时标签会被裁掉 |
| 类目标签 14pt、正文色 | `lbl_size` / `cat_color` | 类名是横排的第一关键字，比数值更需要字号 |

**两个可调项值得单独调**：`dl_size`（投影远时调到 18）、`num_fmt`（整数用 `"0"`，
不要让 640 显示成 `640.000`）。

---

## 3. 缺失值：标「未参评」，绝不画成 0

这是本文件最重要的一条。`values` 里放 `None` 表示这一格**没有数据**：

```python
kit.column_chart(s, x, y, w, h,
                 ["仅词面", "＋语义", "＋图谱", "＋融合"],
                 (0.41, 0.63, 0.58, None),          # 最后一个未参评
                 [kit.ACC, kit.ACC, BAND, kit.ACC])
# → 第四格不画柱，位置上是灰色斜体「未参评」
```

**为什么不能画成 0 柱**：0 柱会被读成「测出来是 0」。而「没测」和「测了是 0」
是两件完全不同的事，画成同一根柱子，等于把「我们没数据」讲成了「我们数据不好」——
这是把缺失误报成结果。灰色 + 斜体是双重编码，黑白打印也认得出来。

matplotlib 侧对应 `figstyle.mark_missing()`，且 `label_bars(ax, bars, values=vals)`
**必须把原始 `values` 传进去**：不传的话它会读到 0 高的占位柱并标成 `0.00`。

---

## 4. 量级差巨大时怎么办

差两个数量级以上时，线性轴会把小的几根压成一条线——读者会以为那几项没数据。
三种解法，按优先级：

1. **拆成两张图**（首选）。两个量纲两张图，各自的坐标轴都从 0 开始。
   代价是版面要两块地，收益是每个数都读得准。
2. **对数轴**。等距离代表等倍数，适合「一个数量级内讲倍数关系」的场景。
   配合 `set_xticks([])` 隐藏刻度、只留网格——十个 `10⁰/10¹/10²` 标签只是噪音。

   ```python
   # 原生
   kit.column_chart(s, ..., log_axis=True)      # 自动按 LOG_HEADROOM=1.6 留标签空间
   # matplotlib
   figstyle.log_axis(ax, "y", base=10, hide_ticks=True)
   figstyle.grid(ax, "y")
   ```
   对数轴的坑：**要求所有非 None 值 > 0**，0 和负数画不出来；标签空间是乘性的
   （顶部留 1.6 倍，不是加 8%），代码里已经按这个算好了。
3. **归一化**（谨慎）。换成「相对基线的倍数」或百分比，量级差消失，
   但**绝对值的信息也消失了**。只在你已经明确说过绝对值的前提下用。

---

## 5. matplotlib 十条技法

`scripts/figstyle.py` 是这十条的实现，示例跑一遍就看到成品：

```powershell
python scripts\figstyle.py        # → assets/figures/*.png（示例数据全部虚构）
```

**① 字体族：中文在前，西文回落 Times New Roman**

```python
figstyle.use_preset()      # 内部：FAMILY = ["KaiTi", "Times New Roman"]
plt.rcParams["axes.unicode_minus"] = False
```

*为什么*：matplotlib 一个文本只能挂一个 family。中文在前，西文自动回落到与
pptx 正文同源的 Times New Roman——图里出现的数字和页面上的数字看起来是同一套字。
`unicode_minus=False` 是必需的：默认用 U+2212 减号，回退字体常常没有这个码位，
负刻度会渲染成方框。`figstyle.py` 只挑**本机确实注册过**的族名，
否则中文全是方框。

**② 白底三件套**

```python
plt.rcParams.update({"figure.facecolor": "white",
                     "axes.facecolor": "white",
                     "savefig.facecolor": "white"})
```

*为什么*：三个都要设。只设前两个的话，不透明背景可能被 `savefig` 的 rcParam
重新涂一层；深色主题下会出现「页面白、图里灰」的对不齐。

**③ `despine()`：top / right 一律隐藏**

```python
figstyle.despine(ax)                      # 只留 left + bottom
figstyle.despine(ax, keep=("bottom",))    # 横向条形图通常只留 bottom
```

*为什么*：没有第二量纲时，top / right 两条边是纯装饰。它们把画面框成盒子，
读者会去读「盒子的边界」这个不存在的信息。

**④ 网格单向 + `set_axisbelow(True)` + 线宽 0.6**

```python
figstyle.grid(ax, "y", lw=0.6)
```

*为什么*：双向网格等于每根柱子都被横竖线切开，柱宽的视觉误差比柱高差还大。
`set_axisbelow(True)` 不设的话网格画在数据之上，线会穿过柱子。

**⑤ 柱用白边分隔**

```python
ax.bar(x, v, width=0.55, color=figstyle.SERIES[0], edgecolor="white", linewidth=0.6)
```

*为什么*：相邻同色柱在投影上会糊成一块；0.6pt 白边是比调整柱宽更省事的分隔手段。

**⑥ 数据标签直接标在柱端，格式自适应**

```python
figstyle.label_bars(ax, bars, values=vals)   # < 100 两位小数，否则取整
```

*为什么*：自适应格式是为了避免同一张图里 `0.74` 和 `1024` 用同一种小数位数
（要么拖一串 0，要么挤成一团）。

**⑦ 顶部 / 右端留白**

```python
figstyle.headroom(ax, ymax, ratio=1.24, xmax=None)   # ylim(0, ymax*1.24)
```

*为什么*：标签画在数据之上，轴上限贴着最高那根柱顶时标签会被裁掉，
而「标签被裁掉一半」在缩略图上看不出来。

**⑧ 对数刻度 + 自定义刻度标签**

```python
figstyle.log_axis(ax, "y", base=10, labels={10: "10", 100: "百", 1000: "千"},
                  hide_ticks=True)
```

*为什么*：量级差图里读者要的是「量级」这个信息本身，十个 `10ⁿ` 标签只是噪音。

**⑨ 缺失值标灰斜体「未参评」**

```python
figstyle.mark_missing(ax, 3, 0.012)          # 位置与 0 柱顶端对齐
```

*为什么*：见第 3 节。灰色 + 斜体双重编码，投影和黑白打印都认得出。

**⑩ 最大的坑：不要用 `tight_layout`，不要用 `bbox_inches="tight"`**

```python
figstyle.finish(fig, "out/fig.png")                    # ✅ 默认 tight=False
fig.tight_layout(); fig.savefig(p, bbox_inches="tight")  # ❌ 两个都不要
```

*为什么*：

- `tight_layout()` 与手工给的 `GridSpec` 边距互相打架——一边要贴边，
  一边要固定边距，最后谁也不听谁的，图与图之间还会挤在一起。
- `bbox_inches="tight"` 按内容裁掉画布四周，于是**同一份代码在不同数据下
  产出不同像素尺寸的比例**。pptx / docx 里按固定比例放置时就被拉伸变形，
  而且闸门一看不出来：清单里的 `ar` 是实测值，它是「对」的，错的是放置比例。

所以边距一律用 `GridSpec(left/right/top/bottom/wspace/hspace)` 手工给，
`figsize` 就是最终比例，落盘后除一下就知道该占多宽。

```python
fig = plt.figure(figsize=(7.2, 3.0))
gs = fig.add_gridspec(1, 3, left=0.07, right=0.985, top=0.88, bottom=0.14, wspace=0.30)
```

---

## 6. 插进 pptx 的最后一步

`finish()` 落的是位图，插进 pptx 时**必须走 `kit.fig()` / `kit.fig_fit()`**：

```python
kit.fig_fit(s, x, y, w, slot_h, str(png))    # 等比 + 双向居中
kit.cap(s, x, y_caption, w, "图 1", "说明文字")
```

`fig_fit` 用 pillow **实测**原图宽高比，不是让调用者手填常数——手填对不上时，
闸门一（清单里的 ar）和闸门二（PDF 里的落位）会同时把正常图判成「拉伸」。
宁可让图小一点，也不要变形。

---

## 7. 三套配色预设

`figstyle.use_preset("default" | "dense" | "mono")`：

| 预设 | 用于 | 特点 |
|---|---|---|
| `default` | 单图 / 双图（最常用） | 10pt 底字，蓝色主色 + 浅蓝次色 + 陶土橙强调 |
| `dense` | 2×2 及以上多子图 | 8.5pt 底字，边距收窄，线宽减一档 |
| `mono` | 打印稿、外部评审 | 全灰阶单色柱，只留 bottom 轴 |

`mono` 存在的理由不是「极简好看」，是**少一种颜色就少一个「是不是颜色在表达立场」
的问题**。图里出现的强调色（陶土橙 `#B5532C`，白底 4.95:1）只用于标注，
不用于分类——分类靠位置和标签，不靠颜色。

---

## 8. 示例图清单

`python scripts/figstyle.py` 生成的五张（数据全部虚构，一眼假）：

| 文件 | 演示什么 | 比例 |
|---|---|---|
| `fig_demo_bar.png` | 柱端标签 + 未参评（技法 1–7） | 1.875 = 6.0/3.2 |
| `fig_demo_log.png` | 对数轴 + 隐藏刻度（技法 8） | 1.875 |
| `fig_demo_panels.png` | 多子图 + GridSpec 手工边距（技法 10） | 2.400 = 7.2/3.0 |
| `fig_demo_mono.png` | `mono` 预设 | 1.875 |
| `fig_demo_latency.png` | 示例 deck 第 14 页用的那张 | 1.875 |

比例全部等于 `figsize` 的比值 = 没被裁切。重新生成后先核这一条：
比例对不上，说明有人打开了 `tight`。
