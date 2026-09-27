---
name: research-ppt-builder
description: 生成中文科研汇报 PPT（组会、项目答辩、开题/中期、学术汇报）时使用。基于用户自己的 .pptx 模板或内置中性母版出稿，自动排版、生成演讲者备注与讲稿，并用四道自动闸门（文字溢出 / 渲染几何 / 对比度 / 版面空隙）验收。内置三套配色主题（academic-blue / scholar-red / minimal-mono）与一套 matplotlib 图表风格预设。触发词包括：组会 PPT、汇报 PPT、答辩 PPT、开题/中期报告 PPT、报告幻灯片、仿某个 pptx 模板出稿、把论文/实验做成幻灯片、科研配图排版。
---

# 中文科研汇报 PPT 生成器

## 什么时候用这个 skill

- 用户要做**中文**科研汇报：组会、项目答辩、开题/中期、学术报告、对外介绍；
- 用户想**基于自己的 pptx 模板**出稿（「按我们组的那个模板做一份」）；
- 用户手上有**科研图片**（实验曲线、架构图、论文原图）要排进幻灯片；
- 用户要**演讲者备注 / 讲稿 / 预设问答**。

不适用：纯英文或纯代码技术分享、营销型演示、要求复杂动画或交互的场合。

## 两条硬纪律

**1. PPT 里没有肉眼能查的文本框。** 文字溢出时 LibreOffice 静默裁掉、PowerPoint 里
直接漏到框外。**四道闸门不是可选项**，改完任何东西都要重跑。

**2. 内容口径纪律。** 不确定的口径不写；自评数据要照实交代未终审 / 未重复 / 假设区间 /
样本量小只看相对排序。详见 `references/narration.md`。

---

## 端到端工作流

### 步骤 1 · 收集用户输入

一次性问清这六项，缺项就问，不要替用户猜：

| 项 | 例子 | 缺了会怎样 |
|---|---|---|
| 使用场景 / 受众 | 组会（导师 + 师兄）/ 答辩（评委）/ 开题 | 页型取舍完全不同 |
| 时长或页数 | 15 分钟 ≈ 12 页 | 决定每页信息密度 |
| 模板 | 自带 `.pptx` / `.potx` 路径，或用内置 | 决定走不走步骤 2 |
| **有无图片及图片清单** | 见步骤 5 三分支 | 决定整份稿的骨架 |
| 配色偏好 | 沿用模板 / 深色 / 指定色 | 决定用哪套内置主题，见下表 |
| 中英文字体 | 中文微软雅黑/楷体；西文默认 Times New Roman | 决定 `theme.json` 的 `fonts` |

**主题三选一**（三套的栅格与字体完全相同，只换色板与母版装饰色）：

| 主题 | 什么时候用 |
|---|---|
| `academic-blue`（默认） | 组会、学术会议、对外介绍。亮投影环境。**没明确要求就用它** |
| `scholar-red` | 开题/中期、答辩、项目评审。正式场合，红色更「像结论」；红底在偏色投影上会偏紫 |
| `minimal-mono` | 打印稿、投稿附件、外部评审；不希望颜色先入为主时。亮暗环境都可 |

```powershell
python scripts\make_neutral_template.py --theme scholar-red   # 首次切换时生成母版
python scripts\build_deck.py --theme scholar-red                # 出稿
```
或代码里 `kit.load_theme(name="minimal-mono")`（必须在建 deck 之前调用）。
色板定义在 `kit.BUILTIN_THEMES`，母版生成脚本读的是同一份。

**图片情况必须让用户显式声明**（有图 / 无图 / 有图但文件名是 `1.png` 这种）。
不要靠目录里有没有图片文件来猜。

### 步骤 2 · 模板接入

- **用内置母版**：不用管，`assets/theme-neutral.pptx`（= `assets/themes/academic-blue.pptx` 的别名）
  零机构标识、零位图，开箱可用。换配色用 `assets/themes/<name>.pptx`，见步骤 1 的主题表。
- **用用户自己的模板**：

  ```powershell
  python scripts\analyze_template.py "D:\…\我的模板.pptx" -o theme.json
  ```

  照 `references/template-intake.md` 补完 `palette` 的语义映射与 `grid` 核对，
  页面代码一个字都不用改。版式名对不上时脚本会**明确报错并列出所有版式名**，
  按提示改名或改 `theme.json` 的 `layouts`——**不要绕过它**。
  `.potx` 与 `.pptx` 同一条路径，可直接用。

  换基座还有一条命令行入口 `build_deck.py --template <路径>`，但它**只换基座文件**：
  `palette/fonts/grid/layouts` 仍以 `theme.json` 为准（版式名对不上要在 `theme.json` 里指，
  `--template` 不碰它）。接新模板走 `analyze_template.py` 那条路。

  模板带校名/校徽/机构照片时，**明确提醒用户自行确认使用范围**。

### 步骤 3 · 内容规划（**先出大纲，确认后才写代码**）

交一份页面大纲给用户确认，格式：

```markdown
| # | 页标题 | 这一页干什么 | 配图 |
|---|---|---|---|
| 1 | 封面 | 眉标 + 标题 + 范围 + 三行元信息 | — |
| 2 | 本场安排 | 四步 + 一句话总纲 + 口径先说清 | — |
| 3 | 检索栈拆成三层 | 公式 + 两栏对照 + 结论带 | — |
```

**用户没确认之前不要开始写页面代码。** 页面骨架错了，重写比确认贵得多。

同时确认**讲稿**：要不要演讲者备注、要不要 markdown 讲稿、要不要预设问答。

### 步骤 4 · 生成

1. 复制 `scripts/build_deck.py` 为你的 `build.py`；
2. 改 `DECK`（标题、页脚、页数）与 `PAGES`（页面函数清单）；
3. 每个页型从 `references/page-recipes.md` 里挑，照几何改文案；
4. 视觉规格照 `references/design-language.md`（色板、栅格、字体三条铁律都在里面）。

```powershell
python scripts\build_deck.py                       # → out/示例汇报_中文科研组会.pptx
python scripts\build_deck.py --zh-kaiti            # 中文改楷体
python scripts\build_deck.py --template "D:\…\我的模板.pptx"   # 只换基座
python scripts\build_deck.py --help                # 全部参数
```

产物三样：`*.pptx`、`*_layout.json`（版面清单，闸门一的输入，不是调试产物）、
`*_讲稿.md`。

### 步骤 5 · 配图（按用户声明的分支走）

完整规则见 `references/image-handling.md`，摘要：

- **无图** → 全部走纯排版页型（数据大数字、对照表、原生柱状图、三栏卡片）。
  完全成立，不要为了「有图」硬凑。
- **有图且已明确命名** → 直接按用户指定的页码与页内位置插入，**不做任何识别或改名**。
- **有图但命名模糊**（`1.png` `image1`） → **自己逐张识别**内容，决定放哪页、什么位置、
  配什么图注，然后**把识别清单交用户确认**，不允许默默决定。

图注三统一：序号 `图 N` 连续、来源集中交代、比例等比不拉伸。
**等比 + 居中是硬要求**，用 `kit.fig()` / `kit.fig_fit()`，别自己算宽高。

**但「有图」不等于「出图」——图能用原生图表就别用图片。**
柱状图、条形图、单系列折线一律走 `kit.column_chart` / `kit.hbar_chart`：
可编辑（听众或用户想改一个数时能直接双击进 Excel）、体积小几十倍、
天然没有比例问题、不触发图片拉伸那两道闸门。
只有需要**多子图、对数轴自定义刻度、未参评标注、误差棒、非柱非线的图型**时，
才用 `scripts/figstyle.py` 出位图；出图后**必须** `tight=False` 落盘
（`bbox_inches="tight"` 会按内容裁画布，插进 pptx 必然被拉变形），
再用 `kit.fig_fit()` 等比放置。分工与十条技法见 `references/figures-and-charts.md`。

### 步骤 6 · 讲稿

讲稿的**唯一来源**是 `kit.SCRIPT`（`{页码: (标题, 计时, 台词)}`，定义在 `kit.py` 里），
同一份内容同时产出 PPT 内嵌备注与 markdown（`kit.notes(s, pn=…)` +
`kit.write_script_md()`）。页面脚本只往 `kit.SCRIPT` 填一处，两边都从它读。
写法、计时、预设问答、口径纪律见 `references/narration.md`。

改页面就要改讲稿：讲稿里的每个数字都要和页面上的对得上。

### 步骤 7 · 四道闸门验收

```powershell
python scripts\verify_all.py
```

soffice 不在 PATH 上时要给全路径：

```powershell
python scripts\verify_all.py --soffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

| 闸门 | 抓什么 |
|---|---|
| `qa_fit.py` | 文字溢出、表格压页脚、图片拉伸、越界（源码几何层，唯一能在渲染前抓到溢出） |
| `qa_pdf.py` | 越界、压页脚、文本行重叠、图片拉伸、稀页（渲染后真实落位） |
| `qa_visual.py` | 逐片段对比度、文字出框（前两道都看不到颜色） |
| `qa_layout.py` | 文字越出组件、大空隙、提前收工、偏空（好不好看） |

**不通过就改，改完必须重跑。** 阈值与调法见 `references/qa-gates.md`——
先查是不是量错了，别直接放宽阈值。

退出码：`0` 四闸全过 / `1` 有闸门失败 / `2` 有闸门因缺 LibreOffice 被 SKIP。
**CI 只看退出码，`2` 不是通过**——它意味着这一稿还没验过渲染几何、对比度与版面。

### 步骤 8 · 交付

交三样，并在交付消息里说清：

1. `*.pptx` —— 汇报用；
2. `*_讲稿.md` —— 提前背稿、预设问答；
3. 渲染 PDF（`out/render/*.pdf`）—— 供用户抽查版面，不必自己逐页看。

---

## 目录索引

| 文件 | 什么时候读 |
|---|---|
| `references/design-language.md` | 排版前必读：色板（含实测对比度与禁用场景）、**三套主题**、栅格常量、字体策略与两条实现铁律 |
| `references/figures-and-charts.md` | 有任何图表时：原生图表 vs matplotlib 出图的分工、原生样式参数、量级差三种解法、matplotlib 十条技法与那个最大的坑 |
| `references/template-intake.md` | 用户自带模板时 |
| `references/page-recipes.md` | 写每一页之前：十五种页型的几何与排法（含时间线、四象限、KPI 数字墙、two_col 两栏页） |
| `references/image-handling.md` | 有图时 |
| `references/narration.md` | 写讲稿、订口径 |
| `references/qa-gates.md` | 闸门报错要调阈值时 |
| `references/pitfalls.md` | 任何异常行为、渲染怪相、结果不对时先查这里 |

## 脚本

| 脚本 | 作用 |
|---|---|
| `scripts/kit.py` | 组件库：色板/栅格/字体（读 `theme.json`）、`txt/rect/panel/rule/fig/fig_fit/head/foot/chip/headbar/table`、`column_chart/hbar_chart/value_labels/missing_marks`、`open_base/start_deck/save_deck/notes/write_script_md`、`scrub_metadata`（擦 docProps） |
| `scripts/build_deck.py` | 十五页示例 deck（纯排版路线 + 四道闸门的回归基线），新稿照它改 |
| `scripts/figstyle.py` | matplotlib 风格预设（`use_preset` / `despine` / `label_bars` / `mark_missing` / `finish`），三套预设；直接跑一次生成 `assets/figures/` 示例图 |
| `scripts/analyze_template.py` | 从任意 `.pptx` 反推视觉语言，输出 `theme.json` |
| `scripts/make_neutral_template.py` | **可复现**地重新生成母版：`--theme academic-blue`（默认）/ `scholar-red` / `minimal-mono`，并打印该主题每个颜色的实测对比度 |
| `scripts/qa_fit.py` `qa_pdf.py` `qa_visual.py` `qa_layout.py` | 四道闸门，阈值都在文件顶部 |
| `scripts/verify_all.py` | 跨平台一键：构建 → 渲染 → 四闸门 |

改 `kit.py` 或母版之后，跑一遍 `verify_all.py`——示例 deck 是回归基线，
全绿说明没把基线搞坏。**换了主题也要跑**：新配色没过对比度时，只有闸门三会拦。
