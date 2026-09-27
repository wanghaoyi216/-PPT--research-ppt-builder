# -*- coding: utf-8 -*-
"""
科研配图的 matplotlib 风格预设
================================

**先决定要不要用它**
--------------------
这个包里的图，优先用**原生 PowerPoint 图表**（`kit.column_chart` / `kit.hbar_chart`）：
可编辑、体积只有几十 KB、数值跟着数据走、不触发图片拉伸那两道闸门。
只有这几种情况才用 matplotlib 出图：需要多子图、需要非柱非线的图型（曲线/箱线/热力）、
需要复杂标注（对数轴自定义刻度、未参评标注、误差棒、参考区间）。
完整分工见 references/figures-and-charts.md。

**它生成的是位图**
------------------
本模块的 `finish()` 落的是 PNG，插进 pptx 之后就是一张图：**不可编辑、不可追溯数据、
改一个数要重新出图**，而且必须按原比例放置。

所以插进 pptx 时**必须走 `kit.fig()` / `kit.fig_fit()`**——它们用 pillow 实测原图宽高比，
等比缩放 + 双向居中。手工按猜测填宽高，闸门二（qa_pdf）会直接报「图片拉伸」；
闸门一（qa_fit）也会拿清单里的 ar 核对。宁可让图小一点，也不要变形。

**最关键的坑：不要用 tight_layout，也不要用 bbox_inches="tight"**
----------------------------------------------------------------
`fig.tight_layout()` 与手工给的 `GridSpec` 子图间距互相打架：一边要贴边，
一边要留出固定边距，最后的结果既不是你要的边距，图与图之间还会挤在一起。
`savefig(bbox_inches="tight")` 更麻烦——它会**按内容裁掉画布四周**，于是同一份代码
在不同数据下产出不同像素尺寸的比例；pptx / docx 里按固定比例放置时就被拉伸变形，
而且这个变形在闸门一之前根本看不出来（清单里的 ar 是实测值，它是"对"的，
错的是你以为的放置比例）。

`finish(fig, path, tight=False)` 的默认就是 `tight=False`：按 `figsize` 原样落盘，
**比例完全由 figsize 决定**，插进 pptx 时除一下就知道该占多宽。
子图边距一律用手工 `GridSpec` 的 `left/right/top/bottom/wspace/hspace` 表达。

三套预设
--------
    default  单图 / 双图，10pt 底字，白底 + 灰蓝系      （最常用）
    dense    多子图（2×2 及以上），8.5pt 底字，边距收窄   （拼版/消融图）
    mono     极简黑白：无彩色、单色柱、只留 bottom 轴      （打印稿、外部评审）

用法::

    import figstyle, matplotlib.pyplot as plt
    figstyle.use_preset("default")                 # 一次就够，后面全继承 rcParams
    fig, ax = plt.subplots(figsize=(6.0, 3.2))     # 比例写死在 figsize 上
    bars = ax.bar(names, vals, color=figstyle.SERIES[0], edgecolor="white", linewidth=0.6)
    figstyle.label_bars(ax, bars)
    figstyle.despine(ax)
    figstyle.grid(ax, "y")
    figstyle.headroom(ax, max(vals))
    figstyle.finish(fig, "out/fig_ablation.png")    # tight 默认 False
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")            # 无显示环境（构建机 / CI）必须先切后端
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter

# ── 字体 ──
# matplotlib 一个文本只能挂一个 family；给列表，新版按字符回退、旧版整体回退。
# 所以先显式把字体文件注册进来，再只挑**本机确实有**的族名，
# 免得回退到 DejaVu Sans —— 那套没有汉字，图上会是一排方框。
_FONT_FILES = [r"C:\Windows\Fonts\simkai.ttf",        # 楷体
               r"C:\Windows\Fonts\times.ttf",         # Times New Roman
               r"C:\Windows\Fonts\timesbd.ttf",
               r"C:\Windows\Fonts\msyh.ttc",          # 微软雅黑
               r"C:\Windows\Fonts\simhei.ttf"]
for _f in _FONT_FILES:
    if Path(_f).exists():
        try:
            font_manager.fontManager.addfont(_f)
        except Exception:          # 字体文件损坏/不受支持：跳过，不该让整张图挂掉
            pass


def _available(names):
    have = {f.name for f in font_manager.fontManager.ttflist}
    return [n for n in names if n in have]


# 中文在前、西文回落到 Times New Roman：学术语境正式，且与 pptx 正文字体同源，
# 图里出现的数字和页面上的数字看起来是同一套字。
FAMILY = _available(["KaiTi", "Microsoft YaHei", "SimHei"]) + ["Times New Roman"]

# ── 色板（与 kit.BUILTIN_THEMES 同源思路，图内颜色也要过对比度）──
INK = "#22262B"        # 主文字 12.9:1 白底
MUTED = "#5A646E"      # 次文字 5.6:1
GRID = "#DDE3E8"       # 网格线（不承字）
AXIS = "#8A939B"       # 轴线与刻度
MISSING = "#9AA1A7"    # 缺失值标注（灰）5.0:1 白底，且用斜体双重编码
SERIES = ["#1F6FA8", "#7FA9C9", "#B5532C", "#2F2F2F"]   # 主 / 次 / 强调 / 深

# 竖排图与 docx / pptx 里按固定比例放置的图，最常用的是 10.5 × 5.46（≈1.92:1）
FIGSIZE = (6.0, 3.2)

# ── 三套预设 ──
# 值都是 matplotlib rcParams 的键；键名写错不报错、只是静默不生效，
# 所以改完预设务必出一张图看一眼。
PRESETS = {
    "default": {
        "font.size": 10, "axes.labelsize": 10, "xtick.labelsize": 9, "ytick.labelsize": 9,
        "legend.fontsize": 9,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.9,
        "axes.labelcolor": INK, "text.color": INK,
        "xtick.color": AXIS, "ytick.color": AXIS,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
        "lines.linewidth": 1.6,
    },
    "dense": {
        # 多子图：字小一号、边距收窄。图多的时候字号差一级，投影上就分不清主次了。
        "font.size": 8.5, "axes.labelsize": 9, "xtick.labelsize": 8, "ytick.labelsize": 8,
        "legend.fontsize": 8,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.7,
        "axes.labelcolor": INK, "text.color": INK,
        "xtick.color": AXIS, "ytick.color": AXIS,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
        "lines.linewidth": 1.3,
    },
    "mono": {
        # 极简黑白：灰阶柱 + 只留 bottom 轴。打印稿和外部评审用，
        # 少一种颜色就少一个「是不是颜色在表达立场」的问题。
        "font.size": 10, "axes.labelsize": 10, "xtick.labelsize": 9, "ytick.labelsize": 9,
        "axes.edgecolor": INK, "axes.linewidth": 0.8,
        "axes.labelcolor": INK, "text.color": INK,
        "xtick.color": INK, "ytick.color": INK,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
        "lines.linewidth": 1.4,
    },
}

_AXIS_TONES = {"default": SERIES[0], "dense": SERIES[0], "mono": "#5A5A5A"}


def use_preset(name: str = "default") -> str:
    """套用一套 rcParams 预设（含字体族与 `axes.unicode_minus=False`）。返回预设名。

    `unicode_minus=False` 是**必需**的：默认会用 U+2212 减号，Times New Roman 之外
    的回退字体常常没有这个码位，负刻度会渲染成方框。关掉它改用 ASCII 连字符。
    """
    if name not in PRESETS:
        raise KeyError(f"没有预设「{name}」。可选：{sorted(PRESETS)}")
    matplotlib.rcParams.update(PRESETS[name])
    matplotlib.rcParams["font.family"] = FAMILY
    matplotlib.rcParams["axes.unicode_minus"] = False
    return name


def despine(ax, keep=("left", "bottom")):
    """只留下 keep 里点名的轴脊，其余一律隐藏。

    为什么砍：top / right 两条边在只有一个系列、没有第二量纲的图里是纯装饰，
    它们把画面框成一个盒子，读者会去读「盒子的边界」这个不存在的信息。
    留 left+bottom 是习惯，横向条形图通常只留 bottom。
    """
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(s in keep)


def grid(ax, which="y", lw=0.6, color=GRID):
    """**单向**网格 + `set_axisbelow(True)`：网格在数据下面，不穿过柱子。

    双向网格等于每根柱子都被横竖线切开，柱宽的视觉误差比柱高差还大。
    `set_axisbelow(True)` 不设的话 matplotlib 默认把网格画在数据之上。
    """
    ax.grid(True, axis=which, color=color, linewidth=lw, linestyle="-")
    ax.set_axisbelow(True)
    return ax


def label_bars(ax, bars, values=None, fmt=None, size=9, pad=0.02, color=None, bold=True):
    """柱端数据标签，格式自适应：小数 < 100 保留两位，≥ 100 取整。

    直接标在柱端比让读者去对照坐标轴快得多——坐标轴已经在 `despine` 里
    砍掉了，没有轴就没有刻度可读。自适应格式是为了避免同一张图里
    「0.74」和「1024」用同一种小数位数（要么长要么全是 0）。

    `values` 传原始数据（可含 None）时按它决定标不标：缺失的那一格画成 0 高占位柱
    只是为了留住槽位，**高度是 0 不代表数据是 0**——不传 `values` 的话它会被
    标成「0.00」，等于把「没测」讲成了「测出来是 0」。要标缺失值就用 `mark_missing`。
    """
    vals = list(values) if values is not None else (
        [b.get_height() for b in bars] if hasattr(bars, "__iter__") else list(bars))
    out = []
    for i, (b, v) in enumerate(zip(bars, vals)):
        if v is None:                       # 缺失值交给 mark_missing，绝不标成 0
            continue
        s = fmt(v) if fmt else (f"{v:.2f}" if abs(v) < 100 else f"{v:.0f}")
        x = b.get_x() + b.get_width() / 2 if hasattr(b, "get_x") else i
        ax.text(x, v + v * pad if v else pad, s, ha="center", va="bottom",
                fontsize=size, fontweight="bold" if bold else "normal",
                color=color or INK)
        out.append(s)
    return out


def mark_missing(ax, x, y, text="未参评", size=9, color=MISSING):
    """把缺失值标成灰色斜体文字，**不要画成 0 柱**。

    这是本模块最要紧的一条纪律：0 柱会被读成「测出来是 0」，
    「没测」和「测了是 0」是两件完全不同的事，画成同一根柱子就是把
    「我们没数据」讲成了「我们数据不好」。斜体 + 灰是双重编码，
    黑白打印出来也认得出来。
    """
    return ax.text(x, y, text, ha="center", va="bottom", fontsize=size,
                   style="italic", color=color)


def headroom(ax, ymax, ratio=1.24, xmax=None):
    """给柱端标签留出顶部空间：`ylim(0, ymax*ratio)`，可选把 x 轴也放开到 `xmax`。

    标签画在数据之上，轴上限贴着最高那根柱顶时标签就被裁掉了——
    而「标签被裁掉一半」这件事在缩略图上看不出来，等到投影才发现。
    """
    ax.set_ylim(0, ymax * ratio)
    if xmax is not None:
        ax.set_xlim(0, xmax)
    return ax


def log_axis(ax, which="x", base=10, labels=None, hide_ticks=False):
    """对数刻度 + 可选的自定义刻度标签。

    `labels` 是 `{刻度值: 标签文字}`；`hide_ticks=True` 时把刻度完全隐藏，
    只留网格 —— 量级差图里读者要的是「量级」这个信息本身，
    十个 10⁰/10¹/10² 的刻度标签只是噪音。
    """
    ax.set_xscale("log", base=base) if which == "x" else ax.set_yscale("log", base=base)
    if labels:
        axis = ax.xaxis if which == "x" else ax.yaxis
        axis.set_major_formatter(FuncFormatter(lambda v, _p: labels.get(v, "")))
    if hide_ticks:
        (ax.set_xticks([]) if which == "x" else ax.set_yticks([]))
    return ax


def finish(fig, path, tight=False, dpi=200):
    """落盘并关图。**tight 默认 False**，理由见模块 docstring。

    传 `tight=True` 之前先想清楚：pptx / docx 里按固定比例放置时，
    被 `bbox_inches="tight"` 裁过的图比例会随数据漂移，然后被拉变形。
    确实需要贴边时，用 `fig.subplots_adjust(...)` 手工调，别用 tight。
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=dpi,
                bbox_inches="tight" if tight else None,
                pad_inches=0.12 if tight else None)
    plt.close(fig)
    print(f"  ✓ {path.name}  ({path.stat().st_size / 1024:.0f} KB)")
    return path


# ══════════════════════════ 示例（python scripts/figstyle.py 可直接跑）══════════════════════════
# 数据全是编的，一眼假：只为演示这套风格长什么样、比例怎么定。
def _demo():
    out = PKG_FIGURES = Path(__file__).resolve().parent.parent / "assets" / "figures"

    # 1. 分组柱 + 未参评：技法 1–7
    use_preset("default")
    labels = ["仅词面", "＋语义重排", "＋图谱扩展", "＋融合排序"]
    vals = [0.41, 0.63, 0.58, None]            # 最后一个是「未参评」，不是 0
    fig, ax = plt.subplots(figsize=FIGSIZE)
    bars = ax.bar(labels, [v or 0 for v in vals], width=0.55,
                  color=[SERIES[0], SERIES[1], SERIES[1], "white"],
                  edgecolor="white", linewidth=0.6)
    label_bars(ax, bars, values=vals)
    mark_missing(ax, 3, 0.012)                  # 位置与 0 柱顶端对齐
    despine(ax)
    grid(ax, "y")
    headroom(ax, 0.63, 1.24)
    finish(fig, PKG_FIGURES / "fig_demo_bar.png")

    # 2. 量级差 → 对数轴 + 隐藏刻度只留网格（技法 8）
    use_preset("default")
    names = ["配置甲", "配置乙", "配置丙", "配置丁"]
    big = [12, 74, 310, 1580]
    fig, ax = plt.subplots(figsize=FIGSIZE)
    ax.bar(names, big, width=0.55, color=SERIES[0], edgecolor="white", linewidth=0.6)
    label_bars(ax, ax.containers[0])
    despine(ax, keep=("bottom",))
    log_axis(ax, "y", base=10, hide_ticks=True)
    grid(ax, "y")
    ax.set_ylim(1, 2600)
    finish(fig, PKG_FIGURES / "fig_demo_log.png")

    # 3. 多子图（dense 预设）——边距用 GridSpec 手工给，不用 tight_layout
    use_preset("dense")
    fig = plt.figure(figsize=(7.2, 3.0))
    gs = fig.add_gridspec(1, 3, left=0.07, right=0.985, top=0.88, bottom=0.14, wspace=0.30)
    for i, (t, v) in enumerate([("召回", [3, 5, 5, 5]), ("排序", [0.41, 0.63, 0.58, 0.71]),
                                ("时延", [12, 47, 19, 143])]):
        axi = fig.add_subplot(gs[0, i])
        bars = axi.bar(range(4), v, width=0.6,
                       color=SERIES[i % len(SERIES)] if i == 0 else "#5A646E",
                       edgecolor="white", linewidth=0.6)
        label_bars(axi, bars, size=7.5)
        despine(axi)
        grid(axi, "y")
        headroom(axi, max(v), 1.30)
        axi.set_xticks([])
        axi.set_title(t, fontsize=9, pad=4)
    finish(fig, PKG_FIGURES / "fig_demo_panels.png")

    # 4. 极简黑白（mono 预设）
    use_preset("mono")
    fig, ax = plt.subplots(figsize=FIGSIZE)
    bars = ax.bar(["甲", "乙", "丙"], [0.72, 0.64, 0.81], width=0.5,
                  color="#5A5A5A", edgecolor="white", linewidth=0.6)
    label_bars(ax, bars)
    despine(ax, keep=("bottom",))
    grid(ax, "y")
    headroom(ax, 0.81, 1.24)
    finish(fig, PKG_FIGURES / "fig_demo_mono.png")

    # 5. 示例 deck 的图版页用的那张（对数轴 + 隐藏刻度）。数据同样全部虚构。
    use_preset("default")
    fig, ax = plt.subplots(figsize=FIGSIZE)
    names = ["词面召回", "语义重排", "图谱扩展", "三路融合"]
    ms = [12, 53, 19, 143]
    ax.bar(names, ms, width=0.55, color=SERIES[0], edgecolor="white", linewidth=0.6)
    label_bars(ax, ax.containers[0], fmt=lambda v: f"{v:.0f} ms")
    despine(ax, keep=("bottom",))
    log_axis(ax, "y", base=10, hide_ticks=True)
    grid(ax, "y")
    ax.set_ylim(1, 320)
    finish(fig, PKG_FIGURES / "fig_demo_latency.png")
    return PKG_FIGURES


if __name__ == "__main__":
    print(f"字体族：{FAMILY}")
    d = _demo()
    for p in sorted(d.glob("*.png")):
        from PIL import Image
        with Image.open(p) as im:
            w, h = im.size
        print(f"  {p.name}  {w}×{h}px  比例 {w / h:.4f}（figsize 决定，未被裁切）")
