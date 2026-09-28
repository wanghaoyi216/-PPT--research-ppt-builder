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

五套预设
--------
    default  单图 / 双图，10pt 底字，白底 + 灰蓝系      （最常用）
    dense    多子图（2×2 及以上），8.5pt 底字，边距收窄   （拼版/消融图）
    mono     极简黑白：无彩色、单色柱、只留 bottom 轴      （打印稿、外部评审）
    scatter  散点 / 相关图：两轴按数据留白、点带描边        （相关、分布、量级）
    dark     深色底：深画布 + 浅字 + 几乎不可见的网格       （配 midnight 这类深色主题）

`default` / `dense` / `mono` 都是**柱图预设**：`label_bars`（标在柱端）、
`headroom`（`ylim` 到 1.24 倍最高柱）、`despine`（砍掉 top/right 只留两轴）
这三招全都假设「柱子从 0 起」。散点没有柱顶、可能带负值，照抄会把点挤到一角
或直接裁掉——散点图请用 `scatter` 预设配 `axes_for()`。

和主题对齐
----------
`use_theme("scholar-red")` 把 kit 的当前主题灌进本模块：**不调用它就是旧行为**，
`python scripts/figstyle.py` 直接跑也不需要 kit。调完之后 `INK` / `SERIES` 这些名字
就地变成主题色板，柱图配色跟着 `--theme` 走，不会再出现「页面是红的、图还是蓝的」。

用法::

    import figstyle, matplotlib.pyplot as plt
    figstyle.use_preset("default")                 # 一次就够，后面全继承 rcParams
    figstyle.use_theme("scholar-red")              # 可选：图和页面共用一套色板
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
import numpy as np
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
# 这一组是**浅色主题的默认值**：没人调 use_theme() 时就是它，行为与旧版完全一致。
# 调了 use_theme() 之后这六个名字会被就地改成主题色板——所以直接引用它们的
# 代码拿到的永远是当前生效的颜色，而不是 import 时的那张快照。
INK = "#22262B"        # 主文字 12.9:1 白底
MUTED = "#5A646E"      # 次文字 5.6:1
GRID = "#DDE3E8"       # 网格线（不承字）
AXIS = "#8A939B"       # 轴线与刻度
MISSING = "#9AA1A7"    # 缺失值标注（灰）5.0:1 白底，且用斜体双重编码
SERIES = ["#1F6FA8", "#7FA9C9", "#B5532C", "#2F2F2F"]   # 主 / 次 / 强调 / 深
EDGE = "white"         # 柱端描边：浅底用白，深底改用页面底色（白描边在深底上是一圈脏边）

# 竖排图与 docx / pptx 里按固定比例放置的图，最常用的是 10.5 × 5.46（≈1.92:1）
FIGSIZE = (6.0, 3.2)

# 当前生效的配色状态。PRESETS 里只锁**几何**（字号、线宽），颜色一律从这里取，
# 于是 use_theme() 换一次色，所有预设同步跟着变，不用逐个预设改字面量。
_STYLE = {
    "face": "white",                       # 画布底（mono 预设强制白底）
    "ink": INK, "muted": MUTED, "grid": GRID,
    "axis": AXIS, "missing": MISSING, "edge": EDGE,
    "series": list(SERIES),
}

# ── 五套预设 ──
# 值都是 matplotlib rcParams 的键；键名写错不报错、只是静默不生效，
# 所以改完预设务必出一张图看一眼。
# 下面写的颜色是**浅色主题的基准值**：use_preset() 会紧接着用 `_style_rc()`
# 按当前主题再覆盖一遍颜色键，因此这里只当作几何量（字号 / 线宽）来维护。
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
    "scatter": {
        # 散点 / 相关图：两轴都要按数据留白，配 axes_for() 用。
        # 点带一圈描边（EDGE）——深底上点和网格颜色接近，不描边会糊成一片。
        "font.size": 10, "axes.labelsize": 10, "xtick.labelsize": 9, "ytick.labelsize": 9,
        "legend.fontsize": 9,
        "axes.edgecolor": AXIS, "axes.linewidth": 0.9,
        "axes.labelcolor": INK, "text.color": INK,
        "xtick.color": AXIS, "ytick.color": AXIS,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
        "lines.linewidth": 1.6,
        "lines.markersize": 5.5,
    },
    "dark": {
        # 深色底（midnight 这类）：深画布 + 浅字 + 几乎看不见的网格。
        # 字色必须够亮：浅预设里的 AXIS（#8A939B）放在深底上只有 2.5:1，投影上一片糊。
        # 网格取主题 SKY（深底上是比画布暗一点的色），比「浅灰线」更不抢数据。
        "font.size": 10, "axes.labelsize": 10, "xtick.labelsize": 9, "ytick.labelsize": 9,
        "legend.fontsize": 9,
        "axes.edgecolor": "#3E5878", "axes.linewidth": 0.9,
        "axes.labelcolor": "#F0F5FB", "text.color": "#F0F5FB",
        "xtick.color": "#8FA3BC", "ytick.color": "#8FA3BC",
        "figure.facecolor": "#101C2E", "axes.facecolor": "#101C2E", "savefig.facecolor": "#101C2E",
        "lines.linewidth": 1.6,
        "lines.markersize": 5.5,
    },
}

# 「主色调」：散点、误差棒、参考线取这个，和柱用同一个色号才算同一套图。
# 深灰（mono）是写死的——mono 的意义就是不给任何一种颜色以表达立场的机会。
_AXIS_TONES = {"default": SERIES[0], "dense": SERIES[0], "mono": "#5A5A5A",
               "scatter": SERIES[0], "dark": SERIES[0]}


def axis_tone(name: str = "default"):
    """当前预设的主色调。取函数而不是在 def 时求值：`_AXIS_TONES` 里的
    `SERIES[0]` 在 use_theme() 之后会被同步更新，直接读 `SERIES[0]` 拿到的
    会是旧的蓝色——「页面是红的、图里的散点是蓝的」正是同一类断裂。
    """
    return _AXIS_TONES.get(name, _STYLE["series"][0])


def _rgb(h):
    h = str(h).lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _lum(h) -> float:
    """WCAG 相对亮度（0=黑，1=白）。公式与 make_neutral_template.py / qa_visual.py 同一套。"""
    def f(c):
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _rgb(h)
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _toward(h, target, t=0.12):
    """把颜色朝 `target` 挪 `t` 比例（t=0 不变，t=1 变成 target）。"""
    a, b = _rgb(h), _rgb(target)
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def _style_rc(name: str) -> dict:
    """按当前 `_STYLE` 生成的**颜色** rcParams。几何量仍在 PRESETS 里。

    `mono` 永远白底、刻度用 ink：它给打印稿和外部评审用，不跟随深色主题——
    把一套黑白图塞进深底页面里，黑白本身就够用了。
    """
    c = _STYLE
    grey = (name == "mono")
    tick = c["ink"] if grey else c["axis"]
    return {
        "figure.facecolor": "white" if grey else c["face"],
        "axes.facecolor": "white" if grey else c["face"],
        "savefig.facecolor": "white" if grey else c["face"],
        "text.color": c["ink"], "axes.labelcolor": c["ink"],
        "axes.edgecolor": c["ink"] if grey else c["axis"],
        "xtick.color": tick, "ytick.color": tick,
    }


def use_preset(name: str = "default") -> str:
    """套用一套 rcParams 预设（含字体族与 `axes.unicode_minus=False`）。返回预设名。

    先套 PRESETS 的几何量，再用 `_style_rc()` 按当前主题覆盖颜色键——
    顺序不能反，否则 use_theme() 换过的颜色会被 PRESETS 里的浅色基准值盖回去。

    `unicode_minus=False` 是**必需**的：默认会用 U+2212 减号，Times New Roman 之外
    的回退字体常常没有这个码位，负刻度会渲染成方框。关掉它改用 ASCII 连字符。
    """
    if name not in PRESETS:
        raise KeyError(f"没有预设「{name}」。可选：{sorted(PRESETS)}")
    matplotlib.rcParams.update(PRESETS[name])
    matplotlib.rcParams.update(_style_rc(name))
    matplotlib.rcParams["font.family"] = FAMILY
    matplotlib.rcParams["axes.unicode_minus"] = False
    return name


def use_theme(name=None, preset=None) -> str:
    """把 kit 的主题色灌进本模块，让配图和页面共用一套色板。返回生效的预设名。

    **不调用它就是一切照旧**：`_STYLE` 的初值就是上面写死的浅色配色，
    本模块可以完全脱离 kit 独立使用。

    映射关系：
        BAND  → SERIES[0]   主题的「主张色」，也就是柱图的主色
        BROWN → SERIES[1]   对比补色。第二位就放它，是因为双系列柱图最常见，
                            而 BROWN 是主题里唯一跳出主色族的暖色；midnight 里 BAND
                            和 ACC 都是浅蓝，挤在一起第二根柱会看不清
        ACC   → SERIES[2]   同族浅色
        BORDER→ SERIES[3]   同族深一档
        INK / MUTED → 图内主文字 / 次文字
        SKY   → 网格线      SKY 是主题里唯一的「淡色罩」角色，深底上它自动变深，
                            网格于是仍然是淡淡的；BORDER 是全强度描边色，做网格会太吵
        MUTED → 轴线与刻度（深底上略微朝底色压一点，避免太亮）

    `background` 亮度低于 0.35 就判定为深色主题，`preset` 不指定时自动切到 `dark`。
    """
    global INK, MUTED, GRID, AXIS, MISSING, SERIES, EDGE
    try:
        import kit
    except ImportError as e:          # figstyle 要能脱离 kit 独立跑
        raise ImportError(
            "use_theme() 需要 kit.py 与本文件同目录（scripts/）。"
            "不调用 use_theme() 时本模块可以完全独立使用。") from e
    th = kit.load_theme(name=name) if name else (kit._ACTIVE_THEME or
                                                 kit.apply_theme({"name": "theme-neutral"}))
    p = th["palette"]
    bg = "#" + (th.get("background") or "FFFFFF")     # kit 里存的是不带 # 的六位十六进制
    band, acc, brown, border = ("#" + p[k] for k in ("BAND", "ACC", "BROWN", "BORDER"))
    ink, muted, sky = ("#" + p[k] for k in ("INK", "MUTED", "SKY"))
    dark = _lum(bg) < 0.35            # 深底：字要亮、网格要淡、描边要用底色
    if preset is None:
        preset = "dark" if dark else "default"
    INK, MUTED = ink, muted
    GRID = sky
    AXIS = _toward(muted, bg, 0.25) if dark else muted
    MISSING = _toward(muted, bg, 0.14)
    SERIES = [band, brown, acc, border]
    EDGE = bg if dark else "white"
    _STYLE.update(face=bg, ink=INK, muted=MUTED, grid=GRID, axis=AXIS,
                  missing=MISSING, edge=EDGE, series=list(SERIES))
    # 四个彩色预设的主色调都跟着新 SERIES 走；mono 的灰不动。
    _AXIS_TONES.update({k: SERIES[0] for k in ("default", "dense", "scatter", "dark")})
    return use_preset(preset)


from_theme = use_theme      # 两个名字都行：from_theme("scholar-red") == use_theme("scholar-red")


def despine(ax, keep=("left", "bottom")):
    """只留下 keep 里点名的轴脊，其余一律隐藏。

    为什么砍：top / right 两条边在只有一个系列、没有第二量纲的图里是纯装饰，
    它们把画面框成一个盒子，读者会去读「盒子的边界」这个不存在的信息。
    留 left+bottom 是习惯，横向条形图通常只留 bottom。
    """
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(s in keep)


def grid(ax, which="y", lw=0.6, color=None):
    """**单向**网格 + `set_axisbelow(True)`：网格在数据下面，不穿过柱子。

    `color` 默认 None → 调用时取 GRID。写成 `color=GRID` 是错的：**默认参数在 def 时
    就绑定**，绑住的是 import 时的蓝色网格，use_theme() 换主题也不生效。
    双向网格等于每根柱子都被横竖线切开，柱宽的视觉误差比柱高差还大。
    `set_axisbelow(True)` 不设的话 matplotlib 默认把网格画在数据之上。
    """
    ax.grid(True, axis=which, color=color or GRID, linewidth=lw, linestyle="-")
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
                color=color or INK)      # INK 取调用时的全局值（use_theme 之后跟着主题走）
        out.append(s)
    return out


def mark_missing(ax, x, y, text="未参评", size=9, color=None):
    """把缺失值标成灰色斜体文字，**不要画成 0 柱**。

    这是本模块最要紧的一条纪律：0 柱会被读成「测出来是 0」，
    「没测」和「测了是 0」是两件完全不同的事，画成同一根柱子就是把
    「我们没数据」讲成了「我们数据不好」。斜体 + 灰是双重编码，
    黑白打印出来也认得出来。

    `color` 默认 None → 取 MISSING（调用时求值，use_theme 之后跟着主题走）。
    """
    return ax.text(x, y, text, ha="center", va="bottom", fontsize=size,
                   style="italic", color=color or MISSING)


def headroom(ax, ymax, ratio=1.24, xmax=None):
    """给柱端标签留出顶部空间：`ylim(0, ymax*ratio)`，可选把 x 轴也放开到 `xmax`。

    标签画在数据之上，轴上限贴着最高那根柱顶时标签就被裁掉了——
    而「标签被裁掉一半」这件事在缩略图上看不出来，等到投影才发现。
    """
    ax.set_ylim(0, ymax * ratio)
    if xmax is not None:
        ax.set_xlim(0, xmax)
    return ax


def axes_for(ax, xs, ys, pad=0.09):
    """给散点图**两个方向**都留白：x / y 各按数据范围放大 `pad` 比例。

    散点图不能用 headroom()：那个函数只对 y 做 `ylim(0, ymax*1.24)`，
    它假设柱子从 0 起——散点可能带负值，而且 x 方向根本没有上限可言，
    照搬会把最左边一列点裁在轴外、把 x 轴压成一条贴脸的线。
    这里给四周各留一档白边，最外圈的点就不会贴着轴脊。
    """
    xs, ys = list(xs), list(ys)
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    dx = (x1 - x0) * pad or 1.0          # 所有 x 相同时给个默认跨度，避免零宽坐标轴
    dy = (y1 - y0) * pad or 1.0
    ax.set_xlim(x0 - dx, x1 + dx)
    ax.set_ylim(y0 - dy, y1 + dy)
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

    # 6. 散点（scatter 预设）：两轴按数据留白，点带描边。不用 headroom。
    use_preset("scatter")
    rng = np.random.default_rng(7)
    x = rng.uniform(2, 40, 26)
    y = 14 + 0.61 * x + rng.normal(0, 7.5, 26)
    fig, ax = plt.subplots(figsize=FIGSIZE)
    ax.scatter(x, y, s=44, color=axis_tone("scatter"), edgecolor=EDGE, linewidth=0.7,
               alpha=0.88)
    ax.axhline(0, color=AXIS, lw=0.7, ls=(0, (4, 3)))
    axes_for(ax, x, y)
    despine(ax)
    grid(ax, "y", lw=0.5)
    ax.set_xlabel("训练步数（k）")
    ax.set_ylabel("验证误差")
    finish(fig, PKG_FIGURES / "fig_demo_scatter.png")

    # 7、8. 与主题对齐：深底（dark 预设）与浅色主题（scholar-red）各一张。
    # 需要 kit.py，独立分发本模块时跳过这两张，不影响前六张。
    try:
        use_theme("midnight")            # 深底 → 自动选 dark 预设
        fig, ax = plt.subplots(figsize=FIGSIZE)
        bars = ax.bar(["基线", "方法一", "方法二", "三路融合"],
                      [0.41, 0.63, 0.58, 0.71], width=0.55,
                      color=[SERIES[0], SERIES[1], SERIES[2], SERIES[3]],
                      edgecolor=EDGE, linewidth=0.6)
        label_bars(ax, bars)
        despine(ax)
        grid(ax, "y", lw=0.6)
        headroom(ax, 0.71, 1.24)
        finish(fig, PKG_FIGURES / "fig_demo_dark.png")

        use_theme("scholar-red")         # 浅色主题 → 默认预设，但柱是红的
        fig, ax = plt.subplots(figsize=FIGSIZE)
        bars = ax.bar(["词面召回", "语义重排", "图谱扩展"], [0.41, 0.63, 0.58],
                      width=0.55, color=SERIES[0], edgecolor=EDGE, linewidth=0.6)
        label_bars(ax, bars)
        despine(ax)
        grid(ax, "y")
        headroom(ax, 0.63, 1.24)
        finish(fig, PKG_FIGURES / "fig_demo_themed.png")
    except ImportError:
        print("  （跳过 fig_demo_dark / fig_demo_themed：未找到 kit.py）")
    return PKG_FIGURES


if __name__ == "__main__":
    print(f"字体族：{FAMILY}")
    d = _demo()
    for p in sorted(d.glob("*.png")):
        from PIL import Image
        with Image.open(p) as im:
            w, h = im.size
        print(f"  {p.name}  {w}×{h}px  比例 {w / h:.4f}（figsize 决定，未被裁切）")
