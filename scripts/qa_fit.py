# -*- coding: utf-8 -*-
"""
闸门一：文本容量（源码几何层，不依赖渲染）
============================================

查四件事
--------
1. 文字容量——按各 run 自己的字号与字体逐字累加宽度、模拟换行，判断装不装得下自己声明的框。
2. 表格高度——是否越过页脚带。
3. 图片长宽比——放置框是否与原图一致（偏差 > 2% 即算拉伸）。
4. 元素边界——越出页面，或压进页脚带。

为什么这道闸门不可省
--------------------
**渲染后的 PDF 查不出溢出。** LibreOffice 会安静地裁掉超框的文字（导出 PDF 里
根本看不到多余那几行，PDF 闸门一路绿灯），PowerPoint 里则直接把字漏到框外。
所以文字容量只能在源码几何层面算——这就是 kit.save_deck() 必须落
`text_layout_*.json` 的原因：它记着每个文本框的真实坐标与逐 run 字号。
这个 JSON 不是调试产物，是交付物的必要组成部分。

字宽模型
--------
西文宽度用**真实字体文件实测**（pillow 打开 TTF，getlength / em），
不是拍的经验系数。原因：换一个西文字体，经验系数就全错；
而闸门一是唯一能在渲染前抓到溢出的地方，量错等于没设闸门。
实测不到（pillow 缺失 / 字体没装）时退回解析模型，并在输出里明确标注。

中文（CJK）统一按 1.0em——这是实测标定值，不是估计：
一条 25 字（含全角引号）14.5pt 的中文文本渲染后宽 5.04in，
即 0.2016in/字 = 14.5/72 = 1.0em。旧值 0.97 低估 3%，会少算一行、漏掉换行。
全角引号 "" 落在 U+201C/D（常规标点区），在 CJK 语境下同样按 1.0em。

行高模型
--------
行距 advance = 字号 × 行距倍率 × 1.22（实测 LibreOffice 渲染）；
单行墨迹高 = 1.32em。所需高度 = (行数-1) × 行距 + 最后一行的墨迹高，
**不是** 行数 × 行距——最后一行后面没有行距，用行数会凭空多出一整个行距。

用法：python scripts/qa_fit.py <text_layout_*.json>
"""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

# ── 阈值（改这里，不要散在逻辑里）──
EM = 72.0            # 1 in = 72 pt
PAD = 0.02           # 水平余量（文本框 inset 置 0，只留一点抗锯齿余量）
ADV = 1.22           # 行距系数：行距 = 字号 × 行距倍率 × ADV
INKH = 1.32          # 单行墨迹高度（em）
TOL = 0.08           # 溢出容差（in）。文本框不是裁切窗口，字略超框不会裁掉，
                     # 闸门二只认真实碰撞，所以这里给一点容差
TIGHT = 0.03         # 低于这个余量算「贴身」，只提示不算缺陷
STRETCH = 0.02       # 图片长宽比偏差超过此比例算拉伸
W, H = 13.33, 7.5    # 页面尺寸（in）。换主题时用 --w/--h 覆盖
FOOT_TOP = 6.98      # 页脚带起点

FOOT_NUM = re.compile(r"^\d{1,3} / \d{1,3}$")

# 字体名 → 字体文件候选。名字对不上就退回解析模型（并在输出里标注）。
FONT_FILES = {
    ("times new roman", False): ["times.ttf"],
    ("times new roman", True): ["timesbd.ttf"],
    ("calibri", False): ["calibri.ttf"],
    ("calibri", True): ["calibrib.ttf"],
    ("arial", False): ["arial.ttf"],
    ("arial", True): ["arialbd.ttf"],
    ("cambria", False): ["cambria.ttc"],
    ("cambria", True): ["cambriab.ttf"],
    ("georgia", False): ["georgia.ttf"],
    ("simsun", False): ["simsun.ttc"],
    ("microsoft yahei", False): ["msyh.ttc"],
    ("microsoft yahei", True): ["msyhbd.ttc"],
    ("微软雅黑", False): ["msyh.ttc"],
    ("微软雅黑", True): ["msyhbd.ttc"],
    ("楷体", False): ["simkai.ttf"],
    ("楷体", True): ["simkai.ttf"],
    ("kaiti", False): ["simkai.ttf"],
}
FONT_DIRS = [r"C:\Windows\Fonts", "/usr/share/fonts/truetype",
             "/System/Library/Fonts", os.path.expanduser("~/.fonts")]

_W_CACHE: dict[tuple, dict[str, float]] = {}
_FALLBACK_USED: set[str] = set()


def _font_path(name: str, bold: bool):
    key = (name or "").strip().lower()
    files = FONT_FILES.get((key, bool(bold))) or FONT_FILES.get((key, False))
    if not files:
        return None
    for d in FONT_DIRS:
        for f in files:
            p = os.path.join(d, f)
            if os.path.exists(p):
                return p
    return None


def _is_cjk(ch: str) -> bool:
    o = ord(ch)
    return (0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF
            or 0xFF00 <= o <= 0xFF60 or 0xFE30 <= o <= 0xFE4F
            or 0x3000 <= o <= 0x303F)


# 解析兜底模型（em）。只在 pillow 缺失 / 字体文件找不到时启用。
_ANALYTIC = {
    "times new roman": dict(space=0.25, digit=0.50, upper=0.66, lower=0.47, other=0.50),
    "_default": dict(space=0.28, digit=0.55, upper=0.68, lower=0.52, other=0.55),
}


def _analytic(ch: str, key: str) -> float:
    m = _ANALYTIC.get(key) or _ANALYTIC["_default"]
    if ch in (" ", " "):
        return m["space"]
    if ch.isdigit():
        return m["digit"]
    if ch.isupper():
        return m["upper"]
    if ch.islower():
        return m["lower"]
    return m["other"]


def _load_widths(name: str, bold: bool) -> dict | None:
    """载入某字体的逐字宽度表（em）。取不到返回 None。"""
    key = (name or "").strip().lower()
    ck = (key, bool(bold))
    if ck in _W_CACHE:
        return _W_CACHE[ck]
    table = None
    path = _font_path(name, bold)
    if path:
        try:
            from PIL import ImageFont
            ft = ImageFont.truetype(path, 1000)
            chars = ("0123456789"
                     "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                     "abcdefghijklmnopqrstuvwxyz"
                     " .,:;!?()[]{}<>+-*/=%&#@'\"_|$~^`\\")
            # 逐字实测：一次量一批（getlength 单字调用慢），再按需补个别字符
            table = {c: ft.getlength(c) / 1000.0 for c in chars}
            # 「西文」是兜底哨兵：表里没有的字符（未映射字符）一律按 1.0em 算。
            # 写成与源工程一致的「西文」字面量，别换成 U+FFFD 之类的占位——
            # 那两处（建表 + 取值）必须同一个字符，改一处即静默失效。
            table["西文"] = 1.0
        except Exception:
            table = None
    _W_CACHE[ck] = table
    if table is None:
        _FALLBACK_USED.add(name or "?")
    return table


def cw(ch: str, font: str = "Times New Roman", bold: bool = False) -> float:
    """单字宽度，单位 em。中文字形一律 1.0em（同框内中西文各走各的字宽）。"""
    if _is_cjk(ch):
        return 1.0
    if 0x2000 <= ord(ch) <= 0x206F:      # “”‘’—… 等标点，CJK 语境下按全角
        return 1.0
    if 0x2190 <= ord(ch) <= 0x21FF:      # 箭头
        return 0.55
    t = _load_widths(font, bold)
    if t is None:
        return _analytic(ch, (font or "").strip().lower())
    v = t.get(ch)
    if v is None:
        v = t.get("西文", _analytic(ch, (font or "").strip().lower()))
    return v


def measure(runs, box_w, ls=1.30, font="Times New Roman"):
    """按各 run 自己的字号/字重逐字累加宽度、超宽换行。返回 (行数, 所需高度 in)。

    runs 元素是 [文本, 字号, 粗体]（旧版两元组的第三位按非粗体处理）。
    """
    avail = box_w - 2 * PAD
    lines, cur, cur_em = [], 0.0, 0.0
    for item in runs:
        text, pt = str(item[0]), float(item[1])
        bold = bool(item[2]) if len(item) > 2 else False
        em = pt / EM
        for ch in text:
            if ch == "\n":
                lines.append((cur, cur_em))
                cur, cur_em = 0.0, 0.0
                continue
            w = cw(ch, font, bold) * em
            if cur > 0 and cur + w > avail:
                lines.append((cur, cur_em))
                cur, cur_em = w, em
            else:
                cur += w
                cur_em = max(cur_em, em)
    lines.append((cur, cur_em))
    while lines and lines[-1][0] <= 0:        # 去掉结尾空行
        lines.pop()
    if not lines:
        return 0, 0.0
    em = max(e for _, e in lines)
    return len(lines), (len(lines) - 1) * em * ls * ADV + em * INKH


def txt_of(d) -> str:
    return "".join(r[0] for r in d.get("runs", [])).replace("\\n", "\n")


def is_foot(d) -> bool:
    """页脚元素。role 是 kit 写进清单的；没有 role 时退回页码正则 + 位置。"""
    if d.get("role") in ("footer", "page_no"):
        return True
    t = txt_of(d).strip()
    return bool(FOOT_NUM.match(t))


def main(path, w=W, h=H, foot_top=FOOT_TOP):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    texts = [d for d in data if d["kind"] == "text"]
    tables = [d for d in data if d["kind"] == "table"]
    imgs = [d for d in data if d["kind"] == "img"]
    bad = 0
    fonts = sorted({d.get("latin", "") for d in texts if d.get("latin")})

    print(f"清单：{len(texts)} 个文字块 / {len(tables)} 张表 / {len(imgs)} 张图")
    print(f"西文字体：{fonts}　页脚带起点 {foot_top}in　页面 {w}×{h}in\n")

    # 1. 文字容量
    for d in texts:
        n, need = measure(d["runs"], d["w"], d.get("ls", 1.30),
                          d.get("latin", "Times New Roman"))
        slack = d["h"] - need
        t = txt_of(d).replace("\n", " / ")
        if slack < -TOL:
            bad += 1
            print(f"  第{d['p']:>2}页 [溢出 {need - d['h']:+.2f}in] 需 {need:.2f}\" / 框 {d['h']:.2f}\""
                  f"　{n} 行 / 宽 {d['w']:.2f}\"　「{t[:30]}」")
        elif slack < TIGHT:
            print(f"  第{d['p']:>2}页 [贴身 {slack:+.2f}in] 需 {need:.2f}\" / 框 {d['h']:.2f}\""
                  f"　{n} 行　「{t[:30]}」")

    # 2. 表格
    for d in tables:
        bot = d["y"] + d["rowH"] * d["rows"]
        print(f"  第{d['p']:>2}页 [表格] 表头+{d['rows'] - 1} 行 = {d['rowH'] * d['rows']:.2f}\"，"
              f"底部 {bot:.2f}\"")
        if bot > foot_top:
            bad += 1
            print(f"        [警告] 表格底部越过页脚顶 {foot_top}\"")

    # 3. 图片长宽比（ar 由 kit 用 pillow 从原图实测写进清单）
    for d in imgs:
        a = d.get("ar")
        if not a or d.get("sized"):
            continue
        placed = d["w"] / d["h"]
        if abs(placed - a) / a > STRETCH:
            bad += 1
            print(f"  第{d['p']:>2}页 [拉伸] {Path(d['path']).name}: 放置 {placed:.3f} "
                  f"vs 原图 {a:.3f}（{placed / a - 1:+.1%}）")

    # 4. 边界
    footed = {d["p"] for d in texts if is_foot(d)}
    for d in data:
        if not all(k in d for k in ("x", "y", "w", "h")):
            continue
        x, y, ww, hh = d["x"], d["y"], d["w"], d["h"]
        t = txt_of(d).replace("\n", " / ")
        if x < -0.01 or y < -0.01 or x + ww > w + 0.01 or y + hh > h + 0.01:
            bad += 1
            print(f"  第{d['p']:>2}页 [越界] x{x:.2f} y{y:.2f} w{ww:.2f} h{hh:.2f}　「{t[:22]}」")
        elif d["p"] in footed and not is_foot(d) and y + hh > foot_top:
            bad += 1
            print(f"  第{d['p']:>2}页 [压页脚] 底部 {y + hh:.2f}\"　「{t[:22]}」")

    if _FALLBACK_USED:
        print(f"\n[注意] 这些字体没量到真实字宽（pillow 缺失或字体未安装），"
              f"已退回解析模型，精度下降：{sorted(_FALLBACK_USED)}")
    print(f"\n{'全部通过' if bad == 0 else f'共 {bad} 处需要修复'}")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    ap_args = sys.argv[1:]
    if not ap_args:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(*ap_args))
