# -*- coding: utf-8 -*-
"""
闸门四：越界与空隙（版面体检）
=============================

专门盯两件肉眼才看得出、但很要命的事：「文字越过组件边界」和「布局不合理、出现大量空隙」。

都读渲染后的 PDF，**以真实墨迹为准**，不看源码里声明的框——声明的框高常大于实际
墨迹，用它会漏判。

检查一：越界
  一条文字属于某个组件（行中心落在组件内），却撑出了组件的边界。
  组件只认真实存在的面板类矩形：足够大、有填充或有描边，排除分隔线、圆点、编号圆。
  比闸门三的「出框」更贴近读者观感：字起点在卡片里、尾字跑到卡片外那种。

检查二：空隙
  在标题带以下、页脚以上的内容区里取所有内容（面板 / 文字 / 图片 / 图表）的纵向并集，
  并集之间的空档是纵向空隙；再按栅格扫出最大的空白矩形。三条判据：
    · 空隙 ≥ VOID_BAD 且不是「节奏」——三条以上高度相近的空隙是并列条目的刻意排布；
    · 内容底部离页脚 ≥ TAIL_BAD —— 这页提前收工了；
    · 内容覆盖率 < MIN_COVER —— 整页偏空。
  低于 VOID_TOL 的空隙只记录不判缺陷：那是正常的块间距。

踩过的三个坑（都在下面注释里）：LibreOffice 每页会导出一个整页白色矩形；
三条以上等高空隙是节奏不是缺陷；覆盖率阈值要按页型定，纯文字页天生只有 45% 左右。

用法：python scripts/qa_layout.py <render/*.pdf>
"""
from __future__ import annotations

import re
import sys

import fitz

# ── 阈值 ──
CT_TOP = 1.66 * 72      # 内容区顶：标题带 + 章节标签之下
CT_BOT = 6.94 * 72      # 内容区底：页脚文字之上

COMP_W = 50.0           # pt：小于此宽的矩形不算组件
COMP_H = 30.0           # pt：小于此高的矩形不算组件
TOL = 2.5               # pt：越界判定容差（面板描边居中于路径，需留余量）
VOID_TOL = 0.30 * 72    # pt：低于此值只记录
VOID_BAD = 0.55 * 72    # pt：达到此值才算「大量空隙」
TAIL_BAD = 0.55 * 72    # pt：内容底部离页脚超过此值 = 提前收工
MIN_COVER = 0.40        # 内容覆盖率下限（纯文字页无大色块，天生偏低）
GRID = 0.10 * 72        # pt：空白矩形栅格
MARGIN_X = 0.55 * 72    # 左页边距条带本来就该是空的，不计入空白

FOOT_NUM = re.compile(r"^\d{1,3} / \d{1,3}$")


def components(page):
    """面板类矩形：够大、有填充或有描边。排除整页白底与页边距条带。"""
    out = []
    W, H = page.rect.width, page.rect.height
    for d in page.get_drawings():
        r = fitz.Rect(d["rect"])
        if r.width < COMP_W or r.height < COMP_H:
            continue
        if r.width > 0.95 * W and r.height > 0.95 * H:
            continue          # LibreOffice 导出的整页白色矩形：把它当组件，
            #                  空隙就永远查不出来（每页都有它）
        fl = d.get("fill")
        if not fl and not d.get("color"):
            continue
        out.append((r, tuple(round(v * 255) for v in fl) if fl else None))
    return out


def lines(page):
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if blk.get("type") != 0:
            continue
        for ln in blk.get("lines", []):
            t = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
            sz = max([sp.get("size", 1) for sp in ln.get("spans", [])], default=1)
            if t:
                out.append((fitz.Rect(ln["bbox"]), t, sz))
    return out


def items(page):
    """内容区里的全部占位：组件 + 文字行 + 图片。"""
    out = [r for r, _c in components(page)]
    for r, _t, _sz in lines(page):
        if not FOOT_NUM.match(_t):
            out.append(r)
    for im in page.get_image_info():
        out.append(fitz.Rect(im["bbox"]))
    return out


def check_overflow(page):
    bad = []
    for r, fill in components(page):
        for tb, t, sz in lines(page):
            cx, cy = tb.x0 + tb.width / 2, tb.y0 + tb.height / 2
            if not (r.y0 <= cy <= r.y1 and r.x0 <= cx <= r.x1):
                continue
            over = []
            if tb.x0 < r.x0 - TOL:
                over.append(f"左{r.x0 - tb.x0:.1f}pt")
            if tb.x1 > r.x1 + TOL:
                over.append(f"右{tb.x1 - r.x1:.1f}pt")
            if tb.y0 < r.y0 - TOL:
                over.append(f"上{r.y0 - tb.y0:.1f}pt")
            if tb.y1 > r.y1 + TOL:
                over.append(f"下{tb.y1 - r.y1:.1f}pt")
            if over:
                bad.append((r, tb, "".join(over), t, sz, fill))
    return bad


def in_area(r, x0=None, x1=None):
    a = max(r.y0, CT_TOP)
    b = min(r.y1, CT_BOT)
    if b - a <= 0:
        return None
    l = r.x0 if x0 is None else max(r.x0, x0)
    rr = r.x1 if x1 is None else min(r.x1, x1)
    return fitz.Rect(l, a, rr, b) if rr > l else None


def vertical_gaps(area_items):
    ys = sorted({(r.y0, r.y1) for r in area_items})
    merged = []
    for a, b in ys:
        if merged and a <= merged[-1][1] + 0.5:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    voids, prev = [], CT_TOP
    for a, b in merged:
        if a - prev > VOID_TOL:
            voids.append((prev, a))
        prev = max(prev, b)
    if CT_BOT - prev > VOID_TOL:
        voids.append((prev, CT_BOT))
    return voids


def largest_blank(area_items, W):
    """版心内的最大空白矩形（左右页边距条带不计入）。单调栈求最大全 0 矩形。"""
    x_lo, x_hi = MARGIN_X, W - MARGIN_X
    nx = int(round((x_hi - x_lo) / GRID))
    ny = int(round((CT_BOT - CT_TOP) / GRID))
    if nx <= 0 or ny <= 0:
        return None
    covered = [[False] * nx for _ in range(ny)]
    for r in area_items:
        x0 = max(0, int((r.x0 - x_lo) / GRID))
        y0 = max(0, int((r.y0 - CT_TOP) / GRID))
        x1 = min(nx, int((r.x1 - x_lo) / GRID) + 1)
        y1 = min(ny, int((r.y1 - CT_TOP) / GRID) + 1)
        for yy in range(max(0, y0), min(ny, y1)):
            for xx in range(max(0, x0), min(nx, x1)):
                covered[yy][xx] = True
    heights = [0] * nx
    bh = bw = bx = by = 0
    for y in range(ny):
        for x in range(nx):
            heights[x] = 0 if covered[y][x] else heights[x] + 1
        stack = []                                  # [(柱高, 起始列)]
        for x in range(nx + 1):
            h = heights[x] if x < nx else 0
            while stack and stack[-1][0] > h:
                hh, xx = stack.pop()
                if (hh, x - xx) > (bh, bw):
                    bh, bw, bx, by = hh, x - xx, xx, y - hh + 1
            stack.append((h, x))
    return (by, bx, bw, bh) if bh else None


def coverage(area_items, W):
    """版心内的面积覆盖率。矩形面积直接相加会重复计算重叠区，展示时截到 100%。"""
    total = sum(r.width * r.height for r in area_items)
    return min(1.0, total / ((W - 2 * MARGIN_X) * (CT_BOT - CT_TOP)))


def is_rhythm(voids):
    """三条以上高度相近的空隙 = 并列条目的节奏，属刻意设计，不判缺陷。"""
    hs = [b - a for a, b in voids]
    if len(hs) < 3:
        return False
    lo, hi = min(hs), max(hs)
    return hi - lo < 0.25 * hi


def main(pdf):
    doc = fitz.open(pdf)
    W = doc[0].rect.width
    n_ov = n_bad = 0
    print(f"页数 {len(doc)}　内容区 y {CT_TOP / 72:.2f}–{CT_BOT / 72:.2f} in\n")
    for pno, page in enumerate(doc, 1):
        its = items(page)
        ov = check_overflow(page)
        area = [r for r in (in_area(i) for i in its) if r]
        voids = vertical_gaps(area)
        blank = largest_blank(area, W)
        cov = coverage(area, W)
        top = min([r.y0 for r in area], default=CT_TOP)
        bot = max([r.y1 for r in area], default=CT_TOP)

        flags = []
        if ov:
            flags.append(f"越界 {len(ov)}")
            n_ov += len(ov)
        rhythm = is_rhythm(voids)
        big = [v for v in voids if v[1] - v[0] >= VOID_BAD]
        if big and not rhythm:
            flags.append("大空隙")
            n_bad += 1
        tail = CT_BOT - bot
        if tail >= TAIL_BAD:
            flags.append(f"提前收工 {tail / 72:.2f}in")
            n_bad += 1
        if cov < MIN_COVER:
            flags.append(f"偏空 {cov:.0%}")
            n_bad += 1

        tag = "⚠ " + "，".join(flags) if flags else "  ok"
        print(f"第{pno:>2}页{tag:22} 内容 y{top / 72:.2f}–{bot / 72:.2f}  "
              f"覆盖率 {cov:.0%}  空隙 {len(voids)}" + ("（节奏）" if rhythm else ""))
        for r, tb, s, t, sz, _fill in ov:
            print(f"      [越界 {s}] {sz:.1f}pt 组件 x{r.x0 / 72:.2f}-{r.x1 / 72:.2f} "
                  f"y{r.y0 / 72:.2f}-{r.y1 / 72:.2f}｜文字 x{tb.x0 / 72:.2f}-{tb.x1 / 72:.2f} "
                  f"y{tb.y0 / 72:.2f}-{tb.y1 / 72:.2f}「{t[:18]}」")
        for a, b in voids:
            mark = "●" if b - a >= VOID_BAD else "·"
            print(f"      {mark} [空隙 {(b - a) / 72:.2f}in]  y{a / 72:.2f}–{b / 72:.2f}")
        if blank:
            by, bx, bw, bh = blank
            print(f"      [最大空白矩形 {(bw * GRID) / 72:.2f}×{(bh * GRID) / 72:.2f}in]  "
                  f"x{(bx * GRID + MARGIN_X) / 72:.2f} y{(by * GRID + CT_TOP) / 72:.2f}")
    print(f"\n越界 {n_ov} 处；布局缺陷 {n_bad} 页"
          + ("（全部通过）" if n_ov == 0 and n_bad == 0 else ""))
    return 0 if n_ov == 0 and n_bad == 0 else 1


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
