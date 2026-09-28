# -*- coding: utf-8 -*-
"""
闸门三：对比度与出框
====================

代替人工看图的两项硬检查：
  1. 对比度——逐文本片段找出它下面那一层填充色，算 WCAG 2.1 对比度比值。
     门槛：大字（≥18pt，或 ≥14pt 粗体）≥ 3.0；小字 ≥ 4.5。低于门槛一律报。
  2. 出框——文字片段是否撑破了自己所在的那张卡片。
     「卡片」只认足够大的填充矩形（≥ CARD_W × CARD_H），免得把页码点、
     编号圆、分隔线这类小图形当成文字的底。

背景色为什么要落到像素上
------------------------
版式自带的标题带、右上角色块可能不是矢量填充而是图片，`get_drawings()` 看不到
它们。不落到像素上，就会把蓝底上的白字判成白底白字。
所以：先查矢量填充，查不到再取**渲染像素的众数**（众数才是底色，字身墨迹是少数）。

自取样像素时踩过的坑（都在下面注释里）：
  · 坐标是英寸，而 pix.width / page.rect.width 这个缩放是把 pt 换成像素
    （page.rect 单位是 pt），换算时得先把英寸乘 72，否则样本全塌到左上角；
  · 取一个像素要用 buf[py*stride + px*n + c] 逐通道读，直接切 buf[off:off+n]
    会跨字节边界、颜色全读错；
  · get_pixmap() 不带参数时返回的 xres 是 96 而实际只有 72px/in，必须显式传 dpi。

用法：python scripts/qa_visual.py <render/*.pdf> [--bg "主题背景色"]
"""
from __future__ import annotations

import re
import sys

import fitz

# ── 阈值 ──
MARGIN = 1.6      # pt，判定「出框」允许的容差
CARD_W = 120.0    # pt，小于此宽的填充矩形不算卡片
CARD_H = 60.0     # pt，小于此高的填充矩形不算卡片
DPI = 150         # 像素兜底的渲染分辨率
PIX_PAD = 0.4     # 采样框向外扩几倍行高
LARGE_MIN = 3.0   # 大字对比度门槛
SMALL_MIN = 4.5   # 小字对比度门槛

FOOT_NUM = re.compile(r"^\d{1,3} / \d{1,3}$")

# 页面底色。默认 FFFFFF（中性 / 三套浅色主题），但**真正的来源是主题的 background
# 字段**（kit.PAGE_BG）——原来这里写死白，是因为整套对比度系统都建立在白底上；
# 深色主题不改这一行，深蓝底上的白字会被当成「白底白字」判成 1:1，全篇误报。
#
# 取值优先级（resolve_bg 里实现）：--bg <hex> 显式 > kit.PAGE_BG > PDF 里的整页填充矩形
# > BG_DEFAULT。最后两级是为了让本脚本**独立运行时也不误报**：单独
# `python qa_visual.py x.pdf` 时 kit 是全新 import、apply_theme 没跑过，
# 这时靠 PDF 自带的那块整页背景矩形自己把底色找出来。
BG_DEFAULT = (255, 255, 255)
BG = BG_DEFAULT


def unhex(c):
    """PyMuPDF 文本颜色是 0xRRGGBB 整数，转成 (r,g,b) 0-255。"""
    return ((c >> 16) & 255, (c >> 8) & 255, c & 255)


def lum(rgb):
    def f(v):
        v = v / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = f(rgb[0]), f(rgb[1]), f(rgb[2])
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def ratio(a, b):
    la, lb = lum(a), lum(b)
    hi, lo = (la, lb) if la >= lb else (lb, la)
    return (hi + 0.05) / (lo + 0.05)


def kit_bg():
    """kit 里当前生效主题的页面底色。

    kit 不可用、或 import 后还没调 apply_theme（PAGE_BG 停在 FFFFFF）时返回 None，
    让调用方走下一级。
    """
    try:
        import kit
    except Exception:
        return None
    try:
        v = str(getattr(kit, "PAGE_BG", "") or "").strip().lstrip("#")
    except Exception:
        return None
    if len(v) != 6 or not re.fullmatch(r"[0-9A-Fa-f]{6}", v):
        return None
    c = unhex(int(v, 16))
    return c if c != BG_DEFAULT else None


def pdf_page_bg(doc):
    """从渲染结果里找页面底色：取最底层那块几乎占满整页的填充矩形的颜色。

    make_neutral_template 把页面底色写成母版的 <p:bg>，LibreOffice 会把它导出成
    一块满页填充矩形，所以答案就在 PDF 里。用众数是因为多页可能有细微差异
    （不同版式的背景块），取出现最多的那个。
    """
    cnt = {}
    for page in doc:
        W, H = page.rect.width, page.rect.height
        for r, c in shapes(page):
            if r.width < 0.92 * W or r.height < 0.92 * H:
                continue          # 只认满页背景，正文页的大卡片不算
            cnt[c] = cnt.get(c, 0) + 1
    if not cnt:
        return None
    c = max(cnt.items(), key=lambda kv: kv[1])[0]
    return c if c != BG_DEFAULT else None


def resolve_bg(explicit="", doc=None):
    """按可靠性从高到低定出本次判定的页面底色，写回模块级 BG 并返回它。

    `doc` 给了才走 PDF 兜底那一级——它要读页面，得在 main() 里开了文档之后调。
    """
    global BG
    if explicit:
        BG = unhex(int(explicit.lstrip("#"), 16))
        return BG
    c = kit_bg()
    if c:
        BG = c
        return BG
    if doc is not None:
        c = pdf_page_bg(doc)
        if c:
            BG = c
    return BG


def shapes(page):
    """按绘制顺序取出所有有填充的矢量矩形；越靠后越在上层。

    对比度判定不按尺寸过滤：文字可能压在很小的形状上（表头单元格、编号圆）。
    """
    out = []
    for d in page.get_drawings():
        fl = d.get("fill")
        if not fl:
            continue
        r = d["rect"]
        if r.width < 0.5 or r.height < 0.5:
            continue
        out.append((r, tuple(int(round(v * 255)) for v in fl)))
    return out


def cards(page):
    return [(r, c) for r, c in shapes(page) if r.width >= CARD_W and r.height >= CARD_H]


def bg_at(rects, pt):
    """取点所在最上层矩形的填充色；没有矩形时回落到页面底色。"""
    for r, col in reversed(rects):
        if r.contains(pt):
            return col
    return BG


def bg_pixel(pix, bbox, pad=PIX_PAD):
    """取该片段周围渲染像素的众数色。

    采样框向外扩 pad 倍**行高**，让背景占多数、字身墨迹成为少数。
    扩展量必须全程用 pt——按宽度扩会让整行标题扫过整页、把无关白底带进来。
    """
    h = max(bbox[3] - bbox[1], 1)
    sx, sy = pix.xres / 72.0, pix.yres / 72.0      # 真正的缩放是 xres/72，不是 width/rect.width
    e = h * pad
    x0 = max(0, int((bbox[0] - e) * sx))
    y0 = max(0, int((bbox[1] - e) * sy))
    x1 = min(pix.width - 1, int((bbox[2] + e) * sx))
    y1 = min(pix.height - 1, int((bbox[3] + e) * sy))
    if x1 <= x0 or y1 <= y0:                      # 片段贴住页边，采不到像素
        return BG
    buf, stride, n = pix.samples, pix.stride, pix.n
    cnt = {}
    for yy in range(y0, y1 + 1):
        off = yy * stride                          # 逐行基址 + 逐像素列偏移，n 是通道数
        for xx in range(x0, x1 + 1):
            o = off + xx * n
            k = tuple(buf[o:o + n])
            cnt[k] = cnt.get(k, 0) + 1
    return max(cnt.items(), key=lambda kv: kv[1])[0]


def host_at(rects, pt):
    """包住该点、面积最小的卡片；用于出框判定。"""
    best = None
    for r, _c in rects:
        if not r.contains(pt):
            continue
        if best is None or (r.width * r.height) < (best.width * best.height):
            best = r
    return best


def main(pdf, bg=""):
    doc = fitz.open(pdf)
    BG = resolve_bg(bg, doc)      # 页面底色跟主题走，见 BG_DEFAULT 处的说明
    if BG != BG_DEFAULT:
        print(f"页面底色 {BG}")
    bad = 0
    worst = []

    for pno, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=DPI)
        allshapes = shapes(page)
        rects = cards(page)
        spans = []
        for blk in page.get_text("dict")["blocks"]:
            if blk["type"] != 0:
                continue
            for ln in blk["lines"]:
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        spans.append(sp)

        page_min = None
        for sp in spans:
            tb = fitz.Rect(sp["bbox"])
            if tb.is_empty or tb.height < 1:
                continue
            txt = sp["text"].strip()
            if FOOT_NUM.match(txt):
                continue
            size, bold = sp["size"], bool(sp["flags"] & 16)
            fg = unhex(sp["color"])

            # 1. 对比度：先查矢量填充，查不到再查渲染像素
            ctr = fitz.Point(tb.x0 + tb.width / 2, tb.y0 + tb.height / 2)
            bg = bg_at(allshapes, ctr)
            if bg == BG:
                bg = bg_pixel(pix, tb)
                if ratio(bg, fg) < 1.4:             # 采到字身墨迹 → 退回页面底色判定
                    bg = BG
            r = ratio(fg, bg)
            page_min = r if page_min is None else min(page_min, r)
            large = size >= 18 or (bold and size >= 14)
            if r < LARGE_MIN or (not large and r < SMALL_MIN):
                bad += 1
                print(f"  第{pno:>2}页 [对比度 {r:.2f}] {size:.1f}pt"
                      f"{'粗' if bold else ''} 字{fg} 底{bg}　「{txt[:26]}」")

            # 2. 出框：文字撑破了自己所在的那张卡片
            host = host_at(rects, ctr)
            if host is None:
                continue
            over = []
            if tb.x0 < host.x0 - MARGIN:
                over.append(f"左{host.x0 - tb.x0:.1f}pt")
            if tb.x1 > host.x1 + MARGIN:
                over.append(f"右{tb.x1 - host.x1:.1f}pt")
            if tb.y1 > host.y1 + MARGIN:
                over.append(f"下{tb.y1 - host.y1:.1f}pt")
            if tb.y0 < host.y0 - MARGIN:
                over.append(f"上{host.y0 - tb.y0:.1f}pt")
            if over:
                bad += 1
                print(f"  第{pno:>2}页 [出框 {'/'.join(over)}] {size:.1f}pt　「{txt[:26]}」")

        worst.append((pno, page_min if page_min is not None else 99.0, len(spans)))

    print(f"\n各页文本片段数：{[w[2] for w in worst]}")
    print("对比度最低的五处：")
    for pno, r, _n in sorted(worst, key=lambda t: t[1])[:5]:
        print(f"  第{pno:>2}页 最低 {r:.2f}")
    print(f"{'全部通过' if bad == 0 else f'共 {bad} 处需要修复'}")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    args = sys.argv[1:]
    bg = ""
    if "--bg" in args:
        i = args.index("--bg")
        bg = args[i + 1]
        del args[i:i + 2]
    if not args:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(args[0], bg))
