# 踩坑清单

按「会静默出错 / 很难查」的顺序排。每条都写清了**为什么**和**怎么验**。

---

## 会静默出错（闸门查不出来或查错了方向）

### 1. `run.font.name` 只写 `<a:latin>`

中文字形要另注入 `<a:ea typeface="微软雅黑"/>`，**而且 `<a:ea>` 必须排在
`<a:latin>` 之后**——顺序反了 PowerPoint 读不到中文字体，用的是主题字体。

后果：Linux/低版本 Office 上整篇中文变成宋体，但你在自己机器上看不出问题。

验：解压 pptx，`grep -c '<a:ea' ppt/slides/slide1.xml`，并确认 `<a:ea` 的位置
在同一个 `<a:rPr>` 里排在 `<a:latin>` 之后。`kit._style()` 已经处理。

### 2. 中文字体不能用默认参数传

```python
def txt(s, ..., ea=EAK):      # ← 错
    ...
kit.EAK = "楷体"               # ← 改不动，def 时就绑定了
```

默认参数在 `def` 时求值，之后改模块级变量一个字也改不到。字体名照写进 XML，
**就是不换**——最坏的一种错：文件看着是改了，实际没改。

正解：一律 `ea=None`，在函数体里 `ea = ea or EAK` 现取。`kit.txt()` 就是这样。

### 3. 渲染后的 PDF 查不出文字溢出

LibreOffice 静默裁切，PowerPoint 直接漏字。所以闸门一必须在**源码几何层**算，
靠 `kit.save_deck()` 落下的 `*_layout.json`。

如果哪天 `qa_pdf` 报「无几何问题」而 pptx 打开一看满页是字——
先去查有没有那张 JSON，不要怀疑闸门二。

### 4. soffice 静默失败

`-env:UserInstallation=file://C:/...`（两个斜杠）时 soffice **不产出任何文件，
exit 仍是 0**。于是 PDF 停在旧 mtime 上，闸门拿旧 PDF 判新代码。

正解：三个斜杠 `file:///C:/...`；转换前清残留进程；渲染前删旧 PDF；
拿到 0 之后等文件落盘。`verify_all.py` 四条都做了。

### 5. 同一段里的多个 run 必须共用一个段落

只有 `\n` 才新建段落。写成「一个 run 一个新段落」，像 `"图 4　＋说明"`
这种双 run 标签会被拆成两行，级联出一片重叠。

`kit.txt()` 里的 `for i, seg in enumerate(text.split("\n"))` 就是对的写法。

### 6. 字宽模型量错 = 闸门失效

西文字宽按「经验系数」算，换个字体就全错。**用真实字体文件实测**
（pillow `ImageFont.truetype(path, 1000).getlength(ch) / 1000`）。
中文按 1.0em（实测标定，不是估计）。

行高同理：所需高度 = `(行数-1) × 行距 + 最后一行的墨迹高`，
用 `行数 × 行距` 会凭空多出一整个行距。

---

## OOXML 层面的顺序约束

### 7. 单元格描边必须在设填充之前

OOXML 要求 `<a:lnL/lnR/lnT/lnB>` 排在 `<a:tcPr>` 里 `<a:solidFill>` 前面，
顺序非法的边框会被静默忽略（表格看起来没有边框，但没有报错）。

`kit.table()` 里顺序是 `_cell_ln(c)` → `c.fill.solid()`，别自己调换。

### 8. 图表藏坐标轴要写 `<c:delete val="1"/>` 且排在 `<c:scaling>` 之后

写在别处不生效，坐标轴照样出现。`kit._hide_axis()` 用 `sc.addnext(d)` 定位。

### 9. 透明度要插 `<a:alpha val="千分比"/>` 进填充的 `srgbClr`

不是 `<a:solidFill><a:alpha .../></a:solidFill>`，位置错了不生效。

### 10. 版式自带的图片不在 `slideLayoutN.xml` 里

关系文件在 `ppt/slideLayouts/_rels/`，路径写错会 KeyError。
更麻烦的是：**一张图可能不是 `<p:pic>`**。带 `blipFill` 的自定义多边形
`<p:sp>` python-pptx 和 `iter('{p}pic')` 都扫不到，
只有解析 rels + 搜 `r:embed` 才找得到——
这正是「我以为封面是空白、其实底下有东西」的原因。

---

## 排版层面的坑

### 11. 封面千万别自己加装饰

`add_slide(版式)` **只克隆占位符**，版式里的普通形状（标题带、色块、封面装饰）
由 LibreOffice / PowerPoint 自动画在下层。任何自绘色块盖上去都会毁掉它。

早期版本在封面上叠了一张网络图加深色遮罩，正好把版式自带的封面元素盖掉了。
删占位符留装饰，一行装饰都不用画。

### 12. 删幻灯片要做 XML 手术

```python
lst = prs.slides._sldIdLst
for sld in list(lst):
    prs.part.drop_rel(sld.get("{...relationships}id"))
    lst.remove(sld)
```

只 `lst.remove()` 不 `drop_rel()`，包里的 part 会变成孤儿，
文件体积暴涨，还可能让 PowerPoint 报「需要修复」。

### 13. 文字重叠要按行框判，不按块框判

PyMuPDF 的 block 是它自己分的组，块框是组内所有行的并集。
并排的左右两栏会被合成一个大框、互相「重叠」——三栏卡片页、图组页全是假阳性。

### 14. 版面按声明框高排会留大空隙

声明框高 1.14in 的描述块，实际墨迹可能只有 0.72in。
按框高留缝，页面上段会空出一条——闸门四的「提前收工 / 大空隙」就是这么来的。
**按实测墨迹估行数**（行数 × 字号 × 行距）再定下一个块的 y。

### 15. 深浅底上的文字要用不同的色

`BORDER` 在 `TINT` 上只有 4.35:1（不过 4.5 的小字线），
`ACC` 在 `BAND` 上只有 1.80:1。所以淡底卡片上的一律字用 `BAND`，
蓝底上的高亮必须用 `SKY` 或白。完整比值表见 `design-language.md`。

### 16. 结尾页也会被闸门四判缺陷

一页只有三行字、顶部空 1.2in、覆盖率 10% → 「大空隙 + 偏空」。

正解是给这类版式配大底色（内置母版的 `末尾幻灯片` 压了一整块 `SKY`），
而不是调低 `MIN_COVER` 阈值。

---

## 素材与版权

### 17. 带机构标识的素材不能进 skill 包

校名、校徽、校园照片。**内置母版是纯抽象几何、零位图**——
`assets/theme-neutral.pptx` 里 `ppt/media/` 是空的，
连 python-pptx 自带模板留下的 `docProps/thumbnail.jpeg` 也被
`make_neutral_template.py` 剥掉了：分发用的资产里不该留任何来历不明的图片。

用户自带模板里的标识是**用户自己的版权**，不阻塞流程，但要明确提醒用户
确认使用范围（见 `template-intake.md`）。

### 18. 别把手填的宽高比常数留在代码里

`ARS = {"fig1.png": 1.868, ...}` 这种表一旦对不上，闸门一和闸门二会同时把
正常图判成「拉伸」，而且很难查是哪张图。用 `kit.ar(path)`（pillow 实测），
或干脆用 `fig()` / `fig_fit()`（自动等比）。

---

## 杂项

### 19. 字体 fallback

LibreOffice 自带的字体可能缺某些字（比如 `℃` U+2103），
渲染 PDF 里会走 fallback 字体，看起来正常；Windows 的 PowerPoint 里会用
系统字体自带的那个字，也不影响放映。**这是已知的良性差异，不用管。**
但闸门二会打印 PDF 内嵌字体列表——看到意料之外的字体时先确认是不是这个原因。

### 20. 演讲者备注

用 `slide.notes_slide.notes_text_frame.text`，第一次访问会自动创建备注页，
不需要手工 `add`。备注不占版面，也不参与任何闸门。
