# Research Deck Builder

Turn research work into a deck you can actually present from. It builds from your own
template, ships the speaker notes, and checks four classes of problems before you
open PowerPoint.

**中文文档 → [README.md](README.md)**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Offline](https://img.shields.io/badge/runtime-local%20%C2%B7%20nothing%20uploaded-2ea44f)](#what-you-need)
[![Themes](https://img.shields.io/badge/3%20built-in%20themes-0072B2)](#three-themes)
[![Gates](https://img.shields.io/badge/4%20validation%20gates-success)](#the-four-gates)
[![License](https://img.shields.io/badge/license-MIT-lightgrey)](LICENSE)

---

## Layout

```
research-ppt-builder/
├── SKILL.md              Operating instructions for an AI agent (8-step workflow)
├── README.md             This file's Chinese counterpart
├── requirements.txt
├── assets/
│   ├── themes/           Three theme masters + palette definitions (pptx / json)
│   └── figures/          Sample matplotlib figures
├── references/           Deep-dive docs, read as needed
│   ├── design-language.md     Palette, grid, type: and why
│   ├── page-recipes.md        Geometry for a dozen page types
│   ├── figures-and-charts.md  Native chart vs matplotlib — which one
│   ├── image-handling.md      Three ways to handle figures
│   ├── template-intake.md     Plugging in your own pptx
│   ├── narration.md           Script writing and data discipline
│   ├── qa-gates.md            What each gate catches, how to tune
│   └── pitfalls.md            Things that bit us (check here first)
├── scripts/
│   ├── build_deck.py         Generate: a 15-page sample deck
│   ├── kit.py                 Component library: text, tables, charts, images, grid
│   ├── make_neutral_template.py  Generate theme masters
│   ├── analyze_template.py    Extract visual language from any pptx
│   ├── figstyle.py            matplotlib style presets
│   ├── qa_fit.py  qa_pdf.py  qa_visual.py  qa_layout.py   The four gates
│   └── verify_all.py          One shot: build → render → four gates
└── docs/images/          Screenshots used below
```

## What it does

**Builds from your template.** Point it at your group's `.pptx` and it opens that
file as the base, inheriting the layouts, palette, and decoration. No tracing
screenshots. A neutral built-in master works too if you have no template.

**Writes the script for you.** Slides and script read from the same source. What
lands in the speaker notes and what lands in the exported markdown are the same
text, so the numbers cannot drift apart.

**Checks four classes of problems after generating.** Text overflow, low contrast,
layout voids, template drift. None of these survive a careful look, so four
snippets of code handle them. The result is an exit code — you can put it in CI.

**Three ways to handle figures.** No images. Images already named (you say where
they go; it neither renames nor second-guesses). Images called `1.png` or `2.jpg`
(it looks at each one, decides where it belongs, and reports back for your sign-off).

**Three themes.** Business blue, scholar red, minimal monochrome. All three pass
all four gates.

**Two ways to chart.** PowerPoint native charts by default — editable, small.
matplotlib when the styling needs to be non-standard, with a preset ready to go.

## What it looks like

Six pages from the 15-page sample. The data is invented; look at the layout.

| Cover | Timeline |
|---|---|
| ![Cover](docs/images/01-cover.png) | ![Timeline](docs/images/11-timeline.png) |

| Layered cards | Comparison table |
|---|---|
| ![Layers](docs/images/03-layers.png) | ![Table](docs/images/04-table.png) |

| Quadrant | Log-scale chart |
|---|---|
| ![Quadrant](docs/images/12-quadrant.png) | ![Log chart](docs/images/14-logchart.png) |

## What you need

| | Requirement | Notes |
|---|---|---|
| Python | ≥ 3.10 | Verified on 3.12 |
| `python-pptx` | ≥ 1.0.2 | Required — generation and gates |
| `pymupdf` `pillow` | — | Required — gates read renders, compute aspect ratios |
| `matplotlib` | ≥ 3.7 | Optional, only for your own figures |
| LibreOffice | any recent build | Required, to render the deck to PDF for the gates. **Not on PATH by default — pass the full path** |

```powershell
pip install -r requirements.txt
```

`SKILL.md` is also a skill entry point, if you install it as an agent skill.

## Five-minute start

```powershell
# 1. Generate the sample deck
python scripts\build_deck.py

# 2. Validate it (LibreOffice needs its full path)
python scripts\verify_all.py --soffice "C:\Program Files\LibreOffice\program\soffice.exe"
```

Three files land in `out/`:

| File | What it is |
|---|---|
| `示例汇报_中文科研组会.pptx` | 15 slides with speaker notes |
| `示例汇报_中文科研组会_讲稿.md` | Per-slide script and anticipated Q&A |
| `out/render/*.pdf` | Rendered PDF, for checking layout |

**Make it yours:**

```powershell
Copy-Item scripts\build_deck.py my_deck.py
```

Edit `DECK` (title, footer, page count) and `PAGES` (which page types you want)
in `my_deck.py`. Copy the geometry from
[`references/page-recipes.md`](references/page-recipes.md) and change the copy.

**Use your own template:**

```powershell
python scripts\analyze_template.py "D:\your-template.pptx" -o theme.json
```

Fill in the palette semantics per
[`references/template-intake.md`](references/template-intake.md). After that, your
page code does not change at all.

**Re-run the gates after every change.** That is the one requirement.

## Common commands

```powershell
# Switch theme
python scripts\verify_all.py --theme scholar-red --soffice "C:\...\soffice.exe"

# Chinese typeface → KaiTi
python scripts\build_deck.py --zh-kaiti

# Validate an existing deck
python scripts\verify_all.py --pptx "D:\my-deck.pptx" --soffice "C:\...\soffice.exe"

# Inspect a template
python scripts\analyze_template.py "D:\your-template.pptx" -o theme.json
```

Exit codes: `0` all passed · `1` a gate failed · `2` environment incomplete
(LibreOffice missing, wrong deck path, …). **It does not fake a pass when
LibreOffice is absent** — it tells you which three gates did not run.

## The four gates

| Gate | Catches | Where it looks |
|---|---|---|
| `qa_fit.py` | Text overflow, tables overrunning the footer, stretched images, out-of-bounds | Source geometry, no rendering needed — fastest |
| `qa_pdf.py` | Out-of-bounds, footer collisions, overlapping lines, distortion, thin pages | Real positions after rendering |
| `qa_visual.py` | Per-run contrast, text escaping its box | The first two gates cannot see color |
| `qa_layout.py` | Text outside components, large gaps, early bail-outs, underfilled pages | Does it look good |

`qa_fit` deserves its own note: it is the only gate that can catch text overflow
*before* rendering. Glyph widths are measured from the real font files, not from
eyeballed coefficients — swap the typeface and eyeballed numbers are guaranteed
to be wrong.

Thresholds are tunable, but **check whether you measured wrong first** rather than
loosening them. See [`references/qa-gates.md`](references/qa-gates.md).

## Three themes

Same grid, same type; only the palette and master decoration differ.

| | academic-blue | scholar-red | minimal-mono |
|---|---|---|---|
| Base | `006DB8` | `8A2B34` | `2F2F2F` + terracotta |
| Good for | Group meetings, external talks | Proposals, defenses, reviews | Printing, external reviewers |
| Projection | Lit rooms | Lit rooms; reds skew purple on poor projectors | Either |

Full palette, measured contrast per color, and where each one must not be used:
[`references/design-language.md`](references/design-language.md).

## How this differs

Most ways of making a group-meeting deck share one property: **layout is vouched
for by your eyes.**

| Approach | Generation | Who catches overflow / contrast / voids |
|---|---|---|
| Ask an AI for `python-pptx` | One shot | Your eyes, over thumbnails — or not at all |
| Download a `.potx` template | Type into placeholders | Your eyes. When it does not fit, drag the box |
| Marp / Quarto | Markdown to slides | Style is locked; you cannot use your own template |
| PowerPoint Designer | Official auto-layout | Visible in the UI, but no executable regression |

Four differences here:

1. **The check can run in CI.** Others hand you a deck that *looked* fine. This one
   hands you a deck a machine went through line by line, and the exit code is the verdict.
2. **Template reuse is inheritance, not imitation.** Decoration in a layout is an
   ordinary shape, so `add_slide(layout)` paints it underneath automatically.
   Not one decoration is drawn by hand, and none can drift.
3. **Chinese-typography traps live in the code.** `<a:ea>` must come *after*
   `<a:latin>` — in the wrong order PowerPoint silently ignores it, so the
   typeface name is written into the XML and nothing changes. You cannot know this
   without hitting it once, so it stays in comments and
   [`references/pitfalls.md`](references/pitfalls.md).
4. **Slides and script share a source.** Both read `kit.SCRIPT`, so the slide
   cannot say A while the narration says B.

The honest cost: **this is not a general slide framework.** It is built for one
job — talking at a projected deck. Posters, vertical layouts, and animation will
fight it. Changing the palette is easy (three built in); changing the grid means
rebuilding every page.

## When something breaks

| Symptom | Read |
|---|---|
| `KeyError: missing layouts [...]` | [`references/template-intake.md`](references/template-intake.md) |
| A gate failed, you want to tune it | [`references/qa-gates.md`](references/qa-gates.md) |
| Images distorted or not found | [`references/image-handling.md`](references/image-handling.md) |
| LibreOffice produces no PDF | `references/qa-gates.md` (several silent-failure traps) |
| Anything else odd | [`references/pitfalls.md`](references/pitfalls.md) |

It feels most natural inside agentic IDEs such as Z Code: sessions persist, the
project directory is wired in directly, and changing one line of page code means
re-running the gates immediately.

## License and caveats

- [MIT](LICENSE).
- All sample data is **invented** and obviously so. Replace it with your real
  numbers, and follow the data discipline in
  [`references/narration.md`](references/narration.md).
- The built-in masters carry no institutional marks and no bitmaps. Sample
  figures in `assets/figures/` are generated by this package.
- **When you bring your own template**, its logos, photos, and institutional
  identity are yours to clear. This package does not judge that for you.
- `docProps` metadata is scrubbed at save time (author, generator, timestamps),
  so generated decks are safe to publish.
