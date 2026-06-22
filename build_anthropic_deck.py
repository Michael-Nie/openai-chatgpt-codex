# -*- coding: utf-8 -*-
"""Build the Anthropic technical-insight deck (PPTX).

Theme: Agent 时代的"自举飞轮" —— 战略定位 × Agent 能力的协同演化
7 information-dense slides, consulting-minimal style, charts-led.
"""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# ---------- palette (consulting-minimal · single warm clay accent) ----------
BG     = RGBColor(0x14, 0x22, 0x38)   # deep slate (header / cover)
BG_CHIP = RGBColor(0x2A, 0x41, 0x5C)  # dark slate chip (loop/process boxes)
INK    = RGBColor(0x1E, 0x26, 0x30)   # body text
CLAY   = RGBColor(0xC1, 0x5F, 0x3C)   # primary accent (Anthropic-warm)
CLAYD  = RGBColor(0x9A, 0x46, 0x2A)   # darker clay
CLAYL  = RGBColor(0xF0, 0xDF, 0xD5)   # light clay fill
SLATE  = RGBColor(0x35, 0x5C, 0x7D)   # secondary accent (2nd category)
SLATEL = RGBColor(0xDD, 0xE6, 0xEC)   # light slate fill
GRAY   = RGBColor(0x6B, 0x72, 0x80)
LGRAY  = RGBColor(0xE5, 0xE7, 0xEB)
CARD   = RGBColor(0xF5, 0xF3, 0xF0)   # warm off-white card
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
AMBER  = RGBColor(0xB4, 0x53, 0x09)
GREEN  = RGBColor(0x2F, 0x80, 0x5F)
RED    = RGBColor(0xB9, 0x1C, 0x1C)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ============================================================ low-level helpers
def set_font(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', name)

def _fill(sp, color):
    if color is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = color

def _line(sp, color, w=1.0):
    if color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = color; sp.line.width = Pt(w)

def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None):
    sp = slide.shapes.add_shape(kind, x, y, w, h)
    _fill(sp, fill); _line(sp, line, line_w)
    sp.shadow.inherit = False
    if radius is not None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp

def rect(slide, x, y, w, h, color, line=None, line_w=1.0):
    return shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=color, line=line, line_w=line_w)

def rrect(slide, x, y, w, h, color, line=None, line_w=1.0, radius=0.08):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=color, line=line, line_w=line_w, radius=radius)

def textbox(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    return tb, tf

def style(p, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT, space_after=4, italic=False):
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    set_font(r)
    return r

def fill_text(sp, paras, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, wrap=True):
    """Put multi-paragraph text inside a shape. paras: list of dicts."""
    tf = sp.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.06); tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.03); tf.margin_bottom = Inches(0.03)
    first = True
    for pr in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = pr.get("align", align)
        p.space_after = Pt(pr.get("space_after", 1)); p.space_before = Pt(0)
        # support multiple runs in one paragraph via "runs"
        if "runs" in pr:
            for rr in pr["runs"]:
                r = p.add_run(); r.text = rr["t"]
                r.font.size = Pt(rr.get("size", 11)); r.font.bold = rr.get("bold", False)
                r.font.italic = rr.get("italic", False)
                r.font.color.rgb = rr.get("color", INK)
                set_font(r)
        else:
            r = p.add_run(); r.text = pr["t"]
            r.font.size = Pt(pr.get("size", 11)); r.font.bold = pr.get("bold", False)
            r.font.italic = pr.get("italic", False)
            r.font.color.rgb = pr.get("color", INK)
            set_font(r)
    return sp

def tb_text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    tb, tf = textbox(slide, x, y, w, h)
    tf.vertical_anchor = anchor
    first = True
    for pr in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = pr.get("align", align)
        p.space_after = Pt(pr.get("space_after", 3)); p.space_before = Pt(pr.get("space_before", 0))
        if "runs" in pr:
            for rr in pr["runs"]:
                r = p.add_run(); r.text = rr["t"]
                r.font.size = Pt(rr.get("size", 11)); r.font.bold = rr.get("bold", False)
                r.font.italic = rr.get("italic", False)
                r.font.color.rgb = rr.get("color", INK)
                set_font(r)
        else:
            r = p.add_run(); r.text = pr["t"]
            r.font.size = Pt(pr.get("size", 11)); r.font.bold = pr.get("bold", False)
            r.font.italic = pr.get("italic", False)
            r.font.color.rgb = pr.get("color", INK)
            set_font(r)
    return tb

def connector(slide, x1, y1, x2, y2, color=CLAY, w=2.0, head=True, tail=False, dash=None):
    cxn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    cxn.line.color.rgb = color; cxn.line.width = Pt(w)
    ln = cxn.line._get_or_add_ln()
    if tail:
        ln.append(ln.makeelement(qn('a:headEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    if head:
        ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    if dash:
        ln.append(ln.makeelement(qn('a:prstDash'), {'val': dash}))
    cxn.shadow.inherit = False
    return cxn

# ============================================================ chrome
def header(slide, kicker, title):
    rect(slide, 0, 0, SW, Inches(1.12), BG)
    rect(slide, 0, Inches(1.12), SW, Pt(3), CLAY)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.13), Inches(12.2), Inches(0.92))
    p = tf.paragraphs[0]; style(p, kicker, 11, CLAYL, bold=True, space_after=2)
    p2 = tf.add_paragraph(); style(p2, title, 23, WHITE, bold=True)

def footer(slide, n, note=""):
    tb, tf = textbox(slide, Inches(0.55), Inches(7.08), Inches(11.2), Inches(0.34))
    p = tf.paragraphs[0]; style(p, note, 8, GRAY)
    tb2, tf2 = textbox(slide, Inches(12.35), Inches(7.08), Inches(0.75), Inches(0.34))
    p2 = tf2.paragraphs[0]; style(p2, str(n), 9, GRAY, align=PP_ALIGN.RIGHT)

DEFAULT_NOTE = "可信度标注:〔强〕官方+多源 ·〔中〕单一媒体/厂商自报 ·〔向〕方向性信号(自报/早期重度用户样本) · 2026 项超模型知识截止,引用前对照官方原文"

def content_slide(kicker, title, n, note=DEFAULT_NOTE):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, WHITE)
    header(s, kicker, title)
    footer(s, n, note)
    return s

def section_label(slide, x, y, w, text, color=CLAY):
    """small uppercase section tag with a leading rule"""
    rect(slide, x, y + Inches(0.02), Inches(0.10), Inches(0.20), color)
    tb_text(slide, x + Inches(0.18), y, w, Inches(0.28),
            [{"t": text, "size": 12.5, "color": color, "bold": True}])

def tag(slide, x, y, text, fill=CLAYL, fg=CLAYD):
    w = Inches(0.30 + 0.085 * len(text))
    sp = rrect(slide, x, y, w, Inches(0.26), fill, radius=0.5)
    fill_text(sp, [{"t": text, "size": 9, "color": fg, "bold": True}])
    return sp

# ============================================================ diagram helpers
def metric_card(slide, x, y, w, h, big, label, sub="", accent=CLAY, big_size=24):
    rrect(slide, x, y, w, h, CARD, line=LGRAY, line_w=0.75, radius=0.08)
    rect(slide, x, y, Inches(0.08), h, accent)
    paras = [{"t": big, "size": big_size, "color": accent, "bold": True, "space_after": 1, "align": PP_ALIGN.LEFT},
             {"t": label, "size": 10.5, "color": INK, "bold": True, "space_after": 0, "align": PP_ALIGN.LEFT}]
    if sub:
        paras.append({"t": sub, "size": 8.5, "color": GRAY, "align": PP_ALIGN.LEFT})
    tb_text(slide, x + Inches(0.20), y + Inches(0.06), w - Inches(0.26), h - Inches(0.10),
            paras, anchor=MSO_ANCHOR.MIDDLE)

def arrow_progression(slide, x, y, w, steps, h=Inches(0.34), fill=CLAYL, fg=CLAYD,
                      size=9.5, gap=Inches(0.16)):
    """Row of small boxes joined by arrows: a -> b -> c."""
    n = len(steps)
    bw = (w - gap * (n - 1)) / n
    cx = x
    for i, s in enumerate(steps):
        sp = rrect(slide, cx, y, bw, h, fill, radius=0.18)
        fill_text(sp, [{"t": s, "size": size, "color": fg, "bold": True}])
        if i < n - 1:
            ax1 = cx + bw + Emu(int(gap * 0.12))
            ax2 = cx + bw + gap - Emu(int(gap * 0.12))
            ymid = y + Emu(int(h / 2))
            connector(slide, ax1, ymid, ax2, ymid, color=fg, w=1.75)
        cx = cx + bw + gap

def chevron_row(slide, x, y, w, h, cards, gap=Inches(0.10)):
    """Row of right-pointing pentagon/chevron stage cards."""
    n = len(cards)
    cw = (w - gap * (n - 1)) / n
    cx = x
    out = []
    for i, c in enumerate(cards):
        kind = MSO_SHAPE.PENTAGON  # arrow-like
        sp = shape(slide, kind, cx, y, cw, h, fill=c.get("fill", BG), radius=None)
        out.append((sp, cx, cw))
        cx = cx + cw + gap
    return out

def loop_cycle(slide, cx, cy, R, nodes, center_paras, node_w, node_h,
               ring_fill=CLAYL, node_fill=WHITE, node_line=CLAY, accent=CLAY):
    """Flywheel: donut ring + N nodes around it + center label + direction arrows."""
    D = int(2 * R)
    # ring (donut)
    ring = shape(slide, MSO_SHAPE.DONUT, Emu(int(cx - R)), Emu(int(cy - R)), Emu(D), Emu(D),
                 fill=ring_fill)
    try:
        ring.adjustments[0] = 0.16
    except Exception:
        pass
    n = len(nodes)
    positions = []
    for i in range(n):
        ang = -90 + i * (360.0 / n)   # start at top, clockwise
        rad = math.radians(ang)
        nx = cx + R * math.cos(rad)
        ny = cy + R * math.sin(rad)
        positions.append((nx, ny, ang))
    # direction arrowheads on the ring (between nodes)
    for i in range(n):
        a0 = -90 + i * (360.0 / n)
        amid = a0 + (360.0 / n) / 2.0
        rad = math.radians(amid)
        tx = cx + R * math.cos(rad)
        ty = cy + R * math.sin(rad)
        tsz = Inches(0.26)
        tri = shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE,
                    Emu(int(tx - tsz / 2)), Emu(int(ty - tsz / 2)), Emu(int(tsz)), Emu(int(tsz)),
                    fill=accent)
        tri.rotation = (amid + 90) % 360   # point clockwise/tangential
    # nodes
    for i, (nx, ny, ang) in enumerate(positions):
        sp = rrect(slide, Emu(int(nx - node_w / 2)), Emu(int(ny - node_h / 2)),
                   Emu(int(node_w)), Emu(int(node_h)), node_fill, line=node_line, line_w=1.5, radius=0.14)
        num = "①②③④⑤⑥⑦"[i]
        fill_text(sp, [{"runs": [{"t": num + " ", "size": 10, "color": accent, "bold": True},
                                 {"t": nodes[i], "size": 10, "color": INK, "bold": True}]}])
    # center
    cd = int(R * 0.95)
    cc = shape(slide, MSO_SHAPE.OVAL, Emu(int(cx - cd / 2)), Emu(int(cy - cd / 2)),
               Emu(cd), Emu(cd), fill=accent)
    fill_text(cc, center_paras)

def figcap(slide, x, y, w, num, title, source):
    """Sell-side-report style figure caption: 图N: title  /  资料来源: ..."""
    tb_text(slide, x, y, w, Inches(0.42), [
        {"runs": [{"t": num + ": ", "size": 8.5, "color": CLAY, "bold": True},
                  {"t": title, "size": 8.5, "color": INK, "bold": True}], "space_after": 1},
        {"t": "资料来源: " + source, "size": 7.5, "color": GRAY},
    ])

# ============================================================ SLIDE 1 — OVERVIEW
def slide1():
    s = content_slide("01 · STRATEGIC × AGENT CO-EVOLUTION", "战略定位演进 × Agent 能力协同演化", 2,
                      note="图1: Anthropic 定位四段迁移 × Agent 能力协同演化 · 资料来源: Anthropic 官方公告、Claude Code / Agent SDK / MCP 文档,本研究整理")
    # thesis one-liner band
    th = rrect(s, Inches(0.55), Inches(1.26), Inches(12.23), Inches(0.50), CLAYL, radius=0.06)
    rect(s, Inches(0.55), Inches(1.26), Inches(0.10), Inches(0.50), CLAY)
    fill_text(th, [{"runs": [
        {"t": "核心论点　", "size": 11, "color": CLAYD, "bold": True},
        {"t": "定位四段迁移 = Agent 能力逐级上移,同一条协同演化曲线;领先本质 = 速度系统 + Agent 协同演化能力,而非单点模型分数。",
         "size": 11, "color": INK}]}],
        anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

    # --- four-stage timeline ---
    section_label(s, Inches(0.55), Inches(1.94), Inches(7), "四段定位迁移  ×  Agent 能力重心  (Co-evolution Timeline)")
    stages = [
        ("01", "安全研究实验室", "Safety Lab", "2021–23", "Chat 调用", "对齐研究 / RLHF / 宪法式 AI;模型即「被调用的对话能力」"),
        ("02", "前沿模型 / 企业 API", "Frontier Model / API", "2024–25", "Tool Use", "Claude 3/3.5/4 + API;模型学会调用工具,能力开始「外接世界」"),
        ("03", "Agent 平台", "Agent Platform", "2025", "Claude Code 自主 Agent", "Agent Loop + Sub-agents + MCP;模型自主执行多步任务"),
        ("04", "组织操作系统", "Organizational OS", "2025末–26", "动态编排 / 托管多 Agent", "Dynamic Workflows + Managed Agents;AI 造 AI,编排成组 Agent"),
    ]
    x0 = Inches(0.55); top = Inches(2.34); gap = Inches(0.22)
    cw = (Inches(12.23) - gap * 3) / 4
    ch = Inches(1.14)
    for i, (num, cn, en, yr, cap, capdesc) in enumerate(stages):
        cx = x0 + (cw + gap) * i
        # stage card
        card = rrect(s, cx, top, cw, ch, CARD, line=LGRAY, line_w=0.75, radius=0.06)
        rect(s, cx, top, cw, Inches(0.06), CLAY)
        tb_text(s, cx + Inches(0.14), top + Inches(0.12), cw - Inches(0.24), ch - Inches(0.2), [
            {"runs": [{"t": num + "  ", "size": 13, "color": CLAY, "bold": True},
                      {"t": yr, "size": 10, "color": GRAY, "bold": True}], "space_after": 2},
            {"t": cn, "size": 14, "color": INK, "bold": True, "space_after": 0},
            {"t": en, "size": 9.5, "color": GRAY, "space_after": 0},
        ])
        # connecting arrow
        if i < 3:
            ymid = top + Emu(int(ch / 2))
            connector(s, cx + cw + Emu(int(gap*0.10)), ymid, cx + cw + gap - Emu(int(gap*0.10)), ymid, color=CLAY, w=2.25)
        # capability layer card (below)
        capy = top + ch + Inches(0.14)
        cap_card = rrect(s, cx, capy, cw, Inches(0.90), SLATEL, line=None, radius=0.06)
        rect(s, cx, capy, Inches(0.06), Inches(0.90), SLATE)
        tb_text(s, cx + Inches(0.14), capy + Inches(0.08), cw - Inches(0.24), Inches(0.78), [
            {"runs": [{"t": "▸ Agent 重心: ", "size": 9, "color": SLATE, "bold": True},
                      {"t": cap, "size": 10, "color": INK, "bold": True}], "space_after": 2},
            {"t": capdesc, "size": 8.5, "color": GRAY, "space_after": 0},
        ])

    # axis hint (right-aligned, in the section-label row)
    tb_text(s, Inches(8.5), Inches(1.94), Inches(4.28), Inches(0.26),
            [{"t": "自主性 / 托管度逐级上移  ▸▸▸", "size": 9, "color": CLAY, "bold": True, "align": PP_ALIGN.RIGHT}])
    figcap(s, Inches(0.55), Inches(4.58), Inches(9.5), "图1",
           "Anthropic 定位四段迁移 × Agent 能力协同演化", "Anthropic 官方公告、Claude Code 文档,本研究整理")

    # --- three trends ---
    section_label(s, Inches(0.55), Inches(5.14), Inches(6), "三条贯穿趋势  (Through-lines)")
    trends = [
        ("①", "自主度上移", ["人写脚本", "模型编排", "自主执行"]),
        ("②", "执行形态延伸", ["前台 harness", "异步/后台", "托管多 Agent"]),
        ("③", "护城河迁移", ["model score", "Agent Runtime", "+ 协议生态(MCP)"]),
    ]
    ty = Inches(5.50); tcw = (Inches(12.23) - Inches(0.4) * 2) / 3
    for i, (num, lab, steps) in enumerate(trends):
        tx = Inches(0.55) + (tcw + Inches(0.4)) * i
        box = rrect(s, tx, ty, tcw, Inches(1.06), WHITE, line=LGRAY, line_w=0.75, radius=0.05)
        tb_text(s, tx + Inches(0.16), ty + Inches(0.10), tcw - Inches(0.3), Inches(0.34),
                [{"runs": [{"t": num + " ", "size": 13, "color": CLAY, "bold": True},
                           {"t": lab, "size": 12.5, "color": INK, "bold": True}]}])
        arrow_progression(s, tx + Inches(0.16), ty + Inches(0.56), tcw - Inches(0.32), steps,
                          h=Inches(0.34), size=8.5)

# ============================================================ extra diagram helpers
def callout_card(slide, x, y, w, h, title, lines, accent=CLAY, fill=CARD,
                 title_size=12, body_size=10, anchor=MSO_ANCHOR.TOP):
    rrect(slide, x, y, w, h, fill, line=LGRAY, line_w=0.75, radius=0.05)
    rect(slide, x, y, Inches(0.08), h, accent)
    paras = [{"t": title, "size": title_size, "color": accent, "bold": True, "space_after": 4, "align": PP_ALIGN.LEFT}]
    for ln in lines:
        if isinstance(ln, dict):
            d = {"space_after": ln.get("space_after", 3), "align": PP_ALIGN.LEFT}
            if "runs" in ln:
                d["runs"] = ln["runs"]
            else:
                d.update({"t": ln["t"], "size": ln.get("size", body_size),
                          "color": ln.get("color", INK), "bold": ln.get("bold", False)})
            paras.append(d)
        else:
            paras.append({"runs": [{"t": "▸ ", "size": body_size, "color": accent, "bold": True},
                                   {"t": ln, "size": body_size, "color": INK}],
                          "space_after": 3, "align": PP_ALIGN.LEFT})
    tb_text(slide, x + Inches(0.22), y + Inches(0.11), w - Inches(0.34), h - Inches(0.20),
            paras, anchor=anchor)

def axis_row(slide, x, y, w, axis_label, segs, caption, h=Inches(0.40)):
    lbl_w = Inches(1.05)
    tb_text(slide, x, y - Inches(0.02), lbl_w, h + Inches(0.04),
            [{"t": axis_label, "size": 11, "color": CLAY, "bold": True}], anchor=MSO_ANCHOR.MIDDLE)
    bx = x + lbl_w; bw = w - lbl_w
    n = len(segs); seg_gap = Inches(0.12)
    sw = (bw - seg_gap * (n - 1)) / n
    fills = [SLATEL, CLAYL, CLAY]; fgs = [SLATE, CLAYD, WHITE]
    cx = bx
    for i, sg in enumerate(segs):
        sp = rrect(slide, cx, y, sw, h, fills[i % 3], radius=0.14)
        fill_text(sp, [{"t": sg, "size": 8.8, "color": fgs[i % 3], "bold": True}])
        if i < n - 1:
            ym = y + Emu(int(h / 2))
            connector(slide, cx + sw + Emu(int(seg_gap * 0.10)), ym,
                      cx + sw + seg_gap - Emu(int(seg_gap * 0.10)), ym, color=CLAY, w=1.5)
        cx = cx + sw + seg_gap
    tb_text(slide, bx, y + h + Inches(0.01), bw, Inches(0.22),
            [{"t": caption, "size": 8, "color": GRAY}])

def flow_down(slide, x, y, w, boxes, box_h, gap):
    cy = y
    for i, b in enumerate(boxes):
        sp = rrect(slide, x, cy, w, box_h, b.get("fill", SLATEL), line=b.get("line"), line_w=1.0, radius=0.06)
        fill_text(sp, [{"runs": [{"t": b.get("pre", ""), "size": b.get("psize", 10), "color": b.get("pcolor", SLATE), "bold": True},
                                 {"t": b["t"], "size": b.get("size", 11), "color": b.get("color", INK), "bold": True}]}])
        if i < len(boxes) - 1:
            cxm = x + Emu(int(w / 2))
            connector(slide, cxm, cy + box_h + Emu(int(gap * 0.12)), cxm, cy + box_h + gap - Emu(int(gap * 0.12)),
                      color=CLAY, w=2.0)
        cy = cy + box_h + gap

def signpost(slide, x, y, w, h, num, title, status, dotcolor):
    rrect(slide, x, y, w, h, WHITE, line=LGRAY, line_w=0.75, radius=0.06)
    rect(slide, x, y, w, Inches(0.07), dotcolor)
    dot = shape(slide, MSO_SHAPE.OVAL, x + Inches(0.16), y + Inches(0.20), Inches(0.20), Inches(0.20), fill=dotcolor)
    tb_text(slide, x + Inches(0.44), y + Inches(0.16), w - Inches(0.54), Inches(0.30),
            [{"t": num, "size": 13, "color": dotcolor, "bold": True}])
    tb_text(slide, x + Inches(0.16), y + Inches(0.52), w - Inches(0.30), h - Inches(0.62),
            [{"t": title, "size": 10, "color": INK, "bold": True, "space_after": 3},
             {"t": status, "size": 8.6, "color": GRAY}])

def bar_chart(slide, x, y, w, h, bars, ymax):
    base_y = y + h
    n = len(bars); gap = Inches(0.22)
    bw = (w - gap * (n - 1)) / n
    connector(slide, x - Inches(0.05), base_y, x + w, base_y, color=GRAY, w=1.0, head=False)
    cx = x
    for (lab, val, disp, color, tg) in bars:
        bh = Emu(int(h * val / ymax))
        rect(slide, cx, base_y - bh, bw, bh, color)
        tb_text(slide, cx - Inches(0.14), base_y - bh - Inches(0.50), bw + Inches(0.28), Inches(0.48),
                [{"t": disp, "size": 11, "color": INK, "bold": True, "align": PP_ALIGN.CENTER, "space_after": 0},
                 {"t": tg, "size": 7.5, "color": GRAY, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.BOTTOM)
        tb_text(slide, cx - Inches(0.14), base_y + Inches(0.04), bw + Inches(0.28), Inches(0.42),
                [{"t": lab, "size": 8.6, "color": GRAY, "align": PP_ALIGN.CENTER}])
        cx = cx + bw + gap

# ============================================================ SLIDE 2 — STACK & PARADIGM
def slide2():
    s = content_slide("02 · STACK & PARADIGM · 能力谱系三轴", "Anthropic Agent 栈与能力范式:可控可验证的 Runtime", 3,
                      note="图2: Agent 能力谱系三轴与 Claude Code Runtime · 资料来源: Claude Code / Agent SDK / Managed Agents 文档;指标引自《How AI Is Transforming Work at Anthropic》(2025-12) ·〔强〕")
    # ---- left column: three-axis capability spectrum
    section_label(s, Inches(0.55), Inches(1.28), Inches(6), "能力谱系三轴  ·  Capability Spectrum")
    LX, LW = Inches(0.55), Inches(6.05)
    axis_row(s, LX, Inches(1.78), LW, "自主度", ["Static Workflow", "Dynamic Workflows", "Autonomous Agent"],
             "人写流程 → 运行时模型驱动编排(Agent SDK)→ 自主执行")
    axis_row(s, LX, Inches(2.58), LW, "托管度", ["Local Harness", "Async / 后台", "Managed Agents"],
             "本地 harness → 异步/后台长任务 → 托管(Anthropic 跑 harness+沙箱)")
    axis_row(s, LX, Inches(3.38), LW, "协作", ["单 Agent", "Sub-agent 编排", "多 Agent 编排"],
             "单体 → 子 Agent 独立上下文 → Dynamic Workflows 千级并行")
    callout_card(s, LX, Inches(4.28), LW, Inches(0.98), "设计哲学:可控可验证 Runtime,而非黑箱托管", [
        {"runs": [{"t": "权限边界 + 人在环验证 + 沙箱/审计 ", "size": 9.5, "color": INK, "bold": True},
                  {"t": "让自主度可以安全上移;能力沉淀进可被检验的 Agent Loop。", "size": 9.5, "color": INK}]},
    ], accent=CLAY, title_size=11.5)

    # ---- right column: Claude Code architecture
    RX, RW = Inches(6.85), Inches(5.93)
    section_label(s, RX, Inches(1.28), Inches(6), "Claude Code 架构  ·  Agent Runtime")
    arrow_progression(s, RX, Inches(1.80), RW, ["目标", "工具调用", "验证", "上下文管理"],
                      h=Inches(0.42), size=10.5, fill=BG_CHIP, fg=WHITE)
    tb_text(s, RX, Inches(2.32), RW, Inches(0.24),
            [{"runs": [{"t": "↺ ", "size": 10, "color": CLAY, "bold": True},
                       {"t": "循环至目标达成 · ~92% 上下文自动压缩(compaction)", "size": 9, "color": GRAY}]}])
    callout_card(s, RX, Inches(2.66), RW, Inches(0.80), "Sub-agents · 子 Agent", [
        {"t": "独立上下文窗口 + 自定系统提示 + 专属工具/权限;主 Agent 委派、子 Agent 回传摘要(隔离上下文)。", "size": 9.5}],
        accent=SLATE, title_size=11)
    callout_card(s, RX, Inches(3.52), RW, Inches(0.80), "权限边界 · Permission Boundary", [
        {"t": "ask / auto-accept / plan / auto 多档;checkpoints 可回滚、allowlist 预批、组织-项目-个人分层授权。", "size": 9.5}],
        accent=SLATE, title_size=11)
    callout_card(s, RX, Inches(4.38), RW, Inches(0.88), "多入口 · 一次改进多界面释放", [
        {"t": "终端 CLI · VS Code / JetBrains · 桌面 App · Web/iOS(2025-10)· GitHub Actions/GitLab CI · Slack;另有 Agent SDK(可编程)与 MCP(统一工具接口)。", "size": 9.5}],
        accent=SLATE, title_size=11)

    # ---- bottom band: capability-trajectory metrics
    section_label(s, Inches(0.55), Inches(5.46), Inches(9),
                  "可观测的能力跃迁(Anthropic 内部 · 2025 年 2 月 → 8 月)")
    bw = (Inches(12.23) - Inches(0.27) * 2) / 3
    bx = Inches(0.55); by = Inches(5.84)
    metric_card(s, bx, by, bw, Inches(1.05), "9.8 → 21.2", "连续动作 / 链式工具调用",
                "无人工干预的连续工具调用(峰值 +116%)", accent=CLAY, big_size=22)
    metric_card(s, bx + bw + Inches(0.27), by, bw, Inches(1.05), "3.2 → 3.8", "任务复杂度",
                "1–5 分制(1=基础编辑,5=专家级任务)", accent=CLAY, big_size=22)
    metric_card(s, bx + (bw + Inches(0.27)) * 2, by, bw, Inches(1.05), "14% → 37%", "「实现新功能」占比",
                "指 Claude Code 使用中投入新功能的比例(非「功能由 AI 造」)", accent=CLAY, big_size=22)

# ============================================================ SLIDE 3 — FLYWHEEL
def slide3():
    s = content_slide("03 · THE BOOTSTRAPPING FLYWHEEL · AI building AI", "自举飞轮:编码是 AI 造 AI 的高可信闭环", 4,
                      note="图4: 自举飞轮与内部研究证据 · 资料来源: Anthropic《How AI Is Transforming Work at Anthropic》(2025-12)、《When AI builds itself》(2026-06) · 框架: 早期可测加速,非自主 RSI")
    # ---- left: flywheel
    section_label(s, Inches(0.55), Inches(1.28), Inches(5.6), "自举飞轮  ·  Bootstrapping Flywheel")
    nodes = ["模型变强", "Claude Code 变强", "内部研发提速", "真实反馈回流", "反哺下一代模型"]
    loop_cycle(s, Inches(3.25), Inches(3.72), Inches(1.34), nodes,
               [{"t": "自举飞轮", "size": 12.5, "color": WHITE, "bold": True, "space_after": 1},
                {"t": "AI building AI", "size": 8.5, "color": WHITE}],
               node_w=Inches(1.92), node_h=Inches(0.64))

    # ---- right: why coding is a high-trust loop + triple identity
    RX, RW = Inches(6.10), Inches(6.68)
    section_label(s, RX, Inches(1.28), Inches(6.6), "为何编码 = 高可信闭环")
    callout_card(s, RX, Inches(1.70), RW, Inches(1.92), "可验证性:编码闭环的核心", [
        "清晰输入输出:issue / spec 进,补丁 + 通过的测试出",
        {"runs": [{"t": "▸ ", "size": 10, "color": CLAY, "bold": True},
                  {"t": "可自动验证:", "size": 10, "color": INK, "bold": True},
                  {"t": "「代码能跑、测试是否通过」即确定性评分(Demystifying evals, 2026-01)", "size": 10, "color": INK}]},
        {"runs": [{"t": "▸ ", "size": 10, "color": CLAY, "bold": True},
                  {"t": "失败轨迹 + 测试结果 = 训练/评测信号", "size": 10, "color": INK, "bold": True},
                  {"t": "(Code RL:reward signals & verifiers)", "size": 10, "color": INK}]},
        {"runs": [{"t": "▸ ", "size": 10, "color": CLAY, "bold": True},
                  {"t": "Dario:", "size": 10, "color": INK, "bold": True},
                  {"t": "可验证任务(编码)1–2 年内可达;不可验证任务才是真不确定性", "size": 10, "color": INK}]},
    ], accent=CLAY, title_size=11.5)
    tb_text(s, RX, Inches(3.78), RW, Inches(0.26),
            [{"t": "编码的三重身份", "size": 11.5, "color": CLAY, "bold": True}])
    tw = (RW - Inches(0.24) * 2) / 3; ty = Inches(4.10)
    triples = [("营收产品", "Claude Code 成营收主轴"),
               ("困难任务分发器", "把真实难题喂给模型"),
               ("模型改进来源", "轨迹/反馈回流训练")]
    for i, (t1, t2) in enumerate(triples):
        cx = RX + (tw + Inches(0.24)) * i
        callout_card(s, cx, ty, tw, Inches(1.02), t1, [{"t": t2, "size": 8.8}],
                     accent=SLATE, title_size=10.5)

    # ---- bottom: evidence band
    section_label(s, Inches(0.55), Inches(5.40), Inches(11),
                  "内部研究证据 ·《How AI Is Transforming Work at Anthropic》(Anthropic, 2025-12)")
    bw = (Inches(12.23) - Inches(0.22) * 3) / 4
    bx = Inches(0.55); by = Inches(5.78)
    metric_card(s, bx, by, bw, Inches(1.06), "132 · 53 · 200K", "样本规模",
                "工程师 · 深度访谈 · Claude Code 会话", accent=SLATE, big_size=17)
    metric_card(s, bx + (bw + Inches(0.22)), by, bw, Inches(1.06), "+50%", "生产率",
                "员工自报 · 全公司平均(用量 28%→59%)", accent=CLAY, big_size=24)
    metric_card(s, bx + (bw + Inches(0.22)) * 2, by, bw, Inches(1.06), "+67%", "merged PR",
                "人均每日 · Engineering 全面采用后", accent=CLAY, big_size=24)
    metric_card(s, bx + (bw + Inches(0.22)) * 3, by, bw, Inches(1.06), "27%", "本不会做的任务",
                "占 Claude 辅助工作的比例(非全部任务)", accent=CLAY, big_size=24)

# ============================================================ SLIDE 4 — EXPANSION
def slide4():
    s = content_slide("04 · EXPANSION · 可验证性是外扩速度的瓶颈变量", "从 Coding Agent 到通用知识工作 Agent", 5,
                      note="图5: 共享栈与编码→通用知识工作外扩 · 资料来源: Anthropic Claude Cowork(2026-01)、Agent Skills/插件、Claude Code/Agent SDK 文档 ·〔向〕外扩节奏为方向性判断")
    section_label(s, Inches(0.55), Inches(1.28), Inches(11), "迁移逻辑:编码可验证性最高 → 先突破 → 再外扩")
    arrow_progression(s, Inches(0.55), Inches(1.70), Inches(12.23),
                      ["编码:可验证性最高", "打穿楔子(先突破)", "共享栈复用", "外扩通用知识工作"],
                      h=Inches(0.48), size=11.5, fill=CLAYL, fg=CLAYD)
    tb_text(s, Inches(0.55), Inches(2.20), Inches(12.23), Inches(0.22),
            [{"runs": [{"t": "图5: ", "size": 8, "color": CLAY, "bold": True},
                       {"t": "编码楔子 → 平台外扩路径(高价值任务 = 端到端自主闭环) · 资料来源: 本研究整理、NVIDIA GTC 场景分层", "size": 8, "color": GRAY}]}])
    # left: carriers
    LX, LW = Inches(0.55), Inches(6.02)
    callout_card(s, LX, Inches(2.46), LW, Inches(1.28), "载体 ① Cowork", [
        {"t": "把 Claude Code 能力带到通用知识工作的桌面级 Agent:授权文件夹后自主读写、跨本地应用与连接器(Drive/Gmail)交付成品。", "size": 9.6},
        {"t": "研究预览 2026-01 · 领域插件(财务 / HR / 工程)。", "size": 9.6, "color": GRAY},
    ], accent=CLAY, title_size=12)
    callout_card(s, LX, Inches(3.86), LW, Inches(1.30), "载体 ② 开源插件 / Agent Skills", [
        {"t": "销售 / 法务 / 财务等角色专精能力包;同一套栈把内部沉淀快速产品化为外部「角色 Agent」。", "size": 9.6},
        {"t": "MCP 统一工具接口让连接器与技能即插即用。", "size": 9.6, "color": GRAY},
    ], accent=CLAY, title_size=12)
    # right: shared-stack diagram
    RX, RW = Inches(6.76), Inches(6.02)
    section_label(s, RX, Inches(2.42), Inches(6), "共享栈优势:同一引擎,内外复用")
    flow_down(s, RX + Inches(0.55), Inches(2.78), RW - Inches(1.1), [
        {"pre": "底层  ", "t": "Model + Agent Loop + MCP", "fill": SLATEL, "pcolor": SLATE, "color": INK},
        {"pre": "沉淀  ", "t": "内部沉淀的 Agent 能力", "fill": CLAYL, "pcolor": CLAYD, "color": INK},
        {"pre": "产品  ", "t": "产品化为外部「角色 Agent」", "fill": CLAY, "pcolor": WHITE, "color": WHITE},
    ], box_h=Inches(0.62), gap=Inches(0.30))
    # bottom: bottleneck + path judgment
    callout_card(s, LX, Inches(5.34), LW, Inches(1.36), "核心瓶颈:非编码领域「可验证性」弱", [
        {"t": "评测标准模糊、反馈回流慢 —— 没有「跑测试即评分」的确定性信号。", "size": 9.8},
        {"t": "→ 可验证性 = 外扩速度的瓶颈变量。", "size": 9.8, "color": AMBER, "bold": True},
    ], accent=AMBER, title_size=11.5)
    callout_card(s, RX, Inches(5.34), RW, Inches(1.36), "路径判断", [
        {"t": "先守编码楔子,再平台扩张;可验证性强的领域优先被 Agent 化。", "size": 9.8},
        {"t": "弱可验证领域:靠人在环 + 评测建设逐步补齐反馈闭环。", "size": 9.8, "color": GRAY},
    ], accent=CLAY, title_size=11.5)

# ============================================================ SLIDE 5 — ENTERPRISE GTM
def slide5():
    s = content_slide("05 · ENTERPRISE GO-TO-MARKET · 卖 Token → 卖组织提速", "面向 ToB 的布局与动作", 6,
                      note="图6: Anthropic 年化收入轨迹与 B 端市场地位 · 资料来源: The Information、Menlo Ventures、Ramp、国泰海通证券(2026-05-30) ·〔强〕2024末~$1B;余为报道/券商口径〔中〕")
    # left: revenue chart
    section_label(s, Inches(0.55), Inches(1.28), Inches(5.4), "年化收入(ARR)轨迹(估计, $B)")
    bar_chart(s, Inches(0.98), Inches(1.92), Inches(4.25), Inches(1.66), [
        ("2024 末", 1, "$1B", CLAY, "〔强〕"),
        ("2025 末", 9, "$9B", SLATE, "〔中〕"),
        ("2026 Q1", 14, "$14B", SLATE, "〔中〕"),
        ("2026 ARR", 30, "~$30B", CLAYD, "〔中·券商〕"),
    ], ymax=33)
    figcap(s, Inches(0.58), Inches(4.06), Inches(5.0), "图6", "Anthropic 年化收入轨迹(多源口径)",
           "The Information、国泰海通证券(2026-05-30)")
    callout_card(s, Inches(0.55), Inches(4.50), Inches(5.05), Inches(2.14), "B 端经济学:高价值任务驱动", [
        {"t": "约 25% 高价值场景创造约 80% 收入(NVIDIA GTC 场景分层)。", "size": 9.2},
        {"t": "单用户变现 ≈ OpenAI 的 8×;企业 API 份额 12%(2023)→40%(2025)居首。", "size": 9.2},
        {"t": "Ramp:美企 Anthropic 占(A+O)AI 支出 60–65%;新增采购 73%→Anthropic。", "size": 9.2},
        {"t": "年付费 >$1M 企业客户 500(2026-02)→1000(2026-04)。", "size": 9.2},
    ], accent=CLAY, title_size=11)
    # right: action stacks
    RX, RW = Inches(5.85), Inches(6.93)
    callout_card(s, RX, Inches(1.28), RW, Inches(1.58), "企业级动作 · Enterprise", [
        {"t": "Claude for Enterprise(2024-09):500K 上下文 · SSO/RBAC/SCIM/审计 · 不用客户数据训练。", "size": 9.2},
        {"t": "HIPAA-ready(BAA)· 医疗与生命科学(2025-10 / 2026-01)· 金融(2025-07)。", "size": 9.2},
        {"t": "Managed Agents 企业级托管 · 行业专精角色 Agent。", "size": 9.2},
    ], accent=CLAY, title_size=11.5)
    callout_card(s, RX, Inches(2.98), RW, Inches(1.46), "开发者 · 生态", [
        {"t": "Agent SDK(原 Claude Code SDK,2025-09 更名):内部框架 → 可编程外部能力。", "size": 9.2},
        {"t": "MCP 开放协议(2024-11;2025-03 OpenAI 采用;2025-12 入 Linux Foundation)。", "size": 9.2},
        {"t": "插件 / Agent Skills 生态:角色与连接器即插即用。", "size": 9.2},
    ], accent=SLATE, title_size=11.5)
    callout_card(s, RX, Inches(4.56), RW, Inches(1.20), "信任即准入 · Trust = Access", [
        {"t": "RSP / System Cards / Transparency Hub —— 安全可信作 ToB 差异化卖点,而非成本。", "size": 9.2},
        {"t": "分发多入口:终端/IDE/API/MCP/插件,单次改进多界面释放。", "size": 9.2, "color": GRAY},
    ], accent=CLAY, title_size=11.5)
    band = rrect(s, RX, Inches(5.88), RW, Inches(0.62), BG, radius=0.06)
    fill_text(band, [{"runs": [{"t": "模式迁移　", "size": 11, "color": CLAYL, "bold": True},
                               {"t": "卖 Token → 卖「组织提速」:价值随高价值任务的完成度计量。", "size": 11, "color": WHITE}]}],
              anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# ============================================================ SLIDE 6 — SAFETY & CONTROLLABILITY
def slide6():
    s = content_slide("06 · SAFETY & CONTROLLABILITY · 边界条件,而非刹车", "Agent 安全与可控性:治理即 Runtime", 7,
                      note="图7: 治理即 Runtime 与人在环责任链 · 资料来源: Anthropic RSP(2023-09 起)、System Cards、Transparency Hub(2026)· ASL-3 安全措施 2025-05 随 Claude Opus 4 启用")
    # top banner: new risk surface
    band = rrect(s, Inches(0.55), Inches(1.30), Inches(12.23), Inches(0.72), CLAYL, radius=0.06)
    rect(s, Inches(0.55), Inches(1.30), Inches(0.10), Inches(0.72), CLAY)
    fill_text(band, [{"runs": [{"t": "新风险面　", "size": 11.5, "color": CLAYD, "bold": True},
                               {"t": "自主度 / 托管度上升 → 权限、监督、可控性、agentic misalignment（代理性错位）", "size": 11.5, "color": INK}]}],
              anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    # left: governance = runtime
    LX, LW = Inches(0.55), Inches(6.02)
    callout_card(s, LX, Inches(2.28), LW, Inches(3.02), "治理即 Runtime", [
        {"runs": [{"t": "RSP(2023-09):", "size": 10, "color": INK, "bold": True},
                  {"t": "预声明能力阈值与发布标准;ASL 分级(ASL-1…4+),ASL-3 安全措施 2025-05 随 Claude Opus 4 启用。", "size": 10, "color": INK}]},
        {"runs": [{"t": "System Cards:", "size": 10, "color": INK, "bold": True},
                  {"t": "随模型发布公开能力与安全评测。", "size": 10, "color": INK}]},
        {"runs": [{"t": "Transparency Hub(2026):", "size": 10, "color": INK, "bold": True},
                  {"t": "集中模型卡 / 安全措施 / 自愿承诺,另有面向非专家的 Model Report。", "size": 10, "color": INK}]},
        {"runs": [{"t": "评测内嵌:", "size": 10, "color": INK, "bold": True},
                  {"t": "把「能不能发」前置成「按什么规则持续发」。", "size": 10, "color": CLAY, "bold": True}]},
    ], accent=CLAY, title_size=12.5)
    # right: control mechanisms
    RX, RW = Inches(6.76), Inches(6.02)
    section_label(s, RX, Inches(2.28), Inches(6), "可控机制 · 人在环责任链")
    arrow_progression(s, RX, Inches(2.70), RW, ["Delegate", "Verify", "Review", "Accountability"],
                      h=Inches(0.46), size=10.5, fill=BG_CHIP, fg=WHITE)
    callout_card(s, RX, Inches(3.42), RW, Inches(1.88), "落地机制", [
        "权限边界:ask / auto-accept / plan / auto 多档授权",
        "人在环验证:委派 → 验证 → 复核 → 可问责",
        "沙箱 / 审计 / checkpoints 回滚 / allowlist 预批",
        {"runs": [{"t": "▸ ", "size": 10, "color": SLATE, "bold": True},
                  {"t": "自主度可上移的前提 = 监督与验证同步增强", "size": 10, "color": INK, "bold": True}]},
    ], accent=SLATE, title_size=12)
    # bottom banner: evolution callback
    band2 = rrect(s, Inches(0.55), Inches(5.52), Inches(12.23), Inches(1.06), SLATEL, radius=0.06)
    rect(s, Inches(0.55), Inches(5.52), Inches(0.10), Inches(1.06), SLATE)
    fill_text(band2, [{"runs": [{"t": "演化呼应　", "size": 11.5, "color": SLATE, "bold": True},
                                {"t": "阶段一的「安全信仰」在阶段四技术化落地 —— 成为速度的边界条件,而非刹车。", "size": 11.5, "color": INK}]}],
              anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)

# ============================================================ SLIDE 7 — OUTLOOK & SIGNPOSTS
def slide7():
    s = content_slide("07 · OUTLOOK & SIGNPOSTS", "未来展望与信号灯", 8,
                      note="图8: 可跟踪信号灯 · 资料来源: Anthropic、Menlo Ventures/Ramp、The Information、TicketTrends、国泰海通证券(2026-05-30) · 颜色为方向性判断")
    LX, LW = Inches(0.55), Inches(6.02)
    RX, RW = Inches(6.76), Inches(6.02)
    callout_card(s, LX, Inches(1.28), LW, Inches(1.80), "技术拐点 · 趋势延伸", [
        {"t": "把「可验证 Agent 闭环」从编码外推到更广知识工作。", "size": 9.8},
        {"t": "自主度 / 托管度继续上移 —— 后台、长程、多 Agent。", "size": 9.8},
        {"t": "Dynamic Workflows:可达千级并行子 Agent(2026-05)。", "size": 9.8, "color": GRAY},
    ], accent=CLAY, title_size=12.5)
    callout_card(s, RX, Inches(1.28), RW, Inches(1.80), "风险边界", [
        "质量债:自主度↑ 而验证/监督不同步;技能退化、弱可验证领域瓶颈",
        {"runs": [{"t": "▸ ", "size": 9.8, "color": AMBER, "bold": True},
                  {"t": "算力约束:已现 Claude Code 限流/分层(2026-04);锁定 ~12.3GW vs OpenAI 激进扩张", "size": 9.8, "color": INK}]},
        {"runs": [{"t": "▸ ", "size": 9.8, "color": AMBER, "bold": True},
                  {"t": "竞品追赶:Codex 下载量 2026-04 反超 Claude Code(GPT-5.5 接入)", "size": 9.8, "color": INK}]},
    ], accent=AMBER, title_size=12.5)
    # signpost dashboard
    section_label(s, Inches(0.55), Inches(3.24), Inches(11), "信号灯(可跟踪) · Signposts")
    posts = [
        ("①", "连续动作 / 自主度", "9.8→21.2 已现", GREEN),
        ("②", "MCP 生态规模", "入 Linux Foundation", GREEN),
        ("③", "企业 API 份额", "40% 居首;Ramp 60–65%", GREEN),
        ("④", "Cowork / 企业 Agent", "2026-01 预览,待放量", AMBER),
        ("⑤", "算力锁定 & 竞品", "~12.3GW;Codex 反超", AMBER),
    ]
    pw = (Inches(12.23) - Inches(0.20) * 4) / 5
    px = Inches(0.55); py = Inches(3.60)
    for i, (num, title, status, col) in enumerate(posts):
        signpost(s, px + (pw + Inches(0.20)) * i, py, pw, Inches(1.30), num, title, status, col)
    figcap(s, Inches(0.55), Inches(4.98), Inches(9.5), "图8",
           "可跟踪信号灯(自主度 / 生态 / 份额 / 渗透 / 算力·竞品)",
           "Anthropic、Menlo/Ramp、The Information、TicketTrends、国泰海通(2026-05-30)")
    # conclusion banner
    band = rrect(s, Inches(0.55), Inches(5.46), Inches(12.23), Inches(1.16), CLAY, radius=0.06)
    fill_text(band, [
        {"runs": [{"t": "结论回扣　", "size": 13, "color": WHITE, "bold": True},
                  {"t": "领先本质 = 速度系统 + Agent 协同演化能力", "size": 13, "color": WHITE, "bold": True}], "space_after": 3},
        {"t": "把模型能力沉淀进可控、可验证的 Agent Runtime,并以「自举飞轮」持续放大 —— 而非单点模型分数领先。",
         "size": 11, "color": WHITE},
    ], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# ============================================================ COVER / 导读 (page 1)
def slide_cover():
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, WHITE)
    # cover banner
    rect(s, 0, 0, SW, Inches(1.92), BG)
    rect(s, 0, 0, Inches(0.22), Inches(1.92), CLAY)
    tb, tf = textbox(s, Inches(0.6), Inches(0.24), Inches(9.5), Inches(1.6))
    p = tf.paragraphs[0]; style(p, "ANTHROPIC 技术洞察  ·  深度研究 (Insight Report)", 11.5, CLAYL, bold=True, space_after=4)
    p = tf.add_paragraph(); style(p, "Agent 时代的“自举飞轮”", 32, WHITE, bold=True, space_after=3)
    p = tf.add_paragraph(); style(p, "战略定位 × Agent 能力的协同演化　|　From Coding Agent to Organizational OS", 12.5, LGRAY)
    # 核心判断 chip (评级-style)
    chip = rrect(s, Inches(10.4), Inches(0.46), Inches(2.35), Inches(1.02), CLAY, radius=0.10)
    fill_text(chip, [{"t": "核心判断 / Call", "size": 9.5, "color": CLAYL, "bold": True, "space_after": 2},
                     {"t": "结构性领先", "size": 15, "color": WHITE, "bold": True, "space_after": 1},
                     {"t": "边界 = 算力 × 可验证性", "size": 8.5, "color": WHITE}])
    # 核心论点 strip
    th = rect(s, 0, Inches(1.92), SW, Inches(0.62), CLAYL)
    rect(s, 0, Inches(1.92), Inches(0.22), Inches(0.62), CLAY)
    fill_text(th, [{"runs": [
        {"t": "核心论点　", "size": 12, "color": CLAYD, "bold": True},
        {"t": "定位四段迁移由 Agent 能力跃迁驱动并反向重塑;真正的领先是速度系统 + Agent 协同演化能力 —— 把模型沉淀进可控可验证的 Agent Runtime,并以「自举飞轮」持续放大。",
         "size": 11, "color": INK}]}], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    # 核心观点 (导读)
    section_label(s, Inches(0.55), Inches(2.74), Inches(8), "核心观点 · Key Theses(本报告导读)")
    theses = [
        ("定位四段跃迁,每段由 Agent 能力跃迁驱动 — ",
         "安全实验室 → 前沿模型/API → Agent 平台 → 组织操作系统(AI 造 AI);Agent 重心由 Chat 调用上移到动态编排 / 托管多 Agent。", ""),
        ("护城河 = 速度系统 + 自举飞轮,非单点跑分 — ",
         "把模型能力沉淀进可控可验证的 Agent Runtime;编码是高可信闭环,内部研究显示 Claude 写入 >80% 合并代码、人均每日 merged PR +67%、生产率 +50%。", "强"),
        ("价值锚点从「用户规模」转向「高价值任务」 — ",
         "约 25% 高价值场景创造约 80% 收入;Anthropic 单用户变现约为 OpenAI 的 8 倍,企业 API 份额 12%(2023)→40%(2025)居首。", "中·券商"),
        ("编码楔子 → 平台外扩,可验证性是瓶颈变量 — ",
         "Cowork 与角色插件把同一套 Model + Agent Loop + MCP 复制到通用知识工作;非编码领域可验证性弱,制约外扩速度。", ""),
        ("边界条件:安全 × 算力 — ",
         "RSP / System Cards / Transparency Hub 把治理做成 ToB 准入;但算力供给偏紧(已现限流/分层)与 Codex 追赶(2026-04 下载量反超)是关键风险变量。", "中"),
    ]
    paras = []
    for lead, body, tg in theses:
        runs = [{"t": "▍ ", "size": 12, "color": CLAY, "bold": True},
                {"t": lead, "size": 11.5, "color": INK, "bold": True},
                {"t": body, "size": 11.5, "color": INK}]
        if tg:
            runs.append({"t": "  〔" + tg + "〕", "size": 9, "color": GRAY, "italic": True})
        paras.append({"runs": runs, "space_after": 9})
    tb_text(s, Inches(0.55), Inches(3.06), Inches(12.23), Inches(3.42), paras, anchor=MSO_ANCHOR.TOP)
    # 方法与口径 note box
    nb = rrect(s, Inches(0.55), Inches(6.52), Inches(12.23), Inches(0.50), CARD, line=LGRAY, line_w=0.75, radius=0.05)
    fill_text(nb, [{"runs": [
        {"t": "方法与口径　", "size": 8.5, "color": CLAY, "bold": True},
        {"t": "数据经多源交叉核验并按〔强/中/向〕分层标注;Anthropic 内部指标引自《How AI Is Transforming Work at Anthropic》(2025-12);市场/算力/竞品数据引自国泰海通证券《OpenAI:AI 时代的基础设施与超级入口》(2026-05-30,综合 Menlo Ventures / Ramp / The Information)。2026 年项超模型知识截止,引用前对照官方原文。",
         "size": 7.8, "color": GRAY}]}], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT)
    footer(s, 1, "")

slide_cover(); slide1(); slide2(); slide3(); slide4(); slide5(); slide6(); slide7()

out = "/home/user/openai-chatgpt-codex/Anthropic_技术洞察_自举飞轮.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))

# ---------------- layout audit: flag shapes that overflow the canvas ----------------
def audit():
    from pptx import Presentation as _P
    d = _P(out)
    EMU_W, EMU_H = int(SW), int(SH)
    tol = 9525 * 2  # ~2px tolerance
    issues = 0
    for i, sl in enumerate(d.slides, 1):
        for sh in sl.shapes:
            try:
                L, T, W, H = sh.left, sh.top, sh.width, sh.height
            except Exception:
                continue
            if None in (L, T, W, H):
                continue
            if L < -tol or T < -tol or (L + W) > EMU_W + tol or (T + H) > EMU_H + tol:
                issues += 1
                print(f"  [overflow] slide {i}: {sh.shape_type} name={sh.name!r} "
                      f"L={L/914400:.2f} T={T/914400:.2f} R={(L+W)/914400:.2f} B={(T+H)/914400:.2f}")
    print(f"audit: {issues} overflow issue(s); canvas=13.333x7.5in")
audit()
