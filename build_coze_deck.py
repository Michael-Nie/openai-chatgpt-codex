# -*- coding: utf-8 -*-
"""Build the ByteDance Coze (扣子) deep-insight deck (PPTX).

Redesigned per the ppt-master skill: 标题给结论 · 每页换版式且必有视觉锚点 ·
70-25-5 配色 · 思源黑体(免费可商用) · 卡片/流程图/2x2矩阵/大字报数字 · 无标题装饰线。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

# ---------- palette (70-25-5, 低饱和) ----------
NAVY   = RGBColor(0x16, 0x22, 0x3A)   # 深色band/文字
BLUE   = RGBColor(0x2E, 0x6B, 0xE6)   # 扣子蓝 · 主色
BLUED  = RGBColor(0x16, 0x3E, 0x8C)   # 深蓝 · 强调
BLUEL  = RGBColor(0x9D, 0xBB, 0xF2)   # 浅蓝
INK    = RGBColor(0x1F, 0x29, 0x37)   # 正文
GRAYT  = RGBColor(0x6B, 0x72, 0x80)   # 次要/标签
LGRAY  = RGBColor(0xD8, 0xDE, 0xE9)
CARD   = RGBColor(0xF2, 0xF5, 0xFA)   # 卡片底
CARD2  = RGBColor(0xEA, 0xEF, 0xF7)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
TEAL   = RGBColor(0x12, 0x9A, 0x8F)   # 正向 accent
AMBER  = RGBColor(0xC2, 0x6B, 0x09)   # 警示 accent
RED    = RGBColor(0xB0, 0x47, 0x44)   # 威胁(低饱和)

FONT = "思源黑体"            # Source Han Sans (free commercial); Latin runs use FONT_LAT
FONT_LAT = "Arial"          # 渲染校验时曾用 "Noto Sans CJK SC"(即思源黑体)确认版式无误
CHARTS = "/home/user/openai-chatgpt-codex/coze-insight/charts"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
N = [0]  # page counter

def set_font(run, name=FONT):
    run.font.name = FONT_LAT
    rPr = run._r.get_or_add_rPr()
    for tag, face in (('a:latin', FONT_LAT), ('a:ea', name), ('a:cs', name)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set('typeface', face)

def _noshadow(sp):
    sp.shadow.inherit = False

def rect(s, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = Pt(1)
    _noshadow(sp); return sp

def rrect(s, x, y, w, h, color, line=None):
    return rect(s, x, y, w, h, color, line, shape=MSO_SHAPE.ROUNDED_RECTANGLE)

def textbox(s, x, y, w, h, anchor=None):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    if anchor: tf.vertical_anchor = anchor
    return tb, tf

def style(p, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT, space_after=4, italic=False):
    p.alignment = align; p.space_after = Pt(space_after); p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; set_font(r)
    return r

def shape_text(sp, lines, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER):
    tf = sp.text_frame; tf.word_wrap = True
    for m in ("margin_left","margin_right","margin_top","margin_bottom"):
        setattr(tf, m, Inches(0.04))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if isinstance(lines, str): lines = [(lines, size, bold)]
    first = True
    for ln in lines:
        txt, sz, bd = (ln if isinstance(ln, tuple) else (ln, size, bold))
        p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        style(p, txt, sz, color, bold=bd, align=align, space_after=1)
    return sp

# ---------- structural header (no decorative under-line) ----------
def header(s, kicker, title, note="可信度: 强=官方+多源 · 中=单一媒体/厂商自报 | 国内版 coze.cn · 数据截至 2026-06"):
    N[0] += 1
    rect(s, 0, 0, SW, SH, WHITE)
    rect(s, 0, 0, SW, Inches(1.12), NAVY)
    rect(s, Inches(0.55), Inches(0.30), Inches(0.10), Inches(0.52), BLUE)  # accent square (非装饰线)
    tb, tf = textbox(s, Inches(0.78), Inches(0.16), Inches(11.6), Inches(0.92))
    p = tf.paragraphs[0]; style(p, kicker, 11, BLUEL, bold=True, space_after=2)
    p2 = tf.add_paragraph(); style(p2, title, 22, WHITE, bold=True)
    # footer
    tbf, tff = textbox(s, Inches(0.55), Inches(7.06), Inches(11.4), Inches(0.36))
    style(tff.paragraphs[0], note, 8, GRAYT)
    tb2, tf2 = textbox(s, Inches(12.5), Inches(7.04), Inches(0.6), Inches(0.36))
    style(tf2.paragraphs[0], f"{N[0]:02d}", 10, GRAYT, align=PP_ALIGN.RIGHT)
    return s

def slide(kicker, title, note=None):
    s = prs.slides.add_slide(BLANK)
    header(s, kicker, title, note) if note else header(s, kicker, title)
    return s

def card(s, x, y, w, h, title, lines, accent=BLUE, fill=CARD, title_sz=13, body_sz=11, gap=3, side=True):
    rrect(s, x, y, w, h, fill)
    if side: rect(s, x, y, Inches(0.09), h, accent)
    tb, tf = textbox(s, x+Inches(0.22), y+Inches(0.12), w-Inches(0.36), h-Inches(0.2))
    if title:
        style(tf.paragraphs[0], title, title_sz, accent, bold=True, space_after=gap+1)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if (not title and i == 0) else tf.add_paragraph()
        if isinstance(ln, tuple):
            style(p, ln[0], ln[1] if len(ln)>1 else body_sz, ln[2] if len(ln)>2 else INK,
                  bold=ln[3] if len(ln)>3 else False, space_after=gap)
        else:
            style(p, ln, body_sz, INK, space_after=gap)

def bignum(s, x, y, num, label, color=BLUE, num_sz=40, w=Inches(3.2), label_sz=12):
    tb, tf = textbox(s, x, y, w, Inches(1.4))
    style(tf.paragraphs[0], num, num_sz, color, bold=True, space_after=2)
    style(tf.add_paragraph(), label, label_sz, GRAYT, space_after=0)

def chevrons(s, x, y, w, h, items, colors):
    n = len(items); cw = int(w / n); ov = Inches(0.16)
    for i, (it, col) in enumerate(zip(items, colors)):
        cx = x + cw*i - (ov if i > 0 else 0)
        ww = cw + (ov if i > 0 else 0)
        sp = s.shapes.add_shape(MSO_SHAPE.CHEVRON, cx, y, ww, h)
        sp.fill.solid(); sp.fill.fore_color.rgb = col; sp.line.fill.background(); _noshadow(sp)
        shape_text(sp, it, size=12.5, color=WHITE, bold=True)

def rating(n, total=5):
    return "●"*n + "○"*(total-n)

def table_slide(kicker, title, headers, rows, col_w, note=None, fs=11, hfs=11,
                y=Inches(1.45), row_h=Inches(0.42)):
    s = slide(kicker, title, note) if note else slide(kicker, title)
    nrows, ncols = len(rows)+1, len(headers)
    total_w = sum(col_w); x = Inches((13.333-total_w)/2)
    tbl = s.shapes.add_table(nrows, ncols, x, y, Inches(total_w), row_h*nrows).table
    tbl.first_row = True; tbl.horz_banding = True
    for j, cw in enumerate(col_w): tbl.columns[j].width = Inches(cw)
    for j, ht in enumerate(headers):
        c = tbl.cell(0, j); c.fill.solid(); c.fill.fore_color.rgb = NAVY
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Inches(0.07); c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = ht; r.font.size = Pt(hfs); r.font.bold = True
        r.font.color.rgb = WHITE; set_font(r)
    for i, row in enumerate(rows, 1):
        for j, val in enumerate(row):
            c = tbl.cell(i, j); c.fill.solid()
            c.fill.fore_color.rgb = WHITE if i % 2 else CARD
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = Inches(0.07); c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
            p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
            txt, col, bold = val, INK, (j == 0)
            if isinstance(val, tuple):
                txt = val[0]; col = val[1] if len(val) > 1 else INK
                bold = val[2] if len(val) > 2 else (j == 0)
            r = p.add_run(); r.text = txt; r.font.size = Pt(fs)
            r.font.bold = bold; r.font.color.rgb = col; set_font(r)
    return s

def img(s, name, x, y, w=None, h=None):
    kw = {}
    if w is not None: kw["width"] = w
    if h is not None: kw["height"] = h
    return s.shapes.add_picture(os.path.join(CHARTS, name), x, y, **kw)

# ============================================================ 1 · COVER (全图型/居中, 深色)
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, 0, Inches(0.24), SH, BLUE)
# subtle depth blocks
b = rect(s, Inches(9.3), Inches(-1.2), Inches(5.5), Inches(5.5), BLUED, shape=MSO_SHAPE.OVAL)
b2 = rect(s, Inches(10.8), Inches(3.6), Inches(4.2), Inches(4.2), RGBColor(0x1E,0x30,0x52), shape=MSO_SHAPE.OVAL)
tb, tf = textbox(s, Inches(0.95), Inches(1.7), Inches(9.2), Inches(3.0))
style(tf.paragraphs[0], "字节跳动 · 扣子 COZE", 16, BLUEL, bold=True, space_after=8)
style(tf.add_paragraph(), "深度洞察分析", 48, WHITE, bold=True, space_after=8)
style(tf.add_paragraph(), "一个平台，串起字节 AI 的应用层闭环", 18, LGRAY, space_after=2)
tb2, tf2 = textbox(s, Inches(0.95), Inches(5.45), Inches(11.4), Inches(1.7))
style(tf2.paragraphs[0], "综合全景 · 产品 / 技术 / 市场 / 商业 / 生态 / 竞争", 13, BLUEL, bold=True, space_after=8)
style(tf2.add_paragraph(), "聚焦国内版 coze.cn  ·  2026 年 6 月  ·  5 路并行检索 + 多源交叉验证 + 可信度分层标注", 11.5, LGRAY, space_after=5)
style(tf2.add_paragraph(), "说明：本环境对多数中文媒体/官网抓取受限，结论多基于权威页面搜索摘要交叉印证；GitHub 数据为 API 实测。引用前请以官网/官方发布稿复核。", 9, GRAYT)

# ============================================================ 2 · EXEC SUMMARY (2x3 卡片网格)
s = slide("EXECUTIVE SUMMARY", "六个判断：扣子是字节 AI 闭环的「应用层抓手」")
cw, ch, gx, gy = Inches(4.02), Inches(1.72), Inches(0.18), Inches(0.16)
x0, y0 = Inches(0.55), Inches(1.45)
cards2 = [
 ("① 定位在演进", [("零代码 Bot 工厂 → 通用 Agent → ", 11, INK), ("职场 AI + Vibe Coding，两年多次升级", 11, INK)], BLUE),
 ("② 规模领先", [("月活开发者 100万→300万；2.0 称", 11, INK), ("「上千万用户」，自称中国最大", 11, INK)], BLUE),
 ("③ 开源是关键落子", [("Studio+Loop·Apache2.0·约2.1万★", 11, INK), ("抢 Agent「事实标准」、反制 Dify", 11, INK)], BLUE),
 ("④ 三段式火箭变现", [("开源引流 → 企业订阅 → 火山算力/", 11, INK), ("抖音流量；模板分成平台抽 20%", 11, INK)], TEAL),
 ("⑤ 护城河四位一体", [("豆包模型 × 火山算力 × 抖音/飞书", 11, INK), ("流量 × 开源生态", 11, INK)], TEAL),
 ("⑥ 风险也清晰", [("工具型变现天花板、定位摇摆、", 11, AMBER, True), ("同质化、母公司利润压力", 11, AMBER, True)], AMBER),
]
for i, (t, ls, acc) in enumerate(cards2):
    r, c = divmod(i, 3)
    card(s, x0 + c*(cw+gx), y0 + r*(ch+gy), cw, ch, t, ls, accent=acc, title_sz=14, body_sz=11)
tb, tf = textbox(s, Inches(0.55), Inches(6.55), Inches(12.2), Inches(0.5))
style(tf.paragraphs[0], "一句话：扣子的价值不止于工具本身，而在把「模型—应用—分发—算力」串成可变现闭环。", 12.5, NAVY, bold=True)

# ============================================================ 3 · AGENDA (编号卡片目录)
s = slide("CONTENTS", "本报告地图：六大板块")
items3 = [
 ("01", "产品", "定位 · 演进 · 功能架构", BLUE),
 ("02", "技术", "模型底座 · 架构 · 开源", BLUE),
 ("03", "市场", "增长 · 份额 · 底座规模", TEAL),
 ("04", "商业", "定价 · 变现 · 战略", TEAL),
 ("05", "生态", "字节 AI 闭环协同", BLUED),
 ("06", "竞争", "格局 · SWOT · 风险趋势", BLUED),
]
cw, ch, gx, gy = Inches(4.02), Inches(2.45), Inches(0.18), Inches(0.2)
for i, (no, t, sub, acc) in enumerate(items3):
    r, c = divmod(i, 3)
    x = Inches(0.55)+c*(cw+gx); y = Inches(1.55)+r*(ch+gy)
    rrect(s, x, y, cw, ch, CARD)
    rect(s, x, y, cw, Inches(0.10), acc)
    tb, tf = textbox(s, x+Inches(0.28), y+Inches(0.34), cw-Inches(0.5), ch-Inches(0.5))
    style(tf.paragraphs[0], no, 40, acc, bold=True, space_after=4)
    style(tf.add_paragraph(), t, 20, NAVY, bold=True, space_after=3)
    style(tf.add_paragraph(), sub, 12, GRAYT)

# ============================================================ 4 · METHOD & CAVEATS (两张大澄清卡)
s = slide("METHODOLOGY", "先讲清两条最易踩的坑")
card(s, Inches(0.55), Inches(1.5), Inches(6.05), Inches(2.55), "坑① 豆包 ≠ 扣子",
     [("豆包 = C 端 AI 助手 App，月活以亿计；", 12.5, INK),
      ("扣子 = 面向开发者的 Agent 平台，规模以", 12.5, INK),
      ("开发者/智能体数计。", 12.5, INK),
      ("「5700万日活、3.4亿月活」均为豆包，", 12.5, RED, True),
      ("不可挂在扣子名下。", 12.5, RED, True)], accent=RED, title_sz=16)
card(s, Inches(6.75), Inches(1.5), Inches(6.05), Inches(2.55), "坑② 开发者口径在变",
     [("100万「活跃开发者」(2024.12)", 12.5, INK),
      ("→ 300万「月活开发者」(2025.12)", 12.5, INK),
      ("→「上千万用户/1000万开发场景」(2026.01)", 12.5, INK),
      ("三者计量口径不同，构成趋势但", 12.5, AMBER, True),
      ("非同一指标，勿连成单一曲线。", 12.5, AMBER, True)], accent=AMBER, title_sz=16)
card(s, Inches(0.55), Inches(4.35), Inches(12.25), Inches(1.95), "研究方法",
     [("• 5 路并行检索（产品 / 市场 / 商业 / 技术 / 竞争）→ 去重抓取 → 多源交叉核验 → 按可信度排序综合。", 12, INK),
      ("• 可信度分层：强（官方+多源） / 中（单一媒体或厂商自报、无第三方审计） / 弱（已剔除或降级），全程标注。", 12, INK),
      ("• 限制：coze.cn / 火山引擎 / 多数中文媒体抓取受限（403），定价等单源数字标「中」，以官网为准；GitHub 为 API 一手数据。", 12, GRAYT)],
     accent=BLUE, title_sz=14, side=True)

# ============================================================ 5 · POSITIONING (中心聚焦 + 3卡)
s = slide("PRODUCT · POSITIONING", "一句话：零门槛搭 Agent，一键全渠道分发")
rrect(s, Inches(0.55), Inches(1.5), Inches(12.25), Inches(1.35), CARD2)
tb, tf = textbox(s, Inches(0.9), Inches(1.62), Inches(11.6), Inches(1.1), anchor=MSO_ANCHOR.MIDDLE)
style(tf.paragraphs[0], "扣子 = 字节基于自研豆包大模型 + 火山引擎底座的「一站式 AI Agent 开发平台」", 17, NAVY, bold=True, space_after=3)
style(tf.add_paragraph(), "无论是否有编程基础，都能可视化搭建智能体 / AI 应用，并一键发布到多渠道。", 13, INK)
card(s, Inches(0.55), Inches(3.1), Inches(3.97), Inches(3.0), "双版本",
     [("海外 Coze(coze.com)", 12, INK), ("2023.11 上线", 11, GRAYT),
      ("国内 扣子(coze.cn)", 12, INK), ("2024.02 上线", 11, GRAYT),
      ("早期内置「云雀」，", 11, INK), ("现以豆包为主、开放多模型", 11, INK)], accent=BLUE)
card(s, Inches(4.68), Inches(3.1), Inches(3.97), Inches(3.0), "三类用户",
     [("① 个人创作者", 12, INK), ("  私域营销 / 自媒体", 11, GRAYT),
      ("② 中小开发者", 12, INK), ("  快速验证 / 工具搭建", 11, GRAYT),
      ("③ 企业客户", 12, INK), ("  专业版 / 私有化", 11, GRAYT)], accent=TEAL)
card(s, Inches(8.81), Inches(3.1), Inches(3.99), Inches(3.0), "闭环中的角色",
     [("模型(豆包)", 12, GRAYT), ("↓", 12, BLUE, True),
      ("应用构建(扣子) ← 本层", 12, NAVY, True), ("↓", 12, BLUE, True),
      ("分发(抖音/飞书) ↓", 12, GRAYT), ("算力(火山引擎)", 12, GRAYT)], accent=BLUED)

# ============================================================ 6 · TIMELINE (全宽图)
s = slide("PRODUCT · TIMELINE", "两年五级跳：从 Bot 工厂到职场 AI",
          note="日期/口径多源交叉 · 2025.12 与 2026.01 为相邻两次发布 | 数据截至 2026-06")
img(s, "timeline.png", Inches(0.35), Inches(1.7), w=Inches(12.6))
rrect(s, Inches(0.55), Inches(5.7), Inches(12.25), Inches(1.0), CARD)
tb, tf = textbox(s, Inches(0.8), Inches(5.8), Inches(11.8), Inches(0.85), anchor=MSO_ANCHOR.MIDDLE)
style(tf.paragraphs[0], "节奏特征：上线即高速迭代 —— 2024 完成商业化，2025 连推「通用 Agent」与「开源」两大动作，", 12.5, INK, space_after=3)
style(tf.add_paragraph(), "2025 年底起重心转向「职场生产力 + Vibe Coding」，并为 Agent 配置邮箱 / 云电脑 / 云手机（Agent World）。", 12.5, INK)

# ============================================================ 7 · CAPABILITIES (2x3 卡 + 渠道条)
s = slide("PRODUCT · CAPABILITIES", "六大模块拼出「搭得快、发得广」")
mods = [
 ("① 工作流", ["可视化拖拽编排", "LLM/知识库/插件/代码节点"], BLUE),
 ("② 插件", ["200+ 官方插件 + 自定义", "可「MCP 化」发布"], BLUE),
 ("③ 知识库 RAG", ["PDF/网页/飞书等向量化", "检索降幻觉"], BLUE),
 ("④ 记忆体系", ["变量 / 数据库 / 长期记忆", "个性化 + 权限控制"], TEAL),
 ("⑤ 多 Agent 编排", ["多智能体按角色协同", "触发器定时/事件触发"], TEAL),
 ("⑥ 多模态", ["图像流 + Seedance 视频", "语音合成，能看能听"], TEAL),
]
cw, ch, gx, gy = Inches(4.02), Inches(1.5), Inches(0.18), Inches(0.16)
for i, (t, ls, acc) in enumerate(mods):
    r, c = divmod(i, 3)
    card(s, Inches(0.55)+c*(cw+gx), Inches(1.45)+r*(ch+gy), cw, ch, t,
         [(ls[0], 11, INK), (ls[1], 11, GRAYT)], accent=acc, title_sz=13.5)
rect(s, Inches(0.55), Inches(4.85), Inches(12.25), Inches(0.04), LGRAY)
tb, tf = textbox(s, Inches(0.55), Inches(5.0), Inches(12.25), Inches(0.4))
style(tf.paragraphs[0], "一键多端分发（扣子核心优势）", 13, NAVY, bold=True)
chans = ["豆包","飞书","抖音","微信","企业微信","API","Chat SDK","网页/嵌入","扣子商店"]
x = Inches(0.55)
for ch_ in chans:
    w = Inches(0.32 + 0.16*len(ch_))
    p = rrect(s, x, Inches(5.55), w, Inches(0.46), CARD)
    shape_text(p, ch_, size=11.5, color=NAVY, bold=False)
    x = x + w + Inches(0.12)

# ============================================================ 8 · TECH STACK (分层架构图)
s = slide("TECH · ARCHITECTURE", "四层底座：算力 → 模型 → 框架 → 平台")
layers = [
 ("平台层", "扣子 Coze：工作流引擎 · 插件 · 知识库 · 多 Agent · MCP 全面支持", BLUE),
 ("框架层", "CloudWeGo Eino（Go）运行时 + Hertz · 前端 React/TS · 微服务+DDD · RAG: ES + Milvus/VikingDB", BLUED),
 ("模型层", "主力豆包 1.6（256K·区间定价，成本约 R1 的 1/3）· 开放多模型", RGBColor(0x24,0x55,0xB0)),
 ("算力层", "火山引擎：算力 / 方舟模型推理 / 向量库 —— 扣子专业版即由其承载", NAVY),
]
y = Inches(1.55)
for name, detail, col in layers:
    rrect(s, Inches(0.9), y, Inches(2.1), Inches(0.92), col)
    rrect(s, Inches(3.15), y, Inches(9.28), Inches(0.92), CARD)
    tb, tf = textbox(s, Inches(0.9), y, Inches(2.1), Inches(0.92), anchor=MSO_ANCHOR.MIDDLE)
    style(tf.paragraphs[0], name, 15, WHITE, bold=True, align=PP_ALIGN.CENTER)
    tb2, tf2 = textbox(s, Inches(3.35), y, Inches(8.9), Inches(0.92), anchor=MSO_ANCHOR.MIDDLE)
    style(tf2.paragraphs[0], detail, 12, INK)
    y = y + Inches(1.04)
# 多模型 pill row
tb, tf = textbox(s, Inches(0.9), Inches(5.78), Inches(6.0), Inches(0.34))
style(tf.paragraphs[0], "可接模型（SaaS + 开源版）", 12, NAVY, bold=True)
x = Inches(0.9)
for m in ["豆包","DeepSeek","通义千问","智谱GLM","Kimi","MiniMax","Claude/Gemini*"]:
    w = Inches(0.32 + 0.15*len(m))
    p = rrect(s, x, Inches(6.18), w, Inches(0.44), CARD2)
    shape_text(p, m, size=11, color=BLUED, bold=False)
    x = x + w + Inches(0.1)

# ============================================================ 9 · OPEN SOURCE (左图右文 + 大字)
s = slide("TECH · OPEN SOURCE", "开源：最具杠杆的一步 —— 抢 Agent「事实标准」",
          note="Stars/Forks 为 GitHub API 实测(2026-06-22) | 数据截至 2026-06")
img(s, "github_stars.png", Inches(0.5), Inches(1.55), w=Inches(6.4))
bignum(s, Inches(0.7), Inches(4.5), "21,013★", "Coze Studio · Apache 2.0 · 2025.7.26 开源", BLUE, 34, Inches(6.2))
card(s, Inches(7.2), Inches(1.55), Inches(5.6), Inches(1.42), "Coze Studio",
     [("开发平台开源版，脱胎于服务", 11.5, INK), ("「数万企业、数百万开发者」的商业版", 11.5, INK)], accent=BLUE)
card(s, Inches(7.2), Inches(3.12), Inches(5.6), Inches(1.42), "Coze Loop",
     [("AgentOps 全生命周期：", 11.5, INK), ("Prompt 开发 / 评测 / 可观测", 11.5, INK)], accent=TEAL)
card(s, Inches(7.2), Inches(4.69), Inches(5.6), Inches(1.6), "开源 vs 商业",
     [("核心搭建能力开源；", 11.5, INK),
      ("音色/语音、租户管理、企业弹性扩容", 11.5, AMBER, True),
      ("留作商业版变现点。", 11.5, AMBER, True)], accent=AMBER)

# ============================================================ 10 · MARKET GROWTH (大字报 + 图)
s = slide("MARKET · GROWTH", "中国最大的 Agent 开发平台",
          note="官方披露口径（Force 大会 / 扣子 2.0）· 口径随时间演变，详见图注 | 数据截至 2026-06")
img(s, "dev_growth.png", Inches(0.4), Inches(1.55), w=Inches(7.4))
bignum(s, Inches(8.2), Inches(1.7), "300万", "月活开发者（2025.12 Force）", BLUE, 50, Inches(4.6))
card(s, Inches(8.2), Inches(3.15), Inches(4.6), Inches(1.45), "2.0 口径跃升",
     [("2026.01 称服务「上千万用户 /", 11.5, INK), ("1000万真实开发场景」", 11.5, INK)], accent=BLUED)
card(s, Inches(8.2), Inches(4.72), Inches(4.6), Inches(1.55), "第三方定位",
     [("2025 多份选型榜列第一梯队/榜首；", 11.5, INK), ("量子位《年度 AI 100》入 Agent TOP3", 11.5, INK),
      ("（月活另有 458万 口径，待复核）", 10.5, GRAYT)], accent=TEAL)

# ============================================================ 11 · SHARE & FOUNDATION (两图 + 大字)
s = slide("MARKET · FOUNDATION", "底座复利：模型 × 算力的护城河",
          note="IDC 为 MaaS/大模型平台口径，非 Agent 平台细分口径 · 豆包 token 为发布会披露 | 数据截至 2026-06")
img(s, "maas_share.png", Inches(0.3), Inches(1.5), h=Inches(3.35))
img(s, "doubao_tokens.png", Inches(6.95), Inches(1.5), w=Inches(5.9))
bignum(s, Inches(0.65), Inches(5.05), "59.2%", "火山引擎大模型云上调用量份额（IDC，2024 为 46.4%）", BLUE, 34, Inches(6.0))
tb, tf = textbox(s, Inches(6.95), Inches(5.05), Inches(5.9), Inches(1.6))
style(tf.paragraphs[0], "豆包日均 token 两年增约 1000 倍（2024.05 的 1200亿 → 2026.03 的 120万亿），", 11.5, INK, space_after=3)
style(tf.add_paragraph(), "居中国第一、全球前三；IDC MaaS 口径下字节约 16%、列第三。", 11.5, INK, space_after=3)
style(tf.add_paragraph(), "注：尚无「Agent 平台」细分赛道的扣子精确市占，不宜标百分比。", 10.5, AMBER)

# ============================================================ 12 · PRICING (表 + 大字)
table_slide(
 "BUSINESS · PRICING", "四档订阅 + 资源点计费 + 模板分成",
 ["版本", "价格（人民币）", "资源 / 权益"],
 [
  ["个人免费版", "0", "约 500 资源点/天（每日重置）"],
  ["个人进阶版", ("首月 9.9 元（原价 39.9/月）", AMBER), "约 3 万积分/月"],
  ["团队版", ("按席位包年包月（单价未确证）", AMBER), "协同开发 · 跨空间迁移 · 方舟模型"],
  ["企业版", ("约 4980 元/月，年付 8.3 折", AMBER), "约 300 万资源点/月"],
  ["专业版(付费)", "按调用 Token 计费", "每日赠 500 资源点"],
 ],
 col_w=[2.3, 4.5, 5.0],
 note="定价随计费规则 2024–2026 反复调整，「中」者为第三方整理、未经官网确认，以官网为准 | 数据截至 2026-06",
 fs=12, hfs=12.5, row_h=Inches(0.6), y=Inches(1.5))
s = prs.slides[-1]
bignum(s, Inches(0.95), Inches(5.05), "抽成 20%", "模板中心：创作者定价售卖、平台抽 20%（创作者得 80%）", TEAL, 30, Inches(7.0))
tb, tf = textbox(s, Inches(8.2), Inches(5.05), Inches(4.6), Inches(1.5))
style(tf.paragraphs[0], "计费单位：资源点（积分）；模型费 = Token × 单价；", 11.5, INK, space_after=3)
style(tf.add_paragraph(), "扣减序：资源点 > 代金券 > 现金账户。", 11.5, INK, space_after=3)
style(tf.add_paragraph(), "智能体交互已改为仅按模型 Token 计费。", 11.5, INK)

# ============================================================ 13 · COMMERCIAL STRATEGY (火箭 chevron + 双平台)
s = slide("BUSINESS · STRATEGY", "三段式火箭 + 双平台分工")
tb, tf = textbox(s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(0.4))
style(tf.paragraphs[0], "三段式火箭（媒体/分析框架，非官方表述）", 14, NAVY, bold=True)
chevrons(s, Inches(0.55), Inches(1.9), Inches(12.25), Inches(0.95),
         [["① 开源引流", "社区版 Apache 2.0 抢开发者"],
          ["② 企业订阅", "SLA/租户/合规等高级功能付费"],
          ["③ 生态盈利", "火山算力 + 抖音流量分发"]],
         [BLUEL, BLUE, BLUED])
tb, tf = textbox(s, Inches(0.55), Inches(3.15), Inches(12.2), Inches(0.4))
style(tf.paragraphs[0], "双平台分工", 14, NAVY, bold=True)
card(s, Inches(0.55), Inches(3.6), Inches(6.05), Inches(2.0), "扣子 Coze（C 端 / 开发者）",
     [("低代码、易用、广度，含开源版；", 12, INK),
      ("快速验证与创作者生态。", 12, INK),
      ("企业落地：1 控制台 + N 预置 Agent + X 自建", 11.5, GRAYT)], accent=BLUE, title_sz=14)
card(s, Inches(6.75), Inches(3.6), Inches(6.05), Inches(2.0), "HiAgent（企业级 / 火山引擎）",
     [("私有化部署、数据不出内网，", 12, INK),
      ("满足金融/医疗/制造合规；", 12, INK),
      ("宣称 AI 应用开发周期缩短 95%+（厂商口径）", 11.5, GRAYT)], accent=BLUED, title_sz=14)
tb, tf = textbox(s, Inches(0.55), Inches(5.85), Inches(12.2), Inches(0.6))
style(tf.paragraphs[0], "线索：据极客公园，团队 2025 年初设定的商业化目标「半年内即达成」；但无公开营收金额，属定性表述。", 11.5, GRAYT, italic=True)

# ============================================================ 14 · ECOSYSTEM LOOP (闭环 chevron + 回流)
s = slide("ECOSYSTEM", "一条闭环：模型 → 应用 → 分发 → 算力")
chevrons(s, Inches(0.55), Inches(1.7), Inches(12.25), Inches(1.25),
         [["模型", "豆包 / 即梦", "训练与生成能力"],
          ["应用", "扣子 Coze", "让能力可被搭建"],
          ["分发", "抖音 / 飞书 / 豆包", "触达海量场景"],
          ["算力", "火山引擎 + HiAgent", "承载并变现"]],
         [BLUED, BLUE, TEAL, NAVY])
p = rect(s, Inches(0.55), Inches(3.25), Inches(12.25), Inches(0.7), CARD2, shape=MSO_SHAPE.LEFT_ARROW)
shape_text(p, "↺ 数据与收入回流，反哺模型与算力 —— 形成正循环", size=13, color=NAVY, bold=True)
cards14 = [
 ("豆包 App", "C 端入口，月活以亿计，为生态导流", BLUE),
 ("即梦 / Seedance", "图像/视频生成，供扣子工作流调用", TEAL),
 ("飞书 智能伙伴", "办公协同 Agent，B 端落地入口", BLUED),
 ("抖音/头条/番茄", "海量流量分发，扣子独特冷启动优势", NAVY),
]
cw, gx = Inches(2.97), Inches(0.13)
for i, (t, d, acc) in enumerate(cards14):
    card(s, Inches(0.55)+i*(cw+gx), Inches(4.25), cw, Inches(1.55), t, [(d, 11, INK)], accent=acc, title_sz=12.5)
tb, tf = textbox(s, Inches(0.55), Inches(6.0), Inches(12.2), Inches(0.5))
style(tf.paragraphs[0], "这条「模型—应用—分发—算力」闭环，是扣子相对独立 Agent 平台的最大结构性差异。", 12.5, BLUE, bold=True)

# ============================================================ 15 · COMPETITION (评分矩阵表)
table_slide(
 "COMPETITION", "易用与流量领先，护城河却最浅",
 ["平台 / 模型", "易用性", "企业级", "分发生态", "模型开放", "一句话定位"],
 [
  [("扣子 Coze · 豆包", BLUE, True), (rating(5), BLUE), (rating(3), BLUE), (rating(5), BLUE), (rating(4), BLUE),
   ("易用+流量强，开源；护城河偏浅", INK, False)],
  ["腾讯元器 · 混元", (rating(4), NAVY), (rating(2), NAVY), (rating(4), NAVY), (rating(2), NAVY),
   ("偏 C 端创作，打通微信生态")],
  ["腾讯云 ADP · 混元", (rating(3), NAVY), (rating(5), NAVY), (rating(3), NAVY), (rating(3), NAVY),
   ("3.0 偏 To B 严肃生产，强合规")],
  ["文心 AgentBuilder · 文心", (rating(4), NAVY), (rating(3), NAVY), (rating(4), NAVY), (rating(2), NAVY),
   ("核心=百度搜索/App 流量")],
  ["阿里云百炼 · 通义", (rating(3), NAVY), (rating(5), NAVY), (rating(3), NAVY), (rating(5), NAVY),
   ("模型+Agent 双核，企业级强")],
  ["Dify · 开源", (rating(3), NAVY), (rating(4), NAVY), (rating(2), NAVY), (rating(5), NAVY),
   ("工业化流水线，开源最直接对手")],
 ],
 col_w=[3.0, 1.5, 1.5, 1.6, 1.6, 3.6],
 note="●=强 ○=弱；评分为基于公开资料的相对定性判断，非量化基准 | 数据截至 2026-06",
 fs=11, hfs=11.5, row_h=Inches(0.62), y=Inches(1.5))

# ============================================================ 16 · SWOT (2x2 象限)
s = slide("SWOT", "SWOT：优势在生态，软肋在变现")
# faint axes
rect(s, Inches(6.66), Inches(1.5), Pt(1.5), Inches(4.85), LGRAY)
rect(s, Inches(0.55), Inches(3.92), Inches(12.25), Pt(1.5), LGRAY)
qw, qh = Inches(5.95), Inches(2.32)
card(s, Inches(0.55), Inches(1.5), qw, qh, "S 优势",
     [("• 豆包模型（120万亿 token/全球前三）", 11.5, INK), ("• 火山引擎调用份额 59.2%", 11.5, INK),
      ("• 抖音/飞书流量分发独一无二", 11.5, INK), ("• 开源生态（Studio 约 2.1万★）+ 易用", 11.5, INK)],
     accent=TEAL, title_sz=15)
card(s, Inches(6.83), Inches(1.5), qw, qh, "W 劣势",
     [("• 工具型变现天花板、低粘性", 11.5, INK), ("• 两年内定位多次摇摆", 11.5, INK),
      ("• 与竞品功能高度同质化", 11.5, INK), ("• 深绑字节生态 + 国内外版本割裂", 11.5, INK)],
     accent=AMBER, title_sz=15)
card(s, Inches(0.55), Inches(4.03), qw, qh, "O 机会",
     [("• MCP/A2A 标准化（已支持 MCP）", 11.5, INK), ("• 2026 被视为 Agent 规模化落地元年", 11.5, INK),
      ("• 职场 AI / Vibe Coding 新赛道", 11.5, INK), ("• 火山引擎主战场押注 Agent", 11.5, INK)],
     accent=BLUE, title_sz=15)
card(s, Inches(6.83), Inches(4.03), qw, qh, "T 威胁",
     [("• 母公司利润波动加大算力压力", 11.5, INK), ("• 腾讯 ADP3.0 / 阿里百炼 / Dify 夹击", 11.5, INK),
      ("• 模型进步「稀释工作流价值」", 11.5, INK), ("• 监管/合规（出海面临欧盟 AI 法案）", 11.5, INK)],
     accent=RED, title_sz=15)

# ============================================================ 17 · RISKS (4卡)
s = slide("RISKS", "四大风险，诚实披露")
risks = [
 ("变现天花板", ["工具型订阅逻辑挑战大、", "LTV 低、低频低粘性（观点）"], AMBER),
 ("定位摇摆", ["两年内 Bot→通用 Agent→职场 AI", "多次重定位，「走一步看一步」"], AMBER),
 ("高度同质化", ["与腾讯/百度/阿里/Dify 趋同，", "差异主要落在流量入口"], RED),
 ("母公司利润压力", ["报道称字节 2025 净利同比降超70%*", "（官方称会计口径，含期权成本）"], RED),
]
cw, gx = Inches(2.97), Inches(0.13)
for i, (t, ls, acc) in enumerate(risks):
    card(s, Inches(0.55)+i*(cw+gx), Inches(1.6), cw, Inches(2.4), t,
         [(ls[0], 11.5, INK), (ls[1], 11.5, GRAYT)], accent=acc, title_sz=14)
card(s, Inches(0.55), Inches(4.3), Inches(12.25), Inches(1.9), "「模型进步稀释工作流价值」悖论",
     [("• 模型越弱、手工编排的工作流越有意义；模型不断进步，工作流的价值被稀释 —— 这是所有低代码 Agent 平台的共同隐忧。", 12.5, INK),
      ("• 生态绑定是双刃：深绑字节系=冷启动优势，也意味着迁移成本与「被锁定」担忧。", 12.5, INK)],
     accent=AMBER, title_sz=14)

# ============================================================ 18 · TRENDS (趋势 → 扣子)
s = slide("INDUSTRY TRENDS", "踩中两大趋势：Agent 规模化 + MCP 标准化")
trends = [
 ("MCP 成事实标准", "2025.12 捐入 Linux 基金会，多巨头共治；扣子已全面支持", BLUE),
 ("多 Agent 进生产", "2026 多 Agent 进生产、A2A/AP2 协议标准化关键年", TEAL),
 ("规模化生产力元年", "2026 Agent 深入企业核心流程、超越 RPA", BLUED),
 ("火山押注 Agent", "豆包破 50万亿 token 后，火山主战场转向 Agent 落地", NAVY),
]
cw, gx = Inches(2.97), Inches(0.13)
for i, (t, d, acc) in enumerate(trends):
    card(s, Inches(0.55)+i*(cw+gx), Inches(1.6), cw, Inches(2.5), t, [(d, 11, INK)], accent=acc, title_sz=13.5)
rrect(s, Inches(0.55), Inches(4.4), Inches(12.25), Inches(1.85), CARD2)
tb, tf = textbox(s, Inches(0.9), Inches(4.6), Inches(11.6), Inches(1.5), anchor=MSO_ANCHOR.MIDDLE)
style(tf.paragraphs[0], "对扣子的含义", 14, NAVY, bold=True, space_after=4)
style(tf.add_paragraph(), "重心正从 C 端创作迁向「职场/企业生产力 + Vibe Coding」，恰好踩在「Agent 规模化元年 + MCP 标准化」两大趋势上 —— 这是 2.0 / 2.5 押注的主线。", 13, INK)

# ============================================================ 19 · JUDGMENTS & RECS
s = slide("KEY JUDGMENTS", "三条洞察 + 两类行动建议")
card(s, Inches(0.55), Inches(1.5), Inches(7.6), Inches(4.8), "核心洞察",
     [("① 扣子是字节 AI 闭环的「应用层抓手」", 13, NAVY, True),
      ("   价值在串起模型—算力—流量，脱离闭环会低估其战略权重。", 11.5, INK),
      ("", 6, INK),
      ("② 开源是攻守兼备、最具杠杆的一步", 13, NAVY, True),
      ("   对外抢「事实标准」压制 Dify，对内为企业版/云服务引流。", 11.5, INK),
      ("", 6, INK),
      ("③ 真正的战役在「长时程 + 企业生产力」", 13, NAVY, True),
      ("   竞争从「谁更易用」转向「谁能在企业核心流程稳定跑通多 Agent」。", 11.5, INK)],
     accent=BLUE, title_sz=15)
card(s, Inches(8.35), Inches(1.5), Inches(4.45), Inches(2.3), "建议 · 企业用户",
     [("轻量/创作 → 扣子 SaaS", 11.5, INK),
      ("敏感/合规 → HiAgent 或开源自托管", 11.5, INK),
      ("要多模型自由 → 评估 Dify 并行", 11.5, INK)], accent=TEAL, title_sz=14)
card(s, Inches(8.35), Inches(4.0), Inches(4.45), Inches(2.3), "建议 · 观察/投资",
     [("盯三个信号：", 11.5, NAVY, True),
      ("① 企业版付费转化与营收披露", 11.5, INK),
      ("② 开源社区活跃度（Star/贡献/自托管）", 11.5, INK),
      ("③ 抖音电商等字节场景深度变现案例", 11.5, INK)], accent=BLUED, title_sz=14)

# ============================================================ 20 · SOURCES
s = slide("SOURCES & CONFIDENCE", "数据来源与可信度（节选）", note="完整来源见随附研究记录")
card(s, Inches(0.55), Inches(1.5), Inches(6.05), Inches(2.35), "官方 / 一手（强）",
     [("• github.com/coze-dev（Star/版本 API 实测）", 11, INK),
      ("• coze.cn 官方文档 · volcengine 火山引擎文档", 11, INK),
      ("• cloudwego.io（Eino）", 11, INK)], accent=BLUE)
card(s, Inches(6.75), Inches(1.5), Inches(6.05), Inches(2.35), "发布会 / 厂商披露（中-强）",
     [("• Force 大会 2024.12 / 2025.06 / 2025.12", 11, INK),
      ("• 扣子 2.0 发布稿（2026.01）", 11, INK),
      ("• 豆包 1.6 发布（2025.06）", 11, INK)], accent=TEAL)
card(s, Inches(0.55), Inches(4.05), Inches(6.05), Inches(2.25), "权威媒体 / 智库（中）",
     [("• 极客公园《扣子两年生长真相》", 11, INK),
      ("• 量子位（开源拆箱 / 年度 AI 100）", 11, INK),
      ("• 36氪 · InfoQ · 钛媒体 · IDC（份额）", 11, INK)], accent=BLUED)
card(s, Inches(6.75), Inches(4.05), Inches(6.05), Inches(2.25), "可信度处理",
     [("• GitHub 数字为 API 直读（强）", 11, INK),
      ("• 定价/458万月活/首日50万 等单源标「中」", 11, AMBER, True),
      ("• 务必区分「豆包 App」与「扣子平台」数据", 11, RED, True)], accent=AMBER)

# ============================================================ 21 · CLOSING (呼应封面)
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, 0, Inches(0.24), SH, BLUE)
rect(s, Inches(9.3), Inches(3.0), Inches(5.5), Inches(5.5), BLUED, shape=MSO_SHAPE.OVAL)
tb, tf = textbox(s, Inches(0.95), Inches(2.4), Inches(10.5), Inches(2.8))
style(tf.paragraphs[0], "扣子 COZE · 深度洞察", 15, BLUEL, bold=True, space_after=10)
style(tf.add_paragraph(), "一个平台，一条字节 AI 闭环", 36, WHITE, bold=True, space_after=12)
style(tf.add_paragraph(), "模型(豆包) → 应用(扣子) → 分发(抖音/飞书) → 算力(火山引擎)", 15, LGRAY, space_after=6)
style(tf.add_paragraph(), "综合全景洞察 · 2026 年 6 月 · 多源交叉验证 · 引用前请以官网/官方发布稿复核", 11, GRAYT)

out = "/home/user/openai-chatgpt-codex/字节扣子Coze深度洞察分析.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
