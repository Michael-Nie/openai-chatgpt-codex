# -*- coding: utf-8 -*-
"""Generate a technical insight analysis PPT for the GenericAgent project."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------- Theme ----------------
NAVY   = RGBColor(0x0B, 0x2A, 0x42)   # deep navy
BLUE   = RGBColor(0x12, 0x5E, 0x8C)   # primary blue
TEAL   = RGBColor(0x00, 0xB3, 0xA4)   # teal accent
ORANGE = RGBColor(0xFF, 0x7A, 0x33)   # orange accent
LIGHT  = RGBColor(0xF3, 0xF5, 0xF8)   # light panel
GREY   = RGBColor(0x5A, 0x66, 0x73)   # muted text
DARK   = RGBColor(0x1A, 0x24, 0x33)   # dark text
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
CARDBG = RGBColor(0xED, 0xF1, 0xF6)

FONT = "Microsoft YaHei"   # CJK-friendly

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def set_font(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def add_rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             space_after=6, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (text,size,color,bold) tuples."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(2)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for (text, size, color, bold) in para:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.bold = bold
            set_font(r)
    return tb


def header(slide, kicker, title, idx):
    add_rect(slide, 0, 0, SW, Inches(1.15), NAVY)
    add_rect(slide, 0, Inches(1.15), SW, Pt(4), TEAL)
    add_rect(slide, Inches(0.5), Inches(0.30), Pt(6), Inches(0.62), ORANGE)
    add_text(slide, Inches(0.72), Inches(0.18), Inches(10.5), Inches(0.9),
             [[(kicker, 11, TEAL, True)], [(title, 25, WHITE, True)]], space_after=2)
    add_text(slide, Inches(12.2), Inches(0.40), Inches(0.9), Inches(0.5),
             [[(f"{idx:02d}", 20, RGBColor(0x3A,0x5A,0x72), True)]], align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, head, body_lines, accent=TEAL, head_color=None):
    add_rect(slide, x, y, w, h, CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(slide, x, y, Pt(5), h, accent)
    runs = [[(head, 14, head_color or NAVY, True)]]
    for ln in body_lines:
        runs.append([("• ", 11, accent, True), (ln, 11, DARK, False)])
    add_text(slide, x+Inches(0.20), y+Inches(0.12), w-Inches(0.32), h-Inches(0.2),
             runs, space_after=4, line_spacing=1.05)


def footer(slide):
    add_text(slide, Inches(0.5), Inches(7.05), Inches(8), Inches(0.4),
             [[("GenericAgent · 技术洞察分析", 9, GREY, False)]])
    add_text(slide, Inches(11.0), Inches(7.05), Inches(1.9), Inches(0.4),
             [[("lsdefine/genericagent", 9, GREY, False)]], align=PP_ALIGN.RIGHT)


# ================= Slide 1: Cover =================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, 0, SW, SH, NAVY)
# decorative bands
add_rect(s, 0, Inches(5.2), SW, Pt(3), TEAL)
add_rect(s, Inches(0.9), Inches(2.0), Inches(0.18), Inches(2.5), ORANGE)
add_text(s, Inches(1.25), Inches(1.55), Inches(11), Inches(0.6),
         [[("自进化自治智能体框架 · 深度技术洞察", 15, TEAL, True)]])
add_text(s, Inches(1.25), Inches(2.15), Inches(11.2), Inches(1.6),
         [[("GenericAgent", 58, WHITE, True)],
          [("极简自进化 Agent 框架技术分析", 30, RGBColor(0xC9,0xD6,0xE0), True)]],
         space_after=8)
add_text(s, Inches(1.25), Inches(4.7), Inches(11), Inches(0.6),
         [[("约 3,000 行种子代码　·　100 行 Agent 循环　·　9 个原子工具　·　5 层记忆", 14, RGBColor(0xA9,0xBC,0xCC), False)]])
add_text(s, Inches(1.25), Inches(5.45), Inches(11), Inches(1.2),
         [[("12.8k", 26, TEAL, True), ("  Stars      ", 13, RGBColor(0x9A,0xAE,0xBF), False),
           ("1.5k", 26, TEAL, True), ("  Forks      ", 13, RGBColor(0x9A,0xAE,0xBF), False),
           ("MIT", 26, ORANGE, True), ("  License", 13, RGBColor(0x9A,0xAE,0xBF), False)]])
add_text(s, Inches(1.25), Inches(6.6), Inches(11), Inches(0.5),
         [[("分析日期 2026-06　|　Python 88.6% · JavaScript 6.1% · PowerShell/Rust", 11, RGBColor(0x7C,0x90,0xA3), False)]])

# ================= Slide 2: 项目概览 =================
s = prs.slides.add_slide(BLANK)
header(s, "PROJECT OVERVIEW", "项目概览：它到底是什么", 2)
add_text(s, Inches(0.72), Inches(1.45), Inches(12), Inches(0.95),
         [[("GenericAgent 是一个", 14, DARK, False),
           ("赋予任意大模型「系统级电脑控制能力」", 14, BLUE, True),
           ("的极简自治智能体框架。其核心主张是 ", 14, DARK, False),
           ("「不预装技能，而是进化技能」", 14, ORANGE, True),
           ("——每完成一个任务即结晶为可复用 Skill，逐步生长出个性化技能树。", 14, DARK, False)]],
         line_spacing=1.15)
cards = [
    ("能做什么", ["浏览器自动化 / 真实 Chrome 会话", "终端与命令行操作", "文件读写与补丁", "键鼠输入 + 屏幕视觉", "经 ADB 控制移动设备"], TEAL),
    ("怎么做到", ["仅 9 个原子工具作为底座", "约 100 行的 Agent 主循环", "五层结晶式记忆系统", "任务自动结晶为 SOP/Skill", "约 3K 行种子代码生长能力"], BLUE),
    ("为何不同", ["~30K 上下文 vs 对手 200K–1M", "无需 Playwright / LangChain", "跨模型：Claude/Gemini/Kimi…", "登录态/Cookie/指纹全保留", "仓库自身由 Agent 自举完成"], ORANGE),
]
cx = Inches(0.72)
for head, lines, ac in cards:
    card(s, cx, Inches(2.55), Inches(3.95), Inches(3.9), head, lines, ac)
    cx += Inches(4.10)
footer(s)

# ================= Slide 3: 核心理念 =================
s = prs.slides.add_slide(BLANK)
header(s, "DESIGN PHILOSOPHY", "核心设计理念：进化而非预装", 3)
add_rect(s, Inches(0.72), Inches(1.5), Inches(11.9), Inches(1.35), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(1.0), Inches(1.62), Inches(11.4), Inches(1.1),
         [[("“Don't preload skills — evolve them.”", 22, NAVY, True)],
          [("以最小信息密度的代码 + 直接系统控制 + 持续经验结晶，让能力从一颗「种子」有机生长，而非靠插件与依赖堆叠。", 13, GREY, False)]],
         space_after=6)
twin = [
    ("传统重型 Agent 思路", [
        "预置大量插件 / 工具 / SDK", "依赖 Playwright、LangChain 等框架",
        "上下文塞满工具说明与冗余记忆", "能力固定，跨会话不沉淀", "代码体量数十万行（如 530K）"], ORANGE),
    ("GenericAgent 思路", [
        "9 个原子工具自由组合", "用 code_run 现写脚本替代专用工具",
        "30K 上下文，降噪降幻觉降成本", "任务结晶为 Skill，越用越强", "3K 行种子代码自我生长"], TEAL),
]
cx = Inches(0.72)
for head, lines, ac in twin:
    card(s, cx, Inches(3.1), Inches(5.85), Inches(3.4), head, lines, ac)
    cx += Inches(6.05)
footer(s)

# ================= Slide 4: 整体架构 =================
s = prs.slides.add_slide(BLANK)
header(s, "ARCHITECTURE", "整体架构：极简三层心智模型", 4)
layers = [
    ("感知层  Perception", "web_scan 网页感知 · 屏幕视觉 · 文件读取 · 终端输出回收", TEAL),
    ("决策循环  Agent Loop (~100 行)", "LLM 调用 → 解析工具调用 → 分发 do_{tool} → 回收结果 → 拼接下一轮 prompt", BLUE),
    ("行动层  9 Atomic Tools", "code_run · file_read/write/patch · web_scan · web_execute_js · ask_user · 记忆更新", NAVY),
    ("记忆层  5-Layer Memory", "L0 元规则 · L1 洞察索引 · L2 全局事实 · L3 任务 SOP · L4 会话归档", ORANGE),
]
y = Inches(1.55)
for i, (head, body, ac) in enumerate(layers):
    add_rect(s, Inches(0.72), y, Inches(11.9), Inches(1.05), CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, Inches(0.72), y, Inches(2.9), Inches(1.05), ac, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(0.9), y, Inches(2.6), Inches(1.05),
             [[(head, 14, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(3.85), y, Inches(8.6), Inches(1.05),
             [[(body, 13, DARK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.18)
add_text(s, Inches(0.72), Inches(6.85), Inches(12), Inches(0.4),
         [[("设计要点：生成器流式输出（yield from）+ 生命周期 Hook + 工具可经 should_exit 提前退出，循环紧凑且可插拔。", 11, GREY, False)]])

# ================= Slide 5: 九大原子工具 =================
s = prs.slides.add_slide(BLANK)
header(s, "ATOMIC TOOLS", "九大原子工具：最小却完备的底座", 5)
tools = [
    ("code_run", "执行任意 Python / PowerShell", TEAL),
    ("file_read", "读取文件内容", BLUE),
    ("file_write", "创建或覆盖文件", BLUE),
    ("file_patch", "对已有文件打补丁", BLUE),
    ("web_scan", "感知与解析网页内容", TEAL),
    ("web_execute_js", "用 JS 控制真实浏览器", TEAL),
    ("ask_user", "人在环路确认", ORANGE),
    ("update_working_checkpoint", "短期工作记忆", ORANGE),
    ("start_long_term_update", "蒸馏持久记忆", ORANGE),
]
positions = [(0,0),(1,0),(2,0),(0,1),(1,1),(2,1),(0,2),(1,2),(2,2)]
cw, ch = Inches(3.95), Inches(1.45)
x0, y0 = Inches(0.72), Inches(1.6)
for (col, row), (name, desc, ac) in zip(positions, tools):
    x = x0 + col*Inches(4.10)
    y = y0 + row*Inches(1.62)
    add_rect(s, x, y, cw, ch, CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, y, Pt(5), ch, ac)
    add_text(s, x+Inches(0.22), y+Inches(0.16), cw-Inches(0.3), ch-Inches(0.2),
             [[(name+"  ", 14, NAVY, True)], [(desc, 11, GREY, False)]], space_after=3)
add_text(s, Inches(0.72), Inches(6.85), Inches(12), Inches(0.4),
         [[("洞察：以 code_run「现写脚本」取代成百上千的专用工具，是其工具数极少却能力极广的关键。", 11, GREY, False)]])

# ================= Slide 6: 五层记忆 =================
s = prs.slides.add_slide(BLANK)
header(s, "MEMORY SYSTEM", "五层结晶式记忆：越用越聪明", 6)
mem = [
    ("L0", "元规则 Meta Rules", "核心行为约束与系统级红线，最稳定、最高优先级", NAVY),
    ("L1", "洞察索引 Insight Index", "快速路由与召回入口，决定「该调哪条经验」", TEAL),
    ("L2", "全局事实 Global Facts", "跨任务沉淀的稳定知识与环境事实", BLUE),
    ("L3", "任务技能 Skills / SOP", "可复用工作流，相似任务直接一行调用", ORANGE),
    ("L4", "会话归档 Session Archive", "蒸馏后的长程记录，支撑长时序回忆", GREY),
]
y = Inches(1.6)
for tag, name, desc, ac in mem:
    add_rect(s, Inches(0.72), y, Inches(1.0), Inches(0.92), ac, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(0.72), y, Inches(1.0), Inches(0.92),
             [[(tag, 22, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(1.85), y, Inches(10.75), Inches(0.92), CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(2.1), y+Inches(0.1), Inches(10.3), Inches(0.75),
             [[(name, 14, NAVY, True), ("　—　", 12, GREY, False), (desc, 12, DARK, False)]],
             anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.04)
add_text(s, Inches(0.72), Inches(6.9), Inches(12), Inches(0.4),
         [[("对比：分层「浓缩式」记忆优于冗余堆叠或纯向量检索——降噪的同时保住可复用经验。", 11, GREY, False)]])

# ================= Slide 7: Agent 循环 =================
s = prs.slides.add_slide(BLANK)
header(s, "EXECUTION LOOP", "Agent 执行循环：100 行的引擎", 7)
steps = [
    ("①", "调用 LLM", "携带消息历史 + 工具 schema"),
    ("②", "解析工具调用", "从响应中提取 tool calls"),
    ("③", "分发执行", "BaseHandler.dispatch → do_{tool}"),
    ("④", "回收结果", "生成器 yield，exhaust 排空"),
    ("⑤", "拼接 prompt", "合并各工具 next_prompt"),
    ("⑥", "进入下一轮", "追加 user 消息，循环至完成"),
]
cw = Inches(3.85)
for i, (num, t, d) in enumerate(steps):
    col, row = i % 3, i // 3
    x = Inches(0.72) + col*Inches(4.05)
    y = Inches(1.7) + row*Inches(1.9)
    add_rect(s, x, y, cw, Inches(1.6), CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, x+Inches(0.22), y+Inches(0.12), Inches(1.0), Inches(0.8),
             [[(num, 30, TEAL, True)]])
    add_text(s, x+Inches(0.22), y+Inches(0.78), cw-Inches(0.4), Inches(0.7),
             [[(t, 15, NAVY, True)], [(d, 11, GREY, False)]], space_after=2)
add_text(s, Inches(0.72), Inches(6.0), Inches(11.9), Inches(0.9),
         [[("关键机制：", 12, ORANGE, True),
           ("生成器流式输出 · 生命周期 Hook（agent_before/turn_after）· should_exit 提前退出 · 正则后处理压缩冗长代码块。", 12, DARK, False)]],
         line_spacing=1.1)
footer(s)

# ================= Slide 8: 差异化亮点 =================
s = prs.slides.add_slide(BLANK)
header(s, "KEY DIFFERENTIATORS", "四大差异化技术亮点", 8)
hi = [
    ("Token 效率", ["~30K 上下文窗口运行", "对手普遍 200K–1M", "降噪声 / 降幻觉 / 降成本", "评测中约 6× 更低 token"], TEAL),
    ("真实浏览器", ["注入真实 Chrome 会话", "保留登录态/Cookie/指纹", "过 SannySoft / FingerprintJS", "reCAPTCHA v3 得分约 0.9"], BLUE),
    ("自进化", ["每个任务结晶为 Skill", "相似任务一行直接调用", "无需人工干预积累能力", "形成个性化技能树"], ORANGE),
    ("跨平台跨模型", ["Claude/Gemini/Kimi/MiniMax", "Win / macOS / Linux", "无需 Playwright/LangChain", "依赖极少，部署轻量"], NAVY),
]
cx = Inches(0.72)
for head, lines, ac in hi:
    card(s, cx, Inches(1.6), Inches(2.92), Inches(4.7), head, lines, ac)
    cx += Inches(3.02)
footer(s)

# ================= Slide 9: TMWebDriver =================
s = prs.slides.add_slide(BLANK)
header(s, "TMWebDriver", "真实浏览器方案：反检测的护城河", 9)
add_text(s, Inches(0.72), Inches(1.45), Inches(11.9), Inches(0.9),
         [[("TMWebDriver = 本地 WebSocket 服务 + Chrome 扩展，", 14, DARK, False),
           ("直接注入用户真实浏览器", 14, BLUE, True),
           ("，而非 headless 沙箱。由此天然继承真人浏览环境，绕过主流 Bot 检测。", 14, DARK, False)]],
         line_spacing=1.15)
card(s, Inches(0.72), Inches(2.6), Inches(5.85), Inches(3.7), "技术构成", [
    "本地 WebSocket Server 作桥接", "Chrome / Chromium 扩展注入页面",
    "web_scan 感知 + web_execute_js 操作", "无需下载浏览器二进制 / 驱动",
    "登录会话、Cookie、扩展、指纹全保留"], BLUE)
card(s, Inches(6.77), Inches(2.6), Inches(5.85), Inches(3.7), "反检测实测", [
    "通过 SannySoft 检测", "通过 incolumitas 验证",
    "reCAPTCHA v3 ≈ 0.9 人类可信度", "通过 FingerprintJS 分析",
    "对开放网络更具韧性"], TEAL)
footer(s)

# ================= Slide 10: 自进化机制 =================
s = prs.slides.add_slide(BLANK)
header(s, "SELF-EVOLUTION", "自进化机制：从首次摸索到一键复用", 10)
card(s, Inches(0.72), Inches(1.6), Inches(5.85), Inches(3.0), "首次执行（探索）", [
    "安装依赖、现写脚本", "调试、报错、修复",
    "验证功能正确性", "将成功路径蒸馏存为 Skill/SOP"], ORANGE)
card(s, Inches(6.77), Inches(1.6), Inches(5.85), Inches(3.0), "再次执行（复用）", [
    "命中 L1 洞察索引", "直接一行调用已结晶 Skill",
    "跳过试错，稳定高效", "跨任务收敛，能力持续累积"], TEAL)
add_rect(s, Inches(0.72), Inches(4.85), Inches(11.9), Inches(1.55), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(1.0), Inches(4.98), Inches(11.4), Inches(1.35),
         [[("真实结晶案例", 13, NAVY, True)],
          [("读取加密微信数据库 · 生成晨报摘要 · 量化选股筛选 · 配置 OAuth 流程 · 定时任务调度", 12, DARK, False)],
          [("标志性成就：本仓库从安装 Git 到每条 commit message，均由 GenericAgent 自主完成（自举）。", 12, ORANGE, True)]],
         space_after=5, line_spacing=1.1)
footer(s)

# ================= Slide 11: 性能评测 =================
s = prs.slides.add_slide(BLANK)
header(s, "EVALUATION", "性能评测：五维度优势", 11)
dims = [
    ("任务完成 & Token 效率", "完成困难任务的同时 token 消耗显著更低（约 6×）", TEAL),
    ("工具使用效率", "极简原子工具集胜过专用工具堆叠", BLUE),
    ("记忆系统有效性", "分层浓缩记忆 > 冗余堆叠 / 纯向量检索", ORANGE),
    ("自进化能力", "经验蒸馏为可复用 SOP，无需人工干预", TEAL),
    ("网页浏览能力", "极简设计下仍具开放网络韧性", BLUE),
]
y = Inches(1.6)
for i, (t, d, ac) in enumerate(dims):
    add_rect(s, Inches(0.72), y, Inches(7.4), Inches(0.86), CARDBG, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, Inches(0.72), y, Pt(6), Inches(0.86), ac)
    add_text(s, Inches(0.95), y, Inches(7.1), Inches(0.86),
             [[(t, 13, NAVY, True), ("　—　", 11, GREY, False), (d, 11, DARK, False)]],
             anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.98)
# right comparison panel
add_rect(s, Inches(8.35), Inches(1.6), Inches(4.27), Inches(4.78), NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(8.6), Inches(1.78), Inches(3.8), Inches(4.5),
         [[("对比标杆", 15, TEAL, True)],
          [("代码体量", 12, RGBColor(0x9A,0xAE,0xBF), False)],
          [("3K", 30, WHITE, True), ("  vs  ", 12, GREY, False), ("530K 行", 16, ORANGE, True)],
          [("上下文窗口", 12, RGBColor(0x9A,0xAE,0xBF), False)],
          [("~30K", 30, WHITE, True), ("  vs  ", 12, GREY, False), ("200K–1M", 16, ORANGE, True)],
          [("Token 消耗", 12, RGBColor(0x9A,0xAE,0xBF), False)],
          [("约 6× 更低", 22, TEAL, True)],
          [("对比对象：Claude Code · OpenClaw", 10, RGBColor(0x7C,0x90,0xA3), False)]],
         space_after=6, line_spacing=1.0)
footer(s)

# ================= Slide 12: 技术栈 & 部署 =================
s = prs.slides.add_slide(BLANK)
header(s, "STACK & DEPLOY", "技术栈与部署形态", 12)
card(s, Inches(0.72), Inches(1.6), Inches(3.85), Inches(4.6), "核心技术栈", [
    "Python 3.11–3.12", "requests · beautifulsoup4",
    "bottle · aiohttp", "simple-websocket-server",
    "可选 Streamlit / prompt_toolkit", "无 Playwright / LangChain"], TEAL)
card(s, Inches(4.77), Inches(1.6), Inches(3.85), Inches(4.6), "前端与接入", [
    "终端 UI（prompt_toolkit / rich）", "Streamlit Web 界面",
    "桌面应用 v0.1.0", "IM Bot：Telegram / Discord",
    "微信 / QQ / 飞书 / 企微 / 钉钉"], BLUE)
card(s, Inches(8.82), Inches(1.6), Inches(3.80), Inches(4.6), "部署方式", [
    "git clone + venv + API Key", "Windows PowerShell 一键装",
    "Linux/macOS 一键脚本", "隔离 Python 环境内置运行时",
    "依赖极少，资源开销低"], ORANGE)
footer(s)

# ================= Slide 13: 优势与风险 =================
s = prs.slides.add_slide(BLANK)
header(s, "INSIGHT · SWOT", "技术洞察：优势与风险并存", 13)
quad = [
    ("优势 Strengths", ["极简架构、信息密度高", "Token/成本优势显著", "自进化沉淀复用经验", "反检测真实浏览器护城河"], TEAL, Inches(0.72), Inches(1.6)),
    ("机会 Opportunities", ["个人 RPA / 自动化助手", "跨模型、低成本可商用", "技能树生态可沉淀分享", "端侧/移动场景延展"], BLUE, Inches(6.77), Inches(1.6)),
    ("风险 Weaknesses", ["code_run 任意执行=高权限风险", "强依赖模型推理质量", "记忆漂移/污染需治理", "自进化结果可复现性存疑"], ORANGE, Inches(0.72), Inches(4.05)),
    ("挑战 Threats", ["真实浏览器注入的合规/安全边界", "反检测能力或随平台升级失效", "65+ PR/76 Issues 维护压力", "评测多为自报，需第三方验证"], NAVY, Inches(6.77), Inches(4.05)),
]
for head, lines, ac, x, y in quad:
    card(s, x, y, Inches(5.85), Inches(2.3), head, lines, ac)
footer(s)

# ================= Slide 14: 总结洞察 =================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, Inches(0.9), Inches(0.85), Inches(0.18), Inches(1.0), ORANGE)
add_text(s, Inches(1.25), Inches(0.8), Inches(11), Inches(1.1),
         [[("THINKING TAKEAWAYS", 13, TEAL, True)],
          [("技术洞察总结", 34, WHITE, True)]], space_after=4)
takes = [
    ("极简即强大", "用 9 个原子工具 + code_run 替代庞大工具集，证明「能力来自组合而非堆叠」。"),
    ("记忆是护城河", "五层结晶式记忆 + 自进化，让 Agent 越用越强，沉淀成真正的个人资产。"),
    ("效率即体验", "30K 上下文 / 约 6× 更低 token，直接转化为更低成本、更少幻觉、更高成功率。"),
    ("真实优于沙箱", "注入真实浏览器换来反检测与登录态——这是其在开放网络落地的关键差异点。"),
    ("需冷静看待", "高权限执行的安全边界、评测自报性、记忆治理与可复现性，是规模化前必须回答的问题。"),
]
y = Inches(2.05)
for i, (h, d) in enumerate(takes):
    add_rect(s, Inches(1.25), y, Inches(0.55), Inches(0.78), TEAL if i%2==0 else ORANGE, shape=MSO_SHAPE.OVAL)
    add_text(s, Inches(1.25), y, Inches(0.55), Inches(0.78),
             [[(str(i+1), 18, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(2.05), y+Inches(0.02), Inches(10.4), Inches(0.8),
             [[(h+"　", 16, TEAL if i%2==0 else ORANGE, True), (d, 13, RGBColor(0xD5,0xDF,0xE8), False)]],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    y += Inches(0.92)
add_text(s, Inches(1.25), Inches(6.95), Inches(11), Inches(0.4),
         [[("一句话定位：以最小种子代码，长出会自我进化的个人级自治智能体。", 13, RGBColor(0x9A,0xAE,0xBF), True)]])

prs.save("GenericAgent_技术洞察分析.pptx")
print("Saved:", "GenericAgent_技术洞察分析.pptx")
print("Slides:", len(prs.slides._sldIdLst))
