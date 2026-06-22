# -*- coding: utf-8 -*-
"""Generate charts for the ByteDance Coze deep-insight deck."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os

# ---------- Chinese font ----------
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
font_manager.fontManager.addfont(FONT_PATH)
_fp = font_manager.FontProperties(fname=FONT_PATH)
matplotlib.rcParams["font.family"] = _fp.get_name()
matplotlib.rcParams["axes.unicode_minus"] = False

# ---------- palette (match the deck) ----------
NAVY  = "#0C162C"
BLUE  = "#2B6CF6"
CYAN  = "#12B5CB"
PURPLE= "#6E4AF0"
GREEN = "#129A5E"
AMBER = "#E08A1E"
RED   = "#C01C2E"
GRAY  = "#6B7280"
LGRAY = "#D8DEE9"
INK   = "#1F2937"

OUT = "/home/user/openai-chatgpt-codex/coze-insight/charts"
os.makedirs(OUT, exist_ok=True)

def _save(fig, name):
    path = os.path.join(OUT, name)
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", path)

def _clean(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(LGRAY)
    ax.spines["bottom"].set_color(LGRAY)
    ax.tick_params(colors=GRAY, length=0)

# ============================================================ 1. developer / user growth
def dev_growth():
    fig, ax = plt.subplots(figsize=(7.4, 4.0))
    labels = ["2024.12\nForce 冬季大会", "2025.12\nForce 原动力大会", "2026.01\n扣子 2.0 发布"]
    vals = [100, 300, 1000]
    cols = [BLUE, BLUE, PURPLE]
    bars = ax.bar(labels, vals, color=cols, width=0.6, zorder=3)
    notes = ["100万\n活跃开发者", "300万\n月活开发者", "1000万+\n用户/开发场景*"]
    for b, n in zip(bars, notes):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()+25, n,
                ha="center", va="bottom", fontsize=11, fontweight="bold", color=INK)
    ax.set_ylim(0, 1200)
    ax.set_ylabel("规模（万）", fontsize=10, color=GRAY)
    ax.grid(axis="y", color="#EEF1F6", zorder=0)
    _clean(ax)
    ax.text(0, -260, "* 口径随时间演变：从“活跃开发者”→“月活开发者”→“用户/真实开发场景”，非同一计量口径，不可简单连成单一曲线。",
            fontsize=8.2, color=GRAY)
    ax.set_title("扣子开发者 / 用户规模演进（官方披露口径）", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    _save(fig, "dev_growth.png")

# ============================================================ 2. open-source stars
def github_stars():
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    labels = ["Coze Studio\n(开发平台开源版)", "Coze Loop\n(扣子罗盘开源版)"]
    vals = [21013, 5536]
    cols = [BLUE, CYAN]
    bars = ax.barh(labels, vals, color=cols, height=0.55, zorder=3)
    for b, v in zip(bars, vals):
        ax.text(b.get_width()+300, b.get_y()+b.get_height()/2, f"{v:,} ★",
                va="center", fontsize=12, fontweight="bold", color=INK)
    ax.set_xlim(0, 24000)
    ax.invert_yaxis()
    ax.grid(axis="x", color="#EEF1F6", zorder=0)
    _clean(ax)
    ax.set_title("开源仓库 GitHub Stars（Apache 2.0 · GitHub API 实测 2026-06-22）",
                 fontsize=12.5, fontweight="bold", color=NAVY, pad=10)
    _save(fig, "github_stars.png")

# ============================================================ 3. Doubao token volume (log)
def doubao_tokens():
    fig, ax = plt.subplots(figsize=(7.2, 3.9))
    labels = ["2024.05\n豆包发布", "2025.10", "2025.12", "2026.03"]
    vals = [0.12, 30, 63, 120]  # 万亿 tokens / 日
    bars = ax.bar(labels, vals, color=[LGRAY, CYAN, BLUE, PURPLE], width=0.6, zorder=3)
    ax.set_yscale("log")
    txt = ["1200亿", "30万亿", "63万亿", "120万亿"]
    for b, t in zip(bars, txt):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()*1.15, t,
                ha="center", va="bottom", fontsize=11, fontweight="bold", color=INK)
    ax.set_ylim(0.05, 400)
    from matplotlib.ticker import FuncFormatter
    ax.set_yticks([0.1, 1, 10, 100])
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, _: ("%g" % y)))
    ax.yaxis.set_minor_formatter(plt.NullFormatter())
    ax.set_ylabel("日均 Tokens（对数轴）", fontsize=10, color=GRAY)
    ax.grid(axis="y", color="#EEF1F6", which="both", zorder=0)
    _clean(ax)
    ax.set_title("豆包大模型日均 Token 调用量：两年增长约 1000 倍", fontsize=13, fontweight="bold", color=NAVY, pad=12)
    _save(fig, "doubao_tokens.png")

# ============================================================ 4. IDC MaaS share donut
def maas_share():
    fig, ax = plt.subplots(figsize=(6.0, 4.6))
    labels = ["百度智能云 26%", "阿里云 19%", "字节·火山引擎 16%", "腾讯云 10%", "商汤 5%", "其他 24%"]
    vals = [26, 19, 16, 10, 5, 24]
    cols = [LGRAY, LGRAY, BLUE, LGRAY, LGRAY, "#EDF0F5"]
    explode = [0, 0, 0.08, 0, 0, 0]
    w, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                  explode=explode, wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.legend(w, labels, loc="center left", bbox_to_anchor=(0.92, 0.5),
              fontsize=10, frameon=False)
    ax.text(0, 0, "字节\n第三", ha="center", va="center", fontsize=15, fontweight="bold", color=BLUE)
    ax.set_title("中国大模型 MaaS 市场份额 2024H2（IDC）", fontsize=12.5, fontweight="bold", color=NAVY, pad=8)
    _save(fig, "maas_share.png")

# ============================================================ 5. Volcano model-call share
def volcano_share():
    fig, ax = plt.subplots(figsize=(4.6, 4.0))
    labels = ["2024", "2025"]
    vals = [46.4, 59.2]
    bars = ax.bar(labels, vals, color=[CYAN, BLUE], width=0.5, zorder=3)
    for b, v in zip(bars, vals):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()+1.5, f"{v}%",
                ha="center", va="bottom", fontsize=13, fontweight="bold", color=INK)
    ax.set_ylim(0, 75)
    ax.set_ylabel("份额（%）", fontsize=10, color=GRAY)
    ax.grid(axis="y", color="#EEF1F6", zorder=0)
    _clean(ax)
    ax.set_title("火山引擎大模型云上调用量份额（IDC）", fontsize=12, fontweight="bold", color=NAVY, pad=10)
    _save(fig, "volcano_share.png")

# ============================================================ 6. product timeline
def timeline():
    fig, ax = plt.subplots(figsize=(12.4, 3.5))
    events = [
        ("2023.11", "海外版 Coze 上线", BLUE),
        ("2024.02", "国内版扣子(coze.cn)上线", BLUE),
        ("2024.07", "启动商业化收费", CYAN),
        ("2025.04", "扣子空间：首个通用 Agent", PURPLE),
        ("2025.07", "开源 Studio+Loop(Apache2.0)", GREEN),
        ("2025.12", "扣子编程·Vibe Coding\n(300万月活开发者)", BLUE),
        ("2026.01", "扣子 2.0：职场 AI\nSkills/Plan/Coding/Office", PURPLE),
        ("2026.04", "扣子 2.5 / Agent World", AMBER),
    ]
    n = len(events)
    xs = list(range(n))
    ax.plot([-0.4, n-0.6], [0, 0], color=LGRAY, lw=3, zorder=1)
    for i, (date, label, col) in enumerate(events):
        ax.scatter(i, 0, s=190, color=col, zorder=3, edgecolor="white", linewidth=2)
        up = (i % 2 == 0)
        y = 0.42 if up else -0.42
        va = "bottom" if up else "top"
        ax.plot([i, i], [0, y*0.62], color=col, lw=1.4, zorder=2)
        ax.text(i, y, label, ha="center", va=va, fontsize=9.6, color=INK, fontweight="bold")
        ax.text(i, 0.13 if up else -0.13, date, ha="center",
                va="bottom" if up else "top", fontsize=10.5, color=col, fontweight="bold")
    ax.set_xlim(-0.6, n-0.4)
    ax.set_ylim(-1.05, 1.05)
    ax.axis("off")
    _save(fig, "timeline.png")

if __name__ == "__main__":
    dev_growth()
    github_stars()
    doubao_tokens()
    maas_share()
    timeline()
    print("all charts done")
