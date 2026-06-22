# -*- coding: utf-8 -*-
"""Generate charts for the Coze deep-insight deck (ppt-master 图表七步法 style).

七步法: 删网格线 · 配色2-3种 · 字体统一 · 关键数据高亮(品牌色)其余灰 ·
数据标签 · 去冗余图例 · 每图只传达一个观点。
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter
import os

# ---------- 思源黑体 / Noto Sans CJK SC (free commercial) ----------
for cand in ("Noto Sans CJK SC", "Source Han Sans SC", "Source Han Sans CN"):
    try:
        font_manager.findfont(cand, fallback_to_default=False)
        CJK = cand; break
    except Exception:
        CJK = None
if CJK is None:                          # fallback
    fp = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
    font_manager.fontManager.addfont(fp)
    CJK = font_manager.FontProperties(fname=fp).get_name()
matplotlib.rcParams["font.family"] = CJK
matplotlib.rcParams["axes.unicode_minus"] = False

# ---------- palette (70-25-5, low-saturation) ----------
NAVY  = "#16223A"   # text
BLUE  = "#2E6BE6"   # 扣子蓝 / 主色·关键数据
BLUED = "#163E8C"   # 深蓝（强调）
BLUEL = "#9DBBF2"   # 浅蓝（次级）
GRAY  = "#C2C9D6"   # 非关键灰
GRAYT = "#6B7280"   # 标签灰
AMBER = "#E08A1E"   # 强调 5%
INK   = "#1F2937"

OUT = "/home/user/openai-chatgpt-codex/coze-insight/charts"
os.makedirs(OUT, exist_ok=True)

def _save(fig, name):
    p = os.path.join(OUT, name)
    fig.savefig(p, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig); print("saved", p)

def _bare(ax):
    """七步法: 删网格线/边框, 仅留必要轴。"""
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(GRAY)
    ax.tick_params(colors=GRAYT, length=0)
    ax.set_yticks([])

# ============================================================ 1. developer / user growth
def dev_growth():
    fig, ax = plt.subplots(figsize=(7.2, 4.1))
    labels = ["2024.12", "2025.12", "2026.01"]
    vals   = [100, 300, 1000]
    cols   = [BLUEL, BLUE, BLUED]
    notes  = ["100万\n活跃开发者", "300万\n月活开发者", "1000万+\n用户/开发场景*"]
    sub    = ["Force 冬季大会", "Force 原动力大会", "扣子 2.0 发布"]
    bars = ax.bar(labels, vals, color=cols, width=0.58, zorder=3)
    for b, n in zip(bars, notes):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()+22, n, ha="center",
                va="bottom", fontsize=12.5, fontweight="bold", color=NAVY)
    for i, sx in enumerate(sub):
        ax.text(i, -90, sx, ha="center", va="top", fontsize=9.5, color=GRAYT)
    ax.set_ylim(0, 1240)
    _bare(ax); ax.set_xticks([])
    ax.text(-0.46, -210, "* 口径随时间演变：活跃开发者 → 月活开发者 → 用户/真实开发场景，非同一计量口径。",
            fontsize=8.6, color=GRAYT)
    _save(fig, "dev_growth.png")

# ============================================================ 2. open-source stars
def github_stars():
    fig, ax = plt.subplots(figsize=(7.0, 2.7))
    labels = ["Coze Studio\n开发平台开源版", "Coze Loop\n扣子罗盘开源版"]
    vals   = [21013, 5536]
    bars = ax.barh(labels, vals, color=[BLUE, BLUEL], height=0.52, zorder=3)
    for b, v in zip(bars, vals):
        ax.text(b.get_width()+350, b.get_y()+b.get_height()/2, f"{v:,} ★",
                va="center", fontsize=13, fontweight="bold", color=NAVY)
    ax.set_xlim(0, 25000); ax.invert_yaxis()
    for s in ("top", "right", "bottom"): ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(GRAY)
    ax.set_xticks([]); ax.tick_params(colors=NAVY, length=0, labelsize=11)
    _save(fig, "github_stars.png")

# ============================================================ 3. Doubao token volume (log)
def doubao_tokens():
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    labels = ["2024.05", "2025.10", "2025.12", "2026.03"]
    vals   = [0.12, 30, 63, 120]                       # 万亿 / 日
    cols   = [GRAY, BLUEL, BLUE, BLUED]
    txt    = ["1200亿", "30万亿", "63万亿", "120万亿"]
    bars = ax.bar(labels, vals, color=cols, width=0.6, zorder=3)
    ax.set_yscale("log")
    for b, t in zip(bars, txt):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()*1.18, t, ha="center",
                va="bottom", fontsize=12, fontweight="bold", color=NAVY)
    ax.set_ylim(0.05, 420)
    ax.set_yticks([])
    for s in ("top", "right", "left"): ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(GRAY)
    ax.tick_params(colors=GRAYT, length=0, labelsize=11)
    ax.annotate("", xy=(3, 150), xytext=(0, 0.4),
                arrowprops=dict(arrowstyle="-|>", color=AMBER, lw=2, alpha=.55))
    ax.text(1.05, 150, "两年 ≈ 1000 倍", color=AMBER, fontsize=12, fontweight="bold")
    _save(fig, "doubao_tokens.png")

# ============================================================ 4. IDC MaaS share donut
def maas_share():
    fig, ax = plt.subplots(figsize=(5.6, 4.4))
    labels = ["百度智能云 26%", "阿里云 19%", "字节·火山引擎 16%", "腾讯云 10%", "商汤 5%", "其他 24%"]
    vals   = [26, 19, 16, 10, 5, 24]
    cols   = ["#D6DCE6", "#C2C9D6", BLUE, "#CDD4DF", "#DCE1E9", "#E8ECF2"]
    w, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                  explode=[0,0,0.09,0,0,0],
                  wedgeprops=dict(width=0.40, edgecolor="white", linewidth=2))
    ax.legend(w, labels, loc="center left", bbox_to_anchor=(0.96, 0.5),
              fontsize=10.5, frameon=False)
    ax.text(0, 0.10, "16%", ha="center", va="center", fontsize=22, fontweight="bold", color=BLUE)
    ax.text(0, -0.20, "字节 · 第三", ha="center", va="center", fontsize=11, color=GRAYT)
    _save(fig, "maas_share.png")

if __name__ == "__main__":
    dev_growth(); github_stars(); doubao_tokens(); maas_share()
    print("font:", CJK, "| all charts done")
