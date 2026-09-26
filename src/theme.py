# -*- coding: utf-8 -*-
"""TraeWork 风格浅色主题调色板（TRAE 家族设计语言）

主窗口 app.py 与规则弹窗 rules_dialog.py 共用，保证两处观感一致。
使用方式：`import theme` 后引用 theme.xxx。
"""

# 页面/画布底色（浅灰蓝，TraeWork 浅色工作台质感）
PAGE_BG = "#F5F6FA"

# 卡片：白底 + 1px 低对比边框 + 圆角
CARD_BG = "#FFFFFF"
CARD_BORDER = "#E7EAF1"

# 控件轨道 / 浅底填充（分段控件、进度条）
TRACK = "#EDF0F7"
TRACK_DEEP = "#E4E8F3"

# 输入框
INPUT_BG = "#FFFFFF"
INPUT_BORDER = "#DFE3EC"

# 文字层级
TEXT = "#1F2329"        # 主文字
SUBTEXT = "#6E7686"     # 次要/提示文字

# 强调色（TRAE 蓝）与悬停、浅蓝选中态
ACCENT = "#4B7BFF"
ACCENT_HOVER = "#3B66E0"
ACCENT_SOFT = "#DDE6FF"
SEG_TEXT_SEL = "#FFFFFF"    # 分段控件选中文字
SEG_TEXT = "#3C4250"        # 分段控件未选中文字

# 语义色
DANGER = "#E5484D"
DANGER_HOVER = "#C93A3F"

# 按钮常用别名（沿用原 BTN_* 语义，文字改白色以适配浅色主题）
BTN_FG = ACCENT
BTN_HOVER = ACCENT_HOVER
BTN_TEXT = "#FFFFFF"


def style_treeview():
    """全局 ttk.Treeview 浅色样式（主窗口与规则弹窗共用的表格）"""
    from tkinter import ttk

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "Treeview",
        background=CARD_BG, fieldbackground=CARD_BG,
        foreground=TEXT, rowheight=30, borderwidth=0,
        font=("Microsoft YaHei UI", 10))
    style.configure(
        "Treeview.Heading",
        background="#F4F6FB", foreground="#3C4250",
        font=("Microsoft YaHei UI", 10, "bold"),
        borderwidth=0, relief="flat")
    style.map(
        "Treeview",
        background=[("selected", ACCENT_SOFT)],
        foreground=[("selected", "#1C3FAA")])
    style.map("Treeview.Heading", background=[("active", "#E9EDF7")])