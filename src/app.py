# -*- coding: utf-8 -*-
"""GUI 主窗口（CustomTkinter 单窗口紧凑布局）

布局（自上而下）：
① 选择区：PDF 文件夹 + 规则来源显示
② 选项区：命名方式（单选）+ 包含子文件夹（默认包含，原脚本行为）
③ 操作区：识别预览 / 确认并执行 / 停止
④ 进度条 + 进度文本
⑤ 结果表（状态 / 原文件名 / 识别标题 / 新文件名）
⑥ 底部状态栏：统计 + 撤销上次 + 规则管理

线程模型：
- 工作线程执行 扫描/分析/执行/撤销，通过 queue 向主线程发消息
- 主线程 after(80ms) 轮询 queue 刷新 UI（tkinter 非线程安全，禁止在工作线程碰 UI）
- stop_event（threading.Event）实现随时停止
"""

import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

# 修复：默认的 font_shapes 绘制方法在本机（Python 3.14 + 高 DPI 缩放）下
# 会导致滚动条等小尺寸部件无法渲染（画布内容正确但屏幕上不显示）。
# 改用 polygon_shapes（macOS 默认方法，纯画布多边形绘制，不依赖字体）。
from customtkinter.windows.widgets.core_rendering import DrawEngine
DrawEngine.preferred_drawing_method = "polygon_shapes"

from core import batch, rules as rules_mod
from rules_dialog import RulesDialog
import theme

APP_NAME = "PDF重命名"
APP_VERSION = "3.3.0"

# 按钮统一样式（配色来自 theme.py，TraeWork 风格浅色主题）：TRAE 蓝底 + 白色文字
BTN_FG = theme.BTN_FG
BTN_HOVER = theme.BTN_HOVER
BTN_TEXT = theme.BTN_TEXT


def resource_path(rel):
    """定位打包后的资源文件（图标等）

    PyInstaller onedir 模式下资源在 _MEIPASS（_internal 目录）；
    开发模式下回退到项目 build/ 目录。
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, rel)
    return os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "build", rel)


class App(ctk.CTk):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.title("{} v{}".format(APP_NAME, APP_VERSION))
        self.geometry("980x720")
        self.minsize(860, 620)
        self.configure(fg_color=theme.PAGE_BG)

        # 窗口图标（开发模式下图标缺失时用默认图标，不报错）
        try:
            self.iconbitmap(resource_path("app.ico"))
        except Exception:
            pass

        # ---------- 状态 ----------
        self.rules = []
        self.analyze_results = []       # 最近一次预览的分析结果
        self.match_results = []         # 预览中 status=="match" 的子集
        self.split_analysis = None      # 拆分模式的分析结果
        self.mode_now = "rename"        # rename=现有重命名 / split=拆分
        self.by_folder_now = False      # 预览时的命名方式（执行时保持一致）
        self.stop_event = threading.Event()
        self.msg_queue = queue.Queue()
        self.worker = None
        self.t_start = None

        # ---------- 加载规则 ----------
        try:
            self.rules, msg = rules_mod.load_rules_safe()
        except Exception as e:
            self.rules, msg = [], "规则加载失败：{}".format(e)

        self._build_ui()
        self._set_state("idle")
        self.after(80, self._poll_queue)

        if msg:
            self.status_label.configure(text=msg)

    # ============================================================
    # UI 构建
    # ============================================================

    def _build_ui(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        theme.style_treeview()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)   # 结果表区域可拉伸

        # ---------- ① 选择区（卡片） ----------
        box1 = ctk.CTkFrame(self, fg_color=theme.CARD_BG, corner_radius=14,
                            border_width=1, border_color=theme.CARD_BORDER)
        box1.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        box1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(box1, text="PDF 文件夹",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=(16, 8), pady=(12, 4), sticky="w")
        self.folder_entry = ctk.CTkEntry(
            box1, placeholder_text="点击右侧「浏览」选择包含 PDF 的文件夹（含子文件夹）",
            fg_color=theme.INPUT_BG, border_color=theme.INPUT_BORDER,
            border_width=1, placeholder_text_color=theme.SUBTEXT)
        self.folder_entry.grid(row=0, column=1, padx=4, pady=(12, 4), sticky="ew")
        ctk.CTkButton(box1, text="浏览…", width=80,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._pick_folder).grid(row=0, column=2, padx=(4, 16), pady=(12, 4))

        ctk.CTkLabel(box1, text="映射规则",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=(16, 8), pady=(4, 4), sticky="w")
        self.rules_label = ctk.CTkLabel(box1, text=self._rules_summary(),
                                        anchor="w", text_color=theme.SUBTEXT)
        self.rules_label.grid(row=1, column=1, padx=4, pady=(4, 4), sticky="ew")
        ctk.CTkButton(box1, text="规则管理…", width=90,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._open_rules).grid(
            row=1, column=2, padx=(4, 16), pady=(4, 4))

        # 项目名称解析规则（拆分模式下用于从开工报告表格提取「工程名称」）
        # 说明：地域词/业务后缀原本硬编码在源码中，现改为可配置 + 可由
        # 用户的「项目进度表 Excel」自动学习，避免换地区/换业务后失效。
        ctk.CTkLabel(box1, text="项目名称规则",
                     font=ctk.CTkFont(weight="bold")).grid(
            row=2, column=0, padx=(16, 8), pady=(4, 12), sticky="w")
        self.name_cfg_label = ctk.CTkLabel(
            box1, text=self._name_cfg_summary(), anchor="w",
            text_color=theme.SUBTEXT)
        self.name_cfg_label.grid(row=2, column=1, padx=4, pady=(4, 12), sticky="ew")
        ctk.CTkButton(box1, text="从进度表学习…", width=110,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._learn_name_rules).grid(
            row=2, column=2, padx=(4, 16), pady=(4, 12))

        # ---------- ② 选项区（卡片） ----------
        box2 = ctk.CTkFrame(self, fg_color=theme.CARD_BG, corner_radius=14,
                            border_width=1, border_color=theme.CARD_BORDER)
        box2.grid(row=1, column=0, sticky="ew", padx=16, pady=8)

        ctk.CTkLabel(box2, text="处理模式：",
                     font=ctk.CTkFont(size=13)).grid(
            row=0, column=0, padx=(16, 8), pady=(12, 2), sticky="w")
        self.mode_sel = ctk.CTkSegmentedButton(
            box2, values=["重命名文件", "拆分整份PDF"],
            fg_color=theme.TRACK, border_width=0,
            selected_color=BTN_FG, selected_hover_color=BTN_HOVER,
            unselected_color=theme.TRACK, unselected_hover_color="#E3E7F2",
            text_color=theme.SEG_TEXT, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._on_mode_changed)
        self.mode_sel.set("重命名文件")
        self.mode_sel.grid(row=0, column=1, padx=4, pady=(12, 2), sticky="w")
        ctk.CTkLabel(box2, text="拆分模式：需选择包含整份扫描 PDF 的根文件夹（其下为各项目子文件夹）",
                     anchor="w", text_color=theme.SUBTEXT, font=ctk.CTkFont(size=12)).grid(
            row=0, column=2, padx=8, pady=(12, 2), sticky="w")

        ctk.CTkLabel(box2, text="命名方式：",
                     font=ctk.CTkFont(size=13)).grid(
            row=1, column=0, padx=(16, 8), pady=(2, 12), sticky="w")
        self.name_mode = ctk.CTkSegmentedButton(
            box2, values=["仅按映射内容", "文件夹名+映射内容"],
            fg_color=theme.TRACK, border_width=0,
            selected_color=BTN_FG, selected_hover_color=BTN_HOVER,
            unselected_color=theme.TRACK, unselected_hover_color="#E3E7F2",
            text_color=theme.SEG_TEXT, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._on_mode_changed)
        self.name_mode.set("仅按映射内容")
        self.name_mode.grid(row=1, column=1, padx=4, pady=(2, 12), sticky="w")
        self._sync_segment(self.mode_sel)
        self._sync_segment(self.name_mode)

        # ---------- ③ 操作区 ----------
        box3 = ctk.CTkFrame(self, fg_color="transparent")
        box3.grid(row=2, column=0, sticky="ew", padx=16, pady=8)

        self.btn_preview = ctk.CTkButton(box3, text="① 识别预览", width=130,
                                         fg_color=BTN_FG, hover_color=BTN_HOVER,
                                         text_color=BTN_TEXT,
                                         font=ctk.CTkFont(weight="bold"),
                                         command=self._start_analyze)
        self.btn_preview.grid(row=0, column=0, padx=(0, 8))
        self.btn_execute = ctk.CTkButton(box3, text="② 确认并执行重命名", width=160,
                                         fg_color=BTN_FG, hover_color=BTN_HOVER,
                                         text_color=BTN_TEXT,
                                         font=ctk.CTkFont(weight="bold"),
                                         command=self._start_execute)
        self.btn_execute.grid(row=0, column=1, padx=8)
        self.btn_stop = ctk.CTkButton(box3, text="停止", width=80, fg_color="#C0392B",
                                       hover_color="#96281B", text_color="#FFFFFF",
                                       command=self._stop)
        self.btn_stop.grid(row=0, column=2, padx=8)

        # 阶段四 4.2：页级 OCR 缓存管理（查看占用 / 清空）
        self.btn_cache = ctk.CTkButton(box3, text="缓存管理", width=100,
                                       fg_color=theme.TRACK, hover_color="#E3E7F2",
                                       text_color=theme.SEG_TEXT,
                                       command=self._manage_cache)
        self.btn_cache.grid(row=0, column=3, padx=8)

        # 阶段五 R-6：资料完整性核对（应交 vs 实收，导出差异报告）
        self.btn_inventory = ctk.CTkButton(box3, text="完整性核对", width=110,
                                           fg_color=theme.TRACK,
                                           hover_color="#E3E7F2",
                                           text_color=theme.SEG_TEXT,
                                           command=self._run_inventory)
        self.btn_inventory.grid(row=0, column=4, padx=8)

        # ---------- ④ 进度区 ----------
        box4 = ctk.CTkFrame(self, fg_color="transparent")
        box4.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 8))
        box4.grid_columnconfigure(0, weight=1)

        self.progress = ctk.CTkProgressBar(
            box4, height=10, fg_color=theme.TRACK_DEEP,
            progress_color=BTN_FG)
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew")
        self.progress_label = ctk.CTkLabel(
            box4, text="就绪", anchor="w", text_color=theme.SUBTEXT)
        self.progress_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        # ---------- ⑤ 结果表（卡片） ----------
        box5 = ctk.CTkFrame(self, fg_color=theme.CARD_BG, corner_radius=14,
                            border_width=1, border_color=theme.CARD_BORDER)
        box5.grid(row=4, column=0, sticky="nsew", padx=16, pady=8)
        box5.grid_rowconfigure(0, weight=1)
        box5.grid_columnconfigure(0, weight=1)

        columns = ("status", "old", "title", "new")
        self.tree = ttk.Treeview(box5, columns=columns, show="headings",
                                 selectmode="browse")
        self.tree.heading("status", text="状态")
        self.tree.heading("old", text="原文件名")
        self.tree.heading("title", text="识别标题（截取）")
        self.tree.heading("new", text="新文件名")
        # 列总宽（1070px）大于可视区域 → 水平滚动条生效，拖动可查看完整内容
        self.tree.column("status", width=90, minwidth=70, anchor="center")
        self.tree.column("old", width=340, minwidth=160)
        self.tree.column("title", width=320, minwidth=140)
        self.tree.column("new", width=320, minwidth=140)

        # 垂直 + 水平双滚动条
        # 用 CTkScrollbar（自绘样式）替代 ttk.Scrollbar：
        # Windows 11 下 ttk 滚动条存在不显示、无法操作的兼容性问题
        vsb = ctk.CTkScrollbar(box5, orientation="vertical", command=self.tree.yview)
        hsb = ctk.CTkScrollbar(box5, orientation="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=(8, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))
        # Shift + 鼠标滚轮 = 左右滚动（Windows 滚轮 delta 为 ±120）
        self.tree.bind(
            "<Shift-MouseWheel>",
            lambda e: self.tree.xview_scroll(-1 if e.delta > 0 else 1, "units"))

        # 状态颜色标签
        self.tree.tag_configure("match", foreground="#1E8449")
        self.tree.tag_configure("no_match", foreground="#B7950B")
        self.tree.tag_configure("failed", foreground="#C0392B")
        self.tree.tag_configure("renamed", foreground="#1E8449")
        self.tree.tag_configure("cancelled", foreground="gray")

        # ---------- ⑥ 底部状态栏 ----------
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=5, column=0, sticky="ew", padx=16, pady=(4, 16))

        self.status_label = ctk.CTkLabel(bar, text="", anchor="w",
                                         text_color=theme.SUBTEXT)
        self.status_label.grid(row=0, column=0, sticky="ew", padx=16, pady=10)
        bar.grid_columnconfigure(0, weight=1)

        self.btn_undo = ctk.CTkButton(bar, text="撤销上次重命名", width=120,
                                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                                      text_color=BTN_TEXT,
                                      font=ctk.CTkFont(weight="bold"),
                                      command=self._undo)
        self.btn_undo.grid(row=0, column=1, padx=(4, 4))

        # 阶段三 3.2：撤销上次拆分（删除本次拆分新写出的文件）
        self.btn_undo_split = ctk.CTkButton(
            bar, text="撤销上次拆分", width=120,
            fg_color=BTN_FG, hover_color=BTN_HOVER,
            text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
            command=self._undo_split)
        self.btn_undo_split.grid(row=0, column=2, padx=(4, 16))

    # ============================================================
    # 状态与辅助
    # ============================================================

    def _rules_summary(self):
        return "当前已加载 {} 条映射规则（可在「规则管理」中编辑或导入 Excel）".format(len(self.rules))

    def _name_cfg_summary(self):
        """项目名称解析配置摘要（拆分模式下用于提取「工程名称」）"""
        try:
            from core import project_config as PC
            cfg = PC.load_config()
        except Exception as e:
            return "配置加载失败：{}".format(e)
        n_region = len(cfg.get("region_words") or [])
        n_suffix = len(cfg.get("business_suffixes") or [])
        if n_region:
            sample = "、".join((cfg.get("region_words") or [])[:3])
            return "已学习 {} 个地域词（如 {}）· {} 个业务后缀".format(
                n_region, sample, n_suffix)
        return ("使用内置通用语法（{} 个业务后缀，未限定地域）"
                "· 可点右侧按钮从项目进度表自动学习").format(n_suffix)

    def _set_state(self, state):
        """按工作状态启用/禁用按钮"""
        self.btn_preview.configure(state="normal" if state == "idle" else "disabled")
        self.btn_execute.configure(
            state="normal" if state == "preview_ready" else "disabled")
        self.btn_stop.configure(
            state="normal" if state in ("analyzing", "executing", "undoing") else "disabled")
        self.btn_undo.configure(
            state="normal" if state == "idle" and batch.get_last_rename_count() > 0 else "disabled")
        # 阶段三：撤销拆分按钮同理，按「最近一笔拆分」是否有记录决定可用性
        try:
            _has_split = batch.get_last_split_count() > 0
        except Exception:
            _has_split = False
        self.btn_undo_split.configure(
            state="normal" if state == "idle" and _has_split else "disabled")

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if folder:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, folder)

    def _on_mode_changed(self, _value=None):
        """处理模式切换：重命名文件 / 拆分整份PDF"""
        if self.mode_sel.get() == "拆分整份PDF":
            self.mode_now = "split"
            self.btn_preview.configure(text="① 分析拆分")
            self.btn_execute.configure(text="② 执行拆分")
            # 命名方式在拆分模式下同样可选：
            # 仅映射内容 → 拆出的文档仅用标题命名；
            # 文件夹名+映射内容 → 用「项目文件夹名/工程名称 + 标题」命名
            self.folder_entry.configure(
                placeholder_text="点击右侧「浏览」选择根文件夹（内含整份扫描 PDF 与各项目子文件夹）")
        else:
            self.mode_now = "rename"
            self.btn_preview.configure(text="① 识别预览")
            self.btn_execute.configure(text="② 确认并执行重命名")
            self.name_mode.configure(state="normal")
            self.folder_entry.configure(
                placeholder_text="点击右侧「浏览」选择包含 PDF 的文件夹（含子文件夹）")
        self.split_analysis = None
        self.tree.delete(*self.tree.get_children())
        self.progress_label.configure(text="就绪")
        self._set_state("idle")
        self._sync_segment(self.mode_sel)
        self._sync_segment(self.name_mode)

    def _sync_segment(self, seg):
        """分段控件当前选中项白字、未选中项深灰字（选中底色为 TraeWork 蓝）"""
        for value, btn in seg._buttons_dict.items():
            btn.configure(text_color=(
                "#FFFFFF" if value == seg.get() else "#3C4250"))

    def _stop(self):
        self.stop_event.set()
        self.progress_label.configure(text="正在停止…（当前文件处理完后即停）")

    # ============================================================
    # ① 识别预览
    # ============================================================

    def _start_analyze(self):
        if self.mode_now == "split":
            self._start_split_analyze()
            return
        folder = self.folder_entry.get().strip()
        if not folder:
            messagebox.showwarning(APP_NAME, "请先选择 PDF 文件夹")
            return
        if not os.path.isdir(folder):
            messagebox.showerror(APP_NAME, "所选文件夹不存在：{}".format(folder))
            return
        if not self.rules:
            messagebox.showwarning(APP_NAME, "映射规则为空，请先在「规则管理」中添加或导入规则")
            return

        self.by_folder_now = (self.name_mode.get() == "文件夹名+映射内容")
        self.stop_event = threading.Event()
        self.tree.delete(*self.tree.get_children())
        self.progress.set(0)
        self.progress_label.configure(text="正在扫描文件夹…")
        self._set_state("analyzing")
        self.t_start = time.time()

        self.worker = threading.Thread(
            target=self._worker_analyze, args=(folder,), daemon=True)
        self.worker.start()

    def _worker_analyze(self, folder):
        try:
            pdf_files = batch.scan_pdfs(folder)
            if not pdf_files:
                self.msg_queue.put({"type": "error",
                                    "msg": "该文件夹下未发现任何 PDF 文件"})
                return

            def progress_cb(done, total, result):
                self.msg_queue.put({"type": "analyze_progress",
                                    "done": done, "total": total,
                                    "result": result})

            results = batch.analyze_batch(
                pdf_files, self.rules, self.by_folder_now,
                progress_cb=progress_cb, stop_event=self.stop_event)
            self.msg_queue.put({"type": "analyze_done", "results": results})
        except Exception as e:
            self.msg_queue.put({"type": "error", "msg": "识别过程出错：{}".format(e)})

    # ============================================================
    # ② 确认并执行
    # ============================================================

    def _start_execute(self):
        if self.mode_now == "split":
            self._start_split_execute()
            return
        if not self.match_results:
            return
        n = len(self.match_results)
        if not messagebox.askyesno(
                APP_NAME, "将把 {} 个 PDF 按预览结果重命名。\n执行后可通过「撤销上次重命名」回滚。\n\n确定执行吗？".format(n)):
            return

        self.stop_event = threading.Event()
        self.progress.set(0)
        self._set_state("executing")
        self.t_start = time.time()

        self.worker = threading.Thread(target=self._worker_execute, daemon=True)
        self.worker.start()

    def _worker_execute(self):
        try:
            def progress_cb(done, total, item):
                self.msg_queue.put({"type": "exec_progress",
                                    "done": done, "total": total, "item": item})

            renamed, skipped = batch.execute_renames(
                self.match_results, self.by_folder_now,
                progress_cb=progress_cb, stop_event=self.stop_event)
            self.msg_queue.put({"type": "exec_done",
                                "renamed": renamed, "skipped": skipped})
        except Exception as e:
            self.msg_queue.put({"type": "error", "msg": "执行过程出错：{}".format(e)})

    # ============================================================
    # 拆分整份 PDF（新增功能）
    # ============================================================

    def _find_root_pdf(self, root):
        """在根目录下寻找整份扫描 PDF（不含子文件夹），返回路径或 None"""
        for fn in sorted(os.listdir(root)):
            if fn.lower().endswith(".pdf"):
                full = os.path.join(root, fn)
                if os.path.isfile(full):
                    return full
        return None

    def _start_split_analyze(self):
        root = self.folder_entry.get().strip()
        if not root:
            messagebox.showwarning(APP_NAME, "请先选择根文件夹")
            return
        if not os.path.isdir(root):
            messagebox.showerror(APP_NAME, "所选文件夹不存在：{}".format(root))
            return
        if not self.rules:
            messagebox.showwarning(APP_NAME, "映射规则为空，请先在「规则管理」中添加或导入规则")
            return
        pdf = self._find_root_pdf(root)
        if not pdf:
            messagebox.showwarning(
                APP_NAME, "根目录下未发现 PDF 文件。\n请把整份扫描 PDF 放在根文件夹（即各项目子文件夹的上一级）。")
            return

        self.by_folder_now = (self.name_mode.get() == "文件夹名+映射内容")
        self.stop_event = threading.Event()
        self.split_analysis = None
        self.tree.delete(*self.tree.get_children())
        self.progress.set(0)
        self.progress_label.configure(text="正在分析整份 PDF（逐页 OCR 识别，耗时较长）…")
        self._set_state("analyzing")
        self.t_start = time.time()

        self.worker = threading.Thread(
            target=self._worker_split_analyze, args=(pdf, root), daemon=True)
        self.worker.start()

    def _worker_split_analyze(self, pdf_path, root):
        try:
            from core import splitter
        except ImportError as ie:
            self.msg_queue.put({"type": "error",
                                "msg": "缺少 pypdf 依赖，无法执行拆分：{}".format(ie)})
            return
        try:
            analysis = splitter.analyze_pdf(
                pdf_path, self.rules, root,
                progress_cb=lambda d, t, info: self.msg_queue.put(
                    {"type": "split_analyze_progress", "done": d, "total": t, "info": info}),
                stop_event=self.stop_event, by_folder=self.by_folder_now)
            self.msg_queue.put({"type": "split_analyze_done", "analysis": analysis})
        except splitter.SplitStopped:
            self.msg_queue.put({"type": "split_analyze_stopped"})
        except Exception as e:
            self.msg_queue.put({"type": "error", "msg": "拆分分析出错：{}".format(e)})

    def _start_split_execute(self):
        if not self.split_analysis:
            messagebox.showinfo(APP_NAME, "请先执行「① 分析拆分」生成预览")
            return
        try:
            import pypdf  # noqa: F401  确保 pypdf 可用
        except ImportError:
            messagebox.showerror(
                APP_NAME, "缺少 pypdf 依赖库，无法执行拆分。\n请在命令行运行：\n\npip install pypdf")
            return
        projects = self.split_analysis.get("projects", [])
        n_docs = self.split_analysis.get("total_docs", 0)
        n_unmatched = sum(1 for p in projects if p.get("no_folder"))
        n_conflict = sum(1 for p in projects if p.get("matched_folder") and p.get("conflict"))

        if self.by_folder_now:
            naming_note = ("• 命名方式「文件夹名+映射内容」：拆出的文档用"
                           "「项目文件夹名/工程名称 + 标题」命名（如 xx开工报告.pdf）")
        else:
            naming_note = ("• 命名方式「仅按映射内容」：拆出的文档仅用标题命名"
                           "（如 开工报告.pdf）")
        msg = ("将按预览结果拆分整份 PDF（共 {} 页，拆成 {} 个文档），"
               "写入匹配到的项目子文件夹。\n\n"
               "注意：\n"
               "{}\n"
               "• 拆出的 PDF 是新建文件，原整份 PDF 保留不变\n"
               "• 拆分完成后，可继续用「重命名文件」模式处理拆出的文档\n".format(
                   self.split_analysis.get("page_count", 0), n_docs, naming_note))
        extra = []
        if n_unmatched:
            extra.append("• 有 {} 个项目未匹配到子文件夹，将放入「未匹配项目」文件夹".format(n_unmatched))
        if n_conflict:
            extra.append("• 有 {} 个文件夹被多个项目匹配，请执行前在预览中核对".format(n_conflict))
        if extra:
            msg += "\n" + "\n".join(extra) + "\n"
        if not messagebox.askyesno(APP_NAME, msg + "\n确定执行拆分吗？"):
            return

        self.stop_event = threading.Event()
        self.progress.set(0)
        self._set_state("executing")
        self.t_start = time.time()

        self.worker = threading.Thread(target=self._worker_split_execute, daemon=True)
        self.worker.start()

    def _worker_split_execute(self):
        try:
            from core import splitter
            written, skipped = splitter.execute_split(
                self.split_analysis,
                progress_cb=lambda d, t, item: self.msg_queue.put(
                    {"type": "split_exec_progress", "done": d, "total": t, "item": item}),
                stop_event=self.stop_event)
            self.msg_queue.put({"type": "split_exec_done",
                                "written": written, "skipped": skipped})
        except Exception as e:
            self.msg_queue.put({"type": "error", "msg": "执行拆分出错：{}".format(e)})

    # ============================================================
    # 撤销
    # ============================================================

    def _undo(self):
        n = batch.get_last_rename_count()
        if n == 0:
            messagebox.showinfo(APP_NAME, "没有可撤销的重命名记录")
            return
        if not messagebox.askyesno(APP_NAME, "将撤销最近一次批量重命名（共 {} 个文件），恢复为原文件名。\n确定撤销吗？".format(n)):
            return

        self._set_state("undoing")
        self.progress_label.configure(text="正在撤销…")
        self.worker = threading.Thread(target=self._worker_undo, daemon=True)
        self.worker.start()

    def _worker_undo(self):
        try:
            restored, failed = batch.undo_last()
            self.msg_queue.put({"type": "undo_done",
                                "restored": restored, "failed": failed})
        except Exception as e:
            self.msg_queue.put({"type": "error", "msg": "撤销过程出错：{}".format(e)})

    def _undo_split(self):
        """撤销上次拆分（阶段三 3.2）

        安全设计：删除前弹**确认框**，明确告知「只删除本次拆分新写出的文件，
        不动原整份 PDF 与用户已有文件」，并展示数量供用户核对。
        """
        n = batch.get_last_split_count()
        if n == 0:
            messagebox.showinfo(APP_NAME, "没有可撤销的拆分记录")
            return
        msg = ("将删除最近一次拆分**新写出**的 {} 个文件。\n\n"
               "• 不会删除原整份 PDF\n"
               "• 不会删除清单外的任何用户文件\n"
               "• 已被你移动/改名的文件会自动跳过并报告\n\n"
               "确定撤销这次拆分吗？").format(n)
        if not messagebox.askyesno(APP_NAME, msg):
            return

        self._set_state("undoing")
        self.progress_label.configure(text="正在撤销拆分…")
        self.worker = threading.Thread(target=self._worker_undo_split, daemon=True)
        self.worker.start()

    def _worker_undo_split(self):
        try:
            deleted, skipped = batch.undo_last_split()
            self.msg_queue.put({"type": "undo_split_done",
                                "deleted": deleted, "skipped": skipped})
        except Exception as e:
            self.msg_queue.put({"type": "error",
                                "msg": "撤销拆分出错：{}".format(e)})

    # ============================================================
    # 缓存管理（阶段四 4.2 / A4-4）
    # ============================================================

    def _manage_cache(self):
        """查看/清空页级 OCR 缓存

        缓存用于加速「同一份 PDF 的二次分析」（阶段四 A4-1）。
        文件被修改后缓存自动失效（键含大小与 mtime），故清空缓存是安全的，
        只会让下次分析略慢，不影响结果正确性。
        """
        try:
            from core import cache as C
            n, size = C.cache_size()
        except Exception as e:
            messagebox.showerror(APP_NAME, "读取缓存信息失败：{}".format(e))
            return
        mb = size / (1024.0 * 1024.0)
        msg = ("当前缓存：{} 个页面条目，占用 {:.2f} MB\n"
               "位置：{}\n\n"
               "说明：缓存只影响分析速度，不影响识别结果。\n"
               "文件被修改后对应缓存会自动失效。\n\n"
               "确定要清空缓存吗？").format(n, mb, C.get_cache_dir())
        if not messagebox.askyesno(APP_NAME, msg):
            return
        try:
            deleted, freed = C.clear_cache()
            self.status_label.configure(
                text="缓存已清空：{} 个条目，释放 {:.2f} MB".format(
                    deleted, freed / (1024.0 * 1024.0)))
            messagebox.showinfo(APP_NAME,
                                "已清空 {} 个缓存条目（释放 {:.2f} MB）。".format(
                                    deleted, freed / (1024.0 * 1024.0)))
        except Exception as e:
            messagebox.showerror(APP_NAME, "清空缓存失败：{}".format(e))

    # ============================================================
    # 资料完整性核对（阶段五 R-6）
    # ============================================================

    def _run_inventory(self):
        """应交 vs 实收 资料完整性核对

        取值范围：
          - **核对目录**：拆分模式下的根文件夹（其下为各项目子文件夹）；
          - **进度表**：用于生成「应交」清单的项目名称来源。

        输出差异报告（缺失 / 多余 / 命名不一致），并可导出 Excel/CSV。
        """
        root = self.folder_entry.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showinfo(APP_NAME, "请先选择包含各项目子文件夹的根文件夹")
            return

        # 选择项目进度表（用于生成应交清单）
        xlsx = filedialog.askopenfilename(
            title="选择项目进度表（用于生成应交清单）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("全部文件", "*.*")])
        if not xlsx:
            return

        rules_path = None
        try:
            from core import paths as _paths
            rules_path = _paths.get_rules_file()
        except Exception:
            pass

        try:
            from core import inventory as INV
            from core import rules as R

            names = INV.read_progress_names(xlsx)
            if not names:
                messagebox.showwarning(
                    APP_NAME,
                    "未能从所选进度表读到「项目名称」列，请确认表头存在。")
                return

            rules = []
            if rules_path and os.path.exists(rules_path):
                try:
                    rules = R.read_rules(rules_path)
                except Exception:
                    rules = []

            report = INV.build_report(names, None, root, rules=rules)
            s = report.get("summary", {})
            self.status_label.configure(
                text="核对完成：{} 个项目 · 缺失 {} · 多余 {} · 命名不一致 {}".format(
                    s.get("projects", 0), s.get("missing_total", 0),
                    s.get("extra_total", 0), s.get("mismatch_total", 0)))

            summary = ("资料完整性核对完成\n\n"
                       "进度表项目数：{}\n"
                       "应交合计：{}  实收合计：{}\n\n"
                       "缺失：{}　多余：{}　命名不一致：{}\n"
                       "齐全项目：{}　问题项目：{}").format(
                s.get("projects", 0), s.get("required_total", 0),
                s.get("actual_total", 0), s.get("missing_total", 0),
                s.get("extra_total", 0), s.get("mismatch_total", 0),
                s.get("ok_projects", 0), s.get("problem_projects", 0))

            if not messagebox.askyesno(
                    APP_NAME, summary + "\n\n是否导出差异报告？（Excel/CSV）"):
                return
            out = filedialog.asksaveasfilename(
                title="保存差异报告", defaultextension=".xlsx",
                initialfile="资料完整性核对报告.xlsx",
                filetypes=[("Excel 工作簿", "*.xlsx"), ("CSV 文件", "*.csv")])
            if not out:
                return
            if out.lower().endswith(".csv"):
                INV.export_csv(report, out)
            else:
                out = INV.export_xlsx(report, out)
            messagebox.showinfo(APP_NAME, "报告已导出：\n{}".format(out))
        except Exception as e:
            messagebox.showerror(APP_NAME, "完整性核对出错：{}".format(e))

    # ============================================================
    # 规则管理
    # ============================================================

    def _open_rules(self):
        RulesDialog(self, self.rules, on_saved=self._on_rules_saved)

    def _on_rules_saved(self, new_rules):
        self.rules = new_rules
        self.rules_label.configure(text=self._rules_summary())

    def _learn_name_rules(self):
        """从项目进度表 Excel 自动学习「工程名称」解析规则

        背景：拆分模式需要从开工报告表格里提取「工程名称」以匹配项目子文件夹。
        早期版本把「某省/某市/某县」「业扩配套工程」等硬编码在源码中，
        换地区或换业务类型后即失效（且静默失败）。现改为从用户自己的
        项目进度表（含「项目名称」列）自动学习地域词与业务后缀。
        """
        path = filedialog.askopenfilename(
            title="选择项目进度表 Excel（需含「项目名称」列）",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            from core import project_config as PC
            cfg, info = PC.auto_configure(path, save=True)
        except Exception as e:
            messagebox.showerror(APP_NAME, "学习失败：{}".format(e))
            return

        inferred = info.get("inferred") or {}
        if not inferred:
            messagebox.showwarning(
                APP_NAME,
                "未能从该表格中学习到规则。\n\n"
                "请确认表格中存在「项目名称」列，且名称形如：\n"
                "  某省某市某县3225000000000001某某公司10kV业扩配套工程\n"
                "  2026年某市某县低压台区维修项目包31")
            return

        # 让拆分模块立即使用新配置（避免缓存旧值）
        try:
            from core import splitter
            splitter.set_config(cfg)
        except Exception:
            pass
        self.name_cfg_label.configure(text=self._name_cfg_summary())

        regions = "、".join((inferred.get("region_words") or [])[:6]) or "（无）"
        suffixes = "、".join((inferred.get("business_suffixes") or [])[:5]) or "（无）"
        keywords = "、".join(inferred.get("valid_keywords") or []) or "（无）"
        messagebox.showinfo(
            APP_NAME,
            "已学习项目名称规则并保存。\n\n"
            "地域词（{} 个）：{}\n"
            "业务后缀（{} 个）：{}\n"
            "关键词：{}\n\n"
            "配置文件：\n{}".format(
                len(inferred.get("region_words") or []), regions,
                len(inferred.get("business_suffixes") or []), suffixes,
                keywords, info.get("saved") or PC.config_path()))

    # ============================================================
    # 主线程消息泵（刷新 UI）
    # ============================================================

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self._handle_msg(msg)
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _handle_msg(self, msg):
        mtype = msg.get("type")

        if mtype == "analyze_progress":
            done, total, result = msg["done"], msg["total"], msg["result"]
            self.progress.set(done / total if total else 0)
            self.progress_label.configure(
                text="正在识别：{} / {}（预计剩余 {}）".format(
                    done, total, self._eta(done, total)))
            self._insert_result_row(result)

        elif mtype == "analyze_done":
            results = msg["results"]
            self.analyze_results = results
            self.match_results = [r for r in results if r["status"] == "match"]
            n_match = len(self.match_results)
            n_no = sum(1 for r in results if r["status"] == "no_match")
            n_fail = sum(1 for r in results if r["status"] == "failed")
            self.progress.set(1)
            elapsed = int(time.time() - (self.t_start or time.time()))
            self.progress_label.configure(text="识别完成，用时 {} 秒".format(elapsed))
            self.status_label.configure(text=(
                "预览结果：可重命名 {} · 未匹配 {} · 失败 {}（共 {} 个）"
            ).format(n_match, n_no, n_fail, len(results)))
            if n_match == 0:
                self._set_state("idle")
                messagebox.showinfo(APP_NAME, "没有匹配到规则的文件，无需执行重命名。")
            else:
                self._set_state("preview_ready")

        elif mtype == "exec_progress":
            done, total, item = msg["done"], msg["total"], msg["item"]
            self.progress.set(done / total if total else 0)
            self.progress_label.configure(text="正在重命名：{} / {}".format(done, total))
            if item:
                self._mark_row_renamed(item["base_name"], item["new_name"])

        elif mtype == "exec_done":
            renamed, skipped = msg["renamed"], msg["skipped"]
            elapsed = int(time.time() - (self.t_start or time.time()))
            self.progress.set(1)
            self.progress_label.configure(text="执行完成，用时 {} 秒".format(elapsed))
            self.status_label.configure(text=(
                "重命名 {} 个 · 跳过 {} 个 · 用时 {} 秒"
            ).format(len(renamed), len(skipped), elapsed))
            self._set_state("idle")
            messagebox.showinfo(APP_NAME, "重命名完成：成功 {} 个，跳过 {} 个。".format(
                len(renamed), len(skipped)))

        elif mtype == "split_analyze_progress":
            done, total, _info = msg["done"], msg["total"], msg["info"]
            self.progress.set(done / total if total else 0)
            elapsed = int(time.time() - (self.t_start or time.time()))
            self.progress_label.configure(
                text="正在分析拆分：第 {} / {} 页（已用 {} 秒）".format(
                    done, total, elapsed))

        elif mtype == "split_analyze_stopped":
            self.progress_label.configure(text="分析已停止")
            self._set_state("idle")

        elif mtype == "split_analyze_done":
            analysis = msg["analysis"]
            self.split_analysis = analysis
            self.progress.set(1)
            elapsed = int(time.time() - (self.t_start or time.time()))
            self.progress_label.configure(text="拆分分析完成，用时 {} 秒".format(elapsed))

            # 填充预览表
            n_proj = len(analysis.get("projects", []))
            n_docs = analysis.get("total_docs", 0)
            n_unmatched = sum(1 for p in analysis.get("projects", [])
                              if p.get("no_folder"))
            self.status_label.configure(text=(
                "拆分预览：{} 个项目 · {} 个文档 · {} 个未匹配文件夹 · 用时 {} 秒"
            ).format(n_proj, n_docs, n_unmatched, elapsed))
            warnings = analysis.get("warnings", [])
            if warnings:
                self.progress_label.configure(
                    text="拆分分析完成，用时 {} 秒（注意：{}）".format(
                        elapsed, warnings[0]))

            self._render_split_preview(analysis)
            self._set_state("preview_ready")

            if not analysis.get("projects"):
                self._set_state("idle")
                messagebox.showinfo(APP_NAME, "未识别到任何项目（未发现开工报告页）。")
            elif n_docs == 0:
                self._set_state("idle")
                messagebox.showinfo(APP_NAME, "未识别到可拆分的文档。")

        elif mtype == "split_exec_progress":
            done, total, item = msg["done"], msg["total"], msg["item"]
            self.progress.set(done / total if total else 0)
            self.progress_label.configure(text="正在拆分：{} / {}".format(done, total))

        elif mtype == "split_exec_done":
            written, skipped = msg["written"], msg["skipped"]
            elapsed = int(time.time() - (self.t_start or time.time()))
            self.progress.set(1)
            self.progress_label.configure(text="拆分完成，用时 {} 秒".format(elapsed))
            self.status_label.configure(text=(
                "拆分完成：成功 {} 个 · 跳过 {} 个 · 用时 {} 秒"
            ).format(len(written), len(skipped), elapsed))
            self._set_state("idle")
            summary = "拆分完成：成功 {} 个文档，跳过 {} 个。\n\n".format(
                len(written), len(skipped))
            if written:
                summary += "已写入以下文件夹（示例前 3 个）：\n"
                for w in written[:3]:
                    summary += "• {}\n".format(os.path.dirname(w["dst"]))
                if len(written) > 3:
                    summary += "… 共 {} 个\n".format(len(written))
            summary += ("\n原整份 PDF 保留不变。可用「撤销上次拆分」删除本次写出的文件，"
                        "或切换到「重命名文件」模式继续处理拆出的文档。")
            messagebox.showinfo(APP_NAME, summary)

        elif mtype == "undo_done":
            restored, failed = msg["restored"], msg["failed"]
            self.progress.set(0)
            self.progress_label.configure(text="撤销完成")
            if failed:
                detail = "\n".join("{}：{}".format(f["file"], f["info"])
                                   for f in failed[:10])
                self.status_label.configure(text="撤销：恢复 {} 个，失败 {} 个".format(
                    restored, len(failed)))
                messagebox.showwarning(APP_NAME,
                                       "撤销完成：恢复 {} 个，失败 {} 个。\n\n{}".format(
                                           restored, len(failed), detail))
            else:
                self.status_label.configure(text="撤销完成：已恢复 {} 个文件的原文件名".format(restored))
                messagebox.showinfo(APP_NAME, "已恢复 {} 个文件的原文件名。".format(restored))
            self._set_state("idle")

        elif mtype == "undo_split_done":
            deleted, skipped = msg["deleted"], msg["skipped"]
            self.progress.set(0)
            self.progress_label.configure(text="撤销拆分完成")
            if skipped:
                detail = "\n".join("{}：{}".format(s["file"], s["info"])
                                   for s in skipped[:10])
                self.status_label.configure(
                    text="撤销拆分：删除 {} 个，跳过 {} 个".format(deleted, len(skipped)))
                messagebox.showwarning(
                    APP_NAME,
                    "撤销拆分完成：已删除 {} 个文件，跳过 {} 个。\n\n{}".format(
                        deleted, len(skipped), detail))
            else:
                self.status_label.configure(
                    text="撤销拆分完成：已删除 {} 个文件".format(deleted))
                messagebox.showinfo(
                    APP_NAME, "已删除本次拆分新写出的 {} 个文件。\n"
                              "原整份 PDF 与其它文件未受影响。".format(deleted))
            self._set_state("idle")

        elif mtype == "error":
            self.progress_label.configure(text="出错")
            self._set_state("idle")
            messagebox.showerror(APP_NAME, msg["msg"])

    # ============================================================
    # 表格操作
    # ============================================================

    def _render_split_preview(self, analysis):
        """在结果表中渲染拆分预览：
        每项目一行摘要 + 该项目下每个文档一行
        """
        self.tree.delete(*self.tree.get_children())
        root = analysis.get("root", "")
        for proj in analysis.get("projects", []):
            name = proj.get("name") or "(未识别工程名称)"
            folder = proj.get("matched_folder")
            if folder:
                folder_txt = "✓ {}".format(folder)
            else:
                folder_txt = "⚠ 未匹配 → 未匹配项目/" + (proj.get("name") or "未知工程")
            summary = "【项目】{} ｜ {} ｜ {} 页".format(
                name, folder_txt, proj.get("page_range", "-"))
            self.tree.insert(
                "", "end",
                values=("项目", summary, "", "P{}-P{}".format(
                    proj.get("start_page", 0), proj.get("end_page", 0))),
                tags=("match" if folder else "no_match",))
            for doc in proj.get("docs", []):
                mapping = doc.get("mapping") or "未识别标题"
                pages = doc.get("pages", "")
                file_name = doc.get("file_name") or (mapping + ".pdf")
                if folder:
                    target = os.path.join(root, folder, file_name)
                else:
                    target = os.path.join(
                        root, "未匹配项目", (proj.get("name") or "未知工程"),
                        file_name)
                self.tree.insert(
                    "", "end", values=("文档", mapping, pages, target),
                    tags=("match" if folder else "no_match",))
            for w in analysis.get("warnings", []):
                self.tree.insert("", "end", values=("警告", w, "", ""),
                                 tags=("no_match",))

    _STATUS_TEXT = {"match": "待改名", "no_match": "未匹配",
                    "failed": "失败", "cancelled": "已取消"}

    def _insert_result_row(self, result):
        status = result["status"]
        title = (result.get("title") or "").replace("\n", " ")[:40]
        new_name = result.get("new_name") or ""
        if status == "failed":
            new_name = result.get("error", "")[:60]
        elif status == "no_match":
            new_name = "（保持原名）"
        self.tree.insert("", "end", values=(
            self._STATUS_TEXT.get(status, status),
            result["base_name"], title, new_name), tags=(status,))

    def _mark_row_renamed(self, base_name, new_name):
        """把预览表中对应行状态改为「已改名」"""
        for item in self.tree.get_children(""):
            vals = self.tree.item(item, "values")
            if vals and vals[1] == base_name:
                self.tree.item(item, values=("已改名", vals[1], vals[2], new_name),
                               tags=("renamed",))

    def _eta(self, done, total):
        """根据已耗时估算剩余时间"""
        if done <= 0:
            return "计算中…"
        elapsed = time.time() - (self.t_start or time.time())
        remain = elapsed / done * (total - done)
        if remain < 60:
            return "{:.0f} 秒".format(remain)
        return "{:.0f} 分 {:.0f} 秒".format(remain // 60, remain % 60)


def run():
    app = App()
    app.mainloop()
