# -*- coding: utf-8 -*-
"""规则管理弹窗：表格内增删改映射规则，替代手动编辑 Excel

交互设计（选中 → 编辑 → 保存）：
- 点击表格行：内容加载到下方输入框
- 修改输入框后点「保存修改」：更新选中行
- 「新增」：把输入框内容追加为新规则
- 「删除选中」：删除选中行
- 「导入 Excel…」：从外部 Excel 覆盖导入（带回滚保护）
- 「保存并关闭」：写回规则 Excel 文件
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core import rules as rules_mod
import theme

# 按钮统一样式（与主窗口一致，配色来自 theme.py，TraeWork 风格浅色主题）
BTN_FG = theme.BTN_FG
BTN_HOVER = theme.BTN_HOVER
BTN_TEXT = theme.BTN_TEXT


class RulesDialog(ctk.CTkToplevel):
    """规则管理弹窗"""

    def __init__(self, master, rules, on_saved):
        super().__init__(master)
        self.title("规则管理")
        self.geometry("720x560")
        self.minsize(640, 480)
        self.transient(master)      # 跟随主窗口
        self.grab_set()             # 模态：关闭前锁定主窗口

        self.work_rules = [dict(r) for r in rules]   # 编辑副本
        self.on_saved = on_saved
        self.selected_iid = None

        self._build_ui()
        self._refresh_table()
        self.after(100, self.focus_force)

    # ============================================================
    # UI
    # ============================================================

    def _build_ui(self):
        theme.style_treeview()
        self.configure(fg_color=theme.PAGE_BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ---------- 规则表格（卡片） ----------
        box = ctk.CTkFrame(self, fg_color=theme.CARD_BG, corner_radius=14,
                           border_width=1, border_color=theme.CARD_BORDER)
        box.grid(row=0, column=0, sticky="nsew", padx=16, pady=(16, 8))
        box.grid_rowconfigure(0, weight=1)
        box.grid_columnconfigure(0, weight=1)

        columns = ("key", "val")
        self.tree = ttk.Treeview(box, columns=columns, show="headings",
                                 selectmode="browse")
        self.tree.heading("key", text="识别内容（OCR 标题包含的关键词）")
        self.tree.heading("val", text="映射内容（新文件名主体）")
        self.tree.column("key", width=330, minwidth=200)
        self.tree.column("val", width=280, minwidth=160)
        vsb = ctk.CTkScrollbar(box, orientation="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=8)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ---------- 编辑区（卡片） ----------
        edit = ctk.CTkFrame(self, fg_color=theme.CARD_BG, corner_radius=14,
                            border_width=1, border_color=theme.CARD_BORDER)
        edit.grid(row=1, column=0, sticky="ew", padx=16, pady=8)
        edit.grid_columnconfigure(1, weight=1)
        edit.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(edit, text="识别内容",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=(16, 8), pady=(12, 4), sticky="w")
        self.key_entry = ctk.CTkEntry(
            edit, placeholder_text="如：配电网工程开工报告",
            fg_color=theme.INPUT_BG, border_color=theme.INPUT_BORDER,
            border_width=1, placeholder_text_color=theme.SUBTEXT)
        self.key_entry.grid(row=0, column=1, columnspan=3, padx=4, pady=(12, 4), sticky="ew")

        ctk.CTkLabel(edit, text="映射内容",
                     font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=(16, 8), pady=4, sticky="w")
        self.val_entry = ctk.CTkEntry(
            edit, placeholder_text="如：开工报告（留空则匹配后跳过不改名）",
            fg_color=theme.INPUT_BG, border_color=theme.INPUT_BORDER,
            border_width=1, placeholder_text_color=theme.SUBTEXT)
        self.val_entry.grid(row=1, column=1, columnspan=3, padx=4, pady=4, sticky="ew")

        btns = ctk.CTkFrame(edit, fg_color="transparent")
        btns.grid(row=2, column=0, columnspan=4, sticky="ew", padx=16, pady=(8, 16))
        ctk.CTkButton(btns, text="新增", width=90,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._add_rule).grid(row=0, column=0, padx=4)
        ctk.CTkButton(btns, text="保存修改", width=90,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._update_rule).grid(row=0, column=1, padx=4)
        ctk.CTkButton(btns, text="删除选中", width=90,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._delete_rule).grid(row=0, column=2, padx=4)
        ctk.CTkButton(btns, text="导入 Excel…（覆盖现有）", width=160,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._import_excel).grid(row=0, column=3, padx=4)

        # ---------- 底部 ----------
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", padx=16, pady=(4, 16))
        self.info_label = ctk.CTkLabel(bar, text="", anchor="w",
                                       text_color=theme.SUBTEXT)
        self.info_label.grid(row=0, column=0, sticky="ew", padx=4)
        bar.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(bar, text="保存并关闭", width=110,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self._save).grid(row=0, column=1, padx=4)
        ctk.CTkButton(bar, text="取消", width=80,
                      fg_color=BTN_FG, hover_color=BTN_HOVER,
                      text_color=BTN_TEXT, font=ctk.CTkFont(weight="bold"),
                      command=self.destroy).grid(row=0, column=2, padx=4)

    # ============================================================
    # 表格与编辑操作
    # ============================================================

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, rule in enumerate(self.work_rules):
            self.tree.insert("", "end", iid=str(i),
                             values=(rule.get("识别内容", ""),
                                     rule.get("映射内容", "")))
        self.info_label.configure(text="共 {} 条规则".format(len(self.work_rules)))

    def _on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            self.selected_iid = None
            return
        self.selected_iid = sel[0]
        vals = self.tree.item(self.selected_iid, "values")
        self.key_entry.delete(0, "end")
        self.key_entry.insert(0, vals[0] if vals else "")
        self.val_entry.delete(0, "end")
        self.val_entry.insert(0, vals[1] if len(vals) > 1 else "")

    def _add_rule(self):
        key = self.key_entry.get().strip()
        val = self.val_entry.get().strip()
        if not key:
            messagebox.showwarning("规则管理", "「识别内容」不能为空", parent=self)
            return
        self.work_rules.append({"识别内容": key, "映射内容": val})
        self._refresh_table()
        # 选中新加的行
        last = str(len(self.work_rules) - 1)
        self.tree.selection_set(last)
        self.tree.see(last)

    def _update_rule(self):
        if not self.selected_iid:
            messagebox.showwarning("规则管理", "请先在表格中点击要修改的行", parent=self)
            return
        key = self.key_entry.get().strip()
        val = self.val_entry.get().strip()
        if not key:
            messagebox.showwarning("规则管理", "「识别内容」不能为空", parent=self)
            return
        i = int(self.selected_iid)
        self.work_rules[i] = {"识别内容": key, "映射内容": val}
        self._refresh_table()
        self.tree.selection_set(str(i))
        self.tree.see(str(i))

    def _delete_rule(self):
        if not self.selected_iid:
            messagebox.showwarning("规则管理", "请先在表格中点击要删除的行", parent=self)
            return
        i = int(self.selected_iid)
        del self.work_rules[i]
        self.selected_iid = None
        self._refresh_table()

    def _import_excel(self):
        path = filedialog.askopenfilename(
            title="选择规则 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            parent=self)
        if not path:
            return
        try:
            new_rules = rules_mod.import_rules(path)
        except Exception as e:
            messagebox.showerror("规则管理", "导入失败：{}".format(e), parent=self)
            return
        if not messagebox.askyesno(
                "规则管理",
                "从该 Excel 读取到 {} 条规则，将覆盖当前编辑的 {} 条。\n确定覆盖吗？".format(
                    len(new_rules), len(self.work_rules)),
                parent=self):
            return
        self.work_rules = new_rules
        self._refresh_table()
        self.info_label.configure(text="已导入 {} 条规则（来自 {}）".format(
            len(new_rules), path.split("/")[-1].split("\\")[-1]))

    def _save(self):
        # 过滤掉识别内容为空的行
        clean = [r for r in self.work_rules if str(r.get("识别内容", "")).strip()]
        if not clean:
            if not messagebox.askyesno(
                    "规则管理", "当前规则为空，保存后需重新导入或添加规则才能识别。\n确定保存吗？",
                    parent=self):
                return
        try:
            rules_mod.save_rules(rules_mod.paths.get_rules_file(), clean)
        except Exception as e:
            messagebox.showerror("规则管理", "保存失败：{}".format(e), parent=self)
            return
        if self.on_saved:
            self.on_saved(clean)
        self.destroy()
