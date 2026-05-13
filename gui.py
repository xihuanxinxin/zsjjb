from __future__ import annotations

import json
import os
import re
import sys
import calendar
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, date, timedelta
from decimal import Decimal

from data_manager import (
    AssetDataManager,
    BILLING_DAYS_COLUMN,
    round_money_half_up,
    sum_parsed_money,
)
import sheet_normalize as sheet_norm
import excel_calc_workbook as excel_calc_wb

class DatabaseAssetGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("资产管理系统 - 数据库资产")
        self.root.geometry("1440x820")
        self.root.minsize(1100, 640)

        self._setup_styles()

        self.data_manager = AssetDataManager()

        self.status_var = tk.StringVar(value="就绪")
        self._normalize_summaries: list = []
        self._normalize_written_paths: list = []
        self._normalize_out_dir = ""
        self._normalize_input_path = ""
        self._norm_formulas_only_mode = False
        self._norm_per_sheet_calc: list = []  # [(sheet_name, df公式后含可选「数据块」列), ...]
        self._last_excel_grid_cache_dir = ""
        self.var_bill_year = tk.IntVar(value=2026)
        self.var_bill_month = tk.IntVar(value=4)
        self.create_widgets()
        self.refresh_table()
        self.root.after(200, self._init_assets_paned_sash)
        self.root.after(550, self._sync_group_tree_height)

    def _init_assets_paned_sash(self):
        """初次把分割条下移，让下方（含按列「服务名称」汇总表等）占更大比例（约 60% 高度给下半区）。"""
        try:
            h = self._assets_paned.winfo_height()
            if h > 200:
                # sash 像素位置 = 上半区高度；数值越小下半区越大
                self._assets_paned.sashpos(0, int(h * 0.36))
        except (tk.TclError, AttributeError):
            pass
        self._sync_group_tree_height()

    def _sync_group_tree_height(self, _event=None):
        """按下方区域实际高度动态设置「服务名称」汇总表可见行数，避免只能看到一行。"""
        if not hasattr(self, "_lf_group") or not hasattr(self, "group_tree"):
            return
        try:
            h = self._lf_group.winfo_height()
        except tk.TclError:
            return
        if h <= 1:
            return
        hint_h = 44
        pad = 32
        rh = 26
        avail = h - hint_h - pad
        lines = max(10, min(120, avail // max(18, rh)))
        if lines != getattr(self, "_group_tree_height_lines", -1):
            self._group_tree_height_lines = lines
            self.group_tree.configure(height=lines)

    def _sync_main_scroll_inner(self, _event=None):
        """主界面外层垂直滚动：内层宽度随 Canvas，高度取内容/可视区较大者，便于 Notebook 正常扩展。"""
        c = getattr(self, "_main_scroll_canvas", None)
        inner = getattr(self, "_main_scroll_inner", None)
        win_id = getattr(self, "_main_scroll_win_id", None)
        if c is None or inner is None or win_id is None:
            return
        try:
            ch = c.winfo_height()
            cw = c.winfo_width()
        except tk.TclError:
            return
        if cw <= 1 or ch <= 1:
            return
        self.root.update_idletasks()
        try:
            req_h = inner.winfo_reqheight()
        except tk.TclError:
            return
        h = max(req_h, ch)
        cache = getattr(self, "_main_scroll_last_geom", None)
        if cache == (cw, h, req_h):
            return
        self._main_scroll_last_geom = (cw, h, req_h)
        c.itemconfigure(win_id, width=cw, height=h)
        c.configure(scrollregion=c.bbox("all"))

    def _on_main_scroll_mousewheel_all(self, event):
        """主区垂直滚动：在整块主界面内滚轮有效；表格/文本等自带滚动的控件不拦截。"""
        inner = getattr(self, "_main_scroll_inner", None)
        c = getattr(self, "_main_scroll_canvas", None)
        if inner is None or c is None:
            return
        w = event.widget
        try:
            if not w.winfo_exists():
                return
        except tk.TclError:
            return
        if isinstance(w, (ttk.Treeview, tk.Text, tk.Listbox)):
            return
        p = w
        while p is not None and p != inner:
            p = p.master if hasattr(p, "master") else None
        if p != inner:
            return
        try:
            c.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except tk.TclError:
            pass

    def _setup_styles(self):
        style = ttk.Style()
        if sys.platform == "win32":
            try:
                style.theme_use("vista")
            except tk.TclError:
                style.theme_use("clam")
        else:
            style.theme_use("clam")

        self._font_ui = ("Microsoft YaHei UI", 10)
        self._font_ui_sm = ("Microsoft YaHei UI", 9)
        self._font_title = ("Microsoft YaHei UI", 13, "bold")
        self._font_head = ("Microsoft YaHei UI", 10, "bold")

        style.configure(".", font=self._font_ui)
        style.configure("Title.TLabel", font=self._font_title)
        style.configure("Sub.TLabel", font=self._font_ui_sm, foreground="#5c5c5c")
        style.configure("TLabelframe.Label", font=self._font_head)
        style.configure("Danger.TButton", foreground="#b3261e", font=self._font_head)
        style.configure("Accent.TButton", font=self._font_head)
        style.configure("Treeview", font=self._font_ui_sm, rowheight=26)
        style.configure("Treeview.Heading", font=self._font_head)
        style.map(
            "Treeview",
            background=[("selected", "#1a73e8")],
            foreground=[("selected", "white")],
        )

    def create_widgets(self):
        content_outer = ttk.Frame(self.root)
        content_outer.pack(fill=tk.BOTH, expand=True)

        self._main_scroll_canvas = tk.Canvas(
            content_outer,
            highlightthickness=0,
            bd=0,
        )
        main_vsb = ttk.Scrollbar(
            content_outer,
            orient=tk.VERTICAL,
            command=self._main_scroll_canvas.yview,
        )
        self._main_scroll_canvas.configure(yscrollcommand=main_vsb.set)
        self._main_scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        body = ttk.Frame(self._main_scroll_canvas, padding=(12, 10, 12, 6))
        self._main_scroll_inner = body
        self._main_scroll_win_id = self._main_scroll_canvas.create_window(
            (0, 0), window=body, anchor="nw"
        )
        body.bind("<Configure>", self._sync_main_scroll_inner)
        self._main_scroll_canvas.bind("<Configure>", self._sync_main_scroll_inner)
        self.root.bind_all("<MouseWheel>", self._on_main_scroll_mousewheel_all)

        header = ttk.Frame(body)
        header.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(header, text="数据库资产", style="Title.TLabel").pack(side=tk.LEFT)
        ttk.Label(
            header,
            text="清洗载入后一键计费；合计费用优先应计费天，空则用使用天再核算天。",
            style="Sub.TLabel",
        ).pack(side=tk.LEFT, padx=(12, 0), pady=(4, 0))

        row_ops = ttk.Frame(body)
        row_ops.pack(fill=tk.X, pady=(0, 6))
        row_ops.columnconfigure(0, weight=1)
        row_ops.columnconfigure(1, weight=1)

        lf_file = ttk.LabelFrame(row_ops, text="数据与导出", padding=(8, 6))
        lf_file.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        f1 = ttk.Frame(lf_file)
        f1.pack(fill=tk.X)
        pad = dict(side=tk.LEFT, padx=4, pady=2)
        ttk.Button(
            f1,
            text="清洗、计算并载入…",
            command=self.normalize_summarize_dialog,
        ).pack(**pad)
        ttk.Button(
            f1,
            text="生成计算版 Excel…",
            command=self.export_calculated_workbook_copy,
        ).pack(**pad)
        ttk.Button(f1, text="导出 Excel…", command=self.export_excel).pack(**pad)

        lf_edit = ttk.LabelFrame(row_ops, text="资产记录", padding=(8, 6))
        lf_edit.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        f2 = ttk.Frame(lf_edit)
        f2.pack(fill=tk.X)
        pad2 = dict(side=tk.LEFT, padx=3, pady=2)
        ttk.Button(f2, text="添加", command=self.add_asset_dialog).pack(**pad2)
        ttk.Button(f2, text="修改", command=self.update_asset_dialog).pack(**pad2)
        ttk.Button(f2, text="删除", command=self.delete_asset).pack(**pad2)
        ttk.Button(f2, text="批量删除", command=self.batch_delete_assets).pack(**pad2)
        ttk.Button(f2, text="清空全部", command=self.clear_all_data, style="Danger.TButton").pack(
            **pad2
        )
        ttk.Separator(f2, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Button(f2, text="刷新", command=self.refresh_table).pack(**pad2)
        ttk.Button(f2, text="统计报表", command=self.show_statistics).pack(**pad2)

        lf_calc = ttk.LabelFrame(body, text="计费计算（可逐项执行或一键完成）", padding=(8, 8))
        lf_calc.pack(fill=tk.X, pady=(0, 6))

        bill_row = ttk.Frame(lf_calc)
        bill_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(bill_row, text="账单月：").pack(side=tk.LEFT, padx=(0, 4))
        tk.Spinbox(
            bill_row,
            from_=2020,
            to=2035,
            width=6,
            textvariable=self.var_bill_year,
        ).pack(side=tk.LEFT)
        ttk.Label(bill_row, text="年").pack(side=tk.LEFT, padx=(2, 6))
        tk.Spinbox(
            bill_row,
            from_=1,
            to=12,
            width=4,
            textvariable=self.var_bill_month,
        ).pack(side=tk.LEFT)
        ttk.Label(bill_row, text="月").pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(
            bill_row,
            text="（清洗 Excel 时从文件名「YYYY年M月」自动识别；影响本期计费起止与天数）",
            style="Sub.TLabel",
        ).pack(side=tk.LEFT)

        calc_grid = ttk.Frame(lf_calc)
        calc_grid.pack(fill=tk.X)
        for c in range(4):
            calc_grid.columnconfigure(c, weight=1)

        calc_buttons = [
            ("计费开始时间", self.calculate_billing_start),
            ("本期计费开始", self.calculate_current_period_start),
            ("本期计费截止", self.calculate_current_period_end),
            ("本月核算天数", self.calculate_month_days),
            ("本月核算金额", self.calculate_month_amount),
            ("应计费天数", self.calculate_total_billing_days),
            ("合计费用", self.calculate_total_cost),
        ]
        for i, (text, cmd) in enumerate(calc_buttons):
            r, c = divmod(i, 4)
            ttk.Button(calc_grid, text=text, command=cmd).grid(
                row=r, column=c, padx=3, pady=3, sticky="ew"
            )

        ttk.Button(
            lf_calc,
            text="一键计算全部",
            command=self.calculate_all,
            style="Accent.TButton",
        ).pack(anchor=tk.E, pady=(6, 0))

        self.main_nb = ttk.Notebook(body)
        self.main_nb.pack(fill=tk.BOTH, expand=True)
        self.main_nb.bind(
            "<<NotebookTabChanged>>",
            lambda _e: self.root.after_idle(self._sync_main_scroll_inner),
        )

        self.tab_assets = ttk.Frame(self.main_nb, padding=(0, 6, 0, 0))
        self.main_nb.add(self.tab_assets, text="资产管理")
        self.tab_assets.columnconfigure(0, weight=1)
        self.tab_assets.rowconfigure(0, weight=1)

        self._assets_paned = ttk.Panedwindow(self.tab_assets, orient=tk.VERTICAL)
        self._assets_paned.grid(row=0, column=0, sticky="nsew")
        pane_top = ttk.Frame(self._assets_paned)
        pane_bottom = ttk.Frame(self._assets_paned)
        self._assets_paned.add(pane_top, weight=3)
        self._assets_paned.add(pane_bottom, weight=5)

        search_frame = ttk.LabelFrame(pane_top, text="搜索", padding=(8, 6))
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.search_fields = {}
        search_items = [
            ("服务目录编号", "服务目录编号"),
            ("服务类型", "服务类型"),
            ("服务名称", "服务名称"),
            ("唯一编号", "唯一编号"),
        ]

        sf = ttk.Frame(search_frame)
        sf.pack(fill=tk.X)
        for i, (label, key) in enumerate(search_items):
            ttk.Label(sf, text=label).grid(row=0, column=i * 2, padx=(4, 2), pady=4, sticky=tk.W)
            entry = ttk.Entry(sf, width=18)
            entry.grid(row=0, column=i * 2 + 1, padx=(0, 12), pady=4, sticky=tk.W)
            self.search_fields[key] = entry

        col_btn = len(search_items) * 2
        ttk.Button(sf, text="搜索", command=self.search_assets, width=10).grid(
            row=0, column=col_btn, padx=4, pady=4
        )
        ttk.Button(sf, text="重置", command=self.reset_search, width=10).grid(
            row=0, column=col_btn + 1, padx=4, pady=4
        )

        pane_top.rowconfigure(1, weight=1)
        pane_top.columnconfigure(0, weight=1)

        table_wrap = ttk.Frame(pane_top)
        table_wrap.grid(row=1, column=0, sticky="nsew", pady=(0, 0))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_wrap, show="headings", selectmode="extended")
        self.tree["columns"] = self.data_manager.asset_types["database"]["columns"]

        col_widths = {
            "唯一编号": 80,
            "服务目录编号": 100,
            "服务类型": 100,
            "服务名称": 120,
            "数量": 60,
            "单位": 80,
            "单价": 100,
            "备注": 120,
            "开通时间": 100,
            "关停时间": 100,
            "计费开始时间": 120,
            "本期计费开始时间": 120,
            "本期计费截至日期": 120,
            "本月核算天数": 80,
            "使用天数": 80,
            "本月核算金额": 120,
            BILLING_DAYS_COLUMN: 140,
            "合计费用": 120,
        }

        for col in self.data_manager.asset_types["database"]["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)

        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar = ttk.Scrollbar(
            table_wrap, orient=tk.VERTICAL, command=self.tree.yview
        )
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar = ttk.Scrollbar(
            table_wrap, orient=tk.HORIZONTAL, command=self.tree.xview
        )
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(
            yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set
        )

        self.tree.tag_configure("gray_row", background="#eef2f7")

        pane_bottom.rowconfigure(0, weight=1)
        pane_bottom.columnconfigure(0, weight=1)

        lf_group = ttk.LabelFrame(
            pane_bottom,
            text="按标准列「服务名称」汇总小计（列名与 数据库资产.csv 一致；仅界面，不写入 CSV） · 可拖动中间分隔条调整上下区域",
            padding=(8, 6),
        )
        self._lf_group = lf_group
        lf_group.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
        lf_group.columnconfigure(0, weight=1)
        lf_group.rowconfigure(0, weight=1)
        gcols = ("svc_name", "nrows", "sum_month", "sum_cost")
        self.group_tree = ttk.Treeview(
            lf_group,
            columns=gcols,
            show="headings",
            height=14,
        )
        self.group_tree.heading("svc_name", text="服务名称")
        self.group_tree.heading("nrows", text="明细行数")
        self.group_tree.heading("sum_month", text="本月核算金额小计")
        self.group_tree.heading("sum_cost", text="合计费用小计")
        self.group_tree.column("svc_name", width=280, anchor=tk.W)
        self.group_tree.column("nrows", width=88, anchor=tk.CENTER)
        self.group_tree.column("sum_month", width=160, anchor=tk.E)
        self.group_tree.column("sum_cost", width=160, anchor=tk.E)
        gvs = ttk.Scrollbar(lf_group, orient=tk.VERTICAL, command=self.group_tree.yview)
        self.group_tree.configure(yscrollcommand=gvs.set)
        self.group_tree.grid(row=0, column=0, sticky="nsew")
        gvs.grid(row=0, column=1, sticky="ns")
        ttk.Label(
            lf_group,
            text="有搜索条件时：按列「服务名称」的汇总小计与下方「总计」均只对当前筛选结果；重置后为全库。",
            style="Sub.TLabel",
            wraplength=880,
            justify=tk.LEFT,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        lf_group.bind("<Configure>", self._sync_group_tree_height)
        self._assets_paned.bind("<Configure>", self._sync_group_tree_height)

        self._sum_banner_wrap = ttk.LabelFrame(
            pane_bottom, text="总计金额（界面汇总，不写入 CSV）", padding=(12, 10)
        )
        self._sum_banner_wrap.grid(row=1, column=0, sticky="ew", pady=(0, 0))
        bf = ttk.Frame(self._sum_banner_wrap)
        bf.pack(fill=tk.X)
        tb = ("Microsoft YaHei UI", 12, "bold")
        ttk.Label(bf, text="本月核算金额合计", font=self._font_ui).pack(side=tk.LEFT)
        self._lbl_total_month = ttk.Label(
            bf, text="￥0.00", font=tb, foreground="#b45309"
        )
        self._lbl_total_month.pack(side=tk.LEFT, padx=(6, 28))
        ttk.Label(bf, text="合计费用总计", font=self._font_ui).pack(side=tk.LEFT)
        self._lbl_total_cost = ttk.Label(
            bf, text="￥0.00", font=tb, foreground="#b45309"
        )
        self._lbl_total_cost.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(
            self._sum_banner_wrap,
            text="对当前统计范围内全部明细行的「本月核算金额」「合计费用」求和；不写入 CSV。",
            style="Sub.TLabel",
            wraplength=880,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(8, 0))

        self.tab_normalize = ttk.Frame(self.main_nb, padding=(8, 6))
        self.main_nb.add(self.tab_normalize, text="各 Sheet 汇总")
        self.tab_normalize.columnconfigure(0, weight=1)
        self.tab_normalize.rowconfigure(5, weight=2)
        self.tab_normalize.rowconfigure(6, weight=1)
        self.tab_normalize.rowconfigure(7, weight=1)
        self.tab_normalize.rowconfigure(8, weight=2)
        self.tab_normalize.rowconfigure(10, weight=1)

        self._norm_hint = ttk.Label(
            self.tab_normalize,
            text=(
                "请先顶栏「清洗、计算并载入」。下方为各 Sheet 公式重算后的小计；"
                "点行可看数据块小计、按服务名称汇总与逐行明细。"
            ),
            style="Sub.TLabel",
            wraplength=920,
            justify=tk.LEFT,
        )
        self._norm_hint.grid(row=0, column=0, sticky="w", pady=(0, 4))
        self._norm_src_label = ttk.Label(
            self.tab_normalize, text="源文件：—", style="Sub.TLabel"
        )
        self._norm_src_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        self._norm_out_label = ttk.Label(
            self.tab_normalize, text="（清洗载入后仅保存 数据库资产.csv）", style="Sub.TLabel"
        )
        self._norm_out_label.grid(row=2, column=0, sticky="w", pady=(2, 2))
        self._norm_grand_label = ttk.Label(
            self.tab_normalize, text="全部合计：—", font=self._font_ui
        )
        self._norm_grand_label.grid(row=3, column=0, sticky="w", pady=(0, 4))

        lf_amt_note = ttk.LabelFrame(self.tab_normalize, text="金额是怎么来的", padding=(8, 6))
        lf_amt_note.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            lf_amt_note,
            text=(
                "清洗载入后：明细写入「资产管理」并按规则重算日期与金额；"
                "本月核算金额＝数量×单价×本月核算天数；"
                "合计费用＝数量×单价×（应计费天数→使用天数→本月核算天数）。"
                "本页列表为各 Sheet 公式重算后的小计，点行可看块小计与按服务名称汇总。"
            ),
            style="Sub.TLabel",
            wraplength=880,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        lf_ns = ttk.LabelFrame(
            self.tab_normalize, text="各 Sheet / 文件（金额一览）", padding=8
        )
        lf_ns.grid(row=5, column=0, sticky="nsew", pady=(0, 6))
        lf_ns.columnconfigure(0, weight=1)
        lf_ns.rowconfigure(0, weight=1)
        ncols = ("sheet", "detail_csv", "rows", "month", "cost")
        self.norm_tree_sheet = ttk.Treeview(
            lf_ns,
            columns=ncols,
            show="headings",
            height=10,
            selectmode="browse",
        )
        self.norm_tree_sheet.heading("sheet", text="Sheet 或 CSV 名称")
        self.norm_tree_sheet.heading("detail_csv", text="口径 / 路径")
        self.norm_tree_sheet.heading("rows", text="明细行数")
        self.norm_tree_sheet.heading("month", text="本月核算金额总计")
        self.norm_tree_sheet.heading("cost", text="合计费用总计")
        self.norm_tree_sheet.column("sheet", width=220, anchor=tk.W)
        self.norm_tree_sheet.column("detail_csv", width=260, anchor=tk.W)
        self.norm_tree_sheet.column("rows", width=72, anchor=tk.CENTER)
        self.norm_tree_sheet.column("month", width=150, anchor=tk.E)
        self.norm_tree_sheet.column("cost", width=150, anchor=tk.E)
        vs_n1 = ttk.Scrollbar(lf_ns, orient=tk.VERTICAL, command=self.norm_tree_sheet.yview)
        self.norm_tree_sheet.configure(yscrollcommand=vs_n1.set)
        self.norm_tree_sheet.grid(row=0, column=0, sticky="nsew")
        vs_n1.grid(row=0, column=1, sticky="ns")

        lf_nb = ttk.LabelFrame(
            self.tab_normalize,
            text="数据块汇总（点击上表某 Sheet 后，此处为公式重算后按块小计）",
            padding=8,
        )
        self._lf_norm_block = lf_nb
        lf_nb.grid(row=6, column=0, sticky="nsew", pady=(0, 6))
        lf_nb.columnconfigure(0, weight=1)
        lf_nb.rowconfigure(0, weight=1)
        bcols = ("block", "brows", "bmonth", "bcost")
        self.norm_tree_block = ttk.Treeview(
            lf_nb, columns=bcols, show="headings", height=8
        )
        self.norm_tree_block.heading("block", text="数据块标题")
        self.norm_tree_block.heading("brows", text="行数")
        self.norm_tree_block.heading("bmonth", text="本月核算金额合计")
        self.norm_tree_block.heading("bcost", text="合计费用合计")
        self.norm_tree_block.column("block", width=400, anchor=tk.W)
        self.norm_tree_block.column("brows", width=72, anchor=tk.CENTER)
        self.norm_tree_block.column("bmonth", width=170, anchor=tk.E)
        self.norm_tree_block.column("bcost", width=170, anchor=tk.E)
        vs_n2 = ttk.Scrollbar(lf_nb, orient=tk.VERTICAL, command=self.norm_tree_block.yview)
        self.norm_tree_block.configure(yscrollcommand=vs_n2.set)
        self.norm_tree_block.grid(row=0, column=0, sticky="nsew")
        vs_n2.grid(row=0, column=1, sticky="ns")
        self.norm_tree_sheet.bind("<<TreeviewSelect>>", self._on_norm_sheet_select)

        lf_ng = ttk.LabelFrame(
            self.tab_normalize,
            text="当前 Sheet：按标准列「服务名称」汇总（列名与主表一致；公式重算后）",
            padding=8,
        )
        lf_ng.grid(row=7, column=0, sticky="nsew", pady=(0, 6))
        lf_ng.columnconfigure(0, weight=1)
        lf_ng.rowconfigure(0, weight=1)
        sgcols = ("svc_name", "nrows", "sum_month", "sum_cost")
        self.norm_tree_svc_group = ttk.Treeview(
            lf_ng,
            columns=sgcols,
            show="headings",
            height=8,
            selectmode="browse",
        )
        self.norm_tree_svc_group.heading("svc_name", text="服务名称")
        self.norm_tree_svc_group.heading("nrows", text="明细行数")
        self.norm_tree_svc_group.heading("sum_month", text="本月核算金额小计")
        self.norm_tree_svc_group.heading("sum_cost", text="合计费用小计")
        self.norm_tree_svc_group.column("svc_name", width=280, anchor=tk.W)
        self.norm_tree_svc_group.column("nrows", width=88, anchor=tk.CENTER)
        self.norm_tree_svc_group.column("sum_month", width=160, anchor=tk.E)
        self.norm_tree_svc_group.column("sum_cost", width=160, anchor=tk.E)
        vs_ng = ttk.Scrollbar(
            lf_ng, orient=tk.VERTICAL, command=self.norm_tree_svc_group.yview
        )
        self.norm_tree_svc_group.configure(yscrollcommand=vs_ng.set)
        self.norm_tree_svc_group.grid(row=0, column=0, sticky="nsew")
        vs_ng.grid(row=0, column=1, sticky="ns")

        lf_nd = ttk.LabelFrame(
            self.tab_normalize,
            text="本 Sheet 明细行（公式重算后与「资产管理」同一套规则；多列可横向滚动）",
            padding=8,
        )
        lf_nd.grid(row=8, column=0, sticky="nsew", pady=(0, 6))
        lf_nd.columnconfigure(0, weight=1)
        lf_nd.rowconfigure(0, weight=1)
        dcols = tuple(self.data_manager.asset_types["database"]["columns"])
        self.norm_tree_sheet_detail = ttk.Treeview(
            lf_nd,
            columns=dcols,
            show="headings",
            height=10,
            selectmode="browse",
        )
        _dw = {
            "唯一编号": 80,
            "服务目录编号": 100,
            "服务类型": 100,
            "服务名称": 120,
            "数量": 60,
            "单位": 80,
            "单价": 100,
            "备注": 120,
            "开通时间": 100,
            "关停时间": 100,
            "计费开始时间": 120,
            "本期计费开始时间": 120,
            "本期计费截至日期": 120,
            "本月核算天数": 80,
            "使用天数": 80,
            "本月核算金额": 120,
            BILLING_DAYS_COLUMN: 140,
            "合计费用": 120,
        }
        for col in dcols:
            self.norm_tree_sheet_detail.heading(col, text=col)
            self.norm_tree_sheet_detail.column(
                col, width=_dw.get(col, 100), anchor=tk.CENTER
            )
        vs_nd = ttk.Scrollbar(
            lf_nd, orient=tk.VERTICAL, command=self.norm_tree_sheet_detail.yview
        )
        hs_nd = ttk.Scrollbar(
            lf_nd, orient=tk.HORIZONTAL, command=self.norm_tree_sheet_detail.xview
        )
        self.norm_tree_sheet_detail.configure(
            yscrollcommand=vs_nd.set, xscrollcommand=hs_nd.set
        )
        self.norm_tree_sheet_detail.grid(row=0, column=0, sticky="nsew")
        vs_nd.grid(row=0, column=1, sticky="ns")
        hs_nd.grid(row=1, column=0, sticky="ew")

        norm_btn_row = ttk.Frame(self.tab_normalize)
        norm_btn_row.grid(row=9, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(
            norm_btn_row,
            text="复制文字摘要到剪贴板",
            command=self._copy_normalize_summary,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            norm_btn_row,
            text="切回资产管理",
            command=lambda: self.main_nb.select(self.tab_assets),
        ).pack(side=tk.RIGHT)

        lf_nf = ttk.LabelFrame(self.tab_normalize, text="对应文件（完整路径）", padding=6)
        lf_nf.grid(row=10, column=0, sticky="nsew", pady=(0, 0))
        lf_nf.columnconfigure(0, weight=1)
        lf_nf.rowconfigure(0, weight=1)
        pcols = ("role", "abspath")
        self.norm_path_tree = ttk.Treeview(
            lf_nf,
            columns=pcols,
            show="headings",
            height=6,
        )
        self.norm_path_tree.heading("role", text="类型")
        self.norm_path_tree.heading("abspath", text="完整路径")
        self.norm_path_tree.column("role", width=100, anchor=tk.W)
        self.norm_path_tree.column("abspath", width=760, anchor=tk.W)
        ps_v = ttk.Scrollbar(lf_nf, orient=tk.VERTICAL, command=self.norm_path_tree.yview)
        ps_h = ttk.Scrollbar(lf_nf, orient=tk.HORIZONTAL, command=self.norm_path_tree.xview)
        self.norm_path_tree.configure(
            yscrollcommand=ps_v.set, xscrollcommand=ps_h.set
        )
        self.norm_path_tree.grid(row=0, column=0, sticky="nsew")
        ps_v.grid(row=0, column=1, sticky="ns")
        ps_h.grid(row=1, column=0, sticky="ew")

        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(10, 4),
            font=self._font_ui_sm,
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.root.after_idle(self._sync_main_scroll_inner)

    def _update_status(self):
        n = len(self.data_manager.data)
        self.status_var.set(f"共 {n} 条记录 · 数据库资产")

    def refresh_table(self):
        """刷新表格显示"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        columns = self.data_manager.asset_types["database"]["columns"]
        gray_cols = ["开通时间", "关停时间"]
        gray_col_indices = [i for i, col in enumerate(columns) if col in gray_cols]

        for _, row in self.data_manager.data.iterrows():
            values = []
            for i, col in enumerate(columns):
                val = row[col]
                if pd.isna(val):
                    values.append("")
                else:
                    values.append(str(val))

            has_gray_data = any(values[i] for i in gray_col_indices)
            tag = "gray_row" if has_gray_data else ""
            self.tree.insert("", tk.END, values=values, tags=(tag,))

        self._update_status()
        self._update_asset_totals_banner()
        self._update_service_name_grouping_from_df(self.data_manager.data)

    @staticmethod
    def _parse_money_value(v) -> float:
        return sheet_norm.parse_money(v)

    @staticmethod
    def _parse_billing_date(val) -> date | None:
        """解析开通/关停等日期；支持 2026/04/08 与 2026-04-08 00:00:00。"""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s0 = str(val).strip()
        if not s0 or s0.lower() in ("nan", "none", "!"):
            return None
        if len(s0) >= 10 and s0[4] in "-/" and s0[7] in "-/":
            s0 = s0[:10].replace("-", "/")
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(s0[:10], fmt).date()
            except ValueError:
                continue
        try:
            return datetime.strptime(str(val).strip(), "%Y年%m月%d日").date()
        except ValueError:
            return None

    def _service_name_group_agg(self, df: pd.DataFrame) -> list:
        """按列「服务名称」聚合：(展示名, 行数, 本月核算合计, 合计费用合计)，按合计费用降序。"""
        if df is None or len(df) == 0 or "服务名称" not in df.columns:
            return []
        agg_rows = []
        disp = df["服务名称"].fillna("").astype(str)
        keys = disp.map(excel_calc_wb._ledger_agg_service_name_key)
        tmp = df.copy()
        tmp["_aggk"] = keys
        tmp["_disp"] = disp
        for _k, g in tmp.groupby("_aggk", sort=False):
            names = [str(x).strip() for x in g["_disp"].tolist() if str(x).strip()]
            display_name = max(names, key=len) if names else "（空）"
            n = len(g)
            m = (
                sum_parsed_money(g["本月核算金额"])
                if "本月核算金额" in g.columns
                else 0.0
            )
            c = sum_parsed_money(g["合计费用"]) if "合计费用" in g.columns else 0.0
            agg_rows.append((display_name, n, m, c))
        agg_rows.sort(key=lambda x: -x[3])
        return agg_rows

    def _update_service_name_grouping_from_df(self, df: pd.DataFrame):
        """按标准列「服务名称」汇总本月核算金额、合计费用（仅界面）。"""
        if not hasattr(self, "group_tree"):
            return
        self.group_tree.delete(*self.group_tree.get_children())
        for display_name, n, m, c in self._service_name_group_agg(df):
            self.group_tree.insert(
                "",
                tk.END,
                values=(display_name, n, f"￥{m:,.2f}", f"￥{c:,.2f}"),
            )

    def _update_norm_sheet_service_grouping(self, df: pd.DataFrame):
        """Sheet 汇总页：当前选中 Sheet 按列「服务名称」汇总（公式后明细）。"""
        if not hasattr(self, "norm_tree_svc_group"):
            return
        self.norm_tree_svc_group.delete(*self.norm_tree_svc_group.get_children())
        for display_name, n, m, c in self._service_name_group_agg(df):
            self.norm_tree_svc_group.insert(
                "",
                tk.END,
                values=(display_name, n, f"￥{m:,.2f}", f"￥{c:,.2f}"),
            )

    def _update_asset_totals_banner(self):
        """根据当前统计范围内明细求和，更新底部汇总（与按「服务名称」汇总表同一范围）。"""
        if not hasattr(self, "_lbl_total_month"):
            return
        conditions = {}
        for key, entry in self.search_fields.items():
            v = entry.get().strip()
            if v:
                conditions[key] = v
        if conditions:
            scope = self.data_manager.search_assets(conditions)
        else:
            scope = self.data_manager.data
        m = (
            sum_parsed_money(scope["本月核算金额"])
            if len(scope) and "本月核算金额" in scope.columns
            else 0.0
        )
        c = (
            sum_parsed_money(scope["合计费用"])
            if len(scope) and "合计费用" in scope.columns
            else 0.0
        )
        self._lbl_total_month.configure(text=f"￥{m:,.2f}")
        self._lbl_total_cost.configure(text=f"￥{c:,.2f}")

    def search_assets(self):
        """搜索资产"""
        conditions = {}
        for key, entry in self.search_fields.items():
            value = entry.get().strip()
            if value:
                conditions[key] = value
        
        result = self.data_manager.search_assets(conditions)
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for _, row in result.iterrows():
            values = []
            for col in self.data_manager.asset_types["database"]["columns"]:
                val = row[col]
                if pd.isna(val):
                    values.append("")
                else:
                    values.append(str(val))
            self.tree.insert("", tk.END, values=values)

        self.status_var.set(
            f"筛选结果 {len(result)} 条（共 {len(self.data_manager.data)} 条在库）"
        )
        self._update_asset_totals_banner()
        self._update_service_name_grouping_from_df(result)

    def reset_search(self):
        """重置搜索条件"""
        for entry in self.search_fields.values():
            entry.delete(0, tk.END)
        self.refresh_table()
    
    def add_asset_dialog(self):
        """添加资产对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加数据库资产")
        dialog.geometry("620x640")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        
        entries = {}
        row = 0
        columns = self.data_manager.asset_types['database']['columns']
        
        for col in columns:
            if col == '唯一编号':
                continue
            
            ttk.Label(dialog, text=col).grid(row=row, column=0, padx=10, pady=5, sticky=tk.W)
            if col in ["备注"]:
                entry = tk.Text(dialog, width=40, height=3, font=self._font_ui_sm)
                entry.grid(row=row, column=1, padx=10, pady=5)
            else:
                entry = ttk.Entry(dialog, width=40)
                entry.grid(row=row, column=1, padx=10, pady=5)
            entries[col] = entry
            row += 1
        
        def submit():
            asset_info = {'唯一编号': self.data_manager.generate_unique_id()}
            for col in columns:
                if col == '唯一编号':
                    continue
                if col == '备注':
                    asset_info[col] = entries[col].get("1.0", tk.END).strip()
                else:
                    asset_info[col] = entries[col].get()
            
            if self.data_manager.add_asset(asset_info):
                messagebox.showinfo("成功", f"资产添加成功！\n唯一编号: {asset_info['唯一编号']}")
                dialog.destroy()
                self.refresh_table()
            else:
                messagebox.showerror("错误", "添加失败")
        
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=row, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="确定", command=submit).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
        
        dialog.grid_columnconfigure(1, weight=1)
    
    def update_asset_dialog(self):
        """修改资产对话框"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一条记录")
            return
        
        index = self.tree.index(selected[0])
        row = self.data_manager.data.iloc[index]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("修改数据库资产")
        dialog.geometry("620x640")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        
        entries = {}
        row_num = 0
        columns = self.data_manager.asset_types['database']['columns']
        
        for col in columns:
            ttk.Label(dialog, text=col).grid(row=row_num, column=0, padx=10, pady=5, sticky=tk.W)
            
            if col == "唯一编号":
                entry = ttk.Entry(dialog, width=40, state="readonly")
                entry.insert(0, str(row[col]))
            elif col == "备注":
                entry = tk.Text(dialog, width=40, height=3, font=self._font_ui_sm)
                entry.insert("1.0", str(row[col]))
            else:
                entry = ttk.Entry(dialog, width=40)
                entry.insert(0, str(row[col]))
            
            entry.grid(row=row_num, column=1, padx=10, pady=5)
            entries[col] = entry
            row_num += 1
        
        def submit():
            asset_info = {}
            for col in columns:
                if col == '备注':
                    asset_info[col] = entries[col].get("1.0", tk.END).strip()
                else:
                    asset_info[col] = entries[col].get()
            
            if self.data_manager.update_asset(index, asset_info):
                messagebox.showinfo("成功", "资产修改成功")
                dialog.destroy()
                self.refresh_table()
            else:
                messagebox.showerror("错误", "修改失败")
        
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=row_num, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="确定", command=submit).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
        
        dialog.grid_columnconfigure(1, weight=1)
    
    def delete_asset(self):
        """删除资产"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一条记录")
            return
        
        index = self.tree.index(selected[0])
        asset_id = self.data_manager.data.iloc[index]['唯一编号']
        
        if messagebox.askyesno("确认", f"确定要删除资产 {asset_id} 吗？"):
            if self.data_manager.delete_asset(index):
                messagebox.showinfo("成功", "资产删除成功")
                self.refresh_table()
            else:
                messagebox.showerror("错误", "删除失败")
    
    def clear_all_data(self):
        """清除所有数据"""
        if len(self.data_manager.data) == 0:
            messagebox.showinfo("提示", "当前没有数据可清除")
            return
        
        if messagebox.askyesno("危险操作", "⚠️ 确定要清除所有数据吗？\n\n此操作将删除所有资产记录，且无法恢复！"):
            if self.data_manager.clear_all_data():
                messagebox.showinfo("成功", "所有数据已清除")
                self.refresh_table()
            else:
                messagebox.showerror("错误", "清除失败")
    
    def batch_delete_assets(self):
        """批量删除选中的资产"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的记录")
            return
        
        count = len(selected)
        if messagebox.askyesno("确认批量删除", f"确定要删除选中的 {count} 条记录吗？"):
            # 获取选中行的索引（需要反向删除以避免索引错乱）
            indices = sorted([self.tree.index(item) for item in selected], reverse=True)
            
            delete_count = 0
            for index in indices:
                if self.data_manager.delete_asset(index):
                    delete_count += 1
            
            messagebox.showinfo("成功", f"已成功删除 {delete_count} 条记录")
            self.refresh_table()
    
    @staticmethod
    def parse_bill_month_from_filename(filename: str) -> tuple[int, int] | None:
        """从供应商表文件名解析账单月，如「2026年4月-安恒开通…」→ (2026, 4)。"""
        if not filename:
            return None
        m = re.search(r"(\d{4})年(\d{1,2})月", str(filename))
        if not m:
            return None
        y, mo = int(m.group(1)), int(m.group(2))
        if 2000 <= y <= 2100 and 1 <= mo <= 12:
            return y, mo
        return None

    def _bill_month_bounds(self) -> tuple[date, date]:
        """账单月首末日（自然月）。"""
        y = int(self.var_bill_year.get())
        mo = int(self.var_bill_month.get())
        mo = max(1, min(12, mo))
        last = calendar.monthrange(y, mo)[1]
        return date(y, mo, 1), date(y, mo, last)

    @staticmethod
    def _parse_datetime_cell(val) -> datetime | None:
        """解析单元格日期时间为 datetime（日初）。"""
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return None
        s0 = str(val).strip()
        if not s0 or s0 == "!" or s0.lower() in ("nan", "none", "nat"):
            return None
        if len(s0) >= 10 and s0[4] in "-/" and s0[7] in "-/":
            s0 = s0[:10].replace("-", "/")
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日"):
            try:
                return datetime.strptime(s0[:10], fmt)
            except ValueError:
                continue
        try:
            return datetime.strptime(s0, "%Y年%m月%d日")
        except ValueError:
            return None

    @staticmethod
    def _inclusive_days(d1: date, d2: date) -> int:
        """含首尾日的日历天数；d1>d2 时为 0。"""
        if d1 > d2:
            return 0
        return (d2 - d1).days + 1

    def calculate_billing_start(self, show_message=True, persist=True):
        """计算所有记录的计费开始时间（开通时间+90天）"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        count = 0
        for index, row in self.data_manager.data.iterrows():
            open_date = row["开通时间"]
            if open_date:
                billing_start = self.data_manager.calculate_billing_start_date(open_date)
                self.data_manager.data.loc[index, "计费开始时间"] = billing_start
                count += 1

        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的计费开始时间")

    def calculate_current_period_start(self, show_message=True, persist=True):
        """本期计费开始时间：账单月首日，或开通落在本月时取开通日（与清单一致）。"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        month_start, month_end = self._bill_month_bounds()
        month_start_dt = datetime.combine(month_start, datetime.min.time())
        month_end_dt = datetime.combine(month_end, datetime.min.time())

        count = 0
        for index, row in self.data_manager.data.iterrows():
            open_dt = self._parse_datetime_cell(row.get("开通时间"))
            current_period_start = ""
            if open_dt:
                if month_start_dt <= open_dt <= month_end_dt:
                    current_period_start = open_dt.strftime("%Y/%m/%d")
                elif open_dt < month_start_dt:
                    current_period_start = month_start.strftime("%Y/%m/%d")

            if current_period_start:
                self.data_manager.data.loc[index, "本期计费开始时间"] = current_period_start
                count += 1

        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的本期计费开始时间")

    def calculate_current_period_end(self, show_message=True, persist=True):
        """本期计费截至日期：无关停 → 计至账单月最后一天（整月）；本月内关停 → 关停前一日；关停早于账单月 → 空。"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        month_start, month_end = self._bill_month_bounds()

        count = 0
        for index, row in self.data_manager.data.iterrows():
            shutdown_dt = self._parse_datetime_cell(row.get("关停时间"))
            period_start_dt = self._parse_datetime_cell(row.get("本期计费开始时间"))
            open_dt = self._parse_datetime_cell(row.get("开通时间"))

            current_period_end = ""

            def _open_in_bill_month() -> bool:
                """与「本期计费开始」一致：开通早于月初或落在本月内，且未晚于月末（本月有计费段）。"""
                if not open_dt:
                    return False
                od = open_dt.date()
                return od <= month_end and (od < month_start or month_start <= od <= month_end)

            if shutdown_dt:
                sd = shutdown_dt.date()
                if sd < month_start:
                    current_period_end = ""
                elif month_start <= sd <= month_end:
                    last_bill = sd - timedelta(days=1)
                    if last_bill < month_start:
                        current_period_end = ""
                    else:
                        current_period_end = last_bill.strftime("%Y/%m/%d")
                else:
                    # 关停晚于本月：仍在用的服务，本月计至月末
                    if period_start_dt and month_start <= period_start_dt.date() <= month_end:
                        current_period_end = month_end.strftime("%Y/%m/%d")
                    elif _open_in_bill_month():
                        current_period_end = month_end.strftime("%Y/%m/%d")
                    else:
                        current_period_end = ""
            else:
                # 无关停：整月计至月末（与「半路关停」截到关停前一日相对）
                if period_start_dt and month_start <= period_start_dt.date() <= month_end:
                    current_period_end = month_end.strftime("%Y/%m/%d")
                elif _open_in_bill_month():
                    current_period_end = month_end.strftime("%Y/%m/%d")

            self.data_manager.data.loc[index, "本期计费截至日期"] = current_period_end
            count += 1

        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的本期计费结束时间")

    @staticmethod
    def _parse_nonneg_days_int(val) -> int:
        """解析「天数」类单元格：非负整数；空或无效为 0（用于本月核算天数、应计费天数等）。"""
        if val is None:
            return 0
        s = str(val).strip().replace(",", "")
        if not s:
            return 0
        try:
            return max(0, int(float(s)))
        except (ValueError, TypeError):
            return 0

    def calculate_month_amount(self, show_message=True, persist=True):
        """计算本月核算金额：数量 × 单价 ×「本月核算天数」（账单月口径：按本月核算天若用满约应收多少）。

        「合计费用」另按「应计费天数」计算，表示实收。表中「单价」按日价/天口径（与常见清单一致）。"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return
        
        count = 0
        for index, row in self.data_manager.data.iterrows():
            quantity = row['数量']
            unit_price = row['单价']
            month_days = self._parse_nonneg_days_int(row.get("本月核算天数"))
            
            # 转换为数值
            try:
                quantity = float(str(quantity).replace(',', ''))
            except:
                quantity = 0
            
            try:
                unit_price = float(str(unit_price).replace('￥', '').replace(',', ''))
            except:
                unit_price = 0
            
            # 本月核算天数为 0 或数量/单价无效时，本月核算金额必须为 ￥0.00
            if quantity > 0 and unit_price > 0 and month_days > 0:
                d = Decimal(str(quantity)) * Decimal(str(unit_price)) * Decimal(month_days)
                month_amount = round_money_half_up(d)
                self.data_manager.data.loc[index, "本月核算金额"] = f"￥{month_amount:,.2f}"
            else:
                self.data_manager.data.loc[index, "本月核算金额"] = "￥0.00"
            count += 1
        
        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的本月核算金额")

    def calculate_total_billing_days(self, show_message=True, persist=True):
        """推算「应计费天数」并写回标准列（与供应商「截至…月…日应计费天数」合并列一致）。

        规则：若计费开始日晚于本期计费截至 → 0（未到收钱日）。
        若（本期计费开始 − 计费开始）> 30 天 → 按整段本期窗口：截至 − 本期开始（含首尾）。
        否则 → 从 max(计费开始, 本期计费开始) 计至本期截至（含首尾）。
        """
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        if BILLING_DAYS_COLUMN not in self.data_manager.data.columns:
            if show_message:
                messagebox.showinfo("提示", f"当前表无「{BILLING_DAYS_COLUMN}」列")
            return

        count = 0
        for index, row in self.data_manager.data.iterrows():
            A = self._parse_datetime_cell(row.get("计费开始时间"))
            S = self._parse_datetime_cell(row.get("本期计费开始时间"))
            E = self._parse_datetime_cell(row.get("本期计费截至日期"))
            pe_raw = str(row.get("本期计费截至日期", "")).strip()

            v = 0
            if A and S and E and pe_raw != "!":
                Ad, Sd, Ed = A.date(), S.date(), E.date()
                if Ad <= Ed:
                    if (Sd - Ad).days > 30:
                        v = self._inclusive_days(Sd, Ed)
                    else:
                        v = self._inclusive_days(max(Ad, Sd), Ed)

            self.data_manager.data.loc[index, BILLING_DAYS_COLUMN] = str(v)
            count += 1

        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo(
                "成功",
                f"已推算并写回 {count} 条记录的「{BILLING_DAYS_COLUMN}」。",
            )

    def _effective_days_for_total_cost(self, row) -> int:
        """合计费用实收所用天数：应计费天数 → 使用天数 → 本月核算天数。

        供应商常见「本月核算天数」按账期满月 30 天，「截至…使用天数」按实际使用 29 天；
        若应计费天列为空却用核算天，会把「本期合计费用」算成与「本月核算金额」同口径，多收一天价。"""
        df = self.data_manager.data

        def _raw_days(col: str) -> str:
            if col not in df.columns:
                return ""
            raw = row.get(col)
            if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                return ""
            s = str(raw).strip()
            if s.lower() in ("nan", "none", "nat"):
                return ""
            if s in ("", "-", "—", "NA", "None"):
                return ""
            return s

        s_bd = _raw_days(BILLING_DAYS_COLUMN)
        if s_bd != "":
            return self._parse_nonneg_days_int(row.get(BILLING_DAYS_COLUMN))
        s_use = _raw_days("使用天数")
        if s_use != "":
            return self._parse_nonneg_days_int(row.get("使用天数"))
        return self._parse_nonneg_days_int(row.get("本月核算天数"))

    def calculate_total_cost(self, show_message=True, persist=True):
        """计算合计费用：数量 × 单价 × 实收天数（应计费天数 → 使用天数 → 本月核算天数）。"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return
        
        count = 0
        for index, row in self.data_manager.data.iterrows():
            quantity = row['数量']
            unit_price = row['单价']
            bill_days = self._effective_days_for_total_cost(row)
            
            # 转换为数值
            try:
                quantity = float(str(quantity).replace(',', ''))
            except:
                quantity = 0
            
            try:
                unit_price = float(str(unit_price).replace('￥', '').replace(',', ''))
            except:
                unit_price = 0
            
            # 实收天数为 0 则合计费用为 ￥0.00
            if quantity > 0 and unit_price > 0 and bill_days > 0:
                d = Decimal(str(quantity)) * Decimal(str(unit_price)) * Decimal(bill_days)
                total_cost = round_money_half_up(d)
                self.data_manager.data.loc[index, "合计费用"] = f"￥{total_cost:,.2f}"
            else:
                self.data_manager.data.loc[index, "合计费用"] = "￥0.00"
            count += 1
        
        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的合计费用")

    def calculate_all(self, show_message=True, persist=True):
        """一键计算全部：依次执行所有计算操作"""
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        self.calculate_billing_start(show_message=False, persist=persist)
        self.calculate_current_period_start(show_message=False, persist=persist)
        self.calculate_current_period_end(show_message=False, persist=persist)
        self.calculate_month_days(show_message=False, persist=persist)
        self.calculate_total_billing_days(show_message=False, persist=persist)
        self.calculate_month_amount(show_message=False, persist=persist)
        self.calculate_total_cost(show_message=False, persist=persist)

        if show_message:
            messagebox.showinfo("成功", "已按顺序完成全部计费相关计算并保存。")

    def calculate_month_days(self, show_message=True, persist=True):
        """本月核算天数：本期计费截至 − 本期计费开始（含首尾）。

        若两列暂无法解析，回退为「账单月 ∩ 开通～止期」交集（与旧逻辑一致，止期含关停前一日）。
        """
        if len(self.data_manager.data) == 0:
            if show_message:
                messagebox.showinfo("提示", "当前没有数据")
            return

        month_start, month_end = self._bill_month_bounds()

        count = 0
        for index, row in self.data_manager.data.iterrows():
            month_days: int | None = None
            ps = self._parse_datetime_cell(row.get("本期计费开始时间"))
            pe = self._parse_datetime_cell(row.get("本期计费截至日期"))
            pe_raw = str(row.get("本期计费截至日期", "")).strip()

            if ps and pe and pe_raw != "!":
                month_days = self._inclusive_days(ps.date(), pe.date())

            if month_days is None:
                open_d = self._parse_billing_date(row.get("开通时间"))
                shut_d = self._parse_billing_date(row.get("关停时间"))
                bill_until = self._parse_billing_date(row.get("本期计费截至日期"))

                month_days = 0
                if open_d:
                    ends = [month_end]
                    if shut_d:
                        if month_start <= shut_d <= month_end:
                            ends.append(shut_d - timedelta(days=1))
                        else:
                            ends.append(shut_d)
                    if bill_until and pe_raw != "!":
                        ends.append(bill_until)
                    last_in_month = min(ends)
                    period_start = max(month_start, open_d)
                    period_end = last_in_month
                    if (
                        period_end >= month_start
                        and period_start <= month_end
                        and period_start <= period_end
                    ):
                        month_days = (period_end - period_start).days + 1

            self.data_manager.data.loc[index, "本月核算天数"] = int(month_days or 0)
            count += 1

        if persist:
            self.data_manager.save_data()
            self.refresh_table()
        if show_message:
            messagebox.showinfo("成功", f"已计算 {count} 条记录的本月核算天数")

    def export_excel(self):
        """导出到Excel"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if file_path:
            success, msg = self.data_manager.export_to_excel(file_path)
            if success:
                messagebox.showinfo("成功", msg)
            else:
                messagebox.showerror("错误", msg)

    def export_calculated_workbook_copy(self):
        """复制原始供应商工作簿，追加「程序计算明细」子表（当前资产管理全量列）。"""
        src = getattr(self, "_normalize_input_path", "") or ""
        if not src or not os.path.isfile(src) or not src.lower().endswith((".xlsx", ".xlsm")):
            src = filedialog.askopenfilename(
                title="选择原始供应商 Excel（将整本复制后再追加明细子表）",
                filetypes=[("Excel", "*.xlsx;*.xlsm"), ("所有文件", "*.*")],
            )
        if not src:
            return
        src = os.path.abspath(src)
        if not src.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showwarning("提示", "请选择 .xlsx 或 .xlsm 文件。")
            return

        base, _ext = os.path.splitext(os.path.basename(src))
        dest = filedialog.asksaveasfilename(
            title="计算版保存位置",
            initialfile=f"{base}_计算版.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not dest:
            return
        dest = os.path.abspath(dest)

        if len(self.data_manager.data) == 0:
            messagebox.showwarning("提示", "当前资产管理无数据，请先执行「清洗、计算并载入资产管理」。")
            return

        ok, msg = excel_calc_wb.export_workbook_copy_with_detail_sheet(
            src,
            dest,
            self.data_manager.data,
            fill_source_columns=True,
            per_sheet_calc=getattr(self, "_norm_per_sheet_calc", None) or [],
        )
        if ok:
            messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("未完成", msg)
    
    def _coerce_dates_for_billing(self, df: pd.DataFrame) -> pd.DataFrame:
        """把 2026-02-05 00:00:00 等压成 2026/02/05，便于现有 strptime 规则。"""
        date_cols = (
            "开通时间",
            "关停时间",
            "计费开始时间",
            "本期计费开始时间",
            "本期计费截至日期",
        )
        out = df.copy()

        def conv(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return ""
            s0 = str(v).strip()
            if not s0 or s0.lower() in ("nan", "none"):
                return ""
            if len(s0) >= 10 and s0[4] in "-/" and s0[7] in "-/":
                return s0[:10].replace("-", "/")
            return s0

        for col in date_cols:
            if col in out.columns:
                out[col] = out[col].map(conv)
        return out

    def _load_cleaned_into_assets_replace(self, combined_df: pd.DataFrame) -> None:
        """用清洗后的明细整表替换「资产管理」内存与 CSV（仅标准列）。"""
        cols = self.data_manager.asset_types["database"]["columns"]
        n = len(combined_df)
        piece = {}
        for c in cols:
            if c in combined_df.columns:
                piece[c] = combined_df[c].reset_index(drop=True)
            else:
                piece[c] = pd.Series([""] * n, dtype=object)
        aligned = pd.DataFrame(piece, columns=cols)
        for i in range(len(aligned)):
            aligned.iat[i, aligned.columns.get_loc("唯一编号")] = f"DB{(i + 1):04d}"
        self.data_manager.data = aligned.fillna("")
        self.data_manager.save_data()

    def _aligned_asset_from_clean_df(self, clean_df: pd.DataFrame) -> pd.DataFrame:
        """将清洗结果对齐为「数据库资产」标准列（不写库），用于按 Sheet 内存试算。"""
        cols = self.data_manager.asset_types["database"]["columns"]
        n = len(clean_df)
        if n == 0:
            return pd.DataFrame(columns=cols)
        piece = {}
        for c in cols:
            if c in clean_df.columns:
                piece[c] = clean_df[c].astype(object).reset_index(drop=True)
            else:
                piece[c] = pd.Series([""] * n, dtype=object)
        aligned = pd.DataFrame(piece, columns=cols)
        for i in range(len(aligned)):
            aligned.iat[i, aligned.columns.get_loc("唯一编号")] = f"X{i + 1:05d}"
        return aligned.fillna("")

    def _apply_billing_rules_in_memory(self, aligned: pd.DataFrame) -> pd.DataFrame:
        """对给定明细套用与「一键计算全部」相同的规则，不写 CSV、不刷新主表。"""
        if len(aligned) == 0:
            return aligned.copy()
        backup = self.data_manager.data
        try:
            self.data_manager.data = aligned.copy()
            self.calculate_all(show_message=False, persist=False)
            return self.data_manager.data.copy()
        finally:
            self.data_manager.data = backup

    @staticmethod
    def _sum_df_money_column(df: pd.DataFrame, col: str) -> float:
        if df is None or len(df) == 0 or col not in df.columns:
            return 0.0
        return sum_parsed_money(df[col])

    def _formula_block_rows(self, df: pd.DataFrame) -> list:
        """按「数据块」列对公式后的金额做小计（无列则整页一块）。"""
        if df is None or len(df) == 0:
            return []
        mcol, ccol = "本月核算金额", "合计费用"
        if "数据块" in df.columns:
            out = []
            for title, sub in df.groupby(df["数据块"].fillna(""), sort=False):
                t = str(title).strip() if str(title).strip() else "(未命名块)"
                m = sum_parsed_money(sub[mcol])
                c = sum_parsed_money(sub[ccol])
                out.append(
                    {
                        "数据块标题": t,
                        "行数": len(sub),
                        "本月核算金额合计": m,
                        "合计费用合计": c,
                    }
                )
            return out
        m = sum_parsed_money(df[mcol])
        c = sum_parsed_money(df[ccol])
        return [
            {
                "数据块标题": "(整页)",
                "行数": len(df),
                "本月核算金额合计": m,
                "合计费用合计": c,
            }
        ]

    def _fill_norm_sheet_detail_tree(self, df: pd.DataFrame) -> None:
        """将公式后的标准列明细填入「本 Sheet 明细」表（不含「数据块」列）。"""
        t = self.norm_tree_sheet_detail
        t.delete(*t.get_children())
        if df is None or len(df) == 0:
            return
        cols = self.data_manager.asset_types["database"]["columns"]
        disp = df
        for _, row in disp.iterrows():
            vals = []
            for c in cols:
                v = row.get(c, "")
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    vals.append("")
                else:
                    vals.append(str(v).strip())
            t.insert("", tk.END, values=tuple(vals))

    def _sum_formatted_money_column(self, col: str) -> float:
        if col not in self.data_manager.data.columns:
            return 0.0
        return sum_parsed_money(self.data_manager.data[col])

    def _on_norm_sheet_select(self, _event=None):
        self.norm_tree_block.delete(*self.norm_tree_block.get_children())
        if hasattr(self, "norm_tree_sheet_detail"):
            self.norm_tree_sheet_detail.delete(*self.norm_tree_sheet_detail.get_children())
        if hasattr(self, "norm_tree_svc_group"):
            self.norm_tree_svc_group.delete(*self.norm_tree_svc_group.get_children())
        sel = self.norm_tree_sheet.selection()
        if not sel:
            return
        idx = int(sel[0])
        per_sheet = getattr(self, "_norm_per_sheet_calc", None) or []
        if per_sheet and 0 <= idx < len(per_sheet):
            _sn, df = per_sheet[idx]
            for b in self._formula_block_rows(df):
                self.norm_tree_block.insert(
                    "",
                    tk.END,
                    values=(
                        b.get("数据块标题", ""),
                        b.get("行数", ""),
                        f"￥{float(b.get('本月核算金额合计', 0) or 0):,.2f}",
                        f"￥{float(b.get('合计费用合计', 0) or 0):,.2f}",
                    ),
                )
            disp = df.drop(columns=["数据块"], errors="ignore")
            self._fill_norm_sheet_detail_tree(disp)
            self._update_norm_sheet_service_grouping(disp)
            return

        if not self._normalize_summaries:
            return
        if idx < 0 or idx >= len(self._normalize_summaries):
            return
        s = self._normalize_summaries[idx]
        for b in s.blocks:
            self.norm_tree_block.insert(
                "",
                tk.END,
                values=(
                    b.get("数据块标题", ""),
                    b.get("行数", ""),
                    f"￥{float(b.get('本月核算金额合计', 0) or 0):,.2f}",
                    f"￥{float(b.get('合计费用合计', 0) or 0):,.2f}",
                ),
            )

    def _format_norm_clipboard_text(self) -> str:
        parts = []
        per = getattr(self, "_norm_per_sheet_calc", None) or []
        if per:
            parts.append("=== 各 Sheet / 文件（公式重算，与「资产管理」一致）===")
            for sn, df in per:
                nm = self._sum_df_money_column(df, "本月核算金额")
                ct = self._sum_df_money_column(df, "合计费用")
                parts.append(
                    f"【{sn}】明细 {len(df)} 行 · 本月核算 ￥{nm:,.2f} · 合计费用 ￥{ct:,.2f}"
                )
                for b in self._formula_block_rows(df):
                    parts.append(
                        f"  · {b['数据块标题']}: {b['行数']} 行 · "
                        f"本月核算 ￥{float(b['本月核算金额合计']):,.2f} · "
                        f"合计费用 ￥{float(b['合计费用合计']):,.2f}"
                    )
                dto = df.drop(columns=["数据块"], errors="ignore")
                parts.append("  --- 按标准列「服务名称」汇总（公式后）---")
                for dn, nn, mm, cc in self._service_name_group_agg(dto):
                    parts.append(
                        f"    · {dn}: {nn} 行 · 本月核算 ￥{mm:,.2f} · 合计费用 ￥{cc:,.2f}"
                    )
                parts.append("")
        if self._normalize_summaries:
            parts.append("=== 清洗解析汇总（表内原数字，未套公式，仅供参考）===")
            parts.append(sheet_norm.format_summary_lines(self._normalize_summaries))
        return "\n".join(parts).strip()

    def _copy_normalize_summary(self):
        if not self._normalize_summaries and not (
            getattr(self, "_norm_per_sheet_calc", None) or []
        ):
            messagebox.showwarning("提示", "请先执行顶栏「清洗、计算并载入」。")
            return
        text = self._format_norm_clipboard_text()
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("已复制", "文字摘要已复制到剪贴板。")

    @staticmethod
    def _output_file_role(path: str) -> str:
        bn = os.path.basename(path).lower()
        if "_标准明细" in bn and bn.endswith(".csv"):
            return "标准明细CSV"
        if bn.endswith(".json"):
            return "汇总JSON"
        if bn.endswith(".txt"):
            return "汇总文本"
        return "其他输出"

    def _refresh_normalize_tab(
        self,
        summaries,
        written: list,
        out_dir: str,
        input_path: str,
        *,
        switch_to_assets_tab: bool = False,
        assets_rows: int = 0,
        assets_total_month: float = 0.0,
        assets_total_cost: float = 0.0,
        per_sheet_calc=None,
    ):
        """刷新「Sheet 汇总」页文件列表；可选切到「资产管理」。"""
        self._norm_formulas_only_mode = False
        self._normalize_summaries = list(summaries)
        self._normalize_written_paths = list(written)
        self._normalize_out_dir = out_dir
        self._normalize_input_path = os.path.abspath(input_path)

        self._norm_src_label.configure(text=f"源文件：{self._normalize_input_path}")
        if out_dir and str(out_dir).strip():
            self._norm_out_label.configure(
                text=f"输出目录：{os.path.abspath(out_dir)}"
            )
        else:
            self._norm_out_label.configure(
                text="（本次仅更新项目内 数据库资产.csv，未另存其它目录）"
            )

        if switch_to_assets_tab:
            self._norm_hint.configure(
                text=(
                    "上表为每个 Sheet（或单个 CSV）在「公式重算」后的行数与金额小计，与「资产管理」同一套规则；"
                    "点选一行可查看该页数据块小计、按列「服务名称」汇总的小计与下方逐行明细。底部为合并后的全库总计。"
                )
            )
            self._norm_grand_label.configure(
                text=(
                    f"【全部】合并进资产管理：{assets_rows} 行  ·  "
                    f"本月核算金额 ￥{assets_total_month:,.2f}  ·  "
                    f"合计费用 ￥{assets_total_cost:,.2f}"
                )
            )
        else:
            self._norm_hint.configure(
                text="上表为各 Sheet/文件清洗解析结果（未公式重算）；点选一行可查看数据块明细。"
            )
            g_rows = sum(s.row_count for s in summaries)
            g_month = sum_parsed_money(s.total_month_amount for s in summaries)
            g_cost = sum_parsed_money(s.total_cost for s in summaries)
            self._norm_grand_label.configure(
                text=f"解析合计（表内原数字相加）：明细 {g_rows} 行  ·  本月核算 ￥{g_month:,.2f}  ·  合计费用 ￥{g_cost:,.2f}"
            )

        self.norm_tree_sheet.delete(*self.norm_tree_sheet.get_children())
        if hasattr(self, "norm_tree_sheet_detail"):
            self.norm_tree_sheet_detail.delete(*self.norm_tree_sheet_detail.get_children())
        if hasattr(self, "norm_tree_svc_group"):
            self.norm_tree_svc_group.delete(*self.norm_tree_svc_group.get_children())

        if switch_to_assets_tab and per_sheet_calc:
            self._norm_per_sheet_calc = list(per_sheet_calc)
            self.norm_tree_sheet.heading("detail_csv", text="口径")
            for i, (sn, df) in enumerate(per_sheet_calc):
                nm = self._sum_df_money_column(df, "本月核算金额")
                ct = self._sum_df_money_column(df, "合计费用")
                self.norm_tree_sheet.insert(
                    "",
                    tk.END,
                    iid=str(i),
                    values=(
                        sn,
                        "公式重算（与资产管理一致）",
                        len(df),
                        f"￥{nm:,.2f}",
                        f"￥{ct:,.2f}",
                    ),
                )
            if per_sheet_calc:
                self.norm_tree_sheet.selection_set("0")
                self.norm_tree_sheet.see("0")
                self._on_norm_sheet_select()
        elif switch_to_assets_tab:
            self._norm_per_sheet_calc = []
            self.norm_tree_sheet.heading("detail_csv", text="口径")
            self.norm_tree_sheet.insert(
                "",
                tk.END,
                iid="0",
                values=(
                    "「资产管理」合并后（公式重算）",
                    "—",
                    assets_rows,
                    f"￥{assets_total_month:,.2f}",
                    f"￥{assets_total_cost:,.2f}",
                ),
            )
            self.norm_tree_block.delete(*self.norm_tree_block.get_children())
            self.norm_tree_block.insert(
                "",
                tk.END,
                values=("逐行明细与金额在「资产管理」主表", "—", "—", "—"),
            )
            self.norm_tree_sheet.selection_set("0")
            self.norm_tree_sheet.see("0")
        else:
            self._norm_per_sheet_calc = []
            self.norm_tree_sheet.heading("detail_csv", text="口径 / 路径")
            for i, s in enumerate(summaries):
                detail_csv = s.detail_csv_path or ""
                detail_show = os.path.basename(detail_csv) if detail_csv else "—"
                self.norm_tree_sheet.insert(
                    "",
                    tk.END,
                    iid=str(i),
                    values=(
                        s.sheet_name,
                        detail_show,
                        s.row_count,
                        f"￥{s.total_month_amount:,.2f}",
                        f"￥{s.total_cost:,.2f}",
                    ),
                )

            if summaries:
                self.norm_tree_sheet.selection_set("0")
                self.norm_tree_sheet.see("0")
                self._on_norm_sheet_select()
            else:
                self.norm_tree_block.delete(*self.norm_tree_block.get_children())

        self.norm_path_tree.delete(*self.norm_path_tree.get_children())
        self.norm_path_tree.insert(
            "",
            tk.END,
            values=("源文件（本次读取）", self._normalize_input_path),
        )
        for p in written:
            self.norm_path_tree.insert(
                "",
                tk.END,
                values=(self._output_file_role(p), os.path.abspath(p)),
            )

        if switch_to_assets_tab:
            self.main_nb.select(self.tab_assets)
        else:
            self.main_nb.select(self.tab_normalize)

    def _extract_normalize_from_excel_path(self, path: str) -> tuple[list, list]:
        """生成网格缓存并解析工作簿内全部工作表。返回 (clean_dfs, summaries)。"""
        grid_cache_dir = sheet_norm.new_excel_grid_cache_session(path)
        self._last_excel_grid_cache_dir = grid_cache_dir
        sheet_norm.export_excel_workbook_grid_csvs(path, grid_cache_dir)
        pairs = sheet_norm.process_excel_workbook(
            path,
            grid_cache_dir=grid_cache_dir,
            skip_export=True,
        )
        clean_dfs = [p[0] for p in pairs]
        summaries = [p[1] for p in pairs]
        return clean_dfs, summaries

    def normalize_summarize_dialog(self):
        """内存中清洗 → 合并明细 → 写入「资产管理」→ 公式重算 → 仅保存 数据库资产.csv；不导出其它文件。

        Excel 含多个工作表时始终合并各表解析出的明细（不再询问是否只处理第一张表）。
        """
        path = filedialog.askopenfilename(
            title="选择要清洗的 CSV 或 Excel",
            filetypes=[
                ("Excel 与 CSV", "*.xlsx;*.xls;*.csv"),
                ("Excel", "*.xlsx;*.xls"),
                ("CSV", "*.csv"),
            ],
        )
        if not path:
            return

        bm = self.parse_bill_month_from_filename(os.path.basename(path))
        if bm:
            self.var_bill_year.set(bm[0])
            self.var_bill_month.set(bm[1])

        try:
            summaries = []
            clean_dfs = []
            self._last_excel_grid_cache_dir = ""

            if path.lower().endswith(".csv"):
                clean_df, summary = sheet_norm.process_csv_file(path)
                clean_dfs.append(clean_df)
                summaries.append(summary)
            else:
                clean_dfs, summaries = self._extract_normalize_from_excel_path(path)

            combined = pd.concat(clean_dfs, ignore_index=True)
            if len(combined) == 0:
                messagebox.showwarning("提示", "未解析出任何明细行，请检查表格格式。")
                return

            if not messagebox.askyesno(
                "确认载入",
                "将把本次清洗得到的明细合并写入「资产管理」并替换当前数据，\n"
                "再按规则重算日期与金额，保存到 数据库资产.csv。\n\n"
                "是否继续？",
            ):
                return

            per_sheet_named = []
            for clean_df, summary in zip(clean_dfs, summaries):
                cdf = self._coerce_dates_for_billing(clean_df.copy())
                bs = (
                    cdf["数据块"].reset_index(drop=True)
                    if "数据块" in cdf.columns
                    else None
                )
                cdf_b = cdf.drop(columns=["数据块"], errors="ignore")
                al = self._aligned_asset_from_clean_df(cdf_b)
                rc = self._apply_billing_rules_in_memory(al)
                if bs is not None and len(bs) == len(rc):
                    rc = rc.copy()
                    rc["数据块"] = bs.values
                per_sheet_named.append((summary.sheet_name, rc))

            combined = self._coerce_dates_for_billing(
                pd.concat(clean_dfs, ignore_index=True)
            )
            self._load_cleaned_into_assets_replace(combined)
            self.calculate_all(show_message=False)
            self.refresh_table()

            month_ttl = self._sum_formatted_money_column("本月核算金额")
            cost_ttl = self._sum_formatted_money_column("合计费用")
            n = len(self.data_manager.data)

            self._normalize_summaries = list(summaries)
            self._normalize_input_path = os.path.abspath(path)
            self._refresh_normalize_tab(
                summaries,
                [],
                "",
                path,
                switch_to_assets_tab=True,
                assets_rows=n,
                assets_total_month=month_ttl,
                assets_total_cost=cost_ttl,
                per_sheet_calc=per_sheet_named,
            )
            self.status_var.set(
                f"已载入 {n} 条并重算 · 本月核算合计 ￥{month_ttl:,.2f} · 合计费用 ￥{cost_ttl:,.2f}"
            )
            done_extra = ""
            if self._last_excel_grid_cache_dir:
                done_extra = (
                    f"\n\n本次从 Excel 落盘的原始网格 CSV（与直接上传 CSV 同一套解析）位于：\n"
                    f"{self._last_excel_grid_cache_dir}"
                )
            messagebox.showinfo(
                "完成",
                f"已载入 {n} 条并重算保存（数据库资产.csv）。\n\n"
                f"本月核算金额合计：￥{month_ttl:,.2f}\n"
                f"合计费用总计：￥{cost_ttl:,.2f}\n\n"
                "可在顶栏「生成计算版 Excel」导出带回填的工作簿。"
                f"{done_extra}",
            )
        except Exception as e:
            messagebox.showerror("清洗失败", str(e))

    def show_statistics(self):
        """显示统计报表"""
        stats = self.data_manager.get_statistics()

        dialog = tk.Toplevel(self.root)
        dialog.title("数据库资产统计报表")
        dialog.geometry("520x420")
        dialog.transient(self.root)

        outer = ttk.Frame(dialog, padding=14)
        outer.pack(fill=tk.BOTH, expand=True)

        ttk.Label(outer, text="统计报表", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(outer, text=f"资产总数：{stats['total_count']}", font=self._font_ui).pack(
            anchor=tk.W, pady=(10, 4)
        )

        if "total_cost" in stats:
            ttk.Label(
                outer,
                text=f"总费用：￥{stats['total_cost']:,.2f}",
                font=self._font_ui,
            ).pack(anchor=tk.W, pady=2)

        ttk.Label(outer, text="按服务类型", style="Sub.TLabel").pack(anchor=tk.W, pady=(14, 4))
        if "cost_by_type" in stats:
            for service_type, cost in stats["cost_by_type"].items():
                ttk.Label(
                    outer, text=f"  · {service_type}：￥{cost:,.2f}", font=self._font_ui_sm
                ).pack(anchor=tk.W)

        ttk.Label(outer, text="服务名称（前 10 条）", style="Sub.TLabel").pack(
            anchor=tk.W, pady=(12, 4)
        )
        for service_name in stats["service_names"][:10]:
            ttk.Label(outer, text=f"  · {service_name}", font=self._font_ui_sm).pack(
                anchor=tk.W
            )

        ttk.Button(outer, text="保存为文本…", command=lambda: self.save_report(stats)).pack(
            pady=(16, 0), anchor=tk.W
        )
    
    def save_report(self, stats):
        """保存统计报表"""
        report = f"数据库资产统计报表\n"
        report += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        report += f"资产总数: {stats['total_count']}\n"
        if 'total_cost' in stats:
            report += f"总费用: ￥{stats['total_cost']:,.2f}\n\n"
        
        report += "服务类型费用分布:\n"
        if 'cost_by_type' in stats:
            for service_type, cost in stats['cost_by_type'].items():
                report += f"  {service_type}: ￥{cost:,.2f}\n"
        
        report += "\n服务名称列表:\n"
        for service_name in stats['service_names']:
            report += f"  {service_name}\n"
        
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(report)
            messagebox.showinfo("成功", "报表已保存")

def main():
    root = tk.Tk()
    app = DatabaseAssetGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
