from __future__ import annotations

import json
import os
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from reportflow_core import execute_scheme, load_workbook_preview

try:
    import pythoncom
    import win32com.client
except ImportError:  # pragma: no cover
    pythoncom = None
    win32com = None


APP_TITLE = "ReportFlow"
CONFIG_PATH = Path.home() / ".reportflow_desktop.json"
DEFAULT_SETTINGS = {
    "app_preference": "auto",
    "capture_data": True,
    "capture_format": True,
    "capture_charts": True,
    "capture_workbook": True,
    "capture_cross_sheet": True,
    "max_capture_rows": 200,
    "auto_open_result": False,
    "keep_temp_xlsx": False,
}


@dataclass
class WorkbookSnapshot:
    sheet_name: str
    sheet_names: list[str]
    headers: list[str]
    values: list[dict[str, Any]]
    formulas: list[dict[str, Any]]
    hidden_columns: list[str]
    formats: dict[str, dict[str, Any]]
    column_widths: dict[str, float]
    row_heights: dict[int, float]
    charts: list[dict[str, Any]]


class ExcelNativeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x720")
        self.minsize(860, 640)
        self.configure(bg="#0f1115")

        self.excel = None
        self.workbook = None
        self.spreadsheet_app_name = ""
        self.file_path: Path | None = None
        self.baseline: WorkbookSnapshot | None = None
        self.rules = self.empty_rules()
        self.settings = self.load_settings()

        self._build_style()
        self._build_layout()

    @staticmethod
    def empty_rules() -> dict[str, list[dict[str, Any]]]:
        return {
            "cell_edit_rules": [],
            "operation_rules": [],
            "filter_rules": [],
            "sort_rules": [],
            "calculated_fields": [],
            "excel_formula_rules": [],
            "group_rules": [],
            "visual_rules": [],
            "chart_rules": [],
            "workbook_rules": [],
            "cross_sheet_rules": [],
        }

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Microsoft YaHei UI", 10), background="#0b0d12", foreground="#f5f5f0")
        style.configure("Root.TFrame", background="#0b0d12")
        style.configure("Panel.TFrame", background="#161a22")
        style.configure("Inset.TFrame", background="#10131a")
        style.configure("Title.TLabel", background="#0b0d12", foreground="#f5f5f0", font=("Segoe UI", 23, "bold"))
        style.configure("Sub.TLabel", background="#0b0d12", foreground="#9aa3ad")
        style.configure("PanelTitle.TLabel", background="#161a22", foreground="#f5f5f0", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("InsetTitle.TLabel", background="#10131a", foreground="#f5f5f0", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Muted.TLabel", background="#161a22", foreground="#9aa3ad")
        style.configure("InsetMuted.TLabel", background="#10131a", foreground="#9aa3ad")
        style.configure("Accent.TButton", background="#10a37f", foreground="#ffffff", borderwidth=0, padding=(14, 10))
        style.map("Accent.TButton", background=[("active", "#0b7f63")])
        style.configure("Ghost.TButton", background="#222832", foreground="#f5f5f0", borderwidth=0, padding=(12, 9))
        style.map("Ghost.TButton", background=[("active", "#2e3643")])
        style.configure("Soft.TButton", background="#161a22", foreground="#d7dce2", borderwidth=0, padding=(10, 7))
        style.map("Soft.TButton", background=[("active", "#222832")])
        style.configure("TCheckbutton", background="#161a22", foreground="#f5f5f0")
        style.configure("TRadiobutton", background="#161a22", foreground="#f5f5f0")

    def _build_layout(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=(18, 18, 18, 18))
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=0)
        root.columnconfigure(1, weight=1)
        root.columnconfigure(2, weight=0)
        root.rowconfigure(2, weight=1)

        ttk.Label(root, text="ReportFlow", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(root, text="表格原生录制台：打开 Excel/WPS，录制数据、格式、图表和多 Sheet 联动", style="Sub.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 16))

        left = ttk.Frame(root, style="Root.TFrame", width=300)
        left.grid(row=2, column=0, sticky="nsew", padx=(0, 14))
        left.grid_propagate(False)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        actions = ttk.Frame(left, style="Panel.TFrame", padding=(14, 14))
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        actions.columnconfigure(0, weight=1)

        ttk.Label(actions, text="工作流", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))
        ttk.Button(actions, text="1  打开表格并开始录制", style="Accent.TButton", command=self.open_excel).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(actions, text="2  捕获当前操作", style="Ghost.TButton", command=self.capture_rules).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(actions, text="3  执行并生成结果", style="Accent.TButton", command=self.execute_current_scheme).grid(row=3, column=0, sticky="ew")

        utility = ttk.Frame(left, style="Panel.TFrame", padding=(14, 14))
        utility.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        utility.columnconfigure(0, weight=1)
        ttk.Label(utility, text="方案", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))
        ttk.Button(utility, text="一键加载规则", style="Ghost.TButton", command=self.import_scheme).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(utility, text="导出规则", style="Ghost.TButton", command=self.export_scheme).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(utility, text="重新设为起点", style="Soft.TButton", command=self.reset_baseline).grid(row=3, column=0, sticky="ew", pady=(12, 8))
        ttk.Button(utility, text="清空规则", style="Soft.TButton", command=self.clear_rules).grid(row=4, column=0, sticky="ew")

        settings = ttk.Frame(left, style="Root.TFrame")
        settings.grid(row=2, column=0, sticky="ew")
        settings.columnconfigure(0, weight=1)
        ttk.Button(settings, text="设置", style="Ghost.TButton", command=self.open_settings).grid(row=0, column=0, sticky="ew")

        center = ttk.Frame(root, style="Root.TFrame")
        center.grid(row=2, column=1, sticky="nsew", padx=(0, 14))
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=3)
        center.rowconfigure(1, weight=2)

        rules_panel = ttk.Frame(center, style="Panel.TFrame", padding=(14, 14))
        rules_panel.grid(row=0, column=0, sticky="nsew", pady=(0, 12))
        rules_panel.rowconfigure(1, weight=1)
        rules_panel.columnconfigure(0, weight=1)
        ttk.Label(rules_panel, text="已生成规则", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.rule_list = tk.Listbox(
            rules_panel,
            bg="#10131a",
            fg="#f5f5f0",
            selectbackground="#10a37f",
            selectforeground="#ffffff",
            borderwidth=0,
            highlightthickness=0,
            activestyle="none",
            font=("Microsoft YaHei UI", 10),
        )
        self.rule_list.grid(row=1, column=0, sticky="nsew")

        self._build_formula_panel(center)

        right = ttk.Frame(root, style="Root.TFrame", width=300)
        right.grid(row=2, column=2, sticky="nsew")
        right.grid_propagate(False)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        status_panel = ttk.Frame(right, style="Panel.TFrame", padding=(18, 18))
        status_panel.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        status_panel.columnconfigure(0, weight=1)
        ttk.Label(status_panel, text="当前状态", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.status_var = tk.StringVar(value="先打开一个 Excel 文件。之后直接在 Excel 里筛选、排序、删列、写公式。")
        ttk.Label(status_panel, textvariable=self.status_var, style="Muted.TLabel", wraplength=500).grid(row=1, column=0, sticky="ew")

        settings_summary = ttk.Frame(right, style="Panel.TFrame", padding=(18, 18))
        settings_summary.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        settings_summary.columnconfigure(0, weight=1)
        ttk.Label(settings_summary, text="当前设置", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.settings_summary_var = tk.StringVar()
        ttk.Label(settings_summary, textvariable=self.settings_summary_var, style="Muted.TLabel", wraplength=250).grid(row=1, column=0, sticky="ew")
        ttk.Button(settings_summary, text="打开设置", style="Soft.TButton", command=self.open_settings).grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self.refresh_settings_summary()

        guide_panel = ttk.Frame(right, style="Panel.TFrame", padding=(18, 18))
        guide_panel.grid(row=2, column=0, sticky="nsew")
        guide_panel.columnconfigure(0, weight=1)
        guide_panel.rowconfigure(1, weight=1)
        ttk.Label(guide_panel, text="工作流", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        guide = tk.Text(guide_panel, bg="#11141a", fg="#d5d9de", borderwidth=0, highlightthickness=0, font=("Microsoft YaHei UI", 11), wrap=tk.WORD)
        guide.grid(row=1, column=0, sticky="nsew")
        guide.insert("1.0", self.operation_docs(short=True))
        guide.configure(state=tk.DISABLED)

    def _build_formula_panel(self, parent: ttk.Frame) -> None:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=(14, 14))
        panel.grid(row=1, column=0, sticky="nsew")
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(3, weight=1)

        ttk.Label(panel, text="函数查询 / 生成", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.formula_request = tk.Text(panel, height=4, bg="#11141a", fg="#f4f4f0", insertbackground="#f4f4f0", borderwidth=0, highlightthickness=0, font=("Microsoft YaHei UI", 10))
        self.formula_request.grid(row=1, column=0, sticky="ew")
        self.formula_request.insert("1.0", "例如：根据状态列判断是否完成，完成显示1，否则0")

        buttons = ttk.Frame(panel, style="Panel.TFrame")
        buttons.grid(row=2, column=0, sticky="ew", pady=(10, 8))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)
        buttons.columnconfigure(2, weight=1)
        ttk.Button(buttons, text="生成函数", style="Accent.TButton", command=self.generate_formula).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="复制函数", style="Ghost.TButton", command=self.copy_formula).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(buttons, text="写入选中单元格", style="Ghost.TButton", command=self.write_formula_to_excel).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        self.formula_result = tk.Text(panel, height=9, bg="#0f1115", fg="#d7fff1", insertbackground="#d7fff1", borderwidth=0, highlightthickness=0, font=("Consolas", 11))
        self.formula_result.grid(row=3, column=0, sticky="nsew")

    @staticmethod
    def load_settings() -> dict[str, Any]:
        if not CONFIG_PATH.exists():
            return dict(DEFAULT_SETTINGS)
        try:
            loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return {**DEFAULT_SETTINGS, **loaded}
        except Exception:
            return dict(DEFAULT_SETTINGS)

    def save_settings(self) -> None:
        CONFIG_PATH.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")
        self.refresh_settings_summary()

    def refresh_settings_summary(self) -> None:
        if not hasattr(self, "settings_summary_var"):
            return
        labels = {
            "auto": "自动选择",
            "excel": "优先 Excel",
            "wps": "优先 WPS",
        }
        scopes = []
        if self.settings.get("capture_data"):
            scopes.append("数据")
        if self.settings.get("capture_format"):
            scopes.append("格式")
        if self.settings.get("capture_charts"):
            scopes.append("图表")
        if self.settings.get("capture_workbook"):
            scopes.append("工作簿")
        if self.settings.get("capture_cross_sheet"):
            scopes.append("Sheet联动")
        self.settings_summary_var.set(
            f"表格内核：{labels.get(self.settings.get('app_preference'), '自动选择')}\n"
            f"捕获范围：{'、'.join(scopes) or '未启用'}\n"
            f"最大捕获行数：{self.settings.get('max_capture_rows')}\n"
            f"生成后打开：{'是' if self.settings.get('auto_open_result') else '否'}"
        )

    def open_settings(self) -> None:
        window = tk.Toplevel(self)
        window.title("设置")
        window.geometry("560x620")
        window.minsize(500, 560)
        window.configure(bg="#0b0d12")
        window.columnconfigure(0, weight=1)

        frame = ttk.Frame(window, style="Panel.TFrame", padding=(18, 18))
        frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="设置", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 14))

        app_var = tk.StringVar(value=str(self.settings.get("app_preference", "auto")))
        ttk.Label(frame, text="表格内核偏好", style="Muted.TLabel").grid(row=1, column=0, sticky="w")
        app_box = ttk.Frame(frame, style="Panel.TFrame")
        app_box.grid(row=2, column=0, sticky="ew", pady=(6, 14))
        for index, (value, text) in enumerate([("auto", "自动"), ("excel", "Excel"), ("wps", "WPS")]):
            ttk.Radiobutton(app_box, text=text, variable=app_var, value=value).grid(row=0, column=index, sticky="w", padx=(0, 18))

        capture_vars = {
            "capture_data": tk.BooleanVar(value=bool(self.settings.get("capture_data"))),
            "capture_format": tk.BooleanVar(value=bool(self.settings.get("capture_format"))),
            "capture_charts": tk.BooleanVar(value=bool(self.settings.get("capture_charts"))),
            "capture_workbook": tk.BooleanVar(value=bool(self.settings.get("capture_workbook"))),
            "capture_cross_sheet": tk.BooleanVar(value=bool(self.settings.get("capture_cross_sheet"))),
        }
        ttk.Label(frame, text="捕获范围", style="Muted.TLabel").grid(row=3, column=0, sticky="w")
        capture_box = ttk.Frame(frame, style="Panel.TFrame")
        capture_box.grid(row=4, column=0, sticky="ew", pady=(6, 14))
        for row, (key, text) in enumerate([
            ("capture_data", "数据操作"),
            ("capture_format", "格式修改"),
            ("capture_charts", "图表制作"),
            ("capture_workbook", "Sheet 新增/删除"),
            ("capture_cross_sheet", "跨 Sheet 公式"),
        ]):
            ttk.Checkbutton(capture_box, text=text, variable=capture_vars[key]).grid(row=row, column=0, sticky="w", pady=2)

        max_rows_var = tk.StringVar(value=str(self.settings.get("max_capture_rows", 200)))
        ttk.Label(frame, text="最大捕获行数", style="Muted.TLabel").grid(row=5, column=0, sticky="w")
        ttk.Entry(frame, textvariable=max_rows_var).grid(row=6, column=0, sticky="ew", pady=(6, 14))

        auto_open_var = tk.BooleanVar(value=bool(self.settings.get("auto_open_result")))
        keep_temp_var = tk.BooleanVar(value=bool(self.settings.get("keep_temp_xlsx")))
        ttk.Checkbutton(frame, text="生成结果后自动打开", variable=auto_open_var).grid(row=7, column=0, sticky="w", pady=2)
        ttk.Checkbutton(frame, text="保留 WPS .et 临时 xlsx 文件", variable=keep_temp_var).grid(row=8, column=0, sticky="w", pady=2)

        buttons = ttk.Frame(frame, style="Panel.TFrame")
        buttons.grid(row=9, column=0, sticky="ew", pady=(18, 0))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)
        buttons.columnconfigure(2, weight=1)

        def save_and_close() -> None:
            try:
                max_rows = max(20, min(5000, int(max_rows_var.get())))
            except ValueError:
                messagebox.showwarning(APP_TITLE, "最大捕获行数必须是数字")
                return
            self.settings["app_preference"] = app_var.get()
            for key, var in capture_vars.items():
                self.settings[key] = bool(var.get())
            self.settings["max_capture_rows"] = max_rows
            self.settings["auto_open_result"] = bool(auto_open_var.get())
            self.settings["keep_temp_xlsx"] = bool(keep_temp_var.get())
            self.save_settings()
            window.destroy()

        ttk.Button(buttons, text="操作文档", style="Soft.TButton", command=self.show_operation_docs).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="取消", style="Ghost.TButton", command=window.destroy).grid(row=0, column=1, sticky="ew", padx=3)
        ttk.Button(buttons, text="保存设置", style="Accent.TButton", command=save_and_close).grid(row=0, column=2, sticky="ew", padx=(6, 0))

    def show_operation_docs(self) -> None:
        window = tk.Toplevel(self)
        window.title("操作文档")
        window.geometry("680x620")
        window.minsize(560, 480)
        window.configure(bg="#0f1115")
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        text = tk.Text(window, bg="#11141a", fg="#f4f4f0", insertbackground="#f4f4f0", borderwidth=0, highlightthickness=0, font=("Microsoft YaHei UI", 11), wrap=tk.WORD, padx=18, pady=18)
        text.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        text.insert("1.0", self.operation_docs(short=False))
        text.configure(state=tk.DISABLED)

    @staticmethod
    def operation_docs(short: bool = False) -> str:
        if short:
            return (
                "1. 打开 Excel 并开始录制\n"
                "   ReportFlow 会启动真正的 Microsoft Excel。\n\n"
                "2. 在 Excel 里正常操作\n"
                "   支持 Excel/WPS 表格；直接筛选、排序、删除/隐藏列、重命名列、改单元格、写公式列。\n\n"
                "3. 捕获当前操作为规则\n"
                "   左侧会显示识别到的规则。\n\n"
                "4. 导出或一键加载规则\n"
                "   规则文件可以给不同用户复用。\n\n"
            "5. 函数查询/生成\n"
            "   输入需求，生成公式后复制或写入 Excel 当前单元格。\n\n"
            "6. 设置\n"
            "   配置 Excel/WPS 偏好、捕获范围、最大捕获行数和输出行为。"
            )
        return (
            "ReportFlow 操作文档\n\n"
            "一、创建规则\n"
            "1. 点击“打开 Excel 并开始录制”。\n"
            "2. 选择需要处理的 Excel 文件。\n"
            "3. ReportFlow 会优先打开 Microsoft Excel；如果未找到，会尝试打开 WPS 表格。\n"
            "4. 在 Excel/WPS 中按平时习惯操作，例如筛选、排序、删除列、隐藏列、重命名列、修改单元格、新增固定值列、新增或修改公式列。\n"
            "5. 操作完成后回到 ReportFlow，点击“捕获当前操作为规则”。\n"
            "6. 左侧“已生成规则”会显示识别出来的规则。\n\n"
            "二、复用规则\n"
            "1. 点击“一键加载规则”。\n"
            "2. 选择之前导出的 JSON 规则文件。\n"
            "3. 规则会显示在左侧列表中。\n"
            "4. 打开同结构 Excel 后，点击“执行并生成结果”。\n\n"
            "三、函数查询/生成\n"
            "1. 在左侧“函数查询 / 生成”中输入自然语言需求。\n"
            "2. 例如：根据状态列判断是否完成，完成显示1，否则0。\n"
            "3. 点击“生成函数”。\n"
            "4. 可以点击“复制函数”，也可以先在 Excel 里选中单元格，再点击“写入选中单元格”。\n\n"
            "四、当前可识别规则\n"
            "- 删除列\n"
            "- 隐藏列，按删除列处理\n"
            "- 重命名列\n"
            "- 新增空列/固定值列\n"
            "- 调整/保留列\n"
            "- 单元格修改\n"
            "- 新增或修改公式列\n"
            "- 自动筛选条件\n"
            "- 排序字段\n\n"
            "五、注意事项\n"
            "- 录制时请保持目标工作表为当前激活 Sheet。\n"
            "- 规则复用依赖列名，建议同类报表保持表头一致。\n"
            "- WPS 自有 .et 文件执行时会尝试临时另存为 .xlsx。\n"
            "- 函数生成会优先根据当前 Excel 表头猜测单元格引用，复杂公式仍需要人工确认。\n\n"
            "六、设置项\n"
            "- 表格内核偏好：自动、优先 Excel、优先 WPS。\n"
            "- 捕获范围：数据操作、格式修改、图表制作、Sheet 新增/删除、跨 Sheet 公式。\n"
            "- 最大捕获行数：控制格式和单元格差异扫描范围。\n"
            "- 输出行为：生成后自动打开结果，是否保留 WPS 临时 xlsx 文件。"
        )

    def generate_formula(self) -> None:
        description = self.formula_request.get("1.0", tk.END).strip()
        if not description:
            messagebox.showwarning(APP_TITLE, "请输入函数需求描述")
            return
        headers = self.current_headers()
        formula, note = formula_from_description(description, headers)
        self.formula_result.delete("1.0", tk.END)
        self.formula_result.insert("1.0", f"{formula}\n\n说明：{note}")
        self.status_var.set("已生成函数。可以复制，也可以写入当前 Excel 选中的单元格。")

    def copy_formula(self) -> None:
        formula = self.current_generated_formula()
        if not formula:
            messagebox.showwarning(APP_TITLE, "请先生成函数")
            return
        self.clipboard_clear()
        self.clipboard_append(formula)
        self.status_var.set("函数已复制到剪贴板。")

    def write_formula_to_excel(self) -> None:
        formula = self.current_generated_formula()
        if not formula:
            messagebox.showwarning(APP_TITLE, "请先生成函数")
            return
        if self.excel is None:
            messagebox.showwarning(APP_TITLE, "请先打开 Excel")
            return
        try:
            self.excel.ActiveCell.Formula = formula
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"写入 Excel 失败：{exc}")
            return
        self.status_var.set("函数已写入 Excel 当前选中单元格。")

    def current_generated_formula(self) -> str:
        text = self.formula_result.get("1.0", tk.END).strip()
        return text.splitlines()[0].strip() if text else ""

    def current_headers(self) -> list[str]:
        if self.baseline:
            return self.baseline.headers
        if self.workbook is not None:
            try:
                return self.snapshot_active_sheet().headers
            except Exception:
                return []
        return []

    def open_excel(self) -> None:
        if win32com is None or pythoncom is None:
            messagebox.showerror(APP_TITLE, "需要安装 pywin32 才能调用 Excel/WPS 表格。\n请执行：pip install pywin32")
            return
        path = filedialog.askopenfilename(title="选择表格文件", filetypes=[("表格文件", "*.xlsx *.xlsm *.xls *.et"), ("所有文件", "*.*")])
        if not path:
            return
        self.file_path = Path(path)
        try:
            pythoncom.CoInitialize()
            self.excel, self.spreadsheet_app_name = self.create_spreadsheet_app()
            self.excel.Visible = True
            try:
                self.excel.DisplayAlerts = False
            except Exception:
                pass
            self.workbook = self.excel.Workbooks.Open(str(self.file_path))
            self.baseline = self.snapshot_active_sheet()
            self.rules = self.empty_rules()
            self.refresh_rule_list()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"打开表格失败：{exc}")
            return
        self.status_var.set(f"正在使用 {self.spreadsheet_app_name} 录制：{self.file_path.name}\n请在表格软件里正常操作，完成后点“捕获当前操作为规则”。")

    @staticmethod
    def app_candidates(preference: str) -> list[tuple[str, str]]:
        excel = [
            ("Excel.Application", "Microsoft Excel"),
        ]
        wps = [
            ("Ket.Application", "WPS 表格"),
            ("KET.Application", "WPS 表格"),
            ("Et.Application", "WPS 表格"),
            ("ET.Application", "WPS 表格"),
            ("WPS.Application", "WPS Office"),
        ]
        if preference == "excel":
            return excel + wps
        if preference == "wps":
            return wps + excel
        return excel + wps

    def create_spreadsheet_app(self):
        candidates = self.app_candidates(str(self.settings.get("app_preference", "auto")))
        errors = []
        for prog_id, display_name in candidates:
            try:
                return win32com.client.DispatchEx(prog_id), display_name
            except Exception as exc:
                errors.append(f"{prog_id}: {exc}")
        raise RuntimeError("未找到可调用的 Excel 或 WPS 表格 COM 服务。\n" + "\n".join(errors[-3:]))

    def snapshot_active_sheet(self) -> WorkbookSnapshot:
        if self.workbook is None:
            raise ValueError("请先打开 Excel")
        sheet = self.workbook.ActiveSheet
        used = sheet.UsedRange
        values = self._matrix_from_range(used.Value)
        formulas = self._matrix_from_range(used.Formula)
        if not values:
            return WorkbookSnapshot(sheet.Name, self.workbook_sheet_names(), [], [], [], [], {}, {}, {}, [])
        headers = [self._header(value, index) for index, value in enumerate(values[0], start=1)]
        data_values = [self._row_dict(headers, row) for row in values[1:] if self._row_has_value(row)]
        data_formulas = [self._row_dict(headers, row) for row in formulas[1:] if self._row_has_value(row)]
        hidden_columns = self.hidden_columns(sheet, headers)
        if self.settings.get("capture_format"):
            formats = self.capture_formats(sheet, headers, len(data_values))
            column_widths = self.capture_column_widths(sheet, headers)
            row_heights = self.capture_row_heights(sheet, len(data_values))
        else:
            formats, column_widths, row_heights = {}, {}, {}
        charts = self.capture_charts(sheet) if self.settings.get("capture_charts") else []
        return WorkbookSnapshot(sheet.Name, self.workbook_sheet_names(), headers, data_values, data_formulas, hidden_columns, formats, column_widths, row_heights, charts)

    def workbook_sheet_names(self) -> list[str]:
        if self.workbook is None:
            return []
        try:
            return [self.workbook.Worksheets(index).Name for index in range(1, self.workbook.Worksheets.Count + 1)]
        except Exception:
            return []

    @staticmethod
    def hidden_columns(sheet, headers: list[str]) -> list[str]:
        hidden = []
        for index, header in enumerate(headers, start=1):
            try:
                if bool(sheet.Columns(index).Hidden):
                    hidden.append(header)
            except Exception:
                continue
        return hidden

    def capture_formats(self, sheet, headers: list[str], row_count: int) -> dict[str, dict[str, Any]]:
        formats: dict[str, dict[str, Any]] = {}
        max_rows = min(row_count + 1, int(self.settings.get("max_capture_rows", 200)))
        for row_index in range(1, max_rows + 1):
            for col_index, header in enumerate(headers, start=1):
                try:
                    cell = sheet.Cells(row_index, col_index)
                    key = f"{row_index}:{header}"
                    formats[key] = {
                        "number_format": str(cell.NumberFormat) if cell.NumberFormat is not None else "",
                        "font_bold": bool(cell.Font.Bold),
                        "font_color": ole_color_to_hex(cell.Font.Color),
                        "fill_color": ole_color_to_hex(cell.Interior.Color),
                        "horizontal_alignment": int(cell.HorizontalAlignment) if cell.HorizontalAlignment is not None else None,
                    }
                except Exception:
                    continue
        return formats

    @staticmethod
    def capture_column_widths(sheet, headers: list[str]) -> dict[str, float]:
        widths = {}
        for index, header in enumerate(headers, start=1):
            try:
                widths[header] = float(sheet.Columns(index).ColumnWidth)
            except Exception:
                continue
        return widths

    def capture_row_heights(self, sheet, row_count: int) -> dict[int, float]:
        heights = {}
        for index in range(1, min(row_count + 1, int(self.settings.get("max_capture_rows", 200))) + 1):
            try:
                heights[index] = float(sheet.Rows(index).RowHeight)
            except Exception:
                continue
        return heights

    @staticmethod
    def capture_charts(sheet) -> list[dict[str, Any]]:
        charts = []
        try:
            chart_objects = sheet.ChartObjects()
            for index in range(1, chart_objects.Count + 1):
                obj = chart_objects.Item(index)
                chart = obj.Chart
                source = ""
                try:
                    source = chart.SeriesCollection(1).Formula
                except Exception:
                    pass
                charts.append(
                    {
                        "name": str(obj.Name),
                        "chart_type": int(chart.ChartType),
                        "source_formula": source,
                        "left": float(obj.Left),
                        "top": float(obj.Top),
                        "width": float(obj.Width),
                        "height": float(obj.Height),
                    }
                )
        except Exception:
            pass
        return charts

    @staticmethod
    def _matrix_from_range(value: Any) -> list[list[Any]]:
        if value is None:
            return []
        if not isinstance(value, tuple):
            return [[value]]
        if value and not isinstance(value[0], tuple):
            return [list(value)]
        return [list(row) for row in value]

    @staticmethod
    def _header(value: Any, index: int) -> str:
        text = "" if value is None else str(value).strip()
        return text or f"列{index}"

    @staticmethod
    def _row_dict(headers: list[str], row: list[Any]) -> dict[str, Any]:
        return {header: row[index] if index < len(row) else None for index, header in enumerate(headers)}

    @staticmethod
    def _row_has_value(row: list[Any]) -> bool:
        return any(value not in (None, "") for value in row)

    def capture_rules(self) -> None:
        if self.baseline is None:
            messagebox.showwarning(APP_TITLE, "请先打开 Excel 并开始录制")
            return
        try:
            current = self.snapshot_active_sheet()
            self.rules = self.rules_from_snapshots(self.baseline, current)
            self.capture_excel_filter_and_sort(current)
            self.refresh_rule_list()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"捕获失败：{exc}")
            return
        self.status_var.set(f"已从当前 Excel 状态生成 {self.rule_count()} 条规则。可以导出方案或直接执行。")

    def rules_from_snapshots(self, before: WorkbookSnapshot, after: WorkbookSnapshot) -> dict[str, list[dict[str, Any]]]:
        rules = self.empty_rules()
        before_headers = before.headers
        after_headers = after.headers
        if self.settings.get("capture_workbook"):
            for name in before.sheet_names:
                if name not in after.sheet_names:
                    rules["workbook_rules"].append({"action": "delete_sheet", "sheet": name, "_text": f"删除 Sheet：{name}"})
            for name in after.sheet_names:
                if name not in before.sheet_names:
                    rules["workbook_rules"].append({"action": "add_sheet", "sheet": name, "_text": f"新增 Sheet：{name}"})

        if not self.settings.get("capture_data"):
            if self.settings.get("capture_format"):
                self.capture_visual_differences(rules, before, after)
            if self.settings.get("capture_charts"):
                self.capture_chart_differences(rules, before, after)
            return rules

        renamed = self.detect_renamed_columns(before_headers, after_headers)
        renamed_old = {item["old_field"] for item in renamed}
        renamed_new = {item["new_field"] for item in renamed}

        for item in renamed:
            rules["operation_rules"].append({"action": "rename_column", **item, "_text": f"重命名列：{item['old_field']} -> {item['new_field']}"})

        hidden_removed = [header for header in after.hidden_columns if header in before_headers]
        for header in hidden_removed:
            rules["operation_rules"].append({"action": "drop_columns", "fields": [header], "_text": f"隐藏/删除列：{header}"})

        removed = [header for header in before_headers if header not in after_headers and header not in renamed_old and header not in hidden_removed]
        if removed:
            rules["operation_rules"].append({"action": "drop_columns", "fields": removed, "_text": f"删除列：{', '.join(removed)}"})

        added = [header for header in after_headers if header not in before_headers and header not in renamed_new]
        for header in added:
            formula = self.first_formula(after, header)
            if formula:
                rules["excel_formula_rules"].append({"field_name": header, "excel_formula": self.template_formula(formula, after_headers), "_text": f"新增公式列：{header}"})
                continue
            constant = self.constant_column_value(after, header)
            if constant is not None:
                rules["operation_rules"].append({"action": "add_constant_column", "field": header, "value": constant, "_text": f"新增固定值列：{header} = {constant}"})
            else:
                rules["operation_rules"].append({"action": "add_empty_column", "field": header, "_text": f"新增空列：{header}"})

        shared_headers = [header for header in before_headers if header in after_headers and header not in hidden_removed]
        for header in shared_headers:
            before_formula = self.first_formula(before, header)
            after_formula = self.first_formula(after, header)
            if after_formula and before_formula != after_formula:
                rules["excel_formula_rules"].append({"field_name": header, "excel_formula": self.template_formula(after_formula, after_headers), "_text": f"修改公式列：{header}"})
                if "!" in after_formula and self.settings.get("capture_cross_sheet"):
                    rules["cross_sheet_rules"].append({"field_name": header, "formula": after_formula, "_text": f"跨 Sheet 公式：{header}"})

        comparable_rows = min(len(before.values), len(after.values), 500)
        for row_index in range(comparable_rows):
            for header in shared_headers:
                old_formula = before.formulas[row_index].get(header) if row_index < len(before.formulas) else None
                new_formula = after.formulas[row_index].get(header) if row_index < len(after.formulas) else None
                if old_formula != new_formula and isinstance(new_formula, str) and new_formula.startswith("="):
                    continue
                old = before.values[row_index].get(header)
                new = after.values[row_index].get(header)
                if old != new:
                    rules["cell_edit_rules"].append({"row_index": row_index, "field": header, "value": new, "_text": f"修改第 {row_index + 2} 行 {header}"})

        if after_headers != before_headers:
            kept = [header for header in after_headers if header in before_headers or header in renamed_new]
            if kept:
                before_kept = [header for header in before_headers if header in kept]
                if kept != before_kept:
                    rules["operation_rules"].append({"action": "select_columns", "fields": kept, "_text": "调整/保留列顺序"})

        if self.settings.get("capture_format"):
            self.capture_visual_differences(rules, before, after)
        if self.settings.get("capture_charts"):
            self.capture_chart_differences(rules, before, after)
        return rules

    @staticmethod
    def capture_visual_differences(rules: dict[str, list[dict[str, Any]]], before: WorkbookSnapshot, after: WorkbookSnapshot) -> None:
        width_changes = {}
        for header, width in after.column_widths.items():
            if before.column_widths.get(header) != width:
                width_changes[header] = width
        if width_changes:
            rules["visual_rules"].append({"action": "set_column_widths", "widths": width_changes, "_text": f"列宽调整：{len(width_changes)} 列"})

        row_height_changes = {}
        for row_index, height in after.row_heights.items():
            if before.row_heights.get(row_index) != height:
                row_height_changes[row_index] = height
        if row_height_changes:
            rules["visual_rules"].append({"action": "set_row_heights", "heights": row_height_changes, "_text": f"行高调整：{len(row_height_changes)} 行"})

        style_changes = []
        for key, style in after.formats.items():
            if before.formats.get(key) != style:
                row_text, field = key.split(":", 1)
                style_changes.append({"row": int(row_text), "field": field, "style": style})
        if style_changes:
            rules["visual_rules"].append({"action": "set_cell_styles", "changes": style_changes[:1000], "_text": f"格式修改：{len(style_changes)} 处"})

    @staticmethod
    def capture_chart_differences(rules: dict[str, list[dict[str, Any]]], before: WorkbookSnapshot, after: WorkbookSnapshot) -> None:
        before_names = {chart.get("name") for chart in before.charts}
        for chart in after.charts:
            if chart.get("name") not in before_names:
                copied = dict(chart)
                copied["_text"] = f"新增图表：{chart.get('name')}"
                rules["chart_rules"].append(copied)

    @staticmethod
    def detect_renamed_columns(before_headers: list[str], after_headers: list[str]) -> list[dict[str, str]]:
        renamed = []
        for index, old in enumerate(before_headers):
            if index >= len(after_headers):
                continue
            new = after_headers[index]
            if old == new:
                continue
            if old not in after_headers and new not in before_headers:
                renamed.append({"old_field": old, "new_field": new})
        return renamed

    @staticmethod
    def constant_column_value(snapshot: WorkbookSnapshot, header: str) -> Any:
        values = [row.get(header) for row in snapshot.values if row.get(header) not in (None, "")]
        if not values:
            return None
        first = values[0]
        return first if all(value == first for value in values[:50]) else None

    @staticmethod
    def first_formula(snapshot: WorkbookSnapshot, header: str) -> str:
        for row in snapshot.formulas:
            value = row.get(header)
            if isinstance(value, str) and value.startswith("="):
                return value
        return ""

    @staticmethod
    def template_formula(formula: str, headers: list[str]) -> str:
        result = formula
        for index, header in enumerate(headers, start=1):
            col = column_letter(index)
            result = result.replace(f"{col}2", "{" + header + "}")
        return result

    def capture_excel_filter_and_sort(self, current: WorkbookSnapshot) -> None:
        if self.workbook is None:
            return
        sheet = self.workbook.ActiveSheet
        self.capture_filters(sheet, current.headers)
        self.capture_sort(sheet, current.headers)

    def capture_filters(self, sheet, headers: list[str]) -> None:
        try:
            auto_filter = sheet.AutoFilter
            filters = auto_filter.Filters
        except Exception:
            return
        for index, header in enumerate(headers, start=1):
            try:
                item = filters.Item(index)
                if not item.On:
                    continue
                criteria = str(item.Criteria1)
                operator = "contains" if "*" in criteria else "equals"
                value = criteria.replace("=", "").replace("*", "")
                self.rules["filter_rules"].append({"field": header, "operator": operator, "value": value, "_text": f"筛选 {header} = {value}"})
            except Exception:
                continue

    def capture_sort(self, sheet, headers: list[str]) -> None:
        try:
            sort_fields = sheet.Sort.SortFields
            count = sort_fields.Count
        except Exception:
            return
        for index in range(1, count + 1):
            try:
                field = sort_fields.Item(index)
                column_index = field.Key.Column
                header = headers[column_index - 1]
                order = "desc" if int(field.Order) == 2 else "asc"
                self.rules["sort_rules"].append({"field": header, "order": order, "_text": f"排序 {header} {'降序' if order == 'desc' else '升序'}"})
            except Exception:
                continue

    def reset_baseline(self) -> None:
        if self.workbook is None:
            messagebox.showwarning(APP_TITLE, "请先打开 Excel")
            return
        self.baseline = self.snapshot_active_sheet()
        self.rules = self.empty_rules()
        self.refresh_rule_list()
        self.status_var.set("已把当前 Excel 状态设为新的录制起点。")

    def clear_rules(self) -> None:
        self.rules = self.empty_rules()
        self.refresh_rule_list()
        self.status_var.set("规则已清空，Excel 文件不会被改动。")

    def refresh_rule_list(self) -> None:
        self.rule_list.delete(0, tk.END)
        labels = {
            "cell_edit_rules": "单元格",
            "operation_rules": "列处理",
            "filter_rules": "筛选",
            "sort_rules": "排序",
            "calculated_fields": "计算",
            "excel_formula_rules": "公式",
            "group_rules": "汇总",
            "visual_rules": "格式",
            "chart_rules": "图表",
            "workbook_rules": "工作簿",
            "cross_sheet_rules": "Sheet联动",
        }
        for bucket, items in self.rules.items():
            for rule in items:
                self.rule_list.insert(tk.END, f"{labels.get(bucket, bucket)} · {rule.get('_text', rule)}")

    def rule_count(self) -> int:
        return sum(len(items) for items in self.rules.values())

    def scheme_payload(self) -> dict[str, Any]:
        if not self.file_path:
            raise ValueError("请先打开 Excel")
        sheet_name = self.baseline.sheet_name if self.baseline else "Sheet1"
        try:
            preview_columns = load_workbook_preview(self.file_path, sheet_name=sheet_name).columns
        except Exception:
            preview_columns = [
                {"name": header, "type": "unknown"}
                for header in (self.baseline.headers if self.baseline else [])
            ]
        config = {
            "input_sheet": sheet_name,
            "header_row": 1,
            "field_mappings": [
                {"source_column": item["name"], "standard_field": item["name"], "display_name": item["name"], "type": item["type"], "required": False, "aliases": [item["name"]]}
                for item in preview_columns
            ],
            "validation_rules": [],
            "cell_edit_rules": self.clean_rules("cell_edit_rules"),
            "operation_rules": self.clean_rules("operation_rules"),
            "filter_rules": self.clean_rules("filter_rules"),
            "sort_rules": self.clean_rules("sort_rules"),
            "group_rules": self.clean_rules("group_rules"),
            "calculated_fields": self.clean_rules("calculated_fields"),
            "excel_formula_rules": self.clean_rules("excel_formula_rules"),
            "visual_rules": self.clean_rules("visual_rules"),
            "chart_rules": self.clean_rules("chart_rules"),
            "workbook_rules": self.clean_rules("workbook_rules"),
            "cross_sheet_rules": self.clean_rules("cross_sheet_rules"),
            "exception_rules": [],
            "output_config": {"file_type": "xlsx"},
        }
        return {
            "export_type": "reportflow_excel_native_scheme",
            "export_version": "2.0",
            "scheme_name": self.file_path.stem,
            "config_json": config,
        }

    def clean_rules(self, bucket: str) -> list[dict[str, Any]]:
        return [{key: value for key, value in rule.items() if key != "_text"} for rule in self.rules[bucket]]

    def export_scheme(self) -> None:
        try:
            payload = self.scheme_payload()
        except Exception as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        path = filedialog.asksaveasfilename(title="导出方案", defaultextension=".json", filetypes=[("方案 JSON", "*.json")])
        if not path:
            return
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.status_var.set(f"方案已导出：{Path(path).name}")

    def import_scheme(self) -> None:
        path = filedialog.askopenfilename(title="导入方案", filetypes=[("方案 JSON", "*.json")])
        if not path:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            config = payload.get("config_json") or {}
            self.rules = self.empty_rules()
            for bucket in self.rules:
                self.rules[bucket] = [self.with_text(bucket, rule) for rule in config.get(bucket, [])]
            self.refresh_rule_list()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"导入失败：{exc}")
            return
        self.status_var.set(f"方案已导入：{Path(path).name}")

    @staticmethod
    def with_text(bucket: str, rule: dict[str, Any]) -> dict[str, Any]:
        item = dict(rule)
        if bucket == "filter_rules":
            item["_text"] = f"{rule.get('field')} {rule.get('operator')} {rule.get('value')}"
        elif bucket == "sort_rules":
            item["_text"] = f"{rule.get('field')} {rule.get('order')}"
        elif bucket == "operation_rules":
            if rule.get("action") == "rename_column":
                item["_text"] = f"rename {rule.get('old_field')} -> {rule.get('new_field')}"
            else:
                item["_text"] = f"{rule.get('action')} {rule.get('fields') or rule.get('field')}"
        elif bucket == "excel_formula_rules":
            item["_text"] = f"{rule.get('field_name')} = {rule.get('excel_formula')}"
        elif bucket == "visual_rules":
            item["_text"] = f"{rule.get('action')}"
        elif bucket == "chart_rules":
            item["_text"] = f"{rule.get('name') or rule.get('chart_type')}"
        elif bucket == "workbook_rules":
            item["_text"] = f"{rule.get('action')} {rule.get('sheet')}"
        elif bucket == "cross_sheet_rules":
            item["_text"] = f"{rule.get('field_name')} = {rule.get('formula')}"
        else:
            item["_text"] = str(rule)
        return item

    def execute_current_scheme(self) -> None:
        if not self.file_path:
            messagebox.showwarning(APP_TITLE, "请先打开 Excel 文件")
            return
        output = filedialog.asksaveasfilename(title="保存结果 Excel", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not output:
            return
        try:
            source = self.execution_source_path()
            result = execute_scheme(source, self.scheme_payload(), output)
            if source != self.file_path and not self.settings.get("keep_temp_xlsx"):
                try:
                    source.unlink(missing_ok=True)
                except Exception:
                    pass
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"执行失败：{exc}")
            return
        self.status_var.set(f"执行完成：{result['detail_rows']} 行，已保存到 {Path(output).name}")
        if self.settings.get("auto_open_result"):
            try:
                os.startfile(output)
            except Exception:
                pass
        messagebox.showinfo(APP_TITLE, f"执行完成\n结果文件：{output}")

    def execution_source_path(self) -> Path:
        if not self.file_path:
            raise ValueError("请先打开 Excel 文件")
        if self.file_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            return self.file_path
        if self.workbook is None:
            raise ValueError("当前格式需要通过 Excel/WPS 另存为 xlsx 后执行")
        temp_path = self.file_path.with_suffix(".reportflow_temp.xlsx")
        try:
            self.workbook.SaveAs(str(temp_path), 51)
        except Exception:
            self.workbook.SaveCopyAs(str(temp_path))
        return temp_path


def column_letter(index: int) -> str:
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label


def ole_color_to_hex(value: Any) -> str:
    try:
        number = int(value)
    except Exception:
        return ""
    if number < 0:
        return ""
    red = number & 255
    green = (number >> 8) & 255
    blue = (number >> 16) & 255
    return f"{red:02X}{green:02X}{blue:02X}"


def formula_from_description(description: str, headers: list[str]) -> tuple[str, str]:
    text = description.lower()
    refs = header_refs(headers)

    def pick(*keywords: str) -> str:
        for keyword in keywords:
            for header, ref in refs.items():
                if keyword and (keyword.lower() in header.lower() or header.lower() in keyword.lower()):
                    return ref
        for keyword in keywords:
            if keyword in text:
                for header, ref in refs.items():
                    if keyword in header.lower():
                        return ref
        return ""

    amount = pick("金额", "费用", "销售额", "收入")
    qty = pick("数量", "入库", "出库", "库存", "完成数量", "计划数量")
    status = pick("状态", "完成状态", "处理状态")
    start_date = pick("开始", "创建", "下单", "发起")
    end_date = pick("结束", "完成", "关闭")

    if any(word in text for word in ("完成率", "达成率", "比例", "占比")):
        done = pick("完成", "完成数量", "实际")
        plan = pick("计划", "计划数量", "目标")
        if done and plan:
            return f'=IFERROR({done}/{plan},0)', "完成率/达成率，自动避免除零错误。"
        return "=IFERROR(实际完成/计划目标,0)", "未识别到对应列名，请把“实际完成”和“计划目标”替换成单元格。"

    if any(word in text for word in ("库存", "结存", "剩余")) and any(word in text for word in ("入库", "出库", "收入", "发出")):
        inbound = pick("入库", "收入", "收")
        outbound = pick("出库", "发出", "发")
        opening = pick("期初", "上月结存")
        if opening and inbound and outbound:
            return f"={opening}+{inbound}-{outbound}", "期末库存 = 期初 + 入库 - 出库。"
        if inbound and outbound:
            return f"={inbound}-{outbound}", "库存变化 = 入库 - 出库。"

    if any(word in text for word in ("是否完成", "完成显示", "完成则", "状态判断")):
        field = status or pick("完成")
        if field:
            return f'=IF({field}="完成",1,0)', "根据状态列判断，完成返回 1，否则返回 0。"
        return '=IF(A2="完成",1,0)', "未识别状态列，默认使用 A2。"

    if any(word in text for word in ("日期差", "天数", "耗时", "周期", "相差")):
        if start_date and end_date:
            return f"={end_date}-{start_date}", "计算两个日期之间相差天数。"
        return "=结束日期-开始日期", "未识别日期列，请替换为实际单元格。"

    if any(word in text for word in ("求和", "合计", "总和", "sum")):
        field = amount or qty or "A:A"
        return f"=SUM({field})", "对识别到的列求和；如果是整列引用，可按需改为具体区域。"

    if any(word in text for word in ("平均", "均值", "avg", "average")):
        field = amount or qty or "A:A"
        return f"=AVERAGE({field})", "计算平均值。"

    if any(word in text for word in ("最大", "最高", "max")):
        field = amount or qty or "A:A"
        return f"=MAX({field})", "计算最大值。"

    if any(word in text for word in ("最小", "最低", "min")):
        field = amount or qty or "A:A"
        return f"=MIN({field})", "计算最小值。"

    if any(word in text for word in ("计数", "数量", "多少", "count")):
        field = status or qty or "A:A"
        return f"=COUNTA({field})", "统计非空数量。"

    if any(word in text for word in ("包含", "关键字", "文本")):
        field = status or "A2"
        return f'=IF(ISNUMBER(SEARCH("关键字",{field})),1,0)', "判断单元格是否包含关键字。"

    if any(word in text for word in ("同比", "环比", "增长率")):
        current = pick("本期", "本月", "今年", "当前")
        previous = pick("上期", "上月", "去年", "同期")
        if current and previous:
            return f'=IFERROR(({current}-{previous})/{previous},0)', "增长率 = (本期 - 上期) / 上期。"
        return "=IFERROR((本期-上期)/上期,0)", "未识别本期/上期列，请替换为实际单元格。"

    return "=IFERROR(公式主体,\"\")", "暂未完全识别需求，可在生成结果基础上替换字段。建议描述里写清楚列名和判断条件。"


def header_refs(headers: list[str]) -> dict[str, str]:
    result = {}
    for index, header in enumerate(headers, start=1):
        result[header] = f"{column_letter(index)}2"
    return result


def main() -> None:
    app = ExcelNativeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
