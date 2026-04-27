from __future__ import annotations

import json
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


@dataclass
class WorkbookSnapshot:
    sheet_name: str
    headers: list[str]
    values: list[dict[str, Any]]
    formulas: list[dict[str, Any]]


class ExcelNativeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x720")
        self.minsize(860, 640)
        self.configure(bg="#0f1115")

        self.excel = None
        self.workbook = None
        self.file_path: Path | None = None
        self.baseline: WorkbookSnapshot | None = None
        self.rules = self.empty_rules()

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
        }

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("Microsoft YaHei UI", 10), background="#0f1115", foreground="#f4f4f0")
        style.configure("Root.TFrame", background="#0f1115")
        style.configure("Panel.TFrame", background="#171a21")
        style.configure("Title.TLabel", background="#0f1115", foreground="#f4f4f0", font=("Segoe UI", 22, "bold"))
        style.configure("Sub.TLabel", background="#0f1115", foreground="#9aa0a6")
        style.configure("PanelTitle.TLabel", background="#171a21", foreground="#f4f4f0", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Muted.TLabel", background="#171a21", foreground="#9aa0a6")
        style.configure("Accent.TButton", background="#10a37f", foreground="#ffffff", borderwidth=0, padding=(14, 10))
        style.map("Accent.TButton", background=[("active", "#0b7f63")])
        style.configure("Ghost.TButton", background="#20242c", foreground="#f4f4f0", borderwidth=0, padding=(14, 9))
        style.map("Ghost.TButton", background=[("active", "#2a303a")])

    def _build_layout(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=(18, 18, 18, 18))
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=0)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(2, weight=1)

        ttk.Label(root, text="ReportFlow", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(root, text="用真正的 Excel 操作，ReportFlow 负责录制、函数生成和复用规则", style="Sub.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 16))

        left = ttk.Frame(root, style="Root.TFrame", width=380)
        left.grid(row=2, column=0, sticky="nsew", padx=(0, 14))
        left.grid_propagate(False)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        left.rowconfigure(2, weight=1)

        actions = ttk.Frame(left, style="Panel.TFrame", padding=(14, 14))
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)

        ttk.Button(actions, text="打开 Excel 并开始录制", style="Accent.TButton", command=self.open_excel).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Button(actions, text="捕获当前操作为规则", style="Ghost.TButton", command=self.capture_rules).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Button(actions, text="重新设为起点", style="Ghost.TButton", command=self.reset_baseline).grid(row=2, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(actions, text="清空规则", style="Ghost.TButton", command=self.clear_rules).grid(row=2, column=1, sticky="ew", padx=(5, 0))
        ttk.Button(actions, text="执行并生成结果", style="Accent.TButton", command=self.execute_current_scheme).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        rules_panel = ttk.Frame(left, style="Panel.TFrame", padding=(14, 14))
        rules_panel.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        rules_panel.rowconfigure(1, weight=1)
        rules_panel.columnconfigure(0, weight=1)
        ttk.Label(rules_panel, text="已生成规则", style="PanelTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.rule_list = tk.Listbox(
            rules_panel,
            bg="#11141a",
            fg="#f4f4f0",
            selectbackground="#10a37f",
            selectforeground="#ffffff",
            borderwidth=0,
            highlightthickness=0,
            activestyle="none",
            font=("Microsoft YaHei UI", 10),
        )
        self.rule_list.grid(row=1, column=0, columnspan=2, sticky="nsew")
        ttk.Button(rules_panel, text="一键加载规则", style="Ghost.TButton", command=self.import_scheme).grid(row=2, column=0, sticky="ew", padx=(0, 5), pady=(10, 0))
        ttk.Button(rules_panel, text="导出规则", style="Ghost.TButton", command=self.export_scheme).grid(row=2, column=1, sticky="ew", padx=(5, 0), pady=(10, 0))

        self._build_formula_panel(left)

        settings = ttk.Frame(left, style="Root.TFrame")
        settings.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        settings.columnconfigure(0, weight=1)
        settings_button = ttk.Menubutton(settings, text="设置", style="Ghost.TButton")
        settings_button.grid(row=0, column=0, sticky="sw")
        settings_menu = tk.Menu(settings_button, tearoff=False, bg="#171a21", fg="#f4f4f0", activebackground="#10a37f", activeforeground="#ffffff")
        settings_menu.add_command(label="操作文档", command=self.show_operation_docs)
        settings_button["menu"] = settings_menu

        right = ttk.Frame(root, style="Root.TFrame")
        right.grid(row=2, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        status_panel = ttk.Frame(right, style="Panel.TFrame", padding=(18, 18))
        status_panel.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        status_panel.columnconfigure(0, weight=1)
        ttk.Label(status_panel, text="当前状态", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.status_var = tk.StringVar(value="先打开一个 Excel 文件。之后直接在 Excel 里筛选、排序、删列、写公式。")
        ttk.Label(status_panel, textvariable=self.status_var, style="Muted.TLabel", wraplength=500).grid(row=1, column=0, sticky="ew")

        guide_panel = ttk.Frame(right, style="Panel.TFrame", padding=(18, 18))
        guide_panel.grid(row=1, column=0, sticky="nsew")
        guide_panel.columnconfigure(0, weight=1)
        guide_panel.rowconfigure(1, weight=1)
        ttk.Label(guide_panel, text="工作流", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        guide = tk.Text(guide_panel, bg="#11141a", fg="#d5d9de", borderwidth=0, highlightthickness=0, font=("Microsoft YaHei UI", 11), wrap=tk.WORD)
        guide.grid(row=1, column=0, sticky="nsew")
        guide.insert("1.0", self.operation_docs(short=True))
        guide.configure(state=tk.DISABLED)

    def _build_formula_panel(self, parent: ttk.Frame) -> None:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=(14, 14))
        panel.grid(row=2, column=0, sticky="nsew")
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
                "   直接筛选、排序、删除列、改单元格、写公式列。\n\n"
                "3. 捕获当前操作为规则\n"
                "   左侧会显示识别到的规则。\n\n"
                "4. 导出或一键加载规则\n"
                "   规则文件可以给不同用户复用。\n\n"
                "5. 函数查询/生成\n"
                "   在左侧输入需求，生成公式后复制或写入 Excel 当前单元格。"
            )
        return (
            "ReportFlow 操作文档\n\n"
            "一、创建规则\n"
            "1. 点击“打开 Excel 并开始录制”。\n"
            "2. 选择需要处理的 Excel 文件。\n"
            "3. ReportFlow 会打开真正的 Microsoft Excel。\n"
            "4. 在 Excel 中按平时习惯操作，例如筛选、排序、删除列、修改单元格、新增公式列。\n"
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
            "- 调整/保留列\n"
            "- 单元格修改\n"
            "- 新增公式列\n"
            "- 自动筛选条件\n"
            "- 排序字段\n\n"
            "五、注意事项\n"
            "- 录制时请保持目标工作表为当前激活 Sheet。\n"
            "- 规则复用依赖列名，建议同类报表保持表头一致。\n"
            "- 函数生成会优先根据当前 Excel 表头猜测单元格引用，复杂公式仍需要人工确认。"
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
            messagebox.showerror(APP_TITLE, "需要安装 pywin32 才能调用 Microsoft Excel。\n请执行：pip install pywin32")
            return
        path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls")])
        if not path:
            return
        self.file_path = Path(path)
        try:
            pythoncom.CoInitialize()
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = True
            self.excel.DisplayAlerts = False
            self.workbook = self.excel.Workbooks.Open(str(self.file_path))
            self.baseline = self.snapshot_active_sheet()
            self.rules = self.empty_rules()
            self.refresh_rule_list()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"打开 Excel 失败：{exc}")
            return
        self.status_var.set(f"正在录制：{self.file_path.name}\n请在 Excel 里正常操作，完成后点“捕获当前操作为规则”。")

    def snapshot_active_sheet(self) -> WorkbookSnapshot:
        if self.workbook is None:
            raise ValueError("请先打开 Excel")
        sheet = self.workbook.ActiveSheet
        used = sheet.UsedRange
        values = self._matrix_from_range(used.Value)
        formulas = self._matrix_from_range(used.Formula)
        if not values:
            return WorkbookSnapshot(sheet.Name, [], [], [])
        headers = [self._header(value, index) for index, value in enumerate(values[0], start=1)]
        data_values = [self._row_dict(headers, row) for row in values[1:] if self._row_has_value(row)]
        data_formulas = [self._row_dict(headers, row) for row in formulas[1:] if self._row_has_value(row)]
        return WorkbookSnapshot(sheet.Name, headers, data_values, data_formulas)

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

        removed = [header for header in before_headers if header not in after_headers]
        if removed:
            rules["operation_rules"].append({"action": "drop_columns", "fields": removed, "_text": f"删除列：{', '.join(removed)}"})

        added = [header for header in after_headers if header not in before_headers]
        for header in added:
            formula = self.first_formula(after, header)
            if formula:
                rules["excel_formula_rules"].append({"field_name": header, "excel_formula": self.template_formula(formula, after_headers), "_text": f"新增公式列：{header}"})

        shared_headers = [header for header in before_headers if header in after_headers]
        comparable_rows = min(len(before.values), len(after.values), 500)
        for row_index in range(comparable_rows):
            for header in shared_headers:
                old = before.values[row_index].get(header)
                new = after.values[row_index].get(header)
                if old != new:
                    rules["cell_edit_rules"].append({"row_index": row_index, "field": header, "value": new, "_text": f"修改第 {row_index + 2} 行 {header}"})

        if after_headers != before_headers and not removed and not added:
            kept = [header for header in after_headers if header in before_headers]
            if kept:
                rules["operation_rules"].append({"action": "select_columns", "fields": kept, "_text": "调整/保留列顺序"})

        return rules

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
        preview = load_workbook_preview(self.file_path, sheet_name=sheet_name)
        config = {
            "input_sheet": sheet_name,
            "header_row": 1,
            "field_mappings": [
                {"source_column": item["name"], "standard_field": item["name"], "display_name": item["name"], "type": item["type"], "required": False, "aliases": [item["name"]]}
                for item in preview.columns
            ],
            "validation_rules": [],
            "cell_edit_rules": self.clean_rules("cell_edit_rules"),
            "operation_rules": self.clean_rules("operation_rules"),
            "filter_rules": self.clean_rules("filter_rules"),
            "sort_rules": self.clean_rules("sort_rules"),
            "group_rules": self.clean_rules("group_rules"),
            "calculated_fields": self.clean_rules("calculated_fields"),
            "excel_formula_rules": self.clean_rules("excel_formula_rules"),
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
            item["_text"] = f"{rule.get('action')} {rule.get('fields') or rule.get('field')}"
        elif bucket == "excel_formula_rules":
            item["_text"] = f"{rule.get('field_name')} = {rule.get('excel_formula')}"
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
            result = execute_scheme(self.file_path, self.scheme_payload(), output)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"执行失败：{exc}")
            return
        self.status_var.set(f"执行完成：{result['detail_rows']} 行，已保存到 {Path(output).name}")
        messagebox.showinfo(APP_TITLE, f"执行完成\n结果文件：{output}")


def column_letter(index: int) -> str:
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label


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
