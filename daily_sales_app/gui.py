from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .config_store import ConfigStore
from .processor import ProcessingResult, process_daily_sales


NO_ORDER_SENTINEL = "__NO_ORDER__"


class DailySalesApp(tk.Tk):
    def __init__(self, store: ConfigStore, app_dir: Path) -> None:
        super().__init__()
        self.store = store
        self.app_dir = app_dir
        self.title("每日销售数据汇总")
        self.geometry("1180x760")
        self.minsize(1080, 700)

        self.platform_files: dict[str, str] = {}
        self.platform_index = 0

        self.status_text = tk.StringVar(value="请先确认 SKU 对应表，再按顺序上传平台文件。")
        self.current_platform_text = tk.StringVar()
        self.current_platform_path = tk.StringVar()
        self.sku_path_var = tk.StringVar(value=self.store.state.get("sku_mapping_path", ""))
        self.cost_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.rate_var = tk.StringVar(value=str(self.store.state.get("last_eur_to_cny_rate", 8.0)))

        self.platform_name_var = tk.StringVar()
        self.operator_name_var = tk.StringVar()
        self.requires_cost_var = tk.BooleanVar(value=False)

        self._build_ui()
        self._refresh_platforms(reset_index=True)
        self._refresh_config_table()

    @staticmethod
    def normalize_operator_input(value: str) -> str:
        names = []
        for chunk in value.replace(";", "；").split("；"):
            name = chunk.strip()
            if name and name not in names:
                names.append(name)
        return "；".join(names)

    @property
    def platform_configs(self) -> list[dict]:
        return self.store.get_platforms()

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)

        main_tab = ttk.Frame(notebook)
        config_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="每日数据")
        notebook.add(config_tab, text="运营对应店铺")

        self._build_main_tab(main_tab)
        self._build_config_tab(config_tab)

    def _build_main_tab(self, parent: ttk.Frame) -> None:
        top = ttk.Frame(parent)
        top.pack(fill="x", pady=(0, 10))

        ttk.Label(top, text="SKU 对应表").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.sku_path_var, width=88).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(top, text="重新上传 SKU 表", command=self.choose_sku_mapping).grid(row=1, column=1, sticky="ew")
        top.columnconfigure(0, weight=1)

        upload_frame = ttk.LabelFrame(parent, text="按平台顺序上传")
        upload_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(upload_frame, textvariable=self.current_platform_text, font=("Microsoft YaHei UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 6), columnspan=3
        )
        ttk.Entry(upload_frame, textvariable=self.current_platform_path, width=88).grid(
            row=1, column=0, sticky="ew", padx=(12, 8), pady=(0, 12)
        )
        ttk.Button(upload_frame, text="选择当前平台文件", command=self.choose_current_platform_file).grid(
            row=1, column=1, sticky="ew", padx=(0, 8), pady=(0, 12)
        )
        ttk.Button(upload_frame, text="当天无订单", command=self.mark_current_platform_no_order).grid(
            row=1, column=2, sticky="ew", padx=(0, 8), pady=(0, 12)
        )
        ttk.Button(upload_frame, text="上一步", command=self.go_previous_platform).grid(
            row=1, column=3, sticky="ew", padx=(0, 12), pady=(0, 12)
        )
        upload_frame.columnconfigure(0, weight=1)

        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill="both", expand=False, pady=(0, 10))

        ttk.Label(progress_frame, text="上传进度").pack(anchor="w")
        columns = ("platform", "operator", "requires_cost", "path")
        self.progress_tree = ttk.Treeview(progress_frame, columns=columns, show="headings", height=8)
        self.progress_tree.heading("platform", text="平台")
        self.progress_tree.heading("operator", text="运营")
        self.progress_tree.heading("requires_cost", text="成本模式")
        self.progress_tree.heading("path", text="已选文件")
        self.progress_tree.column("platform", width=120, anchor="center")
        self.progress_tree.column("operator", width=240, anchor="center")
        self.progress_tree.column("requires_cost", width=100, anchor="center")
        self.progress_tree.column("path", width=720)
        self.progress_tree.pack(fill="x")

        bottom = ttk.Frame(parent)
        bottom.pack(fill="x", pady=(0, 10))

        ttk.Label(bottom, text="每日成本报表").grid(row=0, column=0, sticky="w")
        ttk.Label(bottom, text="欧元兑人民币汇率").grid(row=0, column=2, sticky="w", padx=(12, 0))
        ttk.Entry(bottom, textvariable=self.cost_path_var, width=72).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(bottom, text="上传成本报表", command=self.choose_cost_report).grid(row=1, column=1, sticky="ew")
        ttk.Entry(bottom, textvariable=self.rate_var, width=12).grid(row=1, column=2, sticky="w", padx=(12, 8))
        ttk.Label(bottom, text="仅在成本表缺少欧元成本时使用").grid(row=1, column=3, sticky="w")

        ttk.Label(bottom, text="导出文件").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(bottom, textvariable=self.output_path_var, width=72).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(bottom, text="选择导出位置", command=self.choose_output_path).grid(row=3, column=1, sticky="ew")
        ttk.Button(bottom, text="生成销售报表", command=self.generate_report).grid(row=3, column=2, sticky="ew", padx=(12, 8))
        bottom.columnconfigure(0, weight=1)

        unmatched_frame = ttk.LabelFrame(parent, text="未匹配 SKU 提示")
        unmatched_frame.pack(fill="both", expand=True)

        unmatched_columns = ("platform", "user_account", "refrence_no")
        self.unmatched_tree = ttk.Treeview(unmatched_frame, columns=unmatched_columns, show="headings", height=12)
        self.unmatched_tree.heading("platform", text="平台")
        self.unmatched_tree.heading("user_account", text="user_account")
        self.unmatched_tree.heading("refrence_no", text="refrence_no")
        self.unmatched_tree.column("platform", width=120, anchor="center")
        self.unmatched_tree.column("user_account", width=220, anchor="center")
        self.unmatched_tree.column("refrence_no", width=260, anchor="center")
        self.unmatched_tree.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(parent, textvariable=self.status_text, foreground="#234").pack(anchor="w")

    def _build_config_tab(self, parent: ttk.Frame) -> None:
        left = ttk.Frame(parent)
        left.pack(side="left", fill="both", expand=True, padx=(0, 12))

        columns = ("platform", "operator", "mode")
        self.config_tree = ttk.Treeview(left, columns=columns, show="headings", height=20)
        self.config_tree.heading("platform", text="平台")
        self.config_tree.heading("operator", text="运营")
        self.config_tree.heading("mode", text="模式")
        self.config_tree.column("platform", width=160, anchor="center")
        self.config_tree.column("operator", width=320, anchor="center")
        self.config_tree.column("mode", width=120, anchor="center")
        self.config_tree.pack(fill="both", expand=True)
        self.config_tree.bind("<<TreeviewSelect>>", self.on_config_select)

        right = ttk.LabelFrame(parent, text="编辑平台")
        right.pack(side="left", fill="y")

        ttk.Button(right, text="新增店铺", command=self.start_add_platform).grid(
            row=0, column=0, sticky="ew", padx=12, pady=(12, 6)
        )

        ttk.Label(right, text="平台名称").grid(row=1, column=0, sticky="w", padx=12, pady=(6, 4))
        ttk.Entry(right, textvariable=self.platform_name_var, width=24).grid(row=2, column=0, sticky="ew", padx=12)

        ttk.Label(right, text="运营名称").grid(row=3, column=0, sticky="w", padx=12, pady=(12, 4))
        ttk.Entry(right, textvariable=self.operator_name_var, width=24).grid(row=4, column=0, sticky="ew", padx=12)

        ttk.Checkbutton(right, text="需要成本报表", variable=self.requires_cost_var).grid(
            row=5, column=0, sticky="w", padx=12, pady=(12, 0)
        )

        ttk.Button(right, text="新增到末尾", command=self.add_platform).grid(row=6, column=0, sticky="ew", padx=12, pady=(18, 6))
        ttk.Button(right, text="更新当前项", command=self.update_selected_platform).grid(row=7, column=0, sticky="ew", padx=12, pady=6)
        ttk.Button(right, text="删除当前项", command=self.delete_selected_platform).grid(row=8, column=0, sticky="ew", padx=12, pady=6)
        ttk.Button(right, text="保存配置", command=self.save_platform_config).grid(row=9, column=0, sticky="ew", padx=12, pady=(18, 6))

        note = (
            "说明:\n"
            "1. 新增平台默认追加到上传顺序末尾。\n"
            "2. 多个运营共同负责同一平台时，用；分隔。\n"
            "3. 需要成本报表的平台，生成前会强制校验。\n"
            "4. 保存后会立即刷新每日上传顺序。"
        )
        ttk.Label(right, text=note, justify="left").grid(row=10, column=0, sticky="w", padx=12, pady=(18, 12))

    def _refresh_platforms(self, reset_index: bool = False) -> None:
        configs = self.platform_configs
        valid_names = {item["name"] for item in configs}
        self.platform_files = {name: path for name, path in self.platform_files.items() if name in valid_names}
        if reset_index:
            self.platform_index = 0
        else:
            self.platform_index = min(self.platform_index, max(len(configs) - 1, 0))
        self._refresh_prompt()
        self._refresh_progress_table()

    def _refresh_prompt(self) -> None:
        configs = self.platform_configs
        if not configs:
            self.current_platform_text.set("请先到“运营对应店铺”中新增平台。")
            self.current_platform_path.set("")
            return
        current = configs[self.platform_index]
        current_name = current["name"]
        self.current_platform_text.set(
            f"当前第 {self.platform_index + 1} / {len(configs)} 个平台: 请上传 {current_name} 文件"
        )
        current_path = self.platform_files.get(current_name, "")
        self.current_platform_path.set("" if current_path == NO_ORDER_SENTINEL else current_path)

    def _refresh_progress_table(self) -> None:
        for item in self.progress_tree.get_children():
            self.progress_tree.delete(item)
        for config in self.platform_configs:
            path = self.platform_files.get(config["name"], "")
            display_path = "当天无订单" if path == NO_ORDER_SENTINEL else (Path(path).name if path else "")
            self.progress_tree.insert(
                "",
                "end",
                values=(
                    config["name"],
                    config["operator"],
                    "需要成本" if config["requires_cost"] else "不需要成本",
                    display_path,
                ),
            )

    def _refresh_config_table(self) -> None:
        for item in self.config_tree.get_children():
            self.config_tree.delete(item)
        for index, config in enumerate(self.platform_configs):
            self.config_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    config["name"],
                    config["operator"],
                    "需要成本" if config["requires_cost"] else "不需要成本",
                ),
            )

    def choose_sku_mapping(self) -> None:
        path = filedialog.askopenfilename(
            title="选择运营对应SKU文件",
            filetypes=[("表格文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.sku_path_var.set(path)
        self.store.state["sku_mapping_path"] = path
        self.store.save()
        self.status_text.set("已更新 SKU 对应表，后续会直接复用这个路径。")

    def choose_current_platform_file(self) -> None:
        configs = self.platform_configs
        if not configs:
            messagebox.showwarning("提示", "请先配置平台。")
            return
        path = filedialog.askopenfilename(
            title="选择平台文件",
            filetypes=[("表格文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        current_name = configs[self.platform_index]["name"]
        self.platform_files[current_name] = path
        self.status_text.set(f"{current_name} 文件已选择。")
        if self.platform_index < len(configs) - 1:
            self.platform_index += 1
        self._refresh_prompt()
        self._refresh_progress_table()

    def mark_current_platform_no_order(self) -> None:
        configs = self.platform_configs
        if not configs:
            messagebox.showwarning("提示", "请先配置平台。")
            return
        current_name = configs[self.platform_index]["name"]
        self.platform_files[current_name] = NO_ORDER_SENTINEL
        self.status_text.set(f"{current_name} 已标记为当天无订单。")
        if self.platform_index < len(configs) - 1:
            self.platform_index += 1
        self._refresh_prompt()
        self._refresh_progress_table()

    def go_previous_platform(self) -> None:
        if self.platform_index == 0:
            self.status_text.set("已经在第一个平台。")
            return
        self.platform_index -= 1
        previous_name = self.platform_configs[self.platform_index]["name"]
        self.current_platform_path.set("")
        self.platform_files.pop(previous_name, None)
        self._refresh_prompt()
        self._refresh_progress_table()
        self.status_text.set(f"已返回 {previous_name}，请重新上传。")

    def choose_cost_report(self) -> None:
        path = filedialog.askopenfilename(
            title="选择每日成本报表",
            filetypes=[("表格文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")],
        )
        if path:
            self.cost_path_var.set(path)

    def choose_output_path(self) -> None:
        initial_dir = self.store.state.get("last_output_dir", "") or str(self.app_dir)
        default_name = f"{(date.today() - timedelta(days=1)):%Y年%m月%d日} 销售情况.xlsx"
        path = filedialog.asksaveasfilename(
            title="选择导出位置",
            initialdir=initial_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        self.output_path_var.set(path)

    def add_platform(self) -> None:
        platform_name = self.platform_name_var.get().strip()
        operator_name = self.normalize_operator_input(self.operator_name_var.get().strip())
        if not platform_name or not operator_name:
            messagebox.showwarning("提示", "平台名称和运营名称都需要填写。")
            return
        self.store.state["platforms"].append(
            {
                "name": platform_name,
                "operator": operator_name,
                "requires_cost": bool(self.requires_cost_var.get()),
            }
        )
        self._refresh_config_table()
        self._clear_config_form()
        self.status_text.set(f"已新增平台 {platform_name}。")

    def update_selected_platform(self) -> None:
        selection = self.config_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选中要更新的平台。")
            return
        platform_name = self.platform_name_var.get().strip()
        operator_name = self.normalize_operator_input(self.operator_name_var.get().strip())
        if not platform_name or not operator_name:
            messagebox.showwarning("提示", "平台名称和运营名称都需要填写。")
            return
        index = int(selection[0])
        self.store.state["platforms"][index] = {
            "name": platform_name,
            "operator": operator_name,
            "requires_cost": bool(self.requires_cost_var.get()),
        }
        self._refresh_config_table()
        self.status_text.set(f"已更新平台 {platform_name}。")

    def delete_selected_platform(self) -> None:
        selection = self.config_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选中要删除的平台。")
            return
        index = int(selection[0])
        del self.store.state["platforms"][index]
        self._refresh_config_table()
        self._clear_config_form()

    def save_platform_config(self) -> None:
        if not self.store.state["platforms"]:
            messagebox.showwarning("提示", "至少需要保留一个平台。")
            return
        for item in self.store.state["platforms"]:
            if not item["name"] or not item["operator"]:
                messagebox.showwarning("提示", "平台名称和运营名称不能为空。")
                return
        self.store.save()
        self._refresh_platforms(reset_index=True)
        self.status_text.set("平台配置已保存，上传顺序已经刷新。")

    def start_add_platform(self) -> None:
        self.config_tree.selection_remove(*self.config_tree.selection())
        self._clear_config_form()
        self.status_text.set("已切换到新增店铺模式。")

    def on_config_select(self, _event: object) -> None:
        selection = self.config_tree.selection()
        if not selection:
            return
        index = int(selection[0])
        config = self.store.state["platforms"][index]
        self.platform_name_var.set(config["name"])
        self.operator_name_var.set(config["operator"])
        self.requires_cost_var.set(bool(config["requires_cost"]))

    def _clear_config_form(self) -> None:
        self.platform_name_var.set("")
        self.operator_name_var.set("")
        self.requires_cost_var.set(False)

    def _validate_before_generate(self) -> tuple[bool, str]:
        sku_path = self.sku_path_var.get().strip()
        if not sku_path:
            return False, "请先上传 SKU 对应表。"
        for config in self.platform_configs:
            if not config["name"] or not config["operator"]:
                return False, "平台配置中存在空的平台或运营名称。"
            if config["name"] not in self.platform_files:
                return False, f"平台 {config['name']} 还没有上传文件。"
        if not self.cost_path_var.get().strip():
            return False, "请上传每日成本报表。"
        if not self.output_path_var.get().strip():
            return False, "请选择导出位置。"
        for config in self.platform_configs:
            if config["requires_cost"] and not config["operator"]:
                return False, f"{config['name']} 被设为需要成本报表，但未绑定运营。"
        return True, ""

    def generate_report(self) -> None:
        valid, message = self._validate_before_generate()
        if not valid:
            messagebox.showwarning("提示", message)
            return

        try:
            rate = float(self.rate_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "欧元兑人民币汇率请填写数字。")
            return

        try:
            self.status_text.set("正在生成销售报表，请稍候...")
            self.update_idletasks()
            result = process_daily_sales(
                platform_configs=self.platform_configs,
                platform_files={
                    name: (NO_ORDER_SENTINEL if path == NO_ORDER_SENTINEL else Path(path))
                    for name, path in self.platform_files.items()
                },
                sku_mapping_path=Path(self.sku_path_var.get().strip()),
                cost_report_path=Path(self.cost_path_var.get().strip()),
                output_path=Path(self.output_path_var.get().strip()),
                eur_to_cny_rate=rate,
            )
        except Exception as exc:
            self.status_text.set("生成失败。")
            messagebox.showerror("出错", str(exc))
            return

        self.store.state["last_output_dir"] = str(Path(self.output_path_var.get().strip()).parent)
        self.store.state["last_eur_to_cny_rate"] = rate
        self.store.save()
        self._show_unmatched_rows(result)
        self.status_text.set(
            f"生成完成: {result.output_path.name}，报表日期为 {result.report_date:%Y-%m-%d}。"
        )
        messagebox.showinfo("完成", f"销售报表已生成:\n{result.output_path}")

    def _show_unmatched_rows(self, result: ProcessingResult) -> None:
        for item in self.unmatched_tree.get_children():
            self.unmatched_tree.delete(item)
        for row in result.unmatched_rows:
            self.unmatched_tree.insert(
                "",
                "end",
                values=(row["platform"], row["user_account"], row["refrence_no"]),
            )
