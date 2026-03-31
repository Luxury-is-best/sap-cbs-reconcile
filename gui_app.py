#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP 与 CBS 桌面对账工具（双模式）：
1) 明细对账（CBS明细 + SAP明细）
2) 总和对账（CBS历史余额 + SAP科余 + 银行科目）
"""

from __future__ import annotations

__version__ = "1.0.0"

import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from backend_reconcile_sap_cbs import (
    load_cbs_detail,
    load_sap_detail,
    reconcile,
    report_as_string,
    write_report,
)
from total_reconcile_service import run_total_reconcile, write_total_report


DETAIL_MODE = "明细对账"
TOTAL_MODE = "总和对账"


class ReconcileGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"SAP和CBS对账助手 v{__version__}")
        self.geometry("760x560")
        self.minsize(700, 520)

        self.mode_var = tk.StringVar(value=DETAIL_MODE)
        self.status_var = tk.StringVar(value="请选择模式并上传文件。")
        self.output_dir_var = tk.StringVar(value=str(Path.cwd()))

        self.file_vars = {
            "detail_cbs": tk.StringVar(),
            "detail_sap": tk.StringVar(),
            "total_cbs": tk.StringVar(),
            "total_sap": tk.StringVar(),
            "total_bank": tk.StringVar(),
        }

        self._build_layout()
        self._refresh_mode_view()

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        mode_box = ttk.LabelFrame(root, text="1) 选择对账模式", padding=10)
        mode_box.pack(fill="x", pady=(0, 10))
        ttk.Radiobutton(mode_box, text=DETAIL_MODE, variable=self.mode_var, value=DETAIL_MODE, command=self._refresh_mode_view).grid(
            row=0, column=0, sticky="w", padx=(0, 20)
        )
        ttk.Radiobutton(mode_box, text=TOTAL_MODE, variable=self.mode_var, value=TOTAL_MODE, command=self._refresh_mode_view).grid(
            row=0, column=1, sticky="w"
        )

        self.files_box = ttk.LabelFrame(root, text="2) 选择输入文件", padding=10)
        self.files_box.pack(fill="x", pady=(0, 10))

        out_box = ttk.LabelFrame(root, text="3) 选择输出目录", padding=10)
        out_box.pack(fill="x", pady=(0, 10))
        ttk.Entry(out_box, textvariable=self.output_dir_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(out_box, text="浏览", command=self._pick_output_dir).grid(row=0, column=1, sticky="e")
        out_box.columnconfigure(0, weight=1)

        run_box = ttk.Frame(root)
        run_box.pack(fill="x", pady=(0, 8))
        ttk.Button(run_box, text="开始对账", command=self._run_reconcile).pack(side="left")
        ttk.Button(run_box, text="打开输出目录", command=self._open_output_dir_hint).pack(side="left", padx=(8, 0))

        status_box = ttk.LabelFrame(root, text="执行状态", padding=10)
        status_box.pack(fill="both", expand=True)
        ttk.Label(status_box, textvariable=self.status_var, wraplength=700, justify="left").pack(anchor="w")

        tips = (
            "说明：\n"
            " - 明细对账：上传 CBS 明细 + SAP 明细；输出明细结果和对账报告。\n"
            " - 总和对账：上传 CBS 历史余额 + SAP 科余 + 银行科目；输出总和结果和对账报告。"
        )
        ttk.Label(status_box, text=tips, justify="left").pack(anchor="w", pady=(10, 0))

    def _render_row(self, parent: ttk.LabelFrame, row: int, title: str, key: str) -> None:
        ttk.Label(parent, text=title, width=18).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=self.file_vars[key]).grid(row=row, column=1, sticky="ew", pady=4, padx=6)
        ttk.Button(parent, text="选择文件", command=lambda: self._pick_file(key)).grid(row=row, column=2, sticky="e", pady=4)
        parent.columnconfigure(1, weight=1)

    def _refresh_mode_view(self) -> None:
        for w in self.files_box.winfo_children():
            w.destroy()
        if self.mode_var.get() == DETAIL_MODE:
            self._render_row(self.files_box, 0, "CBS 明细文件", "detail_cbs")
            self._render_row(self.files_box, 1, "SAP 明细文件", "detail_sap")
        else:
            self._render_row(self.files_box, 0, "CBS 历史余额", "total_cbs")
            self._render_row(self.files_box, 1, "SAP 科余", "total_sap")
            self._render_row(self.files_box, 2, "银行科目", "total_bank")

    def _pick_file(self, key: str) -> None:
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.file_vars[key].set(path)

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)

    def _validate_files(self) -> list[Path]:
        mode = self.mode_var.get()
        required_keys = ["detail_cbs", "detail_sap"] if mode == DETAIL_MODE else ["total_cbs", "total_sap", "total_bank"]
        paths: list[Path] = []
        for key in required_keys:
            raw = self.file_vars[key].get().strip()
            if not raw:
                raise ValueError("请先选择所有必需的输入文件。")
            p = Path(raw).expanduser().resolve()
            if not p.exists():
                raise ValueError(f"文件不存在：{p}")
            if p.suffix.lower() not in {".xlsx", ".xls"}:
                raise ValueError(f"不是Excel文件：{p}")
            paths.append(p)
        return paths

    def _run_reconcile(self) -> None:
        try:
            files = self._validate_files()
            out_dir = Path(self.output_dir_var.get().strip()).expanduser().resolve()
            out_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("参数错误", str(e))
            return

        mode = self.mode_var.get()
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.status_var.set("正在对账，请稍候...")
        self.update_idletasks()

        try:
            if mode == DETAIL_MODE:
                cbs_df = load_cbs_detail(str(files[0]))
                sap_df = load_sap_detail(str(files[1]))
                result_df = reconcile(cbs_df, sap_df)

                excel_path = out_dir / f"CBS明细_对账结果_{now}.xlsx"
                report_path = out_dir / f"对账报告_明细_{now}.txt"
                result_df.to_excel(excel_path, index=False)
                write_report(result_df, str(report_path))
                report_text = report_as_string(result_df)
            else:
                excel_path = out_dir / f"对账差额结果3.0_{now}.xlsx"
                report_path = out_dir / f"总和对账报告_{now}.txt"
                result = run_total_reconcile(
                    cbs_balance_path=files[0],
                    sap_balance_path=files[1],
                    bank_subject_path=files[2],
                    output_excel_path=excel_path,
                )
                report_text = str(result["report"])
                write_total_report(report_text, report_path)

            msg = (
                f"对账完成。\n\n"
                f"模式：{mode}\n"
                f"结果文件：{excel_path}\n"
                f"报告文件：{report_path}\n\n"
                f"报告预览：\n{report_text}"
            )
            self.status_var.set(msg)
            messagebox.showinfo("完成", "对账已完成，结果与报告已输出。")
        except Exception as e:
            self.status_var.set("执行失败，请检查文件内容与列名。")
            detail = f"{e}\n\n{traceback.format_exc(limit=2)}"
            messagebox.showerror("执行失败", detail)

    def _open_output_dir_hint(self) -> None:
        path = self.output_dir_var.get().strip()
        if not path:
            messagebox.showinfo("提示", "请先设置输出目录。")
            return
        messagebox.showinfo("输出目录", f"当前输出目录：\n{path}")


def main() -> None:
    app = ReconcileGui()
    app.mainloop()


if __name__ == "__main__":
    main()
