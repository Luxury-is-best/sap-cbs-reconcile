#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP 与 CBS 桌面对账工具：
一次上传全部输入文件，同时执行明细对账与总和对账；
结果输出为单个 Excel（多 sheet：明细对账 + 总和各表）及合并文本报告。
"""

from __future__ import annotations

__version__ = "1.1.0"

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
)
from total_reconcile_service import run_total_reconcile, write_combined_reconcile_excel


# 固定顺序：明细 2 个 + 总和 3 个（共 5 个文件）
FILE_ROWS: list[tuple[str, str]] = [
    ("CBS 明细", "detail_cbs"),
    ("SAP 明细", "detail_sap"),
    ("CBS 历史余额", "total_cbs"),
    ("SAP 科余", "total_sap"),
    ("银行科目", "total_bank"),
]


class ReconcileGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"SAP和CBS对账助手 v{__version__}")
        self.geometry("760x620")
        self.minsize(700, 560)

        self.status_var = tk.StringVar(value="请按顺序选择全部输入文件后点击开始对账。")
        self.output_dir_var = tk.StringVar(value=str(Path.cwd()))

        self.file_vars = {key: tk.StringVar() for _, key in FILE_ROWS}

        self._build_layout()

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        files_box = ttk.LabelFrame(root, text="1) 上传全部输入文件（按下列顺序，共 5 个）", padding=10)
        files_box.pack(fill="x", pady=(0, 10))
        for row, (title, key) in enumerate(FILE_ROWS):
            ttk.Label(files_box, text=title, width=18).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(files_box, textvariable=self.file_vars[key]).grid(
                row=row, column=1, sticky="ew", pady=4, padx=6
            )
            ttk.Button(files_box, text="选择文件", command=lambda k=key: self._pick_file(k)).grid(
                row=row, column=2, sticky="e", pady=4
            )
        files_box.columnconfigure(1, weight=1)

        bulk_row = len(FILE_ROWS)
        ttk.Button(
            files_box,
            text="一次性多选文件（按上表顺序依次点选 5 个）",
            command=self._pick_all_in_order,
        ).grid(row=bulk_row, column=0, columnspan=3, sticky="w", pady=(8, 0))

        out_box = ttk.LabelFrame(root, text="2) 选择输出目录", padding=10)
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
            " - 程序会同时跑「明细对账」与「总和对账」。\n"
            " - 输出一个 Excel：工作表「明细对账」+ 总和结果的「CBS汇总」「SAP汇总」「总览」。\n"
            " - 另有一份合并的文本报告（明细报告 + 总和报告）。"
        )
        ttk.Label(status_box, text=tips, justify="left").pack(anchor="w", pady=(10, 0))

    def _pick_file(self, key: str) -> None:
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.file_vars[key].set(path)

    def _pick_all_in_order(self) -> None:
        paths = filedialog.askopenfilenames(
            title="按顺序选择 5 个文件：CBS明细 → SAP明细 → CBS历史余额 → SAP科余 → 银行科目",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        if len(paths) != len(FILE_ROWS):
            messagebox.showwarning(
                "数量不对",
                f"需要恰好选择 {len(FILE_ROWS)} 个文件，当前选了 {len(paths)} 个。\n"
                "请按住 Ctrl/Cmd 按顺序多选，或逐个用「选择文件」指定。",
            )
            return
        for (_, key), p in zip(FILE_ROWS, paths):
            self.file_vars[key].set(p)

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)

    def _validate_files(self) -> list[Path]:
        paths: list[Path] = []
        for _, key in FILE_ROWS:
            raw = self.file_vars[key].get().strip()
            if not raw:
                raise ValueError("请先为每一行选择输入文件（共 5 个）。")
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

        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.status_var.set("正在对账，请稍候...")
        self.update_idletasks()

        excel_path = out_dir / f"SAP_CBS_对账结果_{now}.xlsx"
        report_path = out_dir / f"对账报告_合并_{now}.txt"

        try:
            cbs_detail = load_cbs_detail(str(files[0]))
            sap_detail = load_sap_detail(str(files[1]))
            result_df = reconcile(cbs_detail, sap_detail)

            total_result = run_total_reconcile(
                cbs_balance_path=files[2],
                sap_balance_path=files[3],
                bank_subject_path=files[4],
                output_excel_path=None,
            )
            total_sheets = total_result["sheets"]
            assert isinstance(total_sheets, dict)

            write_combined_reconcile_excel(result_df, total_sheets, excel_path)

            detail_report = report_as_string(result_df)
            total_report = str(total_result["report"])
            combined_report = (
                "========== 明细对账报告 ==========\n" + detail_report.strip() + "\n\n\n" + total_report.strip() + "\n"
            )
            Path(report_path).write_text(combined_report, encoding="utf-8")

            msg = (
                f"对账完成。\n\n"
                f"结果 Excel（多 sheet）：{excel_path}\n"
                f"合并报告：{report_path}\n\n"
                f"报告预览：\n{combined_report[:3500]}"
                + ("…" if len(combined_report) > 3500 else "")
            )
            self.status_var.set(msg)
            messagebox.showinfo("完成", "对账已完成：已生成合并 Excel 与报告。")
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
