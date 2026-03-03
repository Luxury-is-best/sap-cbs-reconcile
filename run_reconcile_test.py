#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP 与 CBS 对账 - 可单独使用的调试脚本（不依赖 Flask，不依赖 Web 服务）。

用途：本地指定 CBS/SAP Excel 路径运行对账，输出结果与报告，并可对指定 CBS 行号做匹配断言。
后续对匹配逻辑或调试功能的修改请在本文件中进行，避免改动线上版后端 backend_reconcile_sap_cbs.py。
"""

import argparse
import sys
from pathlib import Path

from backend_reconcile_sap_cbs import (
    PROJECT_DIR,
    fuzzy_find_cbs_file,
    fuzzy_find_sap_file,
    load_cbs_detail,
    load_sap_detail,
    reconcile,
    report_as_string,
    write_report,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="SAP 与 CBS 对账（纯后端测试）：读本地 CBS/SAP Excel，输出对账结果与报告，可选行号断言。"
    )
    parser.add_argument(
        "--cbs",
        default="CBS交易明细列表-20260302-171719.xlsx",
        help="CBS 明细 Excel 路径（默认: CBS交易明细列表-20260302-171719.xlsx）",
    )
    parser.add_argument(
        "--sap",
        default="SAP银行明细.XLSX",
        help="SAP 明细 Excel 路径（默认: SAP银行明细.XLSX）",
    )
    parser.add_argument(
        "-o", "--output",
        default="CBS明细_对账结果.xlsx",
        help="输出 CBS 带匹配状态的 Excel（默认: CBS明细_对账结果.xlsx）",
    )
    parser.add_argument(
        "--report",
        default="对账报告.txt",
        help="对账报告文本文件路径（默认: 对账报告.txt）",
    )
    parser.add_argument(
        "--assert-rows",
        nargs="*",
        type=int,
        metavar="ROW",
        help="断言这些 CBS 行（Excel 行号，从 1 起）在结果中为「已匹配」，用于回归。例如: --assert-rows 2311 2312",
    )
    args = parser.parse_args()

    base = Path(args.cbs).resolve().parent
    cbs_path = (Path(args.cbs) if Path(args.cbs).is_absolute() else base / args.cbs).resolve()
    sap_path = (Path(args.sap) if Path(args.sap).is_absolute() else base / args.sap).resolve()
    out_path = (Path(args.output) if Path(args.output).is_absolute() else base / args.output).resolve()
    report_path = (Path(args.report) if Path(args.report).is_absolute() else base / args.report).resolve()

    # 精准路径不存在时仅在工程文件夹下模糊查找（不翻父级目录）
    if not cbs_path.exists():
        found = fuzzy_find_cbs_file(PROJECT_DIR)
        if found:
            cbs_path = found.resolve()
    if not sap_path.exists():
        found = fuzzy_find_sap_file(PROJECT_DIR)
        if found:
            sap_path = found.resolve()

    print(f"使用的文件 - CBS: {cbs_path.name}  SAP: {sap_path.name}")

    try:
        cbs_df = load_cbs_detail(str(cbs_path))
        sap_df = load_sap_detail(str(sap_path))
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        sys.exit(1)

    result = reconcile(cbs_df, sap_df)
    result.to_excel(out_path, index=False)
    write_report(result, str(report_path))

    print(f"已输出带匹配状态的 CBS 明细: {out_path}")
    print(f"已生成对账报告: {report_path}")
    print(report_as_string(result))

    if args.assert_rows:
        status_col = "SAP匹配状态"
        if status_col not in result.columns:
            print("错误: 结果中无列「SAP匹配状态」", file=sys.stderr)
            sys.exit(1)
        failed = []
        for row_one_based in args.assert_rows:
            idx = row_one_based - 1  # Excel 行号 1-based -> DataFrame 0-based
            if idx < 0 or idx >= len(result):
                failed.append((row_one_based, f"行号超出范围(1~{len(result)})"))
                continue
            status = result[status_col].iloc[idx]
            if status != "已匹配":
                failed.append((row_one_based, status))
        if failed:
            print("断言失败：以下 CBS 行未为「已匹配」:", file=sys.stderr)
            for row_one_based, msg in failed:
                print(f"  第 {row_one_based} 行: {msg}", file=sys.stderr)
            sys.exit(1)
        print(f"断言通过: CBS 第 {', '.join(str(r) for r in args.assert_rows)} 行均为「已匹配」。")


if __name__ == "__main__":
    main()
