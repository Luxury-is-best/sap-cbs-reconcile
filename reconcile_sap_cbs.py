#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP 与 CBS 对账脚本
功能：识别 CBS 明细中未在 SAP 明细中存在的记录，标注匹配状态并生成对账报告。
"""

import argparse
import io
import sys
from pathlib import Path

import pandas as pd


# CBS 明细关键列（按文档）
CBS_DATE_COL = "交易日期"
CBS_SERIAL_COL = "交易流水号"
CBS_ACCOUNT_COL = "账号"
CBS_DEBIT_COL = "借(支出)"
CBS_CREDIT_COL = "贷(收入)"
CBS_BANK_SERIAL_COL = "银行流水号"
CBS_RECON_CODE_COL = "对账码"

# SAP 明细关键列
SAP_DOC_NO_COL = "凭证编号"
SAP_POST_DATE_COL = "过账日期"
SAP_AMOUNT_COL = "公司代码货币价值"
SAP_DR_CR_COL = "借/贷标识"
SAP_ACCOUNT_COL = "总帐科目"

# 工程文件夹（脚本所在目录），模糊查找仅在此目录下进行，不往父级目录查找
PROJECT_DIR = Path(__file__).resolve().parent

# 模糊匹配用的 Excel 扩展名
_EXCEL_SUFFIXES = (".xlsx", ".xls", ".XLSX", ".XLS")


def fuzzy_find_cbs_file(directory: Path) -> Path | None:
    """在指定目录中模糊查找 CBS 明细文件（如 CBS交易明细列表-20260302-171719.xlsx）。
    匹配规则：文件名含 CBS 且含「明细」或「交易」，扩展名为 .xlsx/.xls。多文件时取最新修改时间。
    """
    directory = Path(directory).resolve()
    if not directory.is_dir():
        return None
    candidates = []
    for p in directory.iterdir():
        if not p.is_file() or not (p.suffix in _EXCEL_SUFFIXES):
            continue
        name = p.name
        if "CBS" not in name.upper():
            continue
        if "明细" in name or "交易" in name:
            candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return candidates[0]


def fuzzy_find_sap_file(directory: Path) -> Path | None:
    """在指定目录中模糊查找 SAP 明细文件（如 SAP银行明细.XLSX）。
    匹配规则：文件名含 SAP 且含「明细」或「银行」，扩展名为 .xlsx/.xls。多文件时取最新修改时间。
    """
    directory = Path(directory).resolve()
    if not directory.is_dir():
        return None
    candidates = []
    for p in directory.iterdir():
        if not p.is_file() or not (p.suffix in _EXCEL_SUFFIXES):
            continue
        name = p.name
        if "SAP" not in name.upper():
            continue
        if "明细" in name or "银行" in name:
            candidates.append(p)
    if not candidates:
        return None
    candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return candidates[0]


def _find_col(df: pd.DataFrame, candidates: list) -> str | None:
    """在 DataFrame 中查找存在的列名（支持候选列表）。"""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _normalize_date(series: pd.Series) -> pd.Series:
    """将日期列统一为 YYYY-MM-DD 字符串，便于比较。"""
    out = pd.to_datetime(series, errors="coerce")
    return out.dt.strftime("%Y-%m-%d")


def _normalize_amount(series: pd.Series) -> pd.Series:
    """将金额转为可比较的浮点数。"""
    if series.dtype == object or series.dtype.name == "string":
        s = series.astype(str).str.replace(",", "").str.replace(" ", "").str.strip()
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(series, errors="coerce")


def _cbs_amount_row(row: pd.Series, debit_col: str, credit_col: str) -> tuple[float, str]:
    """取 CBS 一行的金额（非空借或贷）及方向。返回 (绝对值, '借'|'贷')。"""
    d = _normalize_amount(pd.Series([row.get(debit_col, None)])).iloc[0]
    c = _normalize_amount(pd.Series([row.get(credit_col, None)])).iloc[0]
    if pd.notna(d) and (d != 0 or pd.isna(c)):
        return (abs(float(d)), "借")
    if pd.notna(c):
        return (abs(float(c)), "贷")
    return (0.0, "借")


def _sap_amount_row(row: pd.Series, amount_col: str, dr_cr_col: str) -> tuple[float, str]:
    """取 SAP 一行的金额及方向。返回 (绝对值, '借'|'贷')。
    匹配规则：SAP S 对应 CBS 贷(收入)、SAP H 对应 CBS 借(支出)，故 S→贷、H→借以便与 CBS 键一致。
    """
    amt = _normalize_amount(pd.Series([row.get(amount_col, None)])).iloc[0]
    if pd.isna(amt):
        return (0.0, "借")
    dr_cr = str(row.get(dr_cr_col, "")).strip().upper()
    if "S" in dr_cr:
        direction = "贷"  # SAP S → 与 CBS 贷(收入) 匹配
    elif "H" in dr_cr:
        direction = "借"  # SAP H → 与 CBS 借(支出) 匹配
    else:
        dr_cr_raw = str(row.get(dr_cr_col, "")).strip()
        direction = "贷" if "贷" in dr_cr_raw or "C" in dr_cr_raw.upper() else "借"
    return (abs(float(amt)), direction)


def load_cbs_detail(path: str) -> pd.DataFrame:
    """读取 CBS 明细 Excel，取第一个 Sheet。"""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"CBS 明细文件不存在: {path}")
    df = pd.read_excel(path, sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()
    return df


def load_sap_detail(path: str) -> pd.DataFrame:
    """读取 SAP 明细 Excel，取第一个 Sheet。"""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"SAP 明细文件不存在: {path}")
    df = pd.read_excel(path, sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()
    return df


def load_cbs_detail_from_bytes(data: bytes) -> pd.DataFrame:
    """从内存字节流读取 CBS 明细 Excel，供 Web 上传使用。"""
    df = pd.read_excel(io.BytesIO(data), sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()
    return df


def load_sap_detail_from_bytes(data: bytes) -> pd.DataFrame:
    """从内存字节流读取 SAP 明细 Excel，供 Web 上传使用。"""
    df = pd.read_excel(io.BytesIO(data), sheet_name=0)
    df.columns = df.columns.astype(str).str.strip()
    return df


def reconcile(cbs_df: pd.DataFrame, sap_df: pd.DataFrame) -> pd.DataFrame:
    """
    对账逻辑：在 CBS 明细上标注 SAP 匹配状态与未匹配原因。
    仅按金额+方向匹配（SAP S→CBS 贷同额，SAP H→CBS 借同额），匹配上即已匹配，不核账号。
    """
    # 列名兼容
    cbs_date = _find_col(cbs_df, [CBS_DATE_COL, "交易日期"])
    cbs_serial = _find_col(cbs_df, [CBS_SERIAL_COL, "交易流水号"])
    cbs_account = _find_col(cbs_df, [CBS_ACCOUNT_COL, "账号", "账户编号"])
    cbs_debit = _find_col(cbs_df, [CBS_DEBIT_COL, "借(支出)"])
    cbs_credit = _find_col(cbs_df, [CBS_CREDIT_COL, "贷(收入)"])

    sap_doc = _find_col(sap_df, [SAP_DOC_NO_COL, "凭证编号"])
    sap_date = _find_col(sap_df, [SAP_POST_DATE_COL, "过账日期", "记账日期"])
    sap_amount_col = _find_col(sap_df, [SAP_AMOUNT_COL, "公司代码货币价值"])
    sap_dr_cr = _find_col(sap_df, [SAP_DR_CR_COL, "借/贷标识", "借/贷指示"])
    sap_account = _find_col(sap_df, [SAP_ACCOUNT_COL, "总帐科目"])

    missing = []
    if not cbs_date or not cbs_debit or not cbs_credit:
        missing.append("CBS 缺少：交易日期、借(支出)、贷(收入)")
    if not sap_date or not sap_amount_col:
        missing.append("SAP 缺少：过账日期、公司代码货币价值")
    if missing:
        raise ValueError("列名不匹配，请确认表头与文档一致。\n" + "\n".join(missing))

    # 标准化 CBS
    cbs_df = cbs_df.copy()
    cbs_df["_date"] = _normalize_date(cbs_df[cbs_date])
    cbs_df["_amount"] = cbs_df.apply(
        lambda r: _cbs_amount_row(r, cbs_debit, cbs_credit)[0], axis=1
    )
    cbs_df["_direction"] = cbs_df.apply(
        lambda r: _cbs_amount_row(r, cbs_debit, cbs_credit)[1], axis=1
    )

    # 标准化 SAP
    sap_df = sap_df.copy()
    sap_df["_date"] = _normalize_date(sap_df[sap_date])
    sap_df["_amount"] = sap_df.apply(
        lambda r: _sap_amount_row(r, sap_amount_col, sap_dr_cr or "")[0], axis=1
    )
    sap_df["_direction"] = sap_df.apply(
        lambda r: _sap_amount_row(r, sap_amount_col, sap_dr_cr or "")[1], axis=1
    )

    # 构建 SAP 查找结构：仅按 (金额, 方向)
    sap_amount_direction_key = set()
    for _, r in sap_df.iterrows():
        a, dr = r["_amount"], r["_direction"]
        if pd.notna(a):
            sap_amount_direction_key.add((round(a, 2), dr))

    # 对每条 CBS 判定匹配状态：仅按金额+方向匹配
    status_list = []
    reason_list = []

    for i, row in cbs_df.iterrows():
        a, dr = row["_amount"], row["_direction"]
        a_round = round(a, 2) if pd.notna(a) else 0

        if (a_round, dr) in sap_amount_direction_key:
            status_list.append("已匹配")
            reason_list.append("金额+方向一致")
        else:
            status_list.append("未匹配")
            reason_list.append("无金额+方向一致的SAP记录")

    cbs_df["SAP匹配状态"] = status_list
    cbs_df["未匹配原因"] = reason_list
    # 删除辅助列
    cbs_df.drop(columns=["_date", "_amount", "_direction"], inplace=True)
    return cbs_df


def report_as_string(cbs_result: pd.DataFrame) -> str:
    """生成匹配报告并返回字符串（供 Web 接口使用，不写文件）。"""
    total = len(cbs_result)
    status_col = "SAP匹配状态"
    if status_col not in cbs_result.columns:
        return ""
    matched = int((cbs_result[status_col] == "已匹配").sum())
    unmatched = int((cbs_result[status_col] == "未匹配").sum())
    suspected = int((cbs_result[status_col] == "疑似匹配（需人工确认）").sum())

    rate = (matched / total * 100) if total else 0
    lines = [
        "========== SAP 与 CBS 对账报告 ==========",
        f"总条数: {total}",
        f"已匹配: {matched}",
        f"未匹配: {unmatched}",
        f"疑似匹配（需人工确认）: {suspected}",
        f"匹配率（已匹配/总条数）: {rate:.2f}%",
        "",
        "未匹配与疑似匹配记录请在使用脚本输出的 CBS 明细表中，按列「SAP匹配状态」筛选后人工核对。",
    ]
    return "\n".join(lines)


def write_report(cbs_result: pd.DataFrame, report_path: str) -> None:
    """生成匹配报告（文本）并写入文件。"""
    text = report_as_string(cbs_result)
    if text:
        Path(report_path).write_text(text, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(
        description="SAP 与 CBS 对账：标注 CBS 明细中在 SAP 的匹配状态并生成报告。"
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

    base = Path(args.cbs).parent
    cbs_path = Path(args.cbs) if Path(args.cbs).is_absolute() else base / args.cbs
    sap_path = Path(args.sap) if Path(args.sap).is_absolute() else base / args.sap
    out_path = Path(args.output) if Path(args.output).is_absolute() else base / args.output
    report_path = Path(args.report) if Path(args.report).is_absolute() else base / args.report

    # 精准路径不存在时仅在工程文件夹下模糊查找（不翻父级目录）
    if not cbs_path.exists():
        found = fuzzy_find_cbs_file(PROJECT_DIR)
        if found:
            cbs_path = found
    if not sap_path.exists():
        found = fuzzy_find_sap_file(PROJECT_DIR)
        if found:
            sap_path = found

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
