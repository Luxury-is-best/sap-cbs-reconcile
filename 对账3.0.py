from pathlib import Path

from total_reconcile_service import run_total_reconcile, write_total_report


SCRIPT_DIR = Path(__file__).resolve().parent
SOURCE_DIR = SCRIPT_DIR
OUTPUT_DIR = SCRIPT_DIR / "对账差额结果3.0"
OUTPUT_FILE = OUTPUT_DIR / "对账差额结果3.0.xlsx"
OUTPUT_REPORT = OUTPUT_DIR / "总和对账报告.txt"

CBS_BALANCE_FILE = "账户历史余额-20260302-090622.xlsx"
SAP_BALANCE_FILE = "科余030313.XLSX"
BANK_SUBJECT_FILE = "银行科目.xlsx"


def get_data_file(data_dir: Path, kind: str) -> Path:
    """获取 CBS 汇总余额 / SAP 科余 / 银行科目 文件路径。优先使用固定文件名，否则按名称匹配。"""
    if not data_dir.exists():
        raise FileNotFoundError(f"未找到目录：{data_dir}")

    fixed = {
        "cbs": data_dir / CBS_BALANCE_FILE,
        "sap": data_dir / SAP_BALANCE_FILE,
        "bank_subject": data_dir / BANK_SUBJECT_FILE,
    }
    path = fixed.get(kind)
    if path and path.exists():
        return path

    files = [f for f in data_dir.iterdir() if f.is_file() and f.suffix.lower() in {".xlsx", ".xls"}]
    if kind == "sap":
        matched = [f for f in files if "科余" in f.name]
    elif kind == "cbs":
        matched = [f for f in files if "账户历史余额" in f.name]
    elif kind == "bank_subject":
        matched = [f for f in files if "银行科目" in f.name]
    else:
        raise ValueError(f"未知文件类型：{kind}")

    if not matched:
        raise FileNotFoundError(f"未在 {data_dir} 中识别到 {kind} 文件（可放置 {fixed.get(kind, '')}）")
    matched.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return matched[0]


def main() -> None:
    cbs_file = get_data_file(SOURCE_DIR, "cbs")
    sap_file = get_data_file(SOURCE_DIR, "sap")
    bank_subject_file = get_data_file(SOURCE_DIR, "bank_subject")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    result = run_total_reconcile(
        cbs_balance_path=cbs_file,
        sap_balance_path=sap_file,
        bank_subject_path=bank_subject_file,
        output_excel_path=OUTPUT_FILE,
    )
    write_total_report(result["report"], OUTPUT_REPORT)

    print(f"处理完成，结果已输出：{OUTPUT_FILE}")
    print(f"总和对账报告：{OUTPUT_REPORT}")
    print(f"CBS来源文件：{cbs_file.name}")
    print(f"SAP来源文件：{sap_file.name}")
    print(f"银行科目来源文件：{bank_subject_file.name}")


if __name__ == "__main__":
    main()
