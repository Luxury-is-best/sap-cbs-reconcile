from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd


AMOUNT_DECIMALS = 2


def pick_column(df: pd.DataFrame, index: int, preferred_names: list[str]) -> str:
    """优先按列名匹配，匹配不到时按 Excel 列序号取列。"""
    for name in preferred_names:
        if name in df.columns:
            return name
    if 0 <= index < len(df.columns):
        return df.columns[index]
    raise ValueError(f"未找到目标列：index={index}, preferred_names={preferred_names}")


def to_amount(series: pd.Series) -> pd.Series:
    """将金额列转为数值，兼容千分位、括号负数等常见格式。"""
    s = series.astype(str).str.strip()
    negative_mask = s.str.match(r"^\(.*\)$", na=False)
    s = s.str.replace(r"[(),，\s]", "", regex=True)
    num = pd.to_numeric(s, errors="coerce").fillna(0.0)
    num.loc[negative_mask] = -num.loc[negative_mask]
    return num


def normalize_account_text(value) -> str:
    """标准化账号文本，尽量消除 Excel 数字化导致的 .0 问题。"""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s


def extract_bank_card_from_subject(value) -> str:
    """从 SAP 科目名称中提取最长数字串作为银行卡号。"""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    parts = re.findall(r"\d+", text)
    return max(parts, key=len) if parts else ""


_BANK_NAME_PATTERNS = [
    "中国建设银行",
    "中国工商银行",
    "中国农业银行",
    "招商银行",
    "兴业银行",
    "交通银行",
    "浦发银行",
    "民生银行",
    "中信银行",
    "光大银行",
    "华夏银行",
    "广发银行",
    "天津银行",
    "北京银行",
    "邮储银行",
    "中国银行",
    "招行",
]


def extract_bank_type_from_subject(value) -> str:
    """从科目名称（C列）提取银行类别。"""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text.startswith("银行存款-"):
        return ""
    rest = text[len("银行存款-") :].strip()
    rest = re.sub(r"\d{8,}\s*$", "", rest)
    for name in _BANK_NAME_PATTERNS:
        if name in rest:
            return "招商银行" if name == "招行" else name
    return ""


def extract_bank_type_from_branch(value) -> str:
    """从开户行名称提取银行类别。"""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    for name in _BANK_NAME_PATTERNS:
        if name in text:
            return "招商银行" if name == "招行" else name
    return ""


def _sort_by_company_code(df: pd.DataFrame) -> pd.DataFrame:
    if "公司代码" not in df.columns:
        return df.sort_values(df.columns[0])
    df = df.copy()
    num = pd.to_numeric(df["公司代码"], errors="coerce")
    df["_排序键"] = num.fillna(float("inf"))
    df = df.sort_values(["_排序键", "银行卡"]).drop(columns=["_排序键"])
    return df


def reconcile_total(
    cbs_df: pd.DataFrame, sap_df: pd.DataFrame, bank_subject_df: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    """执行总和对账并返回三个工作表数据。"""
    cbs_col_account = pick_column(cbs_df, 1, ["账号"])
    cbs_col_account_name = pick_column(cbs_df, 2, ["账户名称"])
    cbs_col_branch = pick_column(cbs_df, 3, ["开户行名称"])
    cbs_col_balance = pick_column(cbs_df, 5, ["账户余额", "账户余额（折人民币）"])
    cbs_work = cbs_df.copy()
    cbs_work["银行卡"] = cbs_work[cbs_col_account].map(normalize_account_text)
    cbs_work["_余额"] = to_amount(cbs_work[cbs_col_balance])
    cbs_work["_账户名称"] = cbs_work[cbs_col_account_name].fillna("").astype(str).str.strip()
    cbs_work["_开户行"] = cbs_work[cbs_col_branch].fillna("").astype(str)
    cbs_card_to_name = (
        cbs_work.groupby("银行卡", dropna=False)["_账户名称"]
        .first()
        .reset_index()
        .rename(columns={"_账户名称": "账户名称_来自CBS"})
    )
    cbs_card_to_bank = cbs_work.groupby("银行卡", dropna=False)["_开户行"].first().reset_index()
    cbs_card_to_bank["所属银行_来自CBS"] = cbs_card_to_bank["_开户行"].map(
        extract_bank_type_from_branch
    )
    cbs_card_to_bank = cbs_card_to_bank[["银行卡", "所属银行_来自CBS"]]
    cbs_card_to_info = cbs_card_to_name.merge(cbs_card_to_bank, on="银行卡", how="left")

    cbs_account_summary = (
        cbs_work.groupby("银行卡", dropna=False)["_余额"]
        .sum()
        .reset_index()
        .rename(columns={"_余额": "CBS余额"})
        .sort_values("银行卡")
    )
    cbs_account_summary["CBS余额"] = cbs_account_summary["CBS余额"].round(AMOUNT_DECIMALS)

    sap_col_subject = pick_column(sap_df, 8, ["科目号"])
    sap_col_period = pick_column(sap_df, 7, ["期间/年度"])
    sap_col_balance = pick_column(sap_df, 16, ["累计余额"])
    sap_work = sap_df.copy()
    sap_work["银行卡"] = sap_work[sap_col_subject].astype(str).map(extract_bank_card_from_subject)
    sap_work["_累计余额"] = to_amount(sap_work[sap_col_balance])
    period_series = sap_work[sap_col_period].astype(str).str.strip()
    period_mask = period_series.str.match(r"^(\d+)月\s*(\d{4})$", na=False)
    sap_period = sap_work.loc[period_mask].copy()
    if not sap_period.empty:
        months = period_series.loc[period_mask].str.extract(r"^(\d+)月\s*(\d{4})$")
        sap_period["_月"] = pd.to_numeric(months[0], errors="coerce")
        sap_period["_年"] = pd.to_numeric(months[1], errors="coerce")
        max_year = sap_period["_年"].max()
        max_month = sap_period.loc[sap_period["_年"] == max_year, "_月"].max()
        sap_period = sap_period[(sap_period["_年"] == max_year) & (sap_period["_月"] == max_month)]
    if sap_period.empty:
        sap_period = sap_work

    sap_card_summary = (
        sap_period.groupby("银行卡", dropna=False)["_累计余额"]
        .sum()
        .reset_index()
        .rename(columns={"_累计余额": "SAP余额"})
        .sort_values("银行卡")
    )
    sap_card_summary["SAP余额"] = sap_card_summary["SAP余额"].round(AMOUNT_DECIMALS)

    bank_col_subject = pick_column(bank_subject_df, 2, ["科目名称"])
    bank_col_company = pick_column(bank_subject_df, 3, ["公司代码"])
    bank_col_bank_type = pick_column(bank_subject_df, 4, ["银行类别"])
    bank_map_df = bank_subject_df[[bank_col_subject, bank_col_company, bank_col_bank_type]].copy()
    bank_map_df["银行卡"] = bank_map_df[bank_col_subject].map(extract_bank_card_from_subject)
    bank_map_df["公司代码"] = bank_map_df[bank_col_company].fillna("").astype(str).str.strip()
    bank_type_col = bank_map_df[bank_col_bank_type].fillna("").astype(str).str.strip()
    from_subject = bank_map_df[bank_col_subject].map(extract_bank_type_from_subject)
    bank_map_df["所属银行"] = bank_type_col.where(bank_type_col != "", from_subject)
    bank_map_df = bank_map_df[bank_map_df["银行卡"] != ""]
    bank_company_map = (
        bank_map_df.groupby("银行卡", dropna=False)
        .agg(公司代码=("公司代码", "first"), 所属银行=("所属银行", "first"))
        .reset_index()
    )

    sap_col_company = pick_column(sap_df, 4, ["公司代码"])
    sap_card_to_name = (
        sap_period.groupby("银行卡", dropna=False)[sap_col_company]
        .first()
        .reset_index()
        .rename(columns={sap_col_company: "公司名称"})
    )
    bank_company_map = bank_company_map.merge(sap_card_to_name, on="银行卡", how="left")
    bank_company_map["公司名称"] = bank_company_map["公司名称"].fillna("").astype(str)

    company_name_to_code = (
        bank_company_map[bank_company_map["公司名称"] != ""]
        .drop_duplicates("公司名称", keep="first")[["公司名称", "公司代码"]]
        .set_index("公司名称")["公司代码"]
        .to_dict()
    )

    def _fill_from_cbs_and_sort(df: pd.DataFrame, amount_cols: list[str]) -> pd.DataFrame:
        df = df.merge(cbs_card_to_info, on="银行卡", how="left")
        df["账户名称_来自CBS"] = df["账户名称_来自CBS"].fillna("").astype(str)
        df["所属银行_来自CBS"] = df["所属银行_来自CBS"].fillna("").astype(str)
        df["公司名称"] = df["公司名称"].fillna("").astype(str)
        mask_name = df["公司名称"] == ""
        df.loc[mask_name, "公司名称"] = df.loc[mask_name, "账户名称_来自CBS"]
        df["所属银行"] = df["所属银行"].fillna("").astype(str)
        mask_bank = df["所属银行"] == ""
        df.loc[mask_bank, "所属银行"] = df.loc[mask_bank, "所属银行_来自CBS"]
        df["公司代码"] = df["公司代码"].fillna("").astype(str)
        mask_code = df["公司代码"] == ""
        df.loc[mask_code, "公司代码"] = df.loc[mask_code, "公司名称"].map(
            lambda x: company_name_to_code.get(x, "")
        )
        df = df.drop(columns=["账户名称_来自CBS", "所属银行_来自CBS"], errors="ignore")
        return _sort_by_company_code(df[["公司代码", "公司名称", "银行卡", "所属银行"] + amount_cols])

    total_overview = cbs_account_summary.merge(sap_card_summary, on="银行卡", how="outer")
    total_overview = total_overview.merge(bank_company_map, on="银行卡", how="left")
    total_overview["CBS余额"] = total_overview["CBS余额"].fillna(0.0).round(AMOUNT_DECIMALS)
    total_overview["SAP余额"] = total_overview["SAP余额"].fillna(0.0).round(AMOUNT_DECIMALS)
    total_overview["余额差异(CBS-SAP)"] = (
        total_overview["CBS余额"].astype(float) - total_overview["SAP余额"].astype(float)
    ).round(AMOUNT_DECIMALS)
    total_overview = _fill_from_cbs_and_sort(total_overview, ["CBS余额", "SAP余额", "余额差异(CBS-SAP)"])

    cbs_sheet = _fill_from_cbs_and_sort(
        cbs_account_summary.merge(bank_company_map, on="银行卡", how="left"), ["CBS余额"]
    )
    sap_sheet = _fill_from_cbs_and_sort(
        sap_card_summary.merge(bank_company_map, on="银行卡", how="left"), ["SAP余额"]
    )

    return {"CBS汇总": cbs_sheet, "SAP汇总": sap_sheet, "总览": total_overview}


def total_report_as_string(total_overview: pd.DataFrame) -> str:
    """生成总和对账报告文本。"""
    if total_overview.empty:
        return "========== 总和对账报告 ==========\n无数据。"

    diff_col = "余额差异(CBS-SAP)"
    non_zero = total_overview[total_overview[diff_col].round(AMOUNT_DECIMALS) != 0]
    lines = [
        "========== 总和对账报告 ==========",
        f"银行卡总数: {len(total_overview)}",
        f"差异为0条数: {len(total_overview) - len(non_zero)}",
        f"差异非0条数: {len(non_zero)}",
        f"CBS余额合计: {total_overview['CBS余额'].sum():.2f}",
        f"SAP余额合计: {total_overview['SAP余额'].sum():.2f}",
        f"差异合计(CBS-SAP): {total_overview[diff_col].sum():.2f}",
    ]
    if not non_zero.empty:
        top_rows = non_zero.reindex(non_zero[diff_col].abs().sort_values(ascending=False).index).head(10)
        lines.append("")
        lines.append("差异前10项（按绝对值降序）:")
        for _, row in top_rows.iterrows():
            lines.append(
                f"- 公司代码:{row['公司代码']} 公司名称:{row['公司名称']} 银行卡:{row['银行卡']} 差异:{row[diff_col]:.2f}"
            )
    return "\n".join(lines)


def run_total_reconcile(
    cbs_balance_path: str | Path,
    sap_balance_path: str | Path,
    bank_subject_path: str | Path,
    output_excel_path: str | Path | None = None,
) -> dict[str, object]:
    """按文件路径执行总和对账，并可选写出结果 Excel。"""
    cbs_df = pd.read_excel(cbs_balance_path)
    sap_df = pd.read_excel(sap_balance_path)
    bank_subject_df = pd.read_excel(bank_subject_path)
    sheets = reconcile_total(cbs_df, sap_df, bank_subject_df)
    report_text = total_report_as_string(sheets["总览"])

    output_path_obj = None
    if output_excel_path:
        output_path_obj = Path(output_excel_path).resolve()
        output_path_obj.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path_obj, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)

    return {"sheets": sheets, "report": report_text, "output_path": output_path_obj}


def run_total_reconcile_from_bytes(
    cbs_balance_bytes: bytes,
    sap_balance_bytes: bytes,
    bank_subject_bytes: bytes,
) -> dict[str, object]:
    """按上传字节流执行总和对账，返回报告和结果 Excel 字节。"""
    cbs_df = pd.read_excel(io.BytesIO(cbs_balance_bytes))
    sap_df = pd.read_excel(io.BytesIO(sap_balance_bytes))
    bank_subject_df = pd.read_excel(io.BytesIO(bank_subject_bytes))
    sheets = reconcile_total(cbs_df, sap_df, bank_subject_df)
    report_text = total_report_as_string(sheets["总览"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return {"report": report_text, "excel_bytes": buffer.getvalue(), "sheets": sheets}


def write_total_report(report_text: str, report_path: str | Path) -> None:
    Path(report_path).write_text(report_text, encoding="utf-8")
