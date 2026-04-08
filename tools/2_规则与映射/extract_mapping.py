import pandas as pd
import json

def extract_bank_gl_mapping():
    cbs = pd.read_excel("原始文件/CBS交易明细列表-20260302-171719.xlsx")
    sap = pd.read_excel("原始文件/SAP银行明细.XLSX")

    # Clean dates
    sap['凭证日期'] = pd.to_datetime(sap['凭证日期'], unit='ms')
    cbs['借(支出)'] = cbs['借(支出)'].fillna(0)
    cbs['贷(收入)'] = cbs['贷(收入)'].fillna(0)
    cbs['abs_amount'] = cbs[['借(支出)', '贷(收入)']].max(axis=1)
    cbs['交易日期_dt'] = pd.to_datetime(cbs['交易日期'])
    cbs['年月'] = cbs['交易日期_dt'].dt.strftime('%Y-%m')

    sap['abs_amount'] = sap['公司代码货币价值'].abs()

    # We want to map CBS [账号] to SAP [总帐科目, 公司代码]
    mapping = {}
    for idx, cbs_row in cbs.iterrows():
        amt = cbs_row['abs_amount']
        acc = cbs_row['账号']
        if amt == 0: continue
        
        # find matching sap lines
        sap_matches = sap[sap['abs_amount'] == amt]
        if not sap_matches.empty:
            # take first match (heuristic)
            sap_match = sap_matches.iloc[0]
            if acc not in mapping:
                mapping[acc] = {
                    '总账科目': str(int(sap_match['总帐科目'])),
                    '公司代码': str(int(sap_match['公司代码']) if pd.notna(sap_match['公司代码']) else cbs_row['单位编码']),
                    '科目名称': sap_match['总账科目：长文本']
                }

    print("Bank Account Mapping inferred:")
    print(json.dumps(mapping, indent=2, ensure_ascii=False))

extract_bank_gl_mapping()
