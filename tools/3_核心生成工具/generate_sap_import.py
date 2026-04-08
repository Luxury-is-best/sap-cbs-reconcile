import pandas as pd
import json
from datetime import datetime

def generate_import_file():
    cbs_file = "原始文件/CBS交易明细列表-20260302-171719.xlsx"
    cbs = pd.read_excel(cbs_file)
    
    # We will use the mapping we found earlier, plus a fallback for unknown accounts
    # I'll run the exact same logic to get the mapping inline to be safe.
    sap = pd.read_excel("原始文件/SAP银行明细.XLSX")
    sap['abs_amount'] = sap['公司代码货币价值'].abs()
    
    cbs['借(支出)'] = cbs['借(支出)'].fillna(0)
    cbs['贷(收入)'] = cbs['贷(收入)'].fillna(0)
    cbs['abs_amount'] = cbs[['借(支出)', '贷(收入)']].max(axis=1)
    
    mapping = {}
    for idx, cbs_row in cbs.iterrows():
        amt = cbs_row['abs_amount']
        acc = str(cbs_row['账号'])
        if amt == 0 or acc in mapping: continue
        sap_matches = sap[sap['abs_amount'] == amt]
        if not sap_matches.empty:
            sap_match = sap_matches.iloc[0]
            mapping[acc] = {
                '总账科目': str(int(sap_match['总帐科目'])),
                '公司代码': str(int(sap_match['公司代码']) if pd.notna(sap_match['公司代码']) else cbs_row['单位编码']),
            }
            
    # Generate standard SAP import rows
    import_rows = []
    # For grouping into documents, we can assign a virtual doc number per transaction
    doc_no = 1
    
    for idx, row in cbs.iterrows():
        acc = str(row['账号'])
        date_str = str(row['交易日期']).split()[0] # YYYY-MM-DD
        
        income = row['贷(收入)']
        expense = row['借(支出)']
        
        # If no amount, skip
        if income == 0 and expense == 0:
            continue
            
        # Get Company and GL
        if acc in mapping:
            company_code = mapping[acc]['公司代码']
            bank_gl = mapping[acc]['总账科目']
        else:
            company_code = str(int(row['单位编码'])) if pd.notna(row['单位编码']) else '未知'
            bank_gl = '待查银行科目'
            
        # Determine text
        desc = str(row['摘要']) if pd.notna(row['摘要']) else ''
        purpose = str(row['用途']) if pd.notna(row['用途']) else ''
        text = f"{date_str} {desc} {purpose}".strip()
        header_text = "银行流水自动入账"
        
        if income > 0:
            # Bank is Debit (40)
            import_rows.append({
                '虚拟凭证号': doc_no,
                '公司代码': company_code,
                '凭证日期': date_str,
                '过账日期': date_str,
                '凭证类型': 'DZ', # 收款
                '货币': 'CNY',
                '凭证抬头文本': header_text,
                '记账码': '40',
                '总账科目': bank_gl,
                '金额': income,
                '借/贷标识': 'S',
                '文本': text
            })
            # Offsetting is Credit (50)
            import_rows.append({
                '虚拟凭证号': doc_no,
                '公司代码': company_code,
                '凭证日期': date_str,
                '过账日期': date_str,
                '凭证类型': 'DZ',
                '货币': 'CNY',
                '凭证抬头文本': header_text,
                '记账码': '50',
                '总账科目': '9999999999', # Placeholder
                '金额': income,
                '借/贷标识': 'H',
                '文本': text
            })
        elif expense > 0:
            # Bank is Credit (50)
            import_rows.append({
                '虚拟凭证号': doc_no,
                '公司代码': company_code,
                '凭证日期': date_str,
                '过账日期': date_str,
                '凭证类型': 'SA', # 付款或一般
                '货币': 'CNY',
                '凭证抬头文本': header_text,
                '记账码': '50',
                '总账科目': bank_gl,
                '金额': expense,
                '借/贷标识': 'H',
                '文本': text
            })
            # Offsetting is Debit (40)
            import_rows.append({
                '虚拟凭证号': doc_no,
                '公司代码': company_code,
                '凭证日期': date_str,
                '过账日期': date_str,
                '凭证类型': 'SA',
                '货币': 'CNY',
                '凭证抬头文本': header_text,
                '记账码': '40',
                '总账科目': '9999999999', # Placeholder
                '金额': expense,
                '借/贷标识': 'S',
                '文本': text
            })
            
        doc_no += 1
        
    df_import = pd.DataFrame(import_rows)
    output_file = "SAP自动导入表_基于20260302流水.xlsx"
    df_import.to_excel(output_file, index=False)
    print(f"Successfully generated {output_file} with {len(df_import)} rows.")

if __name__ == "__main__":
    generate_import_file()
