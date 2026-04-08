import pandas as pd
import numpy as np

def compare_sap_generated_vs_actual():
    # 1. Load actual SAP
    sap_actual = pd.read_excel("原始文件/SAP银行明细.XLSX")
    sap_actual['abs_amount'] = sap_actual['公司代码货币价值'].abs()
    
    # 2. Load generated SAP
    sap_generated = pd.read_excel("SAP自动导入表_基于20260302流水.xlsx")
    
    # We only want to compare the Bank GL lines from generated
    # Generated bank GL lines are those where '总账科目' != 9999999999
    gen_bank = sap_generated[sap_generated['总账科目'] != 9999999999].copy()
    gen_bank['金额'] = gen_bank['金额'].astype(float)
    
    print(f"Total lines in Original SAP: {len(sap_actual)}")
    print(f"Total bank lines in Generated SAP: {len(gen_bank)}")
    
    # Let's aggregate by Company Code, Bank GL, and abs amount to see if they match
    # Since dates in SAP might be end-of-month, we can just match by Company, Amount, and D/C indicator
    
    # For actual SAP, we need to extract bank GL lines.
    # We can infer bank lines are those starting with '1002'
    sap_actual['总帐科目'] = sap_actual['总帐科目'].fillna(0).astype(str).str.split('.').str[0]
    act_bank = sap_actual[sap_actual['总帐科目'].str.startswith('1002')].copy()
    print(f"Total bank lines in Original SAP (starting with 1002): {len(act_bank)}")
    
    # Count amounts
    act_grouped = act_bank.groupby(['公司代码', '总帐科目', '借/贷标识', 'abs_amount']).size().reset_index(name='act_count')
    gen_grouped = gen_bank.groupby(['公司代码', '总账科目', '借/贷标识', '金额']).size().reset_index(name='gen_count')
    
    # Merge them
    # First, make types consistent
    act_grouped['公司代码'] = act_grouped['公司代码'].astype(float).fillna(0).astype(int).astype(str)
    gen_grouped['公司代码'] = gen_grouped['公司代码'].astype(str)
    act_grouped['总帐科目'] = act_grouped['总帐科目'].astype(str)
    gen_grouped['总账科目'] = gen_grouped['总账科目'].astype(str)
    act_grouped['金额'] = act_grouped['abs_amount'].round(2)
    gen_grouped['金额'] = gen_grouped['金额'].round(2)
    
    merged = pd.merge(
        act_grouped, 
        gen_grouped, 
        left_on=['公司代码', '总帐科目', '借/贷标识', '金额'],
        right_on=['公司代码', '总账科目', '借/贷标识', '金额'],
        how='outer'
    )
    
    merged['act_count'] = merged['act_count'].fillna(0)
    merged['gen_count'] = merged['gen_count'].fillna(0)
    merged['diff'] = merged['gen_count'] - merged['act_count']
    
    matched_perfectly = merged[merged['diff'] == 0]
    extra_in_gen = merged[merged['diff'] > 0]
    missing_in_gen = merged[merged['diff'] < 0]
    
    print("\n--- 核对报告 ---")
    print(f"完全匹配的金额/笔数分组（公司、科目、借贷、金额均一致）: {len(matched_perfectly)} 组")
    print(f"生成的表中多出的笔数分组（即 CBS 中有，但 SAP 银行科目中未找到对应金额）: {len(extra_in_gen)} 组, 共多出 {extra_in_gen['diff'].sum()} 笔")
    print(f"生成的表中缺失的笔数分组（即 SAP 中有，但 CBS 中未找到对应金额）: {len(missing_in_gen)} 组, 共缺失 {-missing_in_gen['diff'].sum()} 笔")
    
    # Save detailed report
    report_file = "SAP与CBS生成结果比对明细.xlsx"
    merged.to_excel(report_file, index=False)
    print(f"\n详细比对明细已保存至: {report_file}")

compare_sap_generated_vs_actual()
