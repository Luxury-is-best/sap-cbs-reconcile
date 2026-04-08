import pandas as pd
import numpy as np

def match_data():
    cbs = pd.read_excel("原始文件/CBS交易明细列表-20260302-171719.xlsx")
    sap = pd.read_excel("原始文件/SAP银行明细.XLSX")

    # Clean SAP dates
    sap['凭证日期'] = pd.to_datetime(sap['凭证日期'], unit='ms')
    sap['过账日期'] = pd.to_datetime(sap['过账日期'], unit='ms')
    
    # Fill NA and get amounts
    cbs['借(支出)'] = cbs['借(支出)'].fillna(0)
    cbs['贷(收入)'] = cbs['贷(收入)'].fillna(0)
    
    # CBS amount: For bank, income is debit (S) in SAP, expense is credit (H) in SAP
    # But let's just match absolute amounts first
    cbs['abs_amount'] = np.where(cbs['借(支出)'] > 0, cbs['借(支出)'], cbs['贷(收入)'])
    
    sap['abs_amount'] = sap['公司代码货币价值'].abs()
    
    # Try to find a specific amount to see how it maps
    print("Finding matching amounts...")
    sample_amounts = [3734.45, 1672.75]
    for amt in sample_amounts:
        print(f"\nLooking for amount {amt} in CBS:")
        matches = cbs[cbs['abs_amount'] == amt]
        for _, row in matches.iterrows():
            print(f"Date: {row['交易日期']}, Account: {row['账号']}, Income: {row['贷(收入)']}, Expense: {row['借(支出)']}, Purpose: {row['用途']}, Summary: {row['摘要']}")

        print(f"\nLooking for amount {amt} in SAP:")
        matches_sap = sap[sap['abs_amount'] == amt]
        for _, row in matches_sap.iterrows():
            print(f"Date: {row['凭证日期']}, Account: {row['总账科目：长文本']}, Text: {row['文本']}, Doc Header: {row['凭证抬头文本']}, Company: {row['公司代码']}")

match_data()
