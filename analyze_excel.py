import pandas as pd
import json
import sys

def analyze_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        print(f"--- Analysis for {file_path} ---")
        print("Columns:")
        print(list(df.columns))
        print("\nFirst 3 rows (as dict):")
        print(df.head(3).to_json(orient='records', force_ascii=False))
        print("\nRow count:", len(df))
        print("-" * 50)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")

cbs_file = "原始文件/CBS交易明细列表-20260302-171719.xlsx"
sap_file = "原始文件/SAP银行明细.XLSX"

analyze_excel(cbs_file)
analyze_excel(sap_file)
