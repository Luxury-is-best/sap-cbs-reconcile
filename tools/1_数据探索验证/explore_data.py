import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def explore_excel_files():
    sap_file = '原始文件/SAP银行明细.XLSX'
    cbs_file = '原始文件/CBS交易明细列表-20260302-171719.xlsx'
    
    print("=== 读取文件 ===")
    try:
        # SAP通常在第一行就是表头
        df_sap = pd.read_excel(sap_file)
        print(f"✅ 成功读取 SAP 文件: {df_sap.shape[0]} 行, {df_sap.shape[1]} 列")
    except Exception as e:
        print(f"❌ 读取 SAP 文件失败: {e}")
        return

    try:
        # CBS可能包含前面的标题行，我们需要找到真正的表头。通常是在前几行。
        # 我们可以尝试跳过前几行直到找到包含特定关键字的行，或者先读取看一下。
        df_cbs_raw = pd.read_excel(cbs_file, header=None, nrows=10)
        
        # 寻找包含 "交易" 或者 "日期" 类似字样的那一行作为表头
        header_row = 0
        for idx, row in df_cbs_raw.iterrows():
            if any('日期' in str(val) or '交易' in str(val) or '金额' in str(val) for val in row.values):
                header_row = idx
                break
                
        df_cbs = pd.read_excel(cbs_file, header=header_row)
        print(f"✅ 成功读取 CBS 文件: {df_cbs.shape[0]} 行, {df_cbs.shape[1]} 列 (表头在第 {header_row} 行)")
    except Exception as e:
        print(f"❌ 读取 CBS 文件失败: {e}")
        return

    print("\n=== SAP 字段概览 ===")
    print("列名:", df_sap.columns.tolist())
    print("\n=== SAP 示例数据 (前2行) ===")
    print(df_sap.head(2).to_dict('records'))

    print("\n=== CBS 字段概览 ===")
    print("列名:", df_cbs.columns.tolist())
    print("\n=== CBS 示例数据 (前2行) ===")
    print(df_cbs.head(2).to_dict('records'))

if __name__ == '__main__':
    explore_excel_files()
