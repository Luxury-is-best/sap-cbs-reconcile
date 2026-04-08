import pandas as pd

def check_sap_expense():
    sap = pd.read_excel("原始文件/SAP银行明细.XLSX")
    sap['abs_amount'] = sap['公司代码货币价值'].abs()
    doc = sap[sap['abs_amount'] == 158427.96]
    print(doc[['公司代码', '行项目', '过账码', '总帐科目', '借/贷标识', '公司代码货币价值', '文本']].to_string())

check_sap_expense()
