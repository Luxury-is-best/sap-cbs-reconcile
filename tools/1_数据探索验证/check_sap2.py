import pandas as pd

def check_sap_doc_full():
    sap = pd.read_excel("原始文件/SAP银行明细.XLSX")
    doc = sap[sap['凭证编号'] == 500000032.0]
    print(f"Total lines in doc 500000032: {len(doc)}")
    credits = doc[doc['借/贷标识'] == 'H']
    print(f"Credit lines:")
    if not credits.empty:
        print(credits[['行项目', '过账码', '总帐科目', '借/贷标识', '公司代码货币价值', '客户', '文本', '总账科目：长文本']].to_string())
    else:
        print("No credit lines found for this document in the SAP file!")
        
check_sap_doc_full()
