import pandas as pd

def generate_rule_template():
    # Define the structure of the rule mapping template
    columns = [
        "规则编号",
        "匹配字段 (如: 对方账户名称, 摘要, 用途)",
        "匹配关键字 (如: 捷付, 房租, 结息)",
        "SAP对方科目 (如: 11220101)",
        "SAP凭证类型 (如: DZ, SA)",
        "SAP生成文本模板 (如: 收捷付营业款, 支付房租)",
        "优先级 (1最高)",
        "备注说明"
    ]
    
    # Add some example rules to help the user understand
    example_rules = [
        {
            "规则编号": "R001",
            "匹配字段 (如: 对方账户名称, 摘要, 用途)": "用途",
            "匹配关键字 (如: 捷付, 房租, 结息)": "捷付入账",
            "SAP对方科目 (如: 11220101)": "1122000000",
            "SAP凭证类型 (如: DZ, SA)": "DZ",
            "SAP生成文本模板 (如: 收捷付营业款, 支付房租)": "收捷付睿通营业款",
            "优先级 (1最高)": 1,
            "备注说明": "捷付睿通营业款自动记入应收账款"
        },
        {
            "规则编号": "R002",
            "匹配字段 (如: 对方账户名称, 摘要, 用途)": "摘要",
            "匹配关键字 (如: 捷付, 房租, 结息)": "结息",
            "SAP对方科目 (如: 11220101)": "6603000000",
            "SAP凭证类型 (如: DZ, SA)": "SA",
            "SAP生成文本模板 (如: 收捷付营业款, 支付房租)": "收到银行结息",
            "优先级 (1最高)": 2,
            "备注说明": "财务费用-利息收入"
        },
        {
            "规则编号": "R003",
            "匹配字段 (如: 对方账户名称, 摘要, 用途)": "用途",
            "匹配关键字 (如: 捷付, 房租, 结息)": "房租",
            "SAP对方科目 (如: 11220101)": "2202000000",
            "SAP凭证类型 (如: DZ, SA)": "SA",
            "SAP生成文本模板 (如: 收捷付营业款, 支付房租)": "支付房租",
            "优先级 (1最高)": 2,
            "备注说明": "应付账款-房租"
        }
    ]
    
    df = pd.DataFrame(example_rules, columns=columns)
    
    # Add 10 empty rows for user to fill
    empty_rows = pd.DataFrame([[None] * len(columns)] * 10, columns=columns)
    df = pd.concat([df, empty_rows], ignore_index=True)
    
    output_file = "业务入账规则模板.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Generated {output_file}")

generate_rule_template()
