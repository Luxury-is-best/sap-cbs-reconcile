import pandas as pd
import warnings

warnings.filterwarnings('ignore')

def infer_mapping_rules():
    sap_file = '原始文件/SAP银行明细.XLSX'
    cbs_file = '原始文件/CBS交易明细列表-20260302-171719.xlsx'
    
    print("加载数据中...")
    df_sap = pd.read_excel(sap_file)
    
    df_cbs_raw = pd.read_excel(cbs_file, header=None, nrows=10)
    header_row = 0
    for idx, row in df_cbs_raw.iterrows():
        if any('交易' in str(val) or '日期' in str(val) for val in row.values):
            header_row = idx
            break
    df_cbs = pd.read_excel(cbs_file, header=header_row)

    # 预处理 SAP
    # 过滤掉金额为空的行
    df_sap = df_sap.dropna(subset=['公司代码货币价值'])
    # 获取日期部分
    df_sap['SAP日期'] = pd.to_datetime(df_sap['凭证日期']).dt.date
    df_sap['SAP金额'] = df_sap['公司代码货币价值']
    df_sap['公司代码'] = df_sap['公司代码'].fillna(0).astype(int).astype(str)

    # 预处理 CBS
    # 日期转为 date
    df_cbs['CBS日期'] = pd.to_datetime(df_cbs['交易日期']).dt.date
    # 金额处理：收入为正，支出为负，统一为一个金额列，方便与SAP比较
    # SAP中：公司代码货币价值，正数通常代表借方（收钱，或者银行存款增加），负数代表贷方（花钱，银行存款减少）
    # 在CBS中：贷(收入) 是收钱，借(支出) 是花钱
    df_cbs['贷(收入)'] = df_cbs['贷(收入)'].fillna(0)
    df_cbs['借(支出)'] = df_cbs['借(支出)'].fillna(0)
    df_cbs['CBS金额'] = df_cbs['贷(收入)'] - df_cbs['借(支出)']
    df_cbs['单位编码'] = df_cbs['单位编码'].fillna(0).astype(int).astype(str)
    
    # 我们尝试寻找1对1的匹配，以及N对1的匹配
    
    print("\n--- 尝试寻找 1对1 完全金额匹配 ---")
    # 为了简单，按 (单位编码, 日期, 金额) 进行分组合并，如果两边都有唯一的记录，就认为是 1对1 匹配
    sap_grouped = df_sap.groupby(['公司代码', 'SAP日期', 'SAP金额'])
    cbs_grouped = df_cbs.groupby(['单位编码', 'CBS日期', 'CBS金额'])
    
    one_to_one_matches = []
    
    for (sap_code, date, amt), sap_group in sap_grouped:
        cbs_key = (sap_code, date, amt)
        if cbs_key in cbs_grouped.groups:
            cbs_group = df_cbs.iloc[cbs_grouped.groups[cbs_key]]
            if len(sap_group) == 1 and len(cbs_group) == 1:
                # 找到完美1对1
                sap_row = sap_group.iloc[0]
                cbs_row = cbs_group.iloc[0]
                
                sap_subject = sap_row['总账科目：长文本']
                sap_text = sap_row['文本']
                cbs_usage = cbs_row['用途']
                cbs_summary = cbs_row['摘要']
                cbs_counterparty = cbs_row['对方账户名称']
                
                one_to_one_matches.append({
                    '公司代码': sap_code,
                    '日期': date,
                    '金额': amt,
                    'SAP总账科目': sap_subject,
                    'SAP文本': sap_text,
                    'CBS用途': cbs_usage,
                    'CBS摘要': cbs_summary,
                    'CBS对方账户': cbs_counterparty
                })
                
    print(f"找到 {len(one_to_one_matches)} 条潜在的 1对1 匹配规律。")
    if len(one_to_one_matches) > 0:
        matches_df = pd.DataFrame(one_to_one_matches)
        # 统计最常见的 映射规则 (CBS摘要/用途 -> SAP科目)
        matches_df['CBS特征'] = matches_df['CBS摘要'].fillna('') + " | " + matches_df['CBS用途'].fillna('')
        rule_stats = matches_df.groupby(['CBS特征', 'SAP总账科目']).size().reset_index(name='频次').sort_values('频次', ascending=False)
        print("\n高频 1对1 映射规则倒推 (Top 10):")
        print(rule_stats.head(10).to_string(index=False))

    print("\n--- 尝试寻找 多对1 (合并记账) 匹配 ---")
    # 假设同一天、同一个单位、相同 CBS摘要 的多笔明细，可能合并成 SAP 的一笔
    # 先对CBS按 (单位编码, 日期, 摘要) 汇总金额
    cbs_sum_grouped = df_cbs.groupby(['单位编码', 'CBS日期', '摘要']).agg(
        总金额=('CBS金额', 'sum'),
        笔数=('CBS金额', 'count'),
        用途组合=('用途', lambda x: ' / '.join([str(i) for i in set(x) if pd.notna(i)]))
    ).reset_index()
    
    # 只看多笔合并的
    cbs_multi = cbs_sum_grouped[cbs_sum_grouped['笔数'] > 1]
    
    multi_matches = []
    sap_lookup = sap_grouped.size() # 方便快速查找 (代码, 日期, 金额) 组合
    
    for _, row in cbs_multi.iterrows():
        key = (row['单位编码'], row['CBS日期'], round(row['总金额'], 2))
        # 在SAP中寻找是否有这一笔总金额
        sap_candidates = df_sap[(df_sap['公司代码'] == key[0]) & (df_sap['SAP日期'] == key[1]) & (df_sap['SAP金额'].round(2) == key[2])]
        if not sap_candidates.empty:
            sap_row = sap_candidates.iloc[0]
            multi_matches.append({
                '公司代码': key[0],
                '日期': key[1],
                'CBS摘要': row['摘要'],
                '合并笔数': row['笔数'],
                'CBS总金额': key[2],
                'SAP科目': sap_row['总账科目：长文本'],
                'SAP文本': sap_row['文本']
            })

    print(f"找到 {len(multi_matches)} 条潜在的 多对1 匹配规律。")
    if len(multi_matches) > 0:
        multi_df = pd.DataFrame(multi_matches)
        print("\n示例 多对1 映射 (如工资合并、报销合并等):")
        print(multi_df.head(10).to_string(index=False))

if __name__ == '__main__':
    infer_mapping_rules()
