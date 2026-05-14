import pandas as pd
import os
import json
from datetime import datetime
from category_rules import assign_category_by_rules, assign_tax_rate


def process_file(file_path, target_directory):
    print("当前文件：", file_path)

    # 读取Excel文件
    xls = pd.ExcelFile(file_path)

    # 读取第一个工作表
    # 根据用户提供的列名：商家昵称, 入帐日期, 入帐时间, 支付流水号, 主订单id, 子订单id, 入帐类型, 收入金额(元), 支出金额(元), 业务描述, 备注, 收/付渠道, 数据创建时间, 数据修改时间
    new_columns = ['商家昵称', '入帐日期', '入帐时间', '支付流水号', '主订单id', '子订单id', '入帐类型', '收入金额(元)', '支出金额(元)', '业务描述', '备注', '收/付渠道', '数据创建时间', '数据修改时间']
    
    df = pd.read_excel(xls, sheet_name=0, header=None)
    
    # 检查第一行是否包含预期的列名，如果包含则说明有表头
    first_row = [str(x).strip() for x in df.iloc[0].tolist()]
    if '商家昵称' in first_row or '支付流水号' in first_row:
        df = pd.read_excel(xls, sheet_name=0, header=0)
        df.columns = df.columns.str.strip()
    else:
        df.columns = new_columns

    # 打印列名以供检查
    print("读取的列名：", df.columns.tolist())

    # 确保金额列是数值类型
    df['收入金额(元)'] = pd.to_numeric(df['收入金额(元)'], errors='coerce').fillna(0)
    df['支出金额(元)'] = pd.to_numeric(df['支出金额(元)'], errors='coerce').fillna(0)

    # 计算净值并添加到新列 (收入 - 支出)
    df['净值'] = df['收入金额(元)'] - df['支出金额(元)']

    json_path = os.path.join(os.path.dirname(__file__), 'rules/group_rules.json')
    with open(json_path, 'r', encoding='utf-8') as f:
        rules = json.load(f)

    # 添加分类列
    df['分类'] = df.apply(lambda row: assign_category_by_rules(row, rules), axis=1)

    # 添加税率列
    df['税率'] = df.apply(assign_tax_rate, axis=1)

    # 调整列顺序，将“净值”列放在“支出金额(元)”后
    columns_order = list(df.columns)
    # 找到“支出金额(元)”的位置
    outflow_index = columns_order.index('支出金额(元)')

    # 将“净值”移动到“支出金额(元)”之后
    net_value_index = columns_order.index('净值')
    columns_order.insert(outflow_index + 1, columns_order.pop(net_value_index))
    df = df[columns_order]

    # 为透视表生成带税率的分类
    def get_pivot_category(row):
        cat = str(row['分类'])
        tax = str(row['税率'])
        target_categories = ['交易收款', '限时红包', '消费券']
        if any(target in cat for target in target_categories) and tax != '':
            return f"{cat}{tax}"
        return cat

    df['_透视分类'] = df.apply(get_pivot_category, axis=1)

    # 创建透视表来计算分类下的净值、行数和总和
    pivot_table = df.pivot_table(values='净值', index='_透视分类', aggfunc=['sum', 'count']).reset_index()
    pivot_table.columns = ['分类', '净值', '行数']
    
    # 移除临时列
    df = df.drop(columns=['_透视分类'])

    # 添加总和行
    total_row = pd.DataFrame({'分类': ['总和'], '净值': [pivot_table['净值'].sum()], '行数': [pivot_table['行数'].sum()]})
    pivot_table = pd.concat([pivot_table, total_row], ignore_index=True)
    
    base_name = os.path.basename(file_path)
    new_file_name = f"{os.path.splitext(base_name)[0]}_整理.xlsx"
    
    new_file_path = os.path.join(target_directory, new_file_name)
    
    # print("新文件名：", new_file_path, base_name)

    # 创建新的工作表 “整理” “透视”
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='整理', index=False)
        pivot_table.to_excel(writer, sheet_name='透视', index=False)

    print(f"处理完成！新文件生成: {new_file_path}")
