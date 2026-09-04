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
    df = pd.read_excel(xls, sheet_name=0, header=4, skipfooter=4)

    # 打印列名以供检查
    print("读取的列名：", df.columns.tolist())

    # 清理列名，去除前后空格
    df.columns = df.columns.str.strip()

    # 计算净值并添加到新列
    df['净值'] = df['收入金额（+元）'] + df['支出金额（-元）']

    # 计算净值并添加到新列
    df['净值'] = df['收入金额（+元）'] + df['支出金额（-元）']

    json_path = os.path.join(os.path.dirname(__file__), 'rules/taocc_rules.json')
    with open(json_path, 'r', encoding='utf-8') as f:
        rules = json.load(f)

    # 添加分类列
    df['分类'] = df.apply(lambda row: assign_category_by_rules(row, rules), axis=1)

    # 添加税率列
    df['税率'] = df.apply(assign_tax_rate, axis=1)

    # 调整列顺序，将“净值”列放在“支出金额（-元）”后、账户余额之前
    columns_order = list(df.columns)
    # 找到“支出金额（-元）”和“账户余额（元）”的位置
    outflow_index = columns_order.index('支出金额（-元）')
    balance_index = columns_order.index('账户余额（元）')

    # 重新排列列顺序
    columns_order.insert(balance_index, columns_order.pop(outflow_index + 1))  # 移动“净值”到“支出金额（-元）”和“账户余额（元）”之间
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
    
    # 将透视表打印到控制台
    print(pivot_table)
