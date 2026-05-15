import pandas as pd
import os
import json
from category_rules import assign_category_by_rules, assign_tax_rate

def process_file(file_path, target_directory):
    print("当前文件：", file_path)
    
    # 读取Excel文件，尝试自动找到包含“商户订单号”的表头行
    df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
    header_row = 4  # 默认使用4，根据之前的逻辑
    for i, row in df_raw.iterrows():
        if '商户订单号' in row.values:
            header_row = i
            break
            
    df = pd.read_excel(file_path, sheet_name=0, header=header_row, skipfooter=0)
    
    print(f"检测到表头在第 {header_row} 行")
    print("读取的列名：", df.columns.tolist())

    # 定义列名变量（注意全角括号）
    col_income = '收入金额（+元）'
    col_expend = '支出金额（-元）'
    col_desc = '业务描述'

    # 检查必要列是否存在
    if col_income not in df.columns or col_expend not in df.columns:
        print(f"警告：未找到标准列名。当前列名：{df.columns.tolist()}")
        return
    df[col_income] = pd.to_numeric(df[col_income], errors='coerce').fillna(0)
    df[col_expend] = pd.to_numeric(df[col_expend], errors='coerce').fillna(0)

    # 2. 计算净值 (收入 + 支出)
    # 支出金额（-元）在源表中已经是负数，所以直接相加即可得到净值
    df['净值'] = df[col_income] + df[col_expend]

    # 加载分类规则
    json_path = os.path.join(os.path.dirname(__file__), 'rules/pdd_rules.json')
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            rules = json.load(f)
        # 3. 按规则进行归类
        df['分类'] = df.apply(lambda row: assign_category_by_rules(row, rules, default=row[col_desc] if col_desc in df.columns else '其他'), axis=1)
    else:
        print(f"警告：未找到规则文件 {json_path}，使用默认分类")
        if col_desc in df.columns:
            df['分类'] = df[col_desc]
        else:
            df['分类'] = '其他'

    # 添加税率列 (参考 taobao.py)
    df['税率'] = df.apply(assign_tax_rate, axis=1)

    # 4. 创建透视表，汇总“净值”和“行数”
    # 为透视表生成带税率的分类 (参考 taobao.py)
    def get_pivot_category(row):
        cat = str(row['分类'])
        tax = str(row.get('税率', ''))
        target_categories = ['交易收款', '限时红包', '多买多省']
        if any(target in cat for target in target_categories) and tax != '':
            return f"{cat}{tax}"
        return cat

    df['_透视分类'] = df.apply(get_pivot_category, axis=1)
    
    pivot_table = df.groupby('_透视分类')['净值'].agg(['sum', 'count']).reset_index()
    pivot_table.columns = ['分类', '净值', '行数']
    
    # 移除临时列
    df = df.drop(columns=['_透视分类'])

    # 5. 调整列顺序，将“净值”、“分类”、“税率”放在“支出金额（-元）”之后
    cols = list(df.columns)
    if col_expend in cols:
        idx = cols.index(col_expend) + 1
        # 先移除这三个新列
        new_cols = ['净值', '分类', '税率']
        for c in new_cols:
            if c in cols:
                cols.remove(c)
        # 在指定位置插入
        for i, c in enumerate(new_cols):
            cols.insert(idx + i, c)
        df = df[cols]

    # 添加总和行
    total_row = pd.DataFrame({
        '分类': ['总和'],
        '净值': [pivot_table['净值'].sum()],
        '行数': [pivot_table['行数'].sum()]
    })
    pivot_table = pd.concat([pivot_table, total_row], ignore_index=True)

    # 输出结果文件
    base_name = os.path.basename(file_path)
    new_file_name = f"{os.path.splitext(base_name)[0]}_整理.xlsx"
    new_file_path = os.path.join(target_directory, new_file_name)
    
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='整理', index=False)
        pivot_table.to_excel(writer, sheet_name='透视', index=False)

    print(f"处理完成！新文件已生成: {new_file_path}")

