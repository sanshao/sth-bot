import pandas as pd
import os

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

    # 2. 计算净值 (收入 - 支出)
    # PDD账单中，支出金额（-元）通常为正数，代表扣除
    df['净值'] = df[col_income] + df[col_expend]

    # 3. 按业务描述进行归类 (用户明确要求)
    if col_desc in df.columns:
        df['分类'] = df[col_desc]
    else:
        df['分类'] = '其他'

    # 4. 创建透视表，汇总“净值”和“行数”
    pivot_table = df.groupby('分类')['净值'].agg(['sum', 'count']).reset_index()
    pivot_table.columns = ['分类', '净值', '行数']
    
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

