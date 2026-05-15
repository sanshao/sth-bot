import pandas as pd
import os

def process_file(file_path, target_directory):
    print("当前文件：", file_path)
    
     # 读取Excel文件
    xls = pd.ExcelFile(file_path)
    
     # 读取第一个工作表
    df = pd.read_excel(xls, sheet_name=0, header=0, skipfooter=0)

    # 1. 将 "费用项" 当分类
    if '费用项' in df.columns:
        df['分类'] = df['费用项']
    else:
        print(f"警告：未找到 '费用项' 列。当前列名：{df.columns.tolist()}")
        return

    # 2. 汇总统计 "金额"
    # 确保金额列是数值类型
    df['金额'] = pd.to_numeric(df['金额'], errors='coerce').fillna(0)
    
    # 财务逻辑：根据“收支方向”处理金额的正负（支出转为负数进行统计）
    if '收支方向' in df.columns:
        def get_calc_amount(row):
            amt = row['金额']
            direction = str(row['收支方向'])
            if '支出' in direction and amt > 0:
                return -amt
            return amt
        df['_calc_amount'] = df.apply(get_calc_amount, axis=1)
    else:
        df['_calc_amount'] = df['金额']

    # 3. 创建透视表，汇总“金额”和“行数”
    pivot_table = df.groupby('分类')['_calc_amount'].agg(['sum', 'count']).reset_index()
    pivot_table.columns = ['分类', '金额', '行数']
    
    # 添加总和行
    total_row = pd.DataFrame({
        '分类': ['总和'],
        '金额': [pivot_table['金额'].sum()],
        '行数': [pivot_table['行数'].sum()]
    })
    pivot_table = pd.concat([pivot_table, total_row], ignore_index=True)
    
    # 移除临时计算列
    df = df.drop(columns=['_calc_amount'])

    # 输出结果文件
    base_name = os.path.basename(file_path)
    new_file_name = f"{os.path.splitext(base_name)[0]}_整理.xlsx"
    new_file_path = os.path.join(target_directory, new_file_name)
    
    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='整理', index=False)
        pivot_table.to_excel(writer, sheet_name='透视', index=False)

    print(f"处理完成！新文件已生成: {new_file_path}")
