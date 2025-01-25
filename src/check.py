import openpyxl

def check_merged_cells(file_path):
    # 打开 Excel 文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # 选择活动工作表

    # 获取所有合并的单元格
    merged_cells = sheet.merged_cells.ranges

    # 打印合并的单元格范围
    if merged_cells:
        print("合并的单元格范围:")
        for merged in merged_cells:
            print(merged)
    else:
        print("没有合并的单元格。")

# 使用示例
file_path = '处理后_您的文件名.xlsx'  # 替换为您的文件路径
check_merged_cells(file_path)