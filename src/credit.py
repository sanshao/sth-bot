import pandas as pd
import os
import json

def pivot_to_json(file_path, sheet_name):
    # 读取指定的工作表
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 转换 DataFrame 中的每一行
    transformed_data = []
    file_name = os.path.basename(file_path).split('.')[0]  # 获取文件名（去后缀）

    for _, row in df.iterrows():
        record = {
            "comments": f"{file_name}-{row['分类']}",  # 替换为实际分类名
            "longText": f"{file_name}-{row['分类']}",  # 替换为实际分类名
            "basePostedCr": abs(row['净值']) if '净值' in row else 0,  # 确保取绝对值
            "basePostedDr": '0',  # 确保取绝对值
            "dir": 'J'
        }
        transformed_data.append(record)

    content = {
        "records": transformed_data
    }
    
    json_list = {
      "agCode": "0000",
      "agCode": "0010",
      "showName": "mytest2",
      "agCate": "Normal",
      "name": "mytest2",
      "content": json.dumps(content, ensure_ascii=False, indent=0)
    }

    # 将转换后的数据转换为 JSON 格式
    json_data = json.dumps([json_list], ensure_ascii=False, indent=0)
    return json_data

current_dir = os.getcwd()
target_directory = os.path.join(current_dir, 'output/20250125_162415')
file_path = os.path.join(target_directory, '天猫-御家专卖支付宝-12月_整理.xlsx')  # 替换为您的文件路径
sheet_name = '透视'  # 替换为透视表的工作表名称
json_output = pivot_to_json(file_path, sheet_name)

# 打印 JSON 数据
print(json_output)


# 保存 JSON 数据到 .tpl 文件
tpl_file_path = os.path.join(target_directory, 'output.tpl')  # 替换为您的目标路径
with open(tpl_file_path, 'w', encoding='utf-8') as tpl_file:
    tpl_file.write(json_output)