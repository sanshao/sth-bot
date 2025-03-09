import pandas as pd
import os

def generate_voucher(file_path, sheet_name):
    # 读取透视表数据
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    file_name = os.path.basename(file_path).split('.')[0]  # 获取文件名（去后缀）

    # 创建一个新的 DataFrame 用于存放凭证数据
    voucher_data = []
    
    for _, row in df.iterrows():
        
        basePostedDr = abs(row['净值']) if '净值' in row else 0
        basePostedCr = abs(row['净值']) if '净值' in row else 0
        
        # 根据透视表中的每一行生成凭证
        voucher_row = {
            "凭证类别": "记",
            "凭证号": "110",  # 根据需要设置凭证号
            "凭证日期": "2023-11-11",  # 根据需要设置凭证日期
            "附单据数": "",  # 根据需要设置附单据数
            "摘要": f"{file_name}-{row['分类']}", 
            "科目编码": "0",  
            "借方金额": basePostedDr,
            "贷方金额": basePostedCr,
            "项目编码": "",
            "项目": "",
            "客户编码": "",
            "客户": "",
            "供应商编码": "",
            "供应商": "",
            "部门编码": "",
            "部门": "",
            "员工编码": "",
            "员工": "",
            "存货编码": "",
            "存货": "",
            "规格型号": "",
            "数量": "",
            "计量单位": "",
            "单价": "",
            "外币金额": "",
            "币种": "",
            "汇率": "",
            "制单人": "余丹",
            "审核人": ""
        }
        voucher_data.append(voucher_row)

    # 将凭证数据转换为 DataFrame
    voucher_df = pd.DataFrame(voucher_data)

    # 创建凭证文件存储路径
    output_file_path = os.path.join(os.getcwd(), f'{file_name}_凭证文件.xlsx')
    # 将凭证数据写入 Excel 文件
    voucher_df.to_excel(output_file_path, index=False)

    print(f"凭证文件已生成：{output_file_path}")

# 使用示例
current_dir = os.getcwd()
target_directory = os.path.join(current_dir, 'output/20250125_162415')
file_path = os.path.join(target_directory, '天猫-御家专卖支付宝-12月_整理.xlsx')  # 替换为您的文件路径
sheet_name = '透视'  # 替换为透视表的工作表名称

generate_voucher(file_path, sheet_name)