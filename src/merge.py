#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具
合并两个xlsx表格，跳过表头和表尾各4行非列表内容
"""

import pandas as pd
import os
import sys
from pathlib import Path


def merge_excel_files(file1_path, file2_path, output_path=None):
    """
    合并两个Excel文件，跳过表头和表尾各4行非列表内容
    
    Args:
        file1_path (str): 第一个Excel文件路径
        file2_path (str): 第二个Excel文件路径
        output_path (str): 输出文件路径，如果为None则自动生成
    
    Returns:
        str: 输出文件路径
    """
    try:
        # 读取第一个Excel文件
        print(f"正在读取第一个文件: {file1_path}")
        df1 = pd.read_excel(file1_path, header=None)
        
        # 读取第二个Excel文件
        print(f"正在读取第二个文件: {file2_path}")
        df2 = pd.read_excel(file2_path, header=None)
        
        # 跳过表头4行，获取数据部分
        data1 = df1.iloc[4:-4] if len(df1) > 8 else df1.iloc[4:]
        data2 = df2.iloc[4:-4] if len(df2) > 8 else df2.iloc[4:]
        
        print(f"第一个文件数据行数: {len(data1)}")
        print(f"第二个文件数据行数: {len(data2)}")
        
        # 合并数据
        merged_data = pd.concat([data1, data2], ignore_index=True)
        print(f"合并后数据行数: {len(merged_data)}")
        
        # 重新构建完整的Excel文件，包含表头和表尾
        # 获取表头（前4行）
        header1 = df1.iloc[:4]
        header2 = df2.iloc[:4]
        
        # 获取表尾（后4行）
        footer1 = df1.iloc[-4:] if len(df1) > 8 else pd.DataFrame()
        footer2 = df2.iloc[-4:] if len(df2) > 8 else pd.DataFrame()
        
        # 使用第一个文件的表头和表尾
        header = header1
        footer = footer1
        
        # 重新设置列名（使用第一行作为列名）
        if len(merged_data) > 0:
            merged_data.columns = header.iloc[0] if len(header) > 0 else range(len(merged_data.columns))
        
        # 生成输出文件路径
        if output_path is None:
            file1_name = Path(file1_path).stem
            file2_name = Path(file2_path).stem
            output_path = f"output/merged_{file1_name}_{file2_name}.xlsx"
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 写入Excel文件
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 写入表头
            if len(header) > 0:
                header.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
            
            # 写入合并的数据
            if len(merged_data) > 0:
                start_row = len(header) + 1 if len(header) > 0 else 1
                merged_data.to_excel(writer, sheet_name='Sheet1', index=False, 
                                   startrow=start_row, header=False)
            
            # 写入表尾
            if len(footer) > 0:
                start_row = len(header) + len(merged_data) + 1 if len(header) > 0 else len(merged_data) + 1
                footer.to_excel(writer, sheet_name='Sheet1', index=False, 
                              startrow=start_row, header=False)
        
        print(f"合并完成！输出文件: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"合并过程中出现错误: {str(e)}")
        return None


def main():
    """主函数"""
    if len(sys.argv) < 3:
        print("使用方法: python merge.py <文件1路径> <文件2路径> [输出文件路径]")
        print("示例: python merge.py resource/file1.xlsx resource/file2.xlsx output/merged.xlsx")
        return
    
    file1_path = sys.argv[1]
    file2_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    # 检查文件是否存在
    if not os.path.exists(file1_path):
        print(f"错误: 文件1不存在 - {file1_path}")
        return
    
    if not os.path.exists(file2_path):
        print(f"错误: 文件2不存在 - {file2_path}")
        return
    
    # 执行合并
    result = merge_excel_files(file1_path, file2_path, output_path)
    
    if result:
        print("✅ 合并成功完成！")
    else:
        print("❌ 合并失败！")


if __name__ == "__main__":
    main()
