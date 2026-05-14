import os
from taobao import process_file as process_taobao_file  # 导入淘宝处理函数
from tmall import process_file as process_tmall_file     # 导入天猫处理函数
from group import process_file as process_group_file    # 导入聚合处理函数
from datetime import datetime

def process_taobao_files_in_directory(directory):
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_directory = os.path.join(current_dir, f'output/整理/{timestamp}')
    
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)
        
    for file_name in os.listdir(directory):
        if "聚合" in file_name and file_name.endswith(".xlsx"):
            file_path = os.path.join(directory, file_name)
            process_group_file(file_path, target_directory)
        elif file_name.startswith("淘") and file_name.endswith(".xlsx"):
            file_path = os.path.join(directory, file_name)
            process_taobao_file(file_path, target_directory)
        elif file_name.startswith("天猫") and file_name.endswith(".xlsx"):
            file_path = os.path.join(directory, file_name)
            process_tmall_file(file_path, target_directory)    

# 使用示例
if __name__ == "__main__":
    current_dir = os.getcwd()
    target_directory = os.path.join(current_dir, 'resource/2026/4')
    process_taobao_files_in_directory(target_directory)