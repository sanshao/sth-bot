import os
from taobao import process_taobao_file  # 导入淘宝处理函数
from tmall import process_tmall_file     # 导入天猫处理函数
from datetime import datetime

def process_taobao_files_in_directory(directory):
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target_directory = os.path.join(current_dir, f'output/{timestamp}')
    
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)
        
    for file_name in os.listdir(directory):
        if file_name.startswith("淘宝") and file_name.endswith(".xlsx"):
            file_path = os.path.join(directory, file_name)
            process_taobao_file(file_path, target_directory)
        elif file_name.startswith("天猫") and file_name.endswith(".xlsx"):
            file_path = os.path.join(directory, file_name)
            process_tmall_file(file_path, target_directory)    

# 使用示例
if __name__ == "__main__":
    current_dir = os.getcwd()
    target_directory = os.path.join(current_dir, 'resource/12月支付宝账单')
    process_taobao_files_in_directory(target_directory)