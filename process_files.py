import os
import zipfile
import shutil
import logging
import re
from process_word import process_single_file

def extract_author_number(filename):
    """从文件名中提取作者数字"""
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

def process_files(input_files, output_dir):
    successful_files = []
    failed_files = []
    
    # 按作者数字排序输入文件
    sorted_input_files = sorted(input_files, key=lambda x: extract_author_number(os.path.basename(x)))
    
    for input_file in sorted_input_files:
        try:
            # 获取文件名（不含扩展名）
            file_name = os.path.splitext(os.path.basename(input_file))[0]
            
            # 为每个成功文件创建独立的输出目录
            success_output_dir = os.path.join(output_dir, file_name)
            os.makedirs(success_output_dir, exist_ok=True)
            
            # 处理文件
            process_single_file(input_file, success_output_dir)
            successful_files.append(input_file)
            
        except Exception as e:
            failed_files.append((input_file, str(e)))
            logging.error(f"处理文件 {input_file} 时发生错误: {str(e)}")
    
    # 处理失败文件
    if failed_files:
        failed_dir = os.path.join(output_dir, "failed")
        os.makedirs(failed_dir, exist_ok=True)
        
        # 创建失败文件的压缩包
        failed_zip_path = os.path.join(failed_dir, "failed_files.zip")
        with zipfile.ZipFile(failed_zip_path, 'w') as failed_zip:
            for failed_file, error in failed_files:
                # 将失败文件复制到失败目录
                failed_file_name = os.path.basename(failed_file)
                shutil.copy2(failed_file, os.path.join(failed_dir, failed_file_name))
                # 添加到压缩包
                failed_zip.write(failed_file, failed_file_name)
                
        # 创建错误日志
        error_log_path = os.path.join(failed_dir, "error_log.txt")
        with open(error_log_path, 'w', encoding='utf-8') as f:
            for failed_file, error in failed_files:
                f.write(f"文件: {failed_file}\n错误: {error}\n\n")
    
    return successful_files, failed_files 