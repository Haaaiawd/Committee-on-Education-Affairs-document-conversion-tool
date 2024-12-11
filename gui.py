import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
from process_word import (
    process_word_file, 
    extract_author_number, 
    extract_author_from_filename,
    extract_images_from_doc
)
import sys
import io
import os
import shutil
import re
from docx import Document

class RedirectText:
    def __init__(self, text_widget, error_only=False):
        self.text_widget = text_widget
        self.error_only = error_only
        self.error_files = set()  # 存储错误文件路径
        self.success_files = set()  # 存储成功文件路径
        self.current_file = None  # 当前正在处理的文件

    def set_current_file(self, filename):
        self.current_file = filename

    def write(self, string):
        if not string.strip():
            return
        # 检查是否是错误信息
        if string.strip().startswith('×') or string.strip().startswith('!'):
            # 如果有当前文件，添加到错误文件集合中
            if self.current_file:
                self.error_files.add(self.current_file)
        # 检查是否是成功信息
        elif string.strip().startswith('✓'):
            # 如果有当前文件，添加到成功文件集合中
            if self.current_file:
                self.success_files.add(self.current_file)
            
        # 显示错误信息，每个文件占一行
        if string.strip().startswith('×') or string.strip().startswith('!'):
            if self.current_file:
                if not string.strip().endswith('.docx'):
                    # 提取作者行数字和作者名
                    author_num = extract_author_number(self.current_file)
                    author_name = extract_author_from_filename(self.current_file)
                    error_msg = f"作者{author_num}({author_name}): {string.strip()}\n"
                else:
                    error_msg = string
            else:
                error_msg = string
            
            self.text_widget.insert('end', error_msg)
            self.text_widget.see('end')
            self.text_widget.update()

    def get_error_files(self):
        return self.error_files

    def get_success_files(self):
        return self.success_files

    def clear_files(self):
        self.error_files.clear()
        self.success_files.clear()
        self.current_file = None

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word档批量处理工具")
        self.root.geometry("800x600")
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 输入文件夹选择
        ttk.Label(main_frame, text="输入文件夹:").grid(row=0, column=0, sticky=tk.W)
        self.input_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.choose_input_dir).grid(row=0, column=2)
        
        # 输出文件夹选择
        ttk.Label(main_frame, text="输出文件夹:").grid(row=1, column=0, sticky=tk.W)
        self.output_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="浏览", command=self.choose_output_dir).grid(row=1, column=2)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        # 转换按钮
        self.convert_btn = ttk.Button(button_frame, text="开始转换", command=self.start_conversion)
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # 打包错误文件按钮
        self.pack_error_btn = ttk.Button(button_frame, text="打包错误文件", command=self.pack_error_files)
        self.pack_error_btn.pack(side=tk.LEFT, padx=5)
        self.pack_error_btn.state(['disabled'])
        
        # 添加提取图片按钮
        self.extract_images_btn = ttk.Button(button_frame, text="提取图片", command=self.extract_images)
        self.extract_images_btn.pack(side=tk.LEFT, padx=5)
        
        # 添加提取标题按钮
        self.extract_titles_btn = ttk.Button(button_frame, text="提取标题", command=self.extract_titles)
        self.extract_titles_btn.pack(side=tk.LEFT, padx=5)
        
        # 进度显示
        self.progress_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.progress_var).grid(row=3, column=0, columnspan=3)
        
        # 错误信息显示区域
        ttk.Label(main_frame, text="错误信息:").grid(row=4, column=0, sticky=tk.W)
        self.log_text = scrolledtext.ScrolledText(main_frame, height=20, width=80)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=5)
        
        # 配置grid权重
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 重定向标准输出到文本框
        self.redirect = RedirectText(self.log_text, error_only=True)
        sys.stdout = self.redirect
        sys.stderr = self.redirect

    def choose_input_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.input_path.set(directory)
            # 默认设置相同的输出目录
            if not self.output_path.get():
                self.output_path.set(directory)

    def choose_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)

    def pack_error_files(self):
        error_files = self.redirect.get_error_files()
        if not error_files:
            self.progress_var.set("没有错误文件需要打包")
            return
        
        try:
            # 创建错误文件文件夹
            error_dir = os.path.join(self.output_path.get(), "错误文件")
            os.makedirs(error_dir, exist_ok=True)
            
            # 检查磁盘空间
            if shutil.disk_usage(error_dir).free < 1024 * 1024 * 100:  # 100MB
                self.progress_var.set("磁盘空间不足")
                return
            
            # 复制错误文件
            copied_count = 0
            for filename in error_files:
                try:
                    src_path = os.path.join(self.input_path.get(), filename)
                    if os.path.exists(src_path):
                        dst_path = os.path.join(error_dir, filename)
                        shutil.copy2(src_path, dst_path)
                        copied_count += 1
                except Exception as e:
                    print(f"× 复制文件 {filename} 时出错：{str(e)}")
                
            self.progress_var.set(f"已将 {copied_count} 个错误文件复制到 {error_dir}")
            # 复制完成后禁用打包按钮
            self.pack_error_btn.state(['disabled'])
            
        except PermissionError:
            self.progress_var.set("没有足够的权限创建或访问目录")
        except Exception as e:
            self.progress_var.set(f"打包错误文件时出错: {str(e)}")

    def start_conversion(self):
        input_dir = self.input_path.get()
        output_dir = self.output_path.get()
        
        if not input_dir or not output_dir:
            self.progress_var.set("请选择输入和输出文件夹")
            return
        
        # 检查目录是否存在
        if not os.path.exists(input_dir):
            self.progress_var.set("输入目录不存在")
            return
        
        # 检查是否有 Word 文件
        input_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
        if not input_files:
            self.progress_var.set("输入目录中没有 Word 文件")
            return
        
        # 清空之前的文件记录
        self.redirect.clear_files()
        
        # 禁用所有按钮
        self.convert_btn.state(['disabled'])
        self.pack_error_btn.state(['disabled'])
        self.progress_var.set("处理中...")
        self.log_text.delete(1.0, tk.END)
        
        # 在新线程中运行转换
        def conversion_thread():
            try:
                # 获取目录中的所有文件并排序
                input_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
                sorted_files = sorted(input_files, key=lambda x: extract_author_number(x))
                
                total_files = len(sorted_files)
                self.progress_var.set(f"开始处理，共 {total_files} 个文件...")
                
                # 处理每个文件
                for index, filename in enumerate(sorted_files, 1):
                    self.progress_var.set(f"正在处理第 {index}/{total_files} 个文件...")
                    input_file = os.path.join(input_dir, filename)
                    
                    # 设置当前处理的文件
                    self.redirect.set_current_file(filename)
                    
                    # 处理单个文件
                    process_word_file(input_file, output_dir)
                
                # 处理完成后显示统计信息
                success_count = len(self.redirect.get_success_files())
                failed_count = len(self.redirect.get_error_files())
                
                summary = f"\n处理完成！\n总计: {total_files} 个文件\n成功: {success_count} 个\n失败: {failed_count} 个\n"
                if failed_count > 0:
                    summary += "\n失败的文件作者数字:\n"
                    for failed_file in self.redirect.get_error_files():
                        author_num = extract_author_number(failed_file)
                        summary += f"作者{author_num}\n"
                
                self.log_text.insert('end', summary)
                self.log_text.see('end')
                
                self.root.after(0, self.conversion_complete)
            except Exception as e:
                self.root.after(0, lambda: self.conversion_error(str(e)))
        
        threading.Thread(target=conversion_thread, daemon=True).start()

    def conversion_complete(self):
        success_count = len(self.redirect.get_success_files())
        failed_count = len(self.redirect.get_error_files())
        self.progress_var.set(f"处理完成，成功: {success_count} 个，失败: {failed_count} 个")
        self.convert_btn.state(['!disabled'])
        if failed_count > 0:
            self.pack_error_btn.state(['!disabled'])

    def conversion_error(self, error_message):
        self.progress_var.set(f"处理出错: {error_message}")
        self.convert_btn.state(['!disabled'])
        # 只检查错误文件
        if self.redirect.get_error_files():
            self.pack_error_btn.state(['!disabled'])

    def extract_images(self):
        input_dir = self.input_path.get()
        output_dir = self.output_path.get()
        if not input_dir or not output_dir:
            self.progress_var.set("请选择输入和输出文件夹")
            return
        
        if not os.path.exists(input_dir):
            self.progress_var.set("输入目录不存在")
            return
        
        # 清空显示框
        self.log_text.delete(1.0, tk.END)
        self.progress_var.set("正在提取图片...")
        self.extract_images_btn.state(['disabled'])
        
        def extract_thread():
            try:
                # 创建图片输出目录
                images_dir = os.path.join(self.output_path.get(), "提取的图片")
                os.makedirs(images_dir, exist_ok=True)
                
                # 获取所有Word文件
                docx_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
                total_images = 0
                
                for filename in docx_files:
                    input_file = os.path.join(input_dir, filename)
                    # 为每个文件创建子文件夹
                    file_images_dir = os.path.join(images_dir, os.path.splitext(filename)[0])
                    os.makedirs(file_images_dir, exist_ok=True)
                    
                    # 提取图片
                    temp_images = extract_images_from_doc(input_file, file_images_dir)
                    if temp_images:
                        self.log_text.insert('end', f"✓ {filename}: 提取了 {len(temp_images)} 张图片\n")
                        total_images += len(temp_images)
                    else:
                        self.log_text.insert('end', f"! {filename}: 未找到图片\n")
                    
                summary = f"\n提取完成！\n总计处理 {len(docx_files)} 个文件\n共提取 {total_images} 张图片\n"
                self.log_text.insert('end', summary)
                self.log_text.see('end')
                
                self.root.after(0, lambda: self.progress_var.set(f"图片提取完成，共 {total_images} 张"))
                self.root.after(0, lambda: self.extract_images_btn.state(['!disabled']))
                
            except Exception as e:
                self.root.after(0, lambda: self.progress_var.set(f"提取图片时出错: {str(e)}"))
                self.root.after(0, lambda: self.extract_images_btn.state(['!disabled']))
        
        threading.Thread(target=extract_thread, daemon=True).start()

    def extract_titles(self):
        input_dir = self.input_path.get()
        if not input_dir:
            self.progress_var.set("请选择输入文件夹")
            return
        
        if not os.path.exists(input_dir):
            self.progress_var.set("输入目录不存在")
            return
        
        # 清空显示框
        self.log_text.delete(1.0, tk.END)
        self.progress_var.set("正在提取标题...")
        self.extract_titles_btn.state(['disabled'])
        
        def extract_thread():
            try:
                # 获取所有Word文件
                docx_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
                titles = []
                
                for filename in docx_files:
                    input_file = os.path.join(input_dir, filename)
                    try:
                        doc = Document(input_file)
                        # 获取第一个非空段落作为标题
                        for para in doc.paragraphs:
                            if para.text.strip():
                                titles.append((filename, para.text.strip()))
                                break
                    except Exception as e:
                        self.log_text.insert('end', f"× {filename}: 提取标题失败 - {str(e)}\n")
                
                # 按文件名排序
                titles.sort(key=lambda x: extract_author_number(x[0]))
                
                # 显示标题
                self.log_text.insert('end', "提取的标题：\n" + "="*50 + "\n\n")
                for filename, title in titles:
                    self.log_text.insert('end', f"【{filename}】\n{title}\n\n")
                
                summary = f"\n提取完成！\n总计处理 {len(docx_files)} 个文件\n成功提取 {len(titles)} 个标题\n"
                self.log_text.insert('end', "="*50 + "\n" + summary)
                self.log_text.see('end')
                
                self.root.after(0, lambda: self.progress_var.set(f"标题提取完成，共 {len(titles)} 个"))
                self.root.after(0, lambda: self.extract_titles_btn.state(['!disabled']))
                
            except Exception as e:
                self.root.after(0, lambda: self.progress_var.set(f"提取标题时出错: {str(e)}"))
                self.root.after(0, lambda: self.extract_titles_btn.state(['!disabled']))
        
        threading.Thread(target=extract_thread, daemon=True).start()

# 添加提取作者数字的函数
def extract_author_number(filename):
    """从文件名中提取作者数字"""
    match = re.search(r'(\d+)', filename)
    return int(match.group(1)) if match else float('inf')

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop() 