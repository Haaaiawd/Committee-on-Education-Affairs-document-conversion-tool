import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
from process_word import process_folder
import sys
import io
import os
import shutil

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
            
        # 显示错误信息
        if string.strip().startswith('×') or string.strip().startswith('!'):
            if self.current_file:
                if not string.strip().endswith('.docx'):
                    error_msg = f"文件 '{self.current_file}' - {string}"
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
        self.root.title("Word文档批量处理工具")
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
        
        # 打包成功文件按钮
        self.pack_success_btn = ttk.Button(button_frame, text="打包成功文件", command=self.pack_success_files)
        self.pack_success_btn.pack(side=tk.LEFT, padx=5)
        self.pack_success_btn.state(['disabled'])
        
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
            
        # 创建错误文件文件夹
        error_dir = os.path.join(self.output_path.get(), "错误文件")
        if not os.path.exists(error_dir):
            os.makedirs(error_dir)
            
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

    def pack_success_files(self):
        success_files = self.redirect.get_success_files()
        if not success_files:
            self.progress_var.set("没有成功文件需要打包")
            return
            
        # 创建成功文件文件夹
        success_dir = os.path.join(self.output_path.get(), "成功文件")
        if not os.path.exists(success_dir):
            os.makedirs(success_dir)
            
        # 复制成功文件
        copied_count = 0
        for filename in success_files:
            try:
                src_path = os.path.join(self.input_path.get(), filename)
                if os.path.exists(src_path):
                    dst_path = os.path.join(success_dir, filename)
                    shutil.copy2(src_path, dst_path)
                    copied_count += 1
            except Exception as e:
                print(f"× 复制文件 {filename} 时出错：{str(e)}")
                
        self.progress_var.set(f"已将 {copied_count} 个成功文件复制到 {success_dir}")
        # 复制完成后禁用按钮
        self.pack_success_btn.state(['disabled'])

    def start_conversion(self):
        input_dir = self.input_path.get()
        output_dir = self.output_path.get()
        
        if not input_dir or not output_dir:
            self.progress_var.set("请选择输入和输出文件夹")
            return
        
        # 清空之前的文件记录
        self.redirect.clear_files()
        
        # 禁用所有按钮
        self.convert_btn.state(['disabled'])
        self.pack_error_btn.state(['disabled'])
        self.pack_success_btn.state(['disabled'])
        self.progress_var.set("处理中...")
        self.log_text.delete(1.0, tk.END)
        
        # 在新线程中运行转换
        def conversion_thread():
            try:
                process_folder(input_dir, output_dir)
                self.root.after(0, self.conversion_complete)
            except Exception as e:
                self.root.after(0, lambda: self.conversion_error(str(e)))
        
        threading.Thread(target=conversion_thread, daemon=True).start()

    def conversion_complete(self):
        self.progress_var.set("处理完成")
        self.convert_btn.state(['!disabled'])
        # 根据处理结果启用相应的按钮
        if self.redirect.get_error_files():
            self.pack_error_btn.state(['!disabled'])
        if self.redirect.get_success_files():
            self.pack_success_btn.state(['!disabled'])

    def conversion_error(self, error_message):
        self.progress_var.set(f"处理出错: {error_message}")
        self.convert_btn.state(['!disabled'])
        # 根据处理结果启用相应的按钮
        if self.redirect.get_error_files():
            self.pack_error_btn.state(['!disabled'])
        if self.redirect.get_success_files():
            self.pack_success_btn.state(['!disabled'])

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop() 