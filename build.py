import PyInstaller.__main__
import sys
import os

# 确保在正确的目录中
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# 配置打包参数
PyInstaller.__main__.run([
    'gui.py',  # 主程序文件
    '--name=Word文档批量处理工具',  # 生成的exe文件名
    '--windowed',  # 使用GUI模式，不显示控制台窗口
    '--onefile',  # 打包成单个exe文件
    '--icon=app.ico',  # 如果有图标的话
    '--add-data=README.md;.',  # 添加README文件
    '--clean',  # 清理临时文件
    '--noconfirm',  # 不询问确认
]) 