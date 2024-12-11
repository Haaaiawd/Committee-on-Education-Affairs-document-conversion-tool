import PyInstaller.__main__
import os
import shutil

# 删除之前的构建文件
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')
if os.path.exists('release'):
    shutil.rmtree('release')

# PyInstaller 配置
PyInstaller.__main__.run([
    'gui.py',                          # 主程序文件
    '--name=Word文档批量处理工具',      # 生成的exe名称
    '--windowed',                      # 使用GUI模式
    '--onefile',                       # 打包成单个exe文件
    '--icon=app.ico',                  # 程序图标
    '--add-data=README.md;.',          # 添加README文件
    '--clean',                         # 清理临时文件
    '--noconfirm',                     # 不确认覆盖
    '--uac-admin',                     # 请求管理员权限
    '--hidden-import=docx',            # 添加隐藏导入
    '--hidden-import=docx2python',
    '--version-file=file_version_info.txt',  # 添加版本信息
])

# 清理 dist 目录中的非 exe 文件
dist_dir = 'dist'
if os.path.exists(dist_dir):
    for item in os.listdir(dist_dir):
        item_path = os.path.join(dist_dir, item)
        if not item.endswith('.exe'):
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)

# 创建发布目录并复制文件
os.makedirs('release', exist_ok=True)

# 复制exe文件到release目录
exe_path = os.path.join('dist', 'Word文档批量处理工具.exe')
if os.path.exists(exe_path):
    # 复制exe文件
    shutil.copy2(
        exe_path,
        os.path.join('release', 'Word文档批量处理工具.exe')
    )
    
    # 复制README文件
    shutil.copy2(
        'README.md',
        os.path.join('release', 'README.md')
    )
    
    print("构建完成！文件已保存到 release 目录")
    print("可以直接运行 release 目录中的 exe 文件")
else:
    print("错误：构建失败，exe文件不存在")

# 清理构建过程中生成的临时文件
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('__pycache__'):
    shutil.rmtree('__pycache__')
for file in os.listdir('.'):
    if file.endswith('.spec'):
        os.remove(file) 