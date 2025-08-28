#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller打包配置文件
用于自定义打包选项和优化
"""

import os
import sys
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent

# PyInstaller配置
PYINSTALLER_CONFIG = {
    'gui': {
        'script': 'mail_exporter_gui.py',
        'name': 'mail_exporter_gui',
        'options': [
            '--onefile',
            '--windowed',
            '--name=mail_exporter_gui',
            '--add-data=README.md;.',
            '--add-data=LICENSE;.',
            '--add-data=requirements.txt;.',
            '--hidden-import=tkinter',
            '--hidden-import=tkinter.ttk',
            '--hidden-import=tkinter.messagebox',
            '--hidden-import=tkinter.filedialog',
            '--hidden-import=threading',
            '--hidden-import=queue',
            '--exclude-module=matplotlib',
            '--exclude-module=numpy',
            '--exclude-module=pandas',
            '--exclude-module=scipy',
            '--exclude-module=PIL',
            '--exclude-module=cv2',
            '--optimize=2',
            '--strip',
            '--noupx',
        ]
    },
    'cli': {
        'script': 'mail_exporter.py',
        'name': 'mail_exporter',
        'options': [
            '--onefile',
            '--console',
            '--name=mail_exporter',
            '--add-data=README.md;.',
            '--add-data=LICENSE;.',
            '--add-data=requirements.txt;.',
            '--exclude-module=tkinter',
            '--exclude-module=matplotlib',
            '--exclude-module=numpy',
            '--exclude-module=pandas',
            '--exclude-module=scipy',
            '--exclude-module=PIL',
            '--exclude-module=cv2',
            '--optimize=2',
            '--strip',
            '--noupx',
        ]
    }
}

def build_command(config_key):
    """生成PyInstaller命令"""
    config = PYINSTALLER_CONFIG[config_key]
    script = config['script']
    options = ' '.join(config['options'])
    return f'pyinstaller {options} {script}'

def clean_build_files():
    """清理构建文件"""
    import shutil
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        dir_path = PROJECT_ROOT / dir_name
        if dir_path.exists():
            print(f"删除目录: {dir_path}")
            shutil.rmtree(dir_path)
    
    import glob
    for pattern in files_to_clean:
        for file_path in glob.glob(str(PROJECT_ROOT / pattern)):
            print(f"删除文件: {file_path}")
            os.remove(file_path)

def main():
    """主函数"""
    import subprocess
    
    print("邮件导出工具 - PyInstaller打包脚本")
    print("=" * 50)
    
    # 清理之前的构建文件
    print("\n1. 清理构建文件...")
    clean_build_files()
    
    # 检查依赖
    print("\n2. 检查依赖...")
    try:
        import PyInstaller
        print(f"PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("错误: 未安装PyInstaller")
        print("请运行: pip install pyinstaller")
        return 1
    
    # 构建GUI程序
    print("\n3. 构建GUI程序...")
    gui_cmd = build_command('gui')
    print(f"执行命令: {gui_cmd}")
    result = subprocess.run(gui_cmd, shell=True, cwd=PROJECT_ROOT)
    if result.returncode != 0:
        print("错误: GUI程序构建失败")
        return 1
    
    # 构建CLI程序
    print("\n4. 构建CLI程序...")
    cli_cmd = build_command('cli')
    print(f"执行命令: {cli_cmd}")
    result = subprocess.run(cli_cmd, shell=True, cwd=PROJECT_ROOT)
    if result.returncode != 0:
        print("错误: CLI程序构建失败")
        return 1
    
    print("\n" + "=" * 50)
    print("构建完成！")
    print("可执行文件位于 dist/ 目录")
    
    # 显示生成的文件
    dist_dir = PROJECT_ROOT / 'dist'
    if dist_dir.exists():
        print("\n生成的文件:")
        for file_path in dist_dir.iterdir():
            if file_path.is_file():
                size_mb = file_path.stat().st_size / (1024 * 1024)
                print(f"  {file_path.name} ({size_mb:.1f} MB)")
    
    return 0

if __name__ == '__main__':
    sys.exit(main())