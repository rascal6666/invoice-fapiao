#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 将GUI程序打包成Windows可执行文件

使用方法：
python build_exe.py
"""

import os
import sys
import subprocess
import shutil

def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller安装成功")
        return True
    except subprocess.CalledProcessError:
        print("❌ PyInstaller安装失败")
        return False

def create_spec_file():
    """创建PyInstaller配置文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'pdfplumber',
        'openai',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'threading',
        'json',
        'datetime',
        'os',
        'sys'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='发票识别器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    with open('invoice_recognizer.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ 配置文件创建成功")

def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # 检查必要文件
    required_files = ['gui_app.py', 'entry.py']
    missing_files = [f for f in required_files if not os.path.exists(f)]
    
    if missing_files:
        print(f"❌ 缺少必要文件: {missing_files}")
        return False
    
    # 创建spec文件
    create_spec_file()
    
    # 运行PyInstaller
    try:
        cmd = [sys.executable, "-m", "PyInstaller", "--clean", "invoice_recognizer.spec"]
        print(f"执行命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        
        if result.returncode == 0:
            print("✅ 构建成功！")
            print("可执行文件位置: dist/发票识别器.exe")
            return True
        else:
            print("❌ 构建失败")
            print("错误输出:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"❌ 构建过程中出现错误: {e}")
        return False

def create_installer():
    """创建安装包（可选）"""
    print("\n是否创建安装包？(y/n): ", end="")
    choice = input().lower().strip()
    
    if choice == 'y':
        print("正在创建安装包...")
        # 这里可以添加创建安装包的代码
        # 例如使用NSIS或其他工具
        print("安装包功能暂未实现")

def main():
    """主函数"""
    print("=== 发票识别器打包工具 ===")
    print()
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("❌ 需要Python 3.7或更高版本")
        return
    
    # 安装PyInstaller
    if not install_pyinstaller():
        return
    
    # 构建可执行文件
    if build_executable():
        print("\n🎉 打包完成！")
        print("可执行文件: dist/发票识别器.exe")
        print("您可以将此文件分发给其他用户使用")
        
        # 询问是否创建安装包
        create_installer()
    else:
        print("\n❌ 打包失败，请检查错误信息")

if __name__ == "__main__":
    main() 