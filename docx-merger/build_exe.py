"""
构建脚本 - 使用PyInstaller将docx_merger.py打包为Windows exe
在Windows环境下运行: python build_exe.py
"""
import PyInstaller.__main__
import sys

PyInstaller.__main__.run([
    'docx_merger.py',
    '--onefile',
    '--windowed',
    '--name=DOCX文件合并工具',
    '--clean',
    '--noconfirm',
])
