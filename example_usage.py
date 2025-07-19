#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票识别批量处理示例

使用方法：
1. 将PDF文件放在指定目录中
2. 运行此脚本处理所有PDF文件并生成Excel表格
"""

from entry import process_directory_to_xlsx
import os

def main():
    # 指定包含PDF文件的目录路径
    pdf_directory = "./pdf_files"  # 可以根据实际情况修改
    
    # 检查目录是否存在
    if not os.path.exists(pdf_directory):
        print(f"目录 {pdf_directory} 不存在，请创建该目录并放入PDF文件")
        return
    
    # 检查目录中是否有PDF文件
    pdf_files = [f for f in os.listdir(pdf_directory) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"在目录 {pdf_directory} 中未找到PDF文件")
        print("请将PDF文件放入该目录后重新运行")
        return
    
    print(f"找到 {len(pdf_files)} 个PDF文件:")
    for pdf_file in pdf_files:
        print(f"  - {pdf_file}")
    
    # 处理所有PDF文件并生成Excel表格
    output_filename = "发票数据汇总.xlsx"
    print(f"\n开始处理，输出文件: {output_filename}")
    
    try:
        process_directory_to_xlsx(pdf_directory, output_filename)
        print("\n✅ 处理完成！")
    except Exception as e:
        print(f"\n❌ 处理过程中出现错误: {e}")

if __name__ == "__main__":
    main() 