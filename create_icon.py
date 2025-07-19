#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建图标文件的脚本

注意：这是一个简单的示例，实际使用时建议使用专业的图标设计工具
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    """创建简单的图标"""
    # 创建一个32x32的图像
    size = 32
    image = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制背景圆形
    draw.ellipse([2, 2, size-2, size-2], fill=(70, 130, 180, 255), outline=(50, 100, 150, 255), width=2)
    
    # 绘制文档图标
    # 文档主体
    draw.rectangle([8, 6, 24, 26], fill=(255, 255, 255, 255), outline=(50, 50, 50, 255))
    # 文档折角
    draw.polygon([(18, 6), (24, 6), (24, 12), (18, 6)], fill=(200, 200, 200, 255))
    
    # 绘制文字线条
    for i in range(3):
        y = 10 + i * 4
        draw.line([(10, y), (22, y)], fill=(100, 100, 100, 255), width=1)
    
    # 保存为ICO文件
    try:
        # 创建不同尺寸的图标
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64)]
        images = []
        
        for s in sizes:
            resized = image.resize(s, Image.Resampling.LANCZOS)
            images.append(resized)
        
        # 保存为ICO文件
        images[0].save('icon.ico', format='ICO', sizes=[(s[0], s[1]) for s in sizes])
        print("✅ 图标文件创建成功: icon.ico")
        
    except Exception as e:
        print(f"❌ 创建图标文件失败: {e}")
        print("请手动创建icon.ico文件或使用默认图标")

if __name__ == "__main__":
    try:
        from PIL import Image, ImageDraw, ImageFont
        create_icon()
    except ImportError:
        print("❌ 需要安装Pillow库来创建图标")
        print("请运行: pip install Pillow")
        print("或者手动创建icon.ico文件") 