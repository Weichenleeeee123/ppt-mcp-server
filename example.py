#!/usr/bin/env python3
"""
PowerPoint编辑器使用示例
演示如何使用PowerPointEditor类创建和编辑PPT
"""

from tool import PowerPointEditor

def main():
    """示例主函数"""
    # 创建PowerPoint编辑器实例
    editor = PowerPointEditor()
    
    # 1. 创建新的演示文稿
    print("1. 创建新演示文稿...")
    result = editor.create_presentation()
    print(f"结果: {result}")
    
    # 2. 添加标题幻灯片
    print("\n2. 添加标题幻灯片...")
    result = editor.add_title_slide("我的演示文稿", "使用Python创建")
    print(f"结果: {result}")
    
    # 3. 添加内容幻灯片
    print("\n3. 添加内容幻灯片...")
    result = editor.add_slide(1)  # 使用标题和内容布局
    print(f"结果: {result}")
    
    # 4. 在第二张幻灯片添加项目符号内容
    print("\n4. 添加项目符号内容...")
    bullet_points = [
        "第一个要点",
        "第二个要点", 
        "第三个要点"
    ]
    result = editor.add_bullet_points(1, "主要内容", bullet_points)
    print(f"结果: {result}")
    
    # 5. 添加另一张幻灯片
    print("\n5. 添加另一张幻灯片...")
    result = editor.add_slide(1)
    print(f"结果: {result}")
    
    # 6. 在第三张幻灯片添加文本框
    print("\n6. 添加文本框...")
    result = editor.add_text_box(
        slide_index=2,
        text="这是一个自定义文本框",
        left=2,
        top=3,
        width=6,
        height=2,
        font_size=24,
        font_color="FF0000"  # 红色
    )
    print(f"结果: {result}")
    
    # 7. 添加形状
    print("\n7. 添加形状...")
    result = editor.add_shape(
        slide_index=2,
        shape_type="rectangle",
        left=1,
        top=1,
        width=3,
        height=1.5,
        fill_color="00FF00"  # 绿色
    )
    print(f"结果: {result}")
    
    # 8. 获取演示文稿信息
    print("\n8. 获取演示文稿信息...")
    result = editor.get_presentation_info()
    print(f"结果: {result}")
    
    # 9. 保存演示文稿
    print("\n9. 保存演示文稿...")
    result = editor.save_presentation("example_presentation.pptx")
    print(f"结果: {result}")
    
    print("\n演示完成！已创建 example_presentation.pptx 文件")

if __name__ == "__main__":
    main()
