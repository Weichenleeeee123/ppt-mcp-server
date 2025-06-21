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

    # 演示新功能
    print("\n=== 演示新功能 ===")

    # 10. 复制幻灯片
    print("\n10. 复制幻灯片...")
    result = editor.duplicate_slide(0)  # 复制第一张幻灯片
    print(f"结果: {result}")

    # 11. 添加表格
    print("\n11. 添加表格...")
    result = editor.add_table(slide_index=3, rows=3, cols=4, left=1, top=2, width=8, height=3)
    print(f"结果: {result}")

    # 12. 设置表格单元格文本
    print("\n12. 设置表格单元格文本...")
    result = editor.set_table_cell_text(slide_index=3, table_index=0, row=0, col=0, text="标题1")
    print(f"结果: {result}")

    result = editor.set_table_cell_text(slide_index=3, table_index=0, row=0, col=1, text="标题2")
    print(f"结果: {result}")

    result = editor.set_table_cell_text(slide_index=3, table_index=0, row=1, col=0, text="数据1")
    print(f"结果: {result}")

    result = editor.set_table_cell_text(slide_index=3, table_index=0, row=1, col=1, text="数据2")
    print(f"结果: {result}")

    # 13. 设置幻灯片背景颜色
    print("\n13. 设置幻灯片背景颜色...")
    result = editor.set_slide_background_color(slide_index=3, color="E6F3FF")  # 浅蓝色
    print(f"结果: {result}")

    # 14. 获取幻灯片形状信息
    print("\n14. 获取幻灯片形状信息...")
    result = editor.get_slide_shapes_info(slide_index=2)
    print(f"结果: {result}")

    # 15. 设置文本格式
    print("\n15. 设置文本格式...")
    result = editor.set_text_formatting(
        slide_index=2,
        shape_index=2,  # 文本框
        font_name="Arial",
        font_size=24,
        font_color="FF0000",  # 红色
        bold=True,
        italic=True
    )
    print(f"结果: {result}")

    # 16. 移动幻灯片
    print("\n16. 移动幻灯片...")
    result = editor.move_slide(from_index=3, to_index=1)  # 将第4张幻灯片移动到第2个位置
    print(f"结果: {result}")

    # 17. 最终保存
    print("\n17. 保存更新后的演示文稿...")
    result = editor.save_presentation("enhanced_presentation.pptx")
    print(f"结果: {result}")

    print("\n增强功能演示完成！已创建 enhanced_presentation.pptx 文件")

if __name__ == "__main__":
    main()
