#!/usr/bin/env python3
"""
PowerPoint编辑器工具类
提供基础的PPT编辑功能，包括添加文本、图片、形状等
"""

import logging
from typing import Any, Dict, List, Optional, TYPE_CHECKING
from pathlib import Path

# 导入PowerPoint相关库
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
except ImportError:
    raise ImportError("请安装python-pptx库: pip install python-pptx")

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PowerPointEditor:
    """PowerPoint编辑器类"""
    
    def __init__(self):
        self.current_presentation: Optional["PresentationType"] = None
        self.current_file_path: Optional[str] = None

    def create_presentation(self) -> Dict[str, Any]:
        """创建新的演示文稿"""
        try:
            self.current_presentation = Presentation()
            self.current_file_path = None
            return {
                "success": True,
                "message": "成功创建新的演示文稿",
                "slides_count": len(self.current_presentation.slides)
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_presentation(self, file_path: str) -> Dict[str, Any]:
        """打开现有的演示文稿"""
        try:
            if not Path(file_path).exists():
                return {"success": False, "error": f"文件不存在: {file_path}"}

            self.current_presentation = Presentation(file_path)
            self.current_file_path = file_path

            return {
                "success": True,
                "message": f"成功打开演示文稿: {file_path}",
                "slides_count": len(self.current_presentation.slides),
                "file_path": file_path
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_presentation(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """保存演示文稿"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            save_path = file_path or self.current_file_path
            if not save_path:
                return {"success": False, "error": "请指定保存路径"}

            self.current_presentation.save(save_path)
            self.current_file_path = save_path

            return {
                "success": True,
                "message": f"成功保存演示文稿: {save_path}",
                "file_path": save_path
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_slide(self, layout_index: int = 1) -> Dict[str, Any]:
        """添加新幻灯片"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            # 获取幻灯片布局
            slide_layouts = self.current_presentation.slide_layouts
            if layout_index >= len(slide_layouts):
                layout_index = 1  # 默认使用标题和内容布局

            layout = slide_layouts[layout_index]
            slide = self.current_presentation.slides.add_slide(layout)

            return {
                "success": True,
                "message": f"成功添加新幻灯片",
                "slide_index": len(self.current_presentation.slides) - 1,
                "total_slides": len(self.current_presentation.slides)
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_text_box(self, slide_index: int, text: str, left: float = 1,
                     top: float = 1, width: float = 8, height: float = 1,
                     font_size: int = 18, font_color: str = "000000") -> Dict[str, Any]:
        """在指定幻灯片添加文本框"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            slides = self.current_presentation.slides
            if slide_index >= len(slides):
                return {"success": False, "error": f"幻灯片索引超出范围: {slide_index}"}

            slide = slides[slide_index]

            # 添加文本框
            left_inches = Inches(left)
            top_inches = Inches(top)
            width_inches = Inches(width)
            height_inches = Inches(height)

            textbox = slide.shapes.add_textbox(left_inches, top_inches, width_inches, height_inches)
            text_frame = textbox.text_frame
            text_frame.text = text

            # 设置字体样式
            paragraph = text_frame.paragraphs[0]
            font = paragraph.font
            font.size = Pt(font_size)

            # 设置字体颜色
            try:
                rgb_color = RGBColor.from_string(font_color)
                font.color.rgb = rgb_color
            except:
                pass  # 如果颜色格式不正确，使用默认颜色

            return {
                "success": True,
                "message": f"成功在幻灯片 {slide_index} 添加文本框",
                "text": text,
                "position": {"left": left, "top": top, "width": width, "height": height}
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_title_slide(self, title: str, subtitle: str = "") -> Dict[str, Any]:
        """添加标题幻灯片"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            # 使用标题幻灯片布局
            title_slide_layout = self.current_presentation.slide_layouts[0]
            slide = self.current_presentation.slides.add_slide(title_slide_layout)            # 设置标题
            title_shape = slide.shapes.title
            if title_shape:
                title_shape.text = title            # 设置副标题
            if subtitle and len(slide.placeholders) > 1:
                try:
                    subtitle_shape = slide.placeholders[1]
                    subtitle_shape.text_frame.text = subtitle  # type: ignore
                except (AttributeError, TypeError):
                    pass  # 如果无法设置副标题，忽略错误

            return {
                "success": True,
                "message": "成功添加标题幻灯片",
                "slide_index": len(self.current_presentation.slides) - 1,
                "title": title,
                "subtitle": subtitle
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_bullet_points(self, slide_index: int, title: str, bullet_points: List[str]) -> Dict[str, Any]:
        """添加带项目符号的内容幻灯片"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            slides = self.current_presentation.slides
            if slide_index >= len(slides):
                return {"success": False, "error": f"幻灯片索引超出范围: {slide_index}"}

            slide = slides[slide_index]

            # 设置标题
            if slide.shapes.title:
                slide.shapes.title.text = title            # 查找内容占位符
            content_placeholder = None
            for shape in slide.placeholders:
                if shape.placeholder_format.idx == 1:  # 内容占位符通常是索引1
                    content_placeholder = shape
                    break
                    
            if content_placeholder:
                try:
                    text_frame = content_placeholder.text_frame  # type: ignore
                    text_frame.clear()  # 清除现有内容

                    for i, point in enumerate(bullet_points):
                        if i == 0:
                            p = text_frame.paragraphs[0]
                        else:
                            p = text_frame.add_paragraph()
                        p.text = point
                        p.level = 0  # 设置为第一级项目符号
                except (AttributeError, TypeError):
                    pass  # 如果无法设置内容，忽略错误

            return {
                "success": True,
                "message": f"成功在幻灯片 {slide_index} 添加项目符号内容",
                "title": title,
                "bullet_points": bullet_points
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_image(self, slide_index: int, image_path: str, left: float = 1,
                  top: float = 2, width: Optional[float] = None, height: Optional[float] = None) -> Dict[str, Any]:
        """在幻灯片中添加图片"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            if not Path(image_path).exists():
                return {"success": False, "error": f"图片文件不存在: {image_path}"}

            slides = self.current_presentation.slides
            if slide_index >= len(slides):
                return {"success": False, "error": f"幻灯片索引超出范围: {slide_index}"}

            slide = slides[slide_index]

            # 添加图片
            left_inches = Inches(left)
            top_inches = Inches(top)

            if width and height:
                width_inches = Inches(width)
                height_inches = Inches(height)
                pic = slide.shapes.add_picture(image_path, left_inches, top_inches, width_inches, height_inches)
            else:
                pic = slide.shapes.add_picture(image_path, left_inches, top_inches)

            return {
                "success": True,
                "message": f"成功在幻灯片 {slide_index} 添加图片",
                "image_path": image_path,
                "position": {"left": left, "top": top, "width": width, "height": height}
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_shape(self, slide_index: int, shape_type: str, left: float = 1,
                  top: float = 1, width: float = 2, height: float = 1,
                  fill_color: str = "0066CC") -> Dict[str, Any]:
        """添加形状"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            slides = self.current_presentation.slides
            if slide_index >= len(slides):
                return {"success": False, "error": f"幻灯片索引超出范围: {slide_index}"}

            slide = slides[slide_index]

            # 形状类型映射
            shape_map = {
                "rectangle": MSO_SHAPE.RECTANGLE,
                "oval": MSO_SHAPE.OVAL,
                "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
                "diamond": MSO_SHAPE.DIAMOND,
                "pentagon": MSO_SHAPE.REGULAR_PENTAGON,
                "hexagon": MSO_SHAPE.HEXAGON,
                "star": MSO_SHAPE.STAR_5_POINT,
                "arrow": MSO_SHAPE.BLOCK_ARC
            }

            if shape_type.lower() not in shape_map:
                return {"success": False, "error": f"不支持的形状类型: {shape_type}"}

            # 添加形状
            left_inches = Inches(left)
            top_inches = Inches(top)
            width_inches = Inches(width)
            height_inches = Inches(height)

            shape = slide.shapes.add_shape(
                shape_map[shape_type.lower()],
                left_inches, top_inches, width_inches, height_inches
            )

            # 设置填充颜色
            try:
                rgb_color = RGBColor.from_string(fill_color)
                shape.fill.solid()
                shape.fill.fore_color.rgb = rgb_color
            except:
                pass  # 如果颜色格式不正确，使用默认颜色

            return {
                "success": True,
                "message": f"成功在幻灯片 {slide_index} 添加 {shape_type} 形状",
                "shape_type": shape_type,
                "position": {"left": left, "top": top, "width": width, "height": height},
                "fill_color": fill_color
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_presentation_info(self) -> Dict[str, Any]:
        """获取当前演示文稿信息"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            slides_info = []
            for i, slide in enumerate(self.current_presentation.slides):
                slide_info = {
                    "index": i,
                    "shapes_count": len(slide.shapes),
                    "has_title": bool(slide.shapes.title and slide.shapes.title.text),
                    "title": slide.shapes.title.text if slide.shapes.title else ""                }
                slides_info.append(slide_info)
            
            return {
                "success": True,
                "file_path": self.current_file_path,
                "slides_count": len(self.current_presentation.slides),
                "slides": slides_info
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_slide(self, slide_index: int) -> Dict[str, Any]:
        """删除指定幻灯片"""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "没有打开的演示文稿"}

            slides = self.current_presentation.slides
            if slide_index >= len(slides):
                return {"success": False, "error": f"幻灯片索引超出范围: {slide_index}"}            # 获取要删除的幻灯片
            slide_to_remove = slides[slide_index]
            
            # 删除幻灯片的正确方法
            # 直接从slides集合中删除
            xml_slides = self.current_presentation.part._element.sldIdLst
            xml_slides.remove(xml_slides[slide_index])
            slides._sldIdLst.remove(slides._sldIdLst[slide_index])

            return {
                "success": True,
                "message": f"成功删除幻灯片 {slide_index}",
                "remaining_slides": len(self.current_presentation.slides)
            }
        except Exception as e:
            return {"success": False, "error": str(e)}