#!/usr/bin/env python3
"""
PowerPoint编辑MCP Server主程序
提供MCP服务器功能，调用tool.py中的PowerPointEditor类
"""

import asyncio
import json
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# 导入PowerPoint编辑器
from tool import PowerPointEditor

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 创建PowerPoint编辑器实例
ppt_editor = PowerPointEditor()

# 创建MCP Server
server = Server("powerpoint-editor")

@server.list_tools()
async def handle_list_tools():
    """列出所有可用的工具"""
    return [
            Tool(
                name="create_presentation",
                description="创建新的PowerPoint演示文稿",
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),
            Tool(
                name="open_presentation",
                description="打开现有的PowerPoint演示文稿",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "PowerPoint文件的路径"
                        }
                    },
                    "required": ["file_path"]
                }
            ),
            Tool(
                name="save_presentation",
                description="保存PowerPoint演示文稿",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "保存文件的路径（可选，如果不提供则保存到当前路径）"
                        }
                    },
                    "required": []
                }
            ),
            Tool(
                name="add_slide",
                description="添加新的幻灯片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "layout_index": {
                            "type": "integer",
                            "description": "幻灯片布局索引（0=标题幻灯片，1=标题和内容，默认为1）",
                            "default": 1
                        }
                    },
                    "required": []
                }
            ),
            Tool(
                name="add_text_box",
                description="在幻灯片中添加文本框",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "text": {
                            "type": "string",
                            "description": "要添加的文本内容"
                        },
                        "left": {
                            "type": "number",
                            "description": "文本框左边距（英寸）",
                            "default": 1
                        },
                        "top": {
                            "type": "number",
                            "description": "文本框上边距（英寸）",
                            "default": 1
                        },
                        "width": {
                            "type": "number",
                            "description": "文本框宽度（英寸）",
                            "default": 8
                        },
                        "height": {
                            "type": "number",
                            "description": "文本框高度（英寸）",
                            "default": 1
                        },
                        "font_size": {
                            "type": "integer",
                            "description": "字体大小",
                            "default": 18
                        },
                        "font_color": {
                            "type": "string",
                            "description": "字体颜色（十六进制，如000000）",
                            "default": "000000"
                        }
                    },
                    "required": ["slide_index", "text"]
                }
            ),
            Tool(
                name="add_title_slide",
                description="添加标题幻灯片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "title": {
                            "type": "string",
                            "description": "幻灯片标题"
                        },
                        "subtitle": {
                            "type": "string",
                            "description": "幻灯片副标题（可选）",
                            "default": ""
                        }
                    },
                    "required": ["title"]
                }
            ),
            Tool(
                name="add_bullet_points",
                description="添加带项目符号的内容幻灯片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "title": {
                            "type": "string",
                            "description": "幻灯片标题"
                        },
                        "bullet_points": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "项目符号列表"
                        }
                    },
                    "required": ["slide_index", "title", "bullet_points"]
                }
            ),
            Tool(
                name="add_image",
                description="在幻灯片中添加图片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "image_path": {
                            "type": "string",
                            "description": "图片文件路径"
                        },
                        "left": {
                            "type": "number",
                            "description": "图片左边距（英寸）",
                            "default": 1
                        },
                        "top": {
                            "type": "number",
                            "description": "图片上边距（英寸）",
                            "default": 2
                        },
                        "width": {
                            "type": "number",
                            "description": "图片宽度（英寸，可选）"
                        },
                        "height": {
                            "type": "number",
                            "description": "图片高度（英寸，可选）"
                        }
                    },
                    "required": ["slide_index", "image_path"]
                }
            ),
            Tool(
                name="add_shape",
                description="在幻灯片中添加形状",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "shape_type": {
                            "type": "string",
                            "description": "形状类型（rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow）"
                        },
                        "left": {
                            "type": "number",
                            "description": "形状左边距（英寸）",
                            "default": 1
                        },
                        "top": {
                            "type": "number",
                            "description": "形状上边距（英寸）",
                            "default": 1
                        },
                        "width": {
                            "type": "number",
                            "description": "形状宽度（英寸）",
                            "default": 2
                        },
                        "height": {
                            "type": "number",
                            "description": "形状高度（英寸）",
                            "default": 1
                        },
                        "fill_color": {
                            "type": "string",
                            "description": "填充颜色（十六进制，如0066CC）",
                            "default": "0066CC"
                        }
                    },
                    "required": ["slide_index", "shape_type"]
                }
            ),
            Tool(
                name="get_presentation_info",
                description="获取当前演示文稿的信息",
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),
            Tool(
                name="delete_slide",
                description="删除指定的幻灯片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "要删除的幻灯片索引（从0开始）"
                        }
                    },
                    "required": ["slide_index"]
                }
            ),
            Tool(
                name="duplicate_slide",
                description="复制指定的幻灯片",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "要复制的幻灯片索引（从0开始）"
                        }
                    },
                    "required": ["slide_index"]
                }
            ),
            Tool(
                name="move_slide",
                description="移动幻灯片位置",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "from_index": {
                            "type": "integer",
                            "description": "源位置索引（从0开始）"
                        },
                        "to_index": {
                            "type": "integer",
                            "description": "目标位置索引（从0开始）"
                        }
                    },
                    "required": ["from_index", "to_index"]
                }
            ),
            Tool(
                name="add_table",
                description="在幻灯片中添加表格",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "rows": {
                            "type": "integer",
                            "description": "表格行数"
                        },
                        "cols": {
                            "type": "integer",
                            "description": "表格列数"
                        },
                        "left": {
                            "type": "number",
                            "description": "表格左边距（英寸）",
                            "default": 1
                        },
                        "top": {
                            "type": "number",
                            "description": "表格上边距（英寸）",
                            "default": 2
                        },
                        "width": {
                            "type": "number",
                            "description": "表格宽度（英寸）",
                            "default": 8
                        },
                        "height": {
                            "type": "number",
                            "description": "表格高度（英寸）",
                            "default": 4
                        }
                    },
                    "required": ["slide_index", "rows", "cols"]
                }
            ),
            Tool(
                name="set_table_cell_text",
                description="设置表格单元格文本",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "table_index": {
                            "type": "integer",
                            "description": "表格索引（从0开始）"
                        },
                        "row": {
                            "type": "integer",
                            "description": "行索引（从0开始）"
                        },
                        "col": {
                            "type": "integer",
                            "description": "列索引（从0开始）"
                        },
                        "text": {
                            "type": "string",
                            "description": "要设置的文本内容"
                        }
                    },
                    "required": ["slide_index", "table_index", "row", "col", "text"]
                }
            ),
            Tool(
                name="set_slide_background_color",
                description="设置幻灯片背景颜色",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "color": {
                            "type": "string",
                            "description": "背景颜色（十六进制，如FF0000）"
                        }
                    },
                    "required": ["slide_index", "color"]
                }
            ),
            Tool(
                name="add_hyperlink",
                description="为形状添加超链接",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "shape_index": {
                            "type": "integer",
                            "description": "形状索引（从0开始）"
                        },
                        "url": {
                            "type": "string",
                            "description": "超链接URL"
                        },
                        "display_text": {
                            "type": "string",
                            "description": "显示文本（可选）"
                        }
                    },
                    "required": ["slide_index", "shape_index", "url"]
                }
            ),
            Tool(
                name="set_text_formatting",
                description="设置文本格式",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        },
                        "shape_index": {
                            "type": "integer",
                            "description": "形状索引（从0开始）"
                        },
                        "font_name": {
                            "type": "string",
                            "description": "字体名称（可选）"
                        },
                        "font_size": {
                            "type": "integer",
                            "description": "字体大小（可选）"
                        },
                        "font_color": {
                            "type": "string",
                            "description": "字体颜色（十六进制，可选）"
                        },
                        "bold": {
                            "type": "boolean",
                            "description": "是否加粗（可选）"
                        },
                        "italic": {
                            "type": "boolean",
                            "description": "是否斜体（可选）"
                        },
                        "underline": {
                            "type": "boolean",
                            "description": "是否下划线（可选）"
                        }
                    },
                    "required": ["slide_index", "shape_index"]
                }
            ),
            Tool(
                name="get_slide_shapes_info",
                description="获取幻灯片中所有形状的信息",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "slide_index": {
                            "type": "integer",
                            "description": "幻灯片索引（从0开始）"
                        }
                    },
                    "required": ["slide_index"]
                }
            )
        ]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict):
    """处理工具调用"""
    try:
        if name == "create_presentation":
            result = ppt_editor.create_presentation()

        elif name == "open_presentation":
            file_path = arguments.get("file_path", "")
            if not file_path:
                result = {"success": False, "error": "缺少必需参数: file_path"}
            else:
                result = ppt_editor.open_presentation(file_path)

        elif name == "save_presentation":
            file_path = arguments.get("file_path")
            result = ppt_editor.save_presentation(file_path)

        elif name == "add_slide":
            layout_index = arguments.get("layout_index", 1)
            result = ppt_editor.add_slide(layout_index)

        elif name == "add_text_box":
            slide_index = arguments.get("slide_index")
            text = arguments.get("text")
            if slide_index is None or text is None:
                result = {"success": False, "error": "缺少必需参数: slide_index 或 text"}
            else:
                left = arguments.get("left", 1)
                top = arguments.get("top", 1)
                width = arguments.get("width", 8)
                height = arguments.get("height", 1)
                font_size = arguments.get("font_size", 18)
                font_color = arguments.get("font_color", "000000")
                result = ppt_editor.add_text_box(slide_index, text, left, top, width, height, font_size, font_color)

        elif name == "add_title_slide":
            title = arguments.get("title")
            if not title:
                result = {"success": False, "error": "缺少必需参数: title"}
            else:
                subtitle = arguments.get("subtitle", "")
                result = ppt_editor.add_title_slide(title, subtitle)

        elif name == "add_bullet_points":
            slide_index = arguments.get("slide_index")
            title = arguments.get("title")
            bullet_points = arguments.get("bullet_points")
            if slide_index is None or not title or not bullet_points:
                result = {"success": False, "error": "缺少必需参数: slide_index, title 或 bullet_points"}
            else:
                result = ppt_editor.add_bullet_points(slide_index, title, bullet_points)

        elif name == "add_image":
            slide_index = arguments.get("slide_index")
            image_path = arguments.get("image_path")
            if slide_index is None or not image_path:
                result = {"success": False, "error": "缺少必需参数: slide_index 或 image_path"}
            else:
                left = arguments.get("left", 1)
                top = arguments.get("top", 2)
                width = arguments.get("width")
                height = arguments.get("height")
                result = ppt_editor.add_image(slide_index, image_path, left, top, width, height)

        elif name == "add_shape":
            slide_index = arguments.get("slide_index")
            shape_type = arguments.get("shape_type")
            if slide_index is None or not shape_type:
                result = {"success": False, "error": "缺少必需参数: slide_index 或 shape_type"}
            else:
                left = arguments.get("left", 1)
                top = arguments.get("top", 1)
                width = arguments.get("width", 2)
                height = arguments.get("height", 1)
                fill_color = arguments.get("fill_color", "0066CC")
                result = ppt_editor.add_shape(slide_index, shape_type, left, top, width, height, fill_color)

        elif name == "get_presentation_info":
            result = ppt_editor.get_presentation_info()

        elif name == "delete_slide":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {"success": False, "error": "缺少必需参数: slide_index"}
            else:
                result = ppt_editor.delete_slide(slide_index)

        elif name == "duplicate_slide":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {"success": False, "error": "缺少必需参数: slide_index"}
            else:
                result = ppt_editor.duplicate_slide(slide_index)

        elif name == "move_slide":
            from_index = arguments.get("from_index")
            to_index = arguments.get("to_index")
            if from_index is None or to_index is None:
                result = {"success": False, "error": "缺少必需参数: from_index 或 to_index"}
            else:
                result = ppt_editor.move_slide(from_index, to_index)

        elif name == "add_table":
            slide_index = arguments.get("slide_index")
            rows = arguments.get("rows")
            cols = arguments.get("cols")
            if slide_index is None or rows is None or cols is None:
                result = {"success": False, "error": "缺少必需参数: slide_index, rows 或 cols"}
            else:
                left = arguments.get("left", 1)
                top = arguments.get("top", 2)
                width = arguments.get("width", 8)
                height = arguments.get("height", 4)
                result = ppt_editor.add_table(slide_index, rows, cols, left, top, width, height)

        elif name == "set_table_cell_text":
            slide_index = arguments.get("slide_index")
            table_index = arguments.get("table_index")
            row = arguments.get("row")
            col = arguments.get("col")
            text = arguments.get("text")
            # 类型检查验证
            required_params = {
                'slide_index': slide_index,
                'table_index': table_index,
                'row': row,
                'col': col,
                'text': text
            }
            
            # 检查None值
            if any(v is None for v in required_params.values()):
                missing = [k for k, v in required_params.items() if v is None]
                result = {"success": False, "error": f"缺少必需参数: {', '.join(missing)}"}
            else:
                try:
                    # 类型断言
                    assert isinstance(slide_index, int), "slide_index必须是整数"
                    assert isinstance(table_index, int), "table_index必须是整数"
                    assert isinstance(row, int), "row必须是整数"
                    assert isinstance(col, int), "col必须是整数"
                    assert isinstance(text, str), "text必须是字符串"
                    
                    result = ppt_editor.set_table_cell_text(
                        slide_index=slide_index,
                        table_index=table_index,
                        row=row,
                        col=col,
                        text=text
                    )
                except AssertionError as e:
                    result = {"success": False, "error": f"参数验证失败: {str(e)}"}
                except Exception as e:
                    result = {"success": False, "error": str(e)}

        elif name == "set_slide_background_color":
            slide_index = arguments.get("slide_index")
            color = arguments.get("color")
            if slide_index is None or not color:
                result = {"success": False, "error": "缺少必需参数: slide_index 或 color"}
            else:
                result = ppt_editor.set_slide_background_color(slide_index, color)

        elif name == "add_hyperlink":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            url = arguments.get("url")
            if slide_index is None or shape_index is None or not url:
                result = {"success": False, "error": "缺少必需参数: slide_index, shape_index 或 url"}
            else:
                display_text = arguments.get("display_text")
                result = ppt_editor.add_hyperlink(slide_index, shape_index, url, display_text)

        elif name == "set_text_formatting":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            if slide_index is None or shape_index is None:
                result = {"success": False, "error": "缺少必需参数: slide_index 或 shape_index"}
            else:
                font_name = arguments.get("font_name")
                font_size = arguments.get("font_size")
                font_color = arguments.get("font_color")
                bold = arguments.get("bold")
                italic = arguments.get("italic")
                underline = arguments.get("underline")
                result = ppt_editor.set_text_formatting(slide_index, shape_index, font_name, font_size, font_color, bold, italic, underline)

        elif name == "get_slide_shapes_info":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {"success": False, "error": "缺少必需参数: slide_index"}
            else:
                result = ppt_editor.get_slide_shapes_info(slide_index)

        else:
            result = {"success": False, "error": f"未知的工具: {name}"}

        # 返回结果
        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

    except Exception as e:
        logger.error(f"工具调用错误: {e}")
        error_result = {"success": False, "error": str(e)}
        return [TextContent(type="text", text=json.dumps(error_result, ensure_ascii=False, indent=2))]


async def main():
    """主函数"""
    # 使用stdio运行服务器
    # 标准MCP服务器运行方式
    import asyncio
    from contextlib import AsyncExitStack
    
    async with AsyncExitStack() as stack:
        streams = await stack.enter_async_context(stdio_server())
        read_stream, write_stream = streams
        
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
