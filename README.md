# PowerPoint编辑MCP Server

这是一个基于MCP (Model Context Protocol) 的PowerPoint编辑服务器，提供了创建和编辑PowerPoint演示文稿的基础功能。

## 项目结构

- `main.py` - MCP服务器主程序，处理MCP协议通信
- `tool.py` - PowerPoint编辑器工具类，包含所有PPT编辑功能
- `example.py` - 使用示例
- `requirements.txt` - 项目依赖
- `mcp_config.json` - MCP客户端配置文件

## 功能特性

- 创建新的PowerPoint演示文稿
- 打开现有的PowerPoint文件
- 保存演示文稿
- 添加和删除幻灯片
- 添加文本框和文本内容
- 添加标题幻灯片
- 添加带项目符号的内容
- 插入图片
- 添加各种形状（矩形、椭圆、三角形等）
- 获取演示文稿信息

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 作为MCP Server运行

```bash
python main.py
```

### 直接使用PowerPointEditor类

```python
from tool import PowerPointEditor

# 创建编辑器实例
editor = PowerPointEditor()

# 创建新演示文稿
editor.create_presentation()

# 添加标题幻灯片
editor.add_title_slide("我的演示文稿", "副标题")

# 保存文件
editor.save_presentation("my_presentation.pptx")
```

### 运行示例

```bash
python example.py
```

## 可用工具

### 1. create_presentation
创建新的PowerPoint演示文稿

### 2. open_presentation
打开现有的PowerPoint文件
- `file_path`: 文件路径

### 3. save_presentation
保存演示文稿
- `file_path`: 保存路径（可选）

### 4. add_slide
添加新幻灯片
- `layout_index`: 布局索引（0=标题幻灯片，1=标题和内容）

### 5. add_text_box
添加文本框
- `slide_index`: 幻灯片索引
- `text`: 文本内容
- `left`, `top`, `width`, `height`: 位置和大小（英寸）
- `font_size`: 字体大小
- `font_color`: 字体颜色（十六进制）

### 6. add_title_slide
添加标题幻灯片
- `title`: 标题
- `subtitle`: 副标题（可选）

### 7. add_bullet_points
添加项目符号内容
- `slide_index`: 幻灯片索引
- `title`: 标题
- `bullet_points`: 项目符号列表

### 8. add_image
添加图片
- `slide_index`: 幻灯片索引
- `image_path`: 图片路径
- `left`, `top`: 位置（英寸）
- `width`, `height`: 大小（英寸，可选）

### 9. add_shape
添加形状
- `slide_index`: 幻灯片索引
- `shape_type`: 形状类型（rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow）
- `left`, `top`, `width`, `height`: 位置和大小（英寸）
- `fill_color`: 填充颜色（十六进制）

### 10. get_presentation_info
获取演示文稿信息

### 11. delete_slide
删除幻灯片
- `slide_index`: 要删除的幻灯片索引

## 注意事项

1. 确保安装了所有必需的依赖包
2. 图片文件路径必须存在且可访问
3. 幻灯片索引从0开始
4. 颜色使用十六进制格式（如：000000表示黑色，FF0000表示红色）
5. 位置和大小单位为英寸

## 错误处理

所有操作都包含错误处理，返回格式为：
```json
{
  "success": true/false,
  "message": "操作结果消息",
  "error": "错误信息（如果有）"
}
```

## 许可证

MIT License
