# PowerPoint编辑MCP Server

这是一个基于MCP (Model Context Protocol) 的PowerPoint编辑服务器，提供了创建和编辑PowerPoint演示文稿的完整功能，包括内容编辑、格式化和专业动画效果。

## ✨ 最新更新

- 🎬 **全新动画系统** - 添加了多种专业过渡动画效果
- 🚀 **一键专业化** - 快速让演示文稿变得专业
- 🎯 **智能推荐** - 优化工具描述，提高AI模型使用率
- 🛠️ **便利函数** - 简化复杂操作，提供直观的参数接口

## 项目结构

- `main.py` - MCP服务器主程序，处理MCP协议通信
- `tool.py` - PowerPoint编辑器工具类，包含所有PPT编辑功能
- `example.py` - 使用示例
- `test_transitions.py` - 过渡动画功能测试
- `transition_improvements_guide.md` - 动画功能改进指南
- `requirements.txt` - 项目依赖
- `mcp_config.json` - MCP客户端配置文件

## 功能特性

### 基础功能
- 创建新的PowerPoint演示文稿
- 打开现有的PowerPoint文件
- 保存演示文稿
- 获取演示文稿信息

### 幻灯片操作
- 添加新幻灯片（支持不同布局）
- 删除幻灯片
- 复制幻灯片
- 移动幻灯片位置
- 设置幻灯片背景颜色

### 内容编辑
- 添加文本框和文本内容
- 添加标题幻灯片
- 添加带项目符号的内容
- 插入图片
- 添加各种形状（矩形、椭圆、三角形等）
- 添加表格
- 设置表格单元格文本

### 格式化功能
- 设置文本格式（字体、大小、颜色、粗体、斜体、下划线）
- 为形状添加超链接
- 获取幻灯片中所有形状的详细信息

### 🎬 专业动画和过渡效果
- **一键专业化** - 快速为整个演示文稿添加专业过渡效果
- **多种动画风格** - 淡入淡出、推入、擦除、分割、缩放、百叶窗、溶解等8种效果
- **智能速度控制** - 快速、中等、慢速三档速度选择
- **自动播放支持** - 支持自动前进和点击前进
- **批量应用** - 一次性为所有幻灯片设置统一动画
- **便利函数** - 提供流畅过渡、动感效果等预设选项

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

### 测试动画功能

```bash
python test_transitions.py
```

## 🎬 动画功能快速开始

```python
from tool import PowerPointEditor

editor = PowerPointEditor()
editor.create_presentation()

# 添加几张幻灯片
editor.add_title_slide("欢迎", "我的演示文稿")
editor.add_title_slide("内容", "主要内容")
editor.add_title_slide("结束", "谢谢观看")

# 一键专业化 - 为所有幻灯片添加淡入淡出效果
editor.make_presentation_professional()

# 或者添加动感效果
# editor.add_dynamic_effects()

# 保存文件
editor.save_presentation("professional_presentation.pptx")
```

## 🛠️ 可用工具

### 🎬 动画和过渡工具（新增）

#### add_slide_animation
为单张幻灯片添加动画过渡效果，让演示更生动有趣
- `slide_index`: 幻灯片索引
- `animation_style`: 动画风格（fade, push, wipe, zoom, split, blinds, dissolve, none）
- `speed`: 动画速度（fast, medium, slow）
- `auto_advance`: 是否自动切换到下一张
- `auto_advance_seconds`: 自动切换延迟时间

#### make_presentation_dynamic
为整个演示文稿添加统一的动画效果，制作专业演示文稿的重要步骤
- `animation_style`: 统一的动画风格（默认fade）
- `speed`: 动画速度（默认medium）

#### make_professional_presentation ⭐
一键让演示文稿变得专业！自动为所有幻灯片添加优雅的淡入淡出过渡效果
- 无参数，一键操作

#### add_smooth_transitions
为演示文稿添加流畅的过渡动画，让幻灯片切换更加自然
- 无参数，预设流畅效果

#### add_dynamic_effects
为演示文稿添加动感的过渡效果，让演示更有活力
- 无参数，预设动感效果

#### get_animation_options
查看所有可用的幻灯片动画效果选项
- 无参数

### 📄 基础工具

#### 1. create_presentation
创建新的PowerPoint演示文稿

#### 2. open_presentation
打开现有的PowerPoint文件
- `file_path`: 文件路径

#### 3. save_presentation
保存演示文稿
- `file_path`: 保存路径（可选）

### 📝 内容编辑工具

#### 4. add_slide
添加新幻灯片
- `layout_index`: 布局索引（0=标题幻灯片，1=标题和内容）

#### 5. add_text_box
添加文本框
- `slide_index`: 幻灯片索引
- `text`: 文本内容
- `left`, `top`, `width`, `height`: 位置和大小（英寸）
- `font_size`: 字体大小
- `font_color`: 字体颜色（十六进制）

#### 6. add_title_slide
添加标题幻灯片
- `title`: 标题
- `subtitle`: 副标题（可选）

#### 7. add_bullet_points
添加项目符号内容
- `slide_index`: 幻灯片索引
- `title`: 标题
- `bullet_points`: 项目符号列表

#### 8. add_image
添加图片
- `slide_index`: 幻灯片索引
- `image_path`: 图片路径
- `left`, `top`: 位置（英寸）
- `width`, `height`: 大小（英寸，可选）

#### 9. add_shape
添加形状
- `slide_index`: 幻灯片索引
- `shape_type`: 形状类型（rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow）
- `left`, `top`, `width`, `height`: 位置和大小（英寸）
- `fill_color`: 填充颜色（十六进制）

#### 10. add_table
添加表格
- `slide_index`: 幻灯片索引
- `rows`: 表格行数
- `cols`: 表格列数
- `left`, `top`, `width`, `height`: 位置和大小（英寸）

#### 11. set_table_cell_text
设置表格单元格文本
- `slide_index`: 幻灯片索引
- `table_index`: 表格索引
- `row`: 行索引
- `col`: 列索引
- `text`: 文本内容

### 🎨 格式化和样式工具

#### 12. set_slide_background_color
设置幻灯片背景颜色
- `slide_index`: 幻灯片索引
- `color`: 背景颜色（十六进制）

#### 13. add_hyperlink
为形状添加超链接
- `slide_index`: 幻灯片索引
- `shape_index`: 形状索引
- `url`: 超链接URL
- `display_text`: 显示文本（可选）

#### 14. set_text_formatting
设置文本格式
- `slide_index`: 幻灯片索引
- `shape_index`: 形状索引
- `font_name`: 字体名称（可选）
- `font_size`: 字体大小（可选）
- `font_color`: 字体颜色（可选）
- `bold`: 是否加粗（可选）
- `italic`: 是否斜体（可选）
- `underline`: 是否下划线（可选）

### 🔧 管理工具

#### 15. get_presentation_info
获取演示文稿信息

#### 16. delete_slide
删除幻灯片
- `slide_index`: 要删除的幻灯片索引

#### 17. duplicate_slide
复制幻灯片
- `slide_index`: 要复制的幻灯片索引

#### 18. move_slide
移动幻灯片位置
- `from_index`: 源位置索引
- `to_index`: 目标位置索引

#### 19. get_slide_shapes_info
获取幻灯片中所有形状的信息
- `slide_index`: 幻灯片索引

### 🎬 传统动画工具（向后兼容）

#### 20. set_slide_transition
设置幻灯片过渡效果（推荐使用新的动画工具）
- `slide_index`: 幻灯片索引
- `transition_type`: 过渡类型（none, fade, push, wipe, split, zoom, blinds, dissolve）
- `duration`: 过渡持续时间（秒）
- `advance_on_click`: 是否点击前进
- `advance_after_time`: 自动前进时间（秒，可选）

#### 21. get_available_transitions
获取可用的过渡效果列表
- 无参数

## 💡 使用技巧

### 让AI更好地使用动画功能

为了让AI模型更主动地使用动画功能，可以在对话中使用这些关键词：

- **"让演示更专业"** → AI会调用 `make_professional_presentation`
- **"添加动画效果"** → AI会使用 `add_slide_animation` 或 `make_presentation_dynamic`
- **"让幻灯片切换更流畅"** → AI会调用 `add_smooth_transitions`
- **"让演示更有活力"** → AI会使用 `add_dynamic_effects`

### 推荐的工作流程

1. **创建内容** - 先添加所有幻灯片和内容
2. **一键专业化** - 使用 `make_professional_presentation()` 快速添加过渡效果
3. **个性化调整** - 根据需要为特定幻灯片设置不同的动画效果
4. **预览和保存** - 保存文件并在PowerPoint中预览效果

## ⚠️ 注意事项

1. 确保安装了所有必需的依赖包（特别是 `lxml` 用于动画功能）
2. 图片文件路径必须存在且可访问
3. 幻灯片索引从0开始
4. 颜色使用十六进制格式（如：000000表示黑色，FF0000表示红色）
5. 位置和大小单位为英寸
6. 动画效果需要在PowerPoint中打开文件才能看到完整效果

## 错误处理

所有操作都包含错误处理，返回格式为：
```json
{
  "success": true/false,
  "message": "操作结果消息",
  "error": "错误信息（如果有）"
}
```

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📄 许可证

MIT License
