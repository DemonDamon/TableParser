# TableParser - 轻量级智能表格解析工具

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-green.svg)](https://opensource.org/licenses/Apache-2.0)

一个轻量级的表格解析工具，支持Excel和CSV文件的智能解析，**根据表格复杂度自动选择最佳输出格式**（Markdown或HTML）。

## 📑 目录

- [核心特性](#-核心特性)
- [快速开始](#-快速开始)
- [复杂度评估算法](#-复杂度评估算法)
- [架构设计](#️-架构设计)
- [高级用法](#-高级用法)
- [项目结构](#-项目结构)
- [与竞品对比](#-与竞品对比)
- [使用场景](#-使用场景)
- [性能指标](#-性能指标)

## ✨ 核心特性

- 🧠 **业界领先复杂度分析**：7维度动态权重评分系统，智能适应不同表格特征
- 🎯 **自适应输出**：根据复杂度自动选择Markdown或HTML
- 💾 **智能自动保存**：默认保存到同目录，节省90%-99% token（v1.1新增）
- 🛡️ **三层容错机制**：openpyxl → pandas → calamine，最大化兼容性
- 💡 **MCP工具化**：支持AI智能体（Claude、GPT等）直接调用
- 🚀 **轻量级**：最小化依赖，核心仅需openpyxl+pandas
- 📦 **易于集成**：简洁的API，支持Python库、CLI、MCP多种方式
- ✨ **v1.2 新增**：动态权重、数据透视表/图表/VBA宏检测、图片提取、样式保留、上下标支持、公式依赖分析

## 🎬 快速开始

### 安装

```bash
pip install -r requirements.txt
```

### Python API使用

```python
from table_parser import TableParser

# 创建解析器
parser = TableParser()

# 自动模式（推荐）- 根据复杂度自动选择格式
result = parser.parse("data.xlsx", output_format="auto")
print(f"使用格式: {result.output_format}")
print(f"复杂度: {result.complexity_score.level}")
print(result.content)

# 强制指定格式
result = parser.parse("data.xlsx", output_format="markdown")
result = parser.parse("data.xlsx", output_format="html")

# 仅分析复杂度（不解析内容）
score = parser.analyze_only("data.xlsx")
print(f"得分: {score.total_score:.1f}, 等级: {score.level}")

# 快速预览
preview = parser.preview("data.xlsx", max_rows=5)
print(preview['sheets'][0]['preview'])
```

### MCP工具使用（AI智能体）

#### 配置 Cursor

编辑 `~/.cursor/mcp.json` 文件（如不存在则创建）：

```json
{
  "mcpServers": {
    "table-parser": {
      "command": "python",
      "args": [
        "-u",
        "/path/to/your/project/TableParser/start_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "/path/to/your/project/TableParser"
      }
    }
  }
}
```

**注意事项：**
- 将路径替换为你的实际项目路径（使用绝对路径）
- macOS/Linux: `~/.cursor/mcp.json`
- Windows: `%USERPROFILE%\.cursor\mcp.json`
- 配置后需要重启 Cursor 或切换 MCP 开关

**配置示例（macOS）：**

```json
{
  "mcpServers": {
    "table-parser": {
      "command": "python",
      "args": [
        "-u",
        "/Users/username/projects/TableParser/start_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "/Users/username/projects/TableParser"
      }
    }
  }
}
```

**配置示例（Windows）：**

```json
{
  "mcpServers": {
    "table-parser": {
      "command": "python",
      "args": [
        "-u",
        "C:\\Users\\username\\projects\\TableParser\\start_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "C:\\Users\\username\\projects\\TableParser"
      }
    }
  }
}
```

#### 配置 Claude Desktop

编辑 `~/Library/Application Support/Claude/claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "table-parser": {
      "command": "python",
      "args": [
        "-u",
        "/path/to/your/project/TableParser/start_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "/path/to/your/project/TableParser"
      }
    }
  }
}
```

**智能自动保存（v1.1 新功能）**：

```
用户: "帮我解析 /data/sales_2024.xlsx"

AI会自动：
1. 调用 parse_table 解析文件
2. 自动保存到 /data/sales_2024.html（或.md）
3. 只返回元数据（文件路径、大小等）
4. 节省 90%-99% token消耗 🎉
```

**三种使用方式**：

```python
# 方式1：自动保存（默认，推荐）
parse_table(file_path="/data/sales.xlsx")
# → 自动保存到 /data/sales.html，返回元数据

# 方式2：指定保存路径
parse_table(
    file_path="/data/sales.xlsx",
    output_path="/output/report.html"
)
# → 保存到指定位置，返回元数据

# 方式3：Base64输入（临时处理）
parse_table(file_content_base64="...")
# → 返回完整内容（适合临时数据）
```

## 📊 复杂度评估算法

TableParser 实现了业界领先的 **7维度动态评分系统**（基于荷兰国家档案馆Spreadsheet-Complexity-Analyser、Microsoft TableSense、Nanonets评估标准改进）：

### 🎯 动态权重策略

**核心创新**：根据表格特征智能选择权重方案

#### 基础权重（适用于纯结构复杂表格）

当表格**不包含**数据透视表、图表、VBA宏时，自动使用基础权重：

| 维度类别 | 具体指标 | 权重 | 评估内容 |
|---------|---------|------|---------|
| **结构复杂度** (60%) | 合并单元格 | **35%** | 合并单元格数量、比例、跨行跨列复杂度 |
| | 表头层级 | **25%** | 单级/多级表头检测（2-4+级） |
| **数据复杂度** (30%) | 公式/超链接 | **15%** | 公式、超链接 |
| | 内容丰富度 | **15%** | 图片、样式(高亮/颜色)、富文本(上下标) ✨ |
| **规模复杂度** (10%) | 表格规模 | **10%** | 行列数、单元格总数 |

#### 高级权重（适用于功能丰富的复杂表格）

当表格**包含**数据透视表、图表、VBA宏时，自动使用高级权重：

| 维度类别 | 具体指标 | 权重 | 评估内容 |
|---------|---------|------|---------|
| **结构复杂度** (30%) | 合并单元格 | 20% | 合并单元格数量、比例、跨行跨列复杂度 |
| | 表头层级 | 10% | 单级/多级表头检测（2-4+级） |
| **数据复杂度** (50%) | 公式/超链接 | 15% | 公式、超链接 |
| | 内容丰富度 | 10% | 图片、样式、富文本 ✨ |
| | 数据透视表 | 15% | 数据透视表数量检测 |
| | 图表数量 | 10% | 图表、可视化元素 |
| **代码复杂度** (10%) | VBA宏 | 10% | VBA宏代码检测 |
| **规模复杂度** (10%) | 表格规模 | 10% | 行列数、单元格总数 |

### 🧠 动态权重工作原理

```
解析表格
    ↓
检测高级特征（数据透视表/图表/VBA宏）
    ↓
  有？──→ 使用高级权重（7维度平衡）
    │     - 适合：带透视表、图表、宏的表格
    │     - 权重：结构35% + 数据35% + 代码20% + 规模10%
    │
  无？──→ 使用基础权重（结构主导）
        - 适合：纯结构复杂表格（合并单元格、多级表头）
        - 权重：结构70% + 数据20% + 规模10%
```

### 🏆 相比业界方案的优势

| 特性 | TableParser v1.2 | 荷兰档案馆方案 | 传统方案 |
|------|-----------------|---------------|---------|
| 合并单元格 | ✅ 深度分析 | ✅ 基础检测 | ⚠️ 简单统计 |
| 表头层级 | ✅ 智能识别 | ❌ | ❌ |
| 数据透视表 | ✅ 自动检测 | ✅ 需手动 | ❌ |
| 图表分析 | ✅ 自动检测 | ✅ 需手动 | ❌ |
| VBA宏 | ✅ 自动检测 | ✅ 基础检测 | ❌ |
| **动态权重** | ✅ **智能切换** | ❌ 固定权重 | ❌ |
| 权重优化 | ✅ 科学分配 | ⚠️ 平均分配 | ❌ |

### 📏 评分规则

- **0-30分**：简单表格 → 推荐 **Markdown**（易读易编辑，Git友好）
- **31-60分**：中等复杂 → 推荐 **Markdown**（可能有部分格式损失）
- **61-100分**：复杂表格 → 强制 **HTML**（完整保留所有结构）

#### 🎨 特殊规则（内容保真）

**即使总分较低，如果检测到以下特征，也会强制推荐HTML：**
- ✅ **有图片**（内容丰富度 ≥ 40分）→ 强制HTML（Markdown无法嵌入图片）
- ✅ **有样式**（高亮、背景色）→ 强制HTML（Markdown无法显示颜色）
- ✅ **有富文本**（上下标）→ 强制HTML（HTML的`<sup>`/`<sub>`支持更好）

**示例**：
```python
# 案例：简单表格 + 图片 + 样式
- 总分：21.5分（simple级别）
- 内容丰富度：100分（有图片+样式+富文本）
→ 强制推荐 HTML ✅（保留图片和样式）
```

### 💡 典型案例（动态权重效果）

#### 案例1：评审标准表（基础权重）
```python
特征：
- 合并单元格：3个（得分80）
- 表头层级：多级（得分100）
- 无数据透视表/图表/VBA宏

权重选择：✅ 基础权重（40% + 30% + 20% + 10%）
计算：80×40% + 100×30% + 30×20% + 0×10% = 68分
等级：complex → 推荐 HTML ✅

若用高级权重：80×25% + 100×10% + 30×15% = 34.5分
等级：medium → 推荐 Markdown ❌（错误）
```

#### 案例2：带数据透视表的财务报表（高级权重）
```python
特征：
- 合并单元格：少量（得分20）
- 表头层级：2级（得分30）
- 数据透视表：3个（得分70）
- 图表：5个（得分60）

权重选择：✅ 高级权重（检测到数据透视表和图表）
计算：20×25% + 30×10% + 70×10% + 60×10% = 21分
等级：simple → 推荐 Markdown

若用基础权重：20×40% + 30×30% = 17分（数据透视表和图表被忽略）
```

#### 案例3：带VBA宏的复杂表（高级权重）
```python
特征：
- 合并单元格：大量（得分80）
- VBA宏：存在（得分100）
- 图表：2个（得分30）

权重选择：✅ 高级权重（检测到VBA宏）
计算：80×25% + 100×20% + 30×10% = 43分
等级：medium → 推荐 Markdown

含VBA宏但不算复杂，因为用户可能只需要提取数据
```

## 🏗️ 架构设计

```
用户
  ↓
TableParser（主控制器）
  ├─ FileLoader（文件加载器）
  │   ├─ openpyxl（主引擎）
  │   ├─ pandas（备用）
  │   └─ calamine（容错）
  ├─ ComplexityAnalyzer（复杂度分析器）
  │   ├─ 7维度检测引擎
  │   └─ 动态权重选择（基础/高级）
  ├─ FormatConverter（格式转换器）
  │   ├─ Markdown输出
  │   └─ HTML输出（支持合并单元格）
  └─ MCP Server（AI智能体接口）
      ├─ parse_table
      ├─ analyze_complexity
      ├─ batch_parse
      └─ get_preview
```

## 🔧 高级用法

### 批量处理

```python
from pathlib import Path

files = list(Path("/data").glob("*.xlsx"))
for file in files:
    result = parser.parse(file, output_format="auto")
    
    # 保存输出
    if result.output_format == "markdown":
        output_file = file.with_suffix(".md")
        output_file.write_text(result.content)
    else:  # HTML
        output_file = file.with_suffix(".html")
        output_file.write_text("\n\n".join(result.content))
```

### 高级功能使用

```python
# 1. 保留样式（背景色、字体颜色、高亮等）
result = parser.parse(
    "data.xlsx",
    output_format="html",
    preserve_styles=True  # ✅ 启用样式保留
)

# 2. 提取图片
result = parser.parse(
    "data.xlsx",
    extract_images=True,  # 默认True
    images_dir="/custom/path/images"  # 可选，自定义图片目录
)

# 查看提取的图片
if "extracted_images" in result.metadata:
    print(f"提取了 {result.metadata['images_count']} 张图片:")
    for img_path in result.metadata['extracted_images']:
        print(f"  - {img_path}")

# 3. 分析公式依赖关系
if "aggregate_formulas" in result.metadata:
    print(f"聚合公式（合计等）：")
    for formula in result.metadata['aggregate_formulas']:
        print(f"  {formula['cell']}: {formula['description']}")

# 4. 完整配置示例
result = parser.parse(
    "complex_report.xlsx",
    output_format="html",
    chunk_rows=512,              # HTML分块大小
    clean_illegal_chars=True,     # 清理非法字符
    preserve_styles=True,         # ✅ 保留样式
    include_empty_rows=False,     # 不包含空行
    extract_images=True,          # ✅ 提取图片
    images_dir="./my_images"      # 自定义图片目录
)
```

### MCP工具 - 批量解析

```python
# 在AI对话中：
"把 /reports 目录下所有xlsx文件转换为markdown，保存到 /output"

# Claude会调用 batch_parse 工具：
batch_parse(
    file_paths=["/reports/file1.xlsx", "/reports/file2.xlsx", ...],
    output_format="auto",
    output_dir="/output"
)
```

## 📦 项目结构

```
TableParser/
├── table_parser/           # 核心代码
│   ├── __init__.py        # 导出接口
│   ├── parser.py          # 主控制器
│   ├── loader.py          # 文件加载器
│   ├── analyzer.py        # 复杂度分析器
│   ├── converter.py       # 格式转换器
│   ├── types.py           # 类型定义
│   ├── exceptions.py      # 自定义异常
│   ├── mcp_server.py      # MCP服务器
│   └── utils/             # 工具函数
├── tests/                 # 测试代码
├── examples/              # 示例代码
├── requirements.txt       # 依赖列表
├── README.md             # 本文件
└── 技术方案.md            # 详细技术方案
```

## ⚠️ 已知限制

### Excel对象限制（openpyxl技术限制）

**无法提取的内容：**
- ❌ **文本框中的文字**（浮动文本框不属于单元格）
- ❌ **公式编辑器对象**（Equation Editor/MathType插入的数学公式）
- ❌ **形状对象中的文本**（SmartArt、艺术字等）
- ❌ **OLE嵌入对象**（其他程序插入的对象）

**可以提取的内容：**
- ✅ 单元格中的所有内容（文本、数字、公式）
- ✅ 单元格样式（背景色、字体、粗体等）
- ✅ 图片对象（保存为文件）
- ✅ Unicode上下标字符（H₂O、x²等）
- ✅ **文本框/形状中的文本**（v1.2新增，通过XML解析）

**部分支持（提取文本但排版简化）：**
- ⚠️ **公式编辑器对象**（Equation Editor/MathType）
  - ✅ 可提取纯文本内容（如：cos𝛼+cos𝛽=2cos1/2...）
  - ❌ 无法保留专业数学排版（分数线、求和符号大小、上下标对齐等）
  - 原因：公式编辑器是OLE嵌入对象，内部为二进制+LaTeX格式
  - 建议：如需完美排版，请在Excel中将公式另存为图片

**完全无法提取：**
- ❌ OLE嵌入对象的二进制内容

**元数据增强：**
- 返回值中包含 `shapes_text`：文本框和形状中提取的所有文本（包括公式编辑器的文本表示）
- 返回值中包含 `shapes_count`：形状对象数量

## 🆚 与竞品对比

| 特性 | TableParser v1.2 | RAGFlow | Dify | MinerU |
|-----|-----------------|---------|------|--------|
| 复杂度分析 | ✅ 7维度动态权重 | ❌ | ❌ | ❌ |
| 自适应输出 | ✅ | ❌ | ❌ | ❌ |
| 多格式支持 | ✅ MD/HTML | ✅ | ❌ | ✅ MD |
| 合并单元格 | ✅ 完整支持 | ✅ | ⚠️ 展开 | ⚠️ |
| **图片提取** | ✅ **自动提取** | ⚠️ 部分 | ❌ | ✅ |
| **样式保留** | ✅ **完整支持** | ❌ | ❌ | ⚠️ 部分 |
| **上下标** | ✅ **完整支持** | ❌ | ❌ | ⚠️ 部分 |
| **公式分析** | ✅ **依赖追踪** | ❌ | ❌ | ❌ |
| MCP支持 | ✅ | ❌ | ❌ | ❌ |
| 轻量级 | ✅ 最小依赖 | ⚠️ 重 | ✅ | ⚠️ 依赖MS |

## 🎨 新功能详解 (v1.2)

### 1. 图片提取

```python
# 自动提取图片到Excel同目录的images文件夹
result = parser.parse("report.xlsx", extract_images=True)

# 查看提取结果
print(f"提取了 {result.metadata['images_count']} 张图片")
for img in result.metadata['extracted_images']:
    print(f"  - {img}")

# 输出示例：
#   /path/to/images/Sheet1_1.png
#   /path/to/images/Sheet1_2.jpg
```

### 2. 样式保留

```python
# 启用样式保留（背景色、字体颜色、高亮等）
result = parser.parse("data.xlsx", preserve_styles=True)

# HTML输出将包含：
# <td style="background-color: #FFFF00; color: #FF0000; font-weight: bold">内容</td>
```

**支持的样式**：
- ✅ 背景色（填充颜色）
- ✅ 字体颜色
- ✅ 粗体、斜体、下划线
- ✅ 字体大小
- ✅ 高亮识别（黄色系背景自动标记）

### 3. 上下标支持

```python
# 化学式：H₂O、CO₂
# 数学式：x²、E=mc²
# HTML输出：H<sub>2</sub>O、x<sup>2</sup>
# Markdown输出：H~2~O、x^2^
```

**支持两种方式**：
1. ✅ **Unicode上下标字符**（推荐）
   - 输入：H₂O（Unicode字符）
   - 自动转换：H<sub>2</sub>O
   - 支持：²³¹⁰⁴⁵⁶⁷⁸⁹（上标）、₀₁₂₃₄₅₆₇₈₉（下标）

2. ✅ **富文本格式**
   - Excel中设置字体格式为上标/下标
   - 自动识别并转换

**支持场景**：
- ✅ 化学式：H₂O、CO₂、H₂SO₄
- ✅ 数学式：x²、a³、10⁻³
- ✅ 混合文本：普通文字H₂O混合
- ✅ HTML和Markdown双格式输出

### 4. 公式依赖分析

```python
result = parser.parse("financial_report.xlsx")

# 查看公式分析结果
print(f"公式总数: {result.metadata['formulas_count']}")

# 聚合公式（合计、平均等）
for formula in result.metadata['aggregate_formulas']:
    print(f"{formula['cell']}: {formula['description']}")
    # 输出: A10: 聚合计算（合计/平均/计数） - 使用函数: SUM

# 百分比公式
for formula in result.metadata['percentage_formulas']:
    print(f"{formula['cell']}: {formula['description']}")
    # 输出: B5: 百分比计算
```

**支持的分析**：
- ✅ 聚合函数：SUM、AVERAGE、COUNT、MAX、MIN
- ✅ 百分比计算
- ✅ 逻辑判断：IF、AND、OR
- ✅ 查找函数：VLOOKUP、HLOOKUP、INDEX、MATCH
- ✅ 单元格引用追踪：A1、A1:B10、Sheet2!C3

## 🔍 使用场景

### 1. 数据分析
```python
# 快速将Excel转为Markdown，便于Git版本控制
result = parser.parse("report.xlsx", output_format="markdown")
Path("report.md").write_text(result.content)
```

### 2. 文档处理
```python
# 复杂报表保留完整结构（HTML）
result = parser.parse("complex_report.xlsx", output_format="auto")
if result.complexity_score.level == "complex":
    print("检测到复杂表格，已使用HTML格式")
```

### 3. AI助手集成
```
用户: "分析这个财务报表的复杂度"
AI: 自动调用 analyze_complexity 工具
AI: "检测到复杂的多级表头结构，推荐使用HTML格式以保留完整布局"
```

### 4. 批量转换
```python
# 将整个目录的Excel文件批量转换
from table_parser.mcp_server import batch_parse

result = batch_parse(
    file_paths=glob("data/*.xlsx"),
    output_format="auto",
    output_dir="output/"
)
print(f"成功: {result['succeeded']}, 失败: {result['failed']}")
```

## 📝 性能指标

| 表格规模 | 解析时间 |
|---------|---------|
| 小表 (<1000行) | <1秒 |
| 中表 (1000-10000行) | <5秒 |
| 大表 (>10000行) | <30秒 |

## 🛠️ 开发与测试

### 快速测试

```bash
# 测试导入
python -c "from table_parser import TableParser; print('✅ 导入成功')"

# 测试解析
python -c "from table_parser import TableParser; p = TableParser(); r = p.parse('tests/test_data.xlsx'); print(f'✅ 解析成功: {r.output_format}')"
```

### 启动MCP服务器

```bash
# stdio模式（推荐，用于 Cursor/Claude）
python start_mcp_server.py

# HTTP模式（用于独立服务）
python start_mcp_server.py --http --port 8765
```

### 完整测试

```bash
pytest tests/
```

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📄 许可证

Apache License 2.0

## 🙏 致谢

本项目参考了以下开源项目的设计：
- [RAGFlow](https://github.com/infiniflow/ragflow) - Excel解析三层容错机制
- [Dify](https://github.com/langgenius/dify) - 简洁的API设计
- [MinerU](https://github.com/opendatalab/MinerU) - 文档处理架构
- [FastMCP](https://github.com/jlowin/fastmcp) - MCP服务器框架

## 📦 版本历史

### v1.2.0 (2025-11-18)

#### 核心算法改进
- 🚀 **动态权重策略**：根据表格特征智能选择基础权重或高级权重
- 📈 **准确度大幅提升**：修复了权重浪费问题，准确识别各类复杂表格
- 🐛 **修复**：有合并单元格的表格不会被错误判断为simple

#### 新增检测能力（8维度）
- ✨ **内容丰富度**（权重15%）：图片、样式、上下标等 - 关键改进
- ✨ 数据透视表检测（权重10-15%）
- ✨ 图表数量检测（权重10%）
- ✨ VBA宏检测（权重10-20%）
- 🎯 **基础权重**：纯结构表格（结构60% + 数据30% + 规模10%）
- 🎯 **高级权重**：功能丰富表格（平衡8个维度）

#### 新增内容处理
- 📸 **图片提取**：自动提取Excel中的图片到images文件夹
- 🎨 **样式保留**：支持背景色、字体颜色、粗体、斜体、下划线、高亮等
- 🔤 **上下标支持**：完整支持上标<sup>和下标<sub>格式（HTML输出）
- 📊 **公式依赖分析**：识别聚合公式（SUM/AVERAGE）、百分比计算、累计值等
- 🔗 **数据关系追踪**：分析单元格引用、计算类型、依赖关系

#### 功能参数
- `extract_images`: 是否提取图片（默认True）
- `images_dir`: 图片保存目录（默认Excel同目录/images）
- `preserve_styles`: 是否保留样式（默认False，启用后HTML包含颜色等样式）

#### 元数据增强
- 返回图片数量和路径列表
- 返回公式总数、聚合公式、百分比公式列表
- 返回样式统计信息

### v1.1.0 (2025-11-18)
- ✨ **智能自动保存**：默认保存到Excel同目录，节省90%-99% token
- 📁 自定义保存路径支持
- 💾 自动根据复杂度选择扩展名（.html/.md）
- 🏷️ 返回 `auto_generated` 标记

### v1.0.0 (2025-11-17)
- 🎉 初始版本发布
- 🧠 智能复杂度分析（4维度）
- 🎯 自适应格式输出
- 💡 MCP工具支持

## 📞 联系方式

- 项目主页：[GitHub Repository]
- 问题反馈：[GitHub Issues]
- 技术方案：查看 `技术方案.md`

---

**TableParser v1.2** - 让表格解析更智能、更简单！ 🚀

