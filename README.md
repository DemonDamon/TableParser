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

- 🧠 **智能复杂度分析**：4维度评分系统，自动判断表格复杂度
- 🎯 **自适应输出**：根据复杂度自动选择Markdown或HTML
- 💾 **智能自动保存**：默认保存到同目录，节省90%-99% token（v1.1新增）
- 🛡️ **三层容错机制**：openpyxl → pandas → calamine，最大化兼容性
- 💡 **MCP工具化**：支持AI智能体（Claude、GPT等）直接调用
- 🚀 **轻量级**：最小化依赖，核心仅需openpyxl+pandas
- 📦 **易于集成**：简洁的API，支持Python库、CLI、MCP多种方式

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

TableParser创新性地实现了4维度评分系统：

| 维度 | 权重 | 评估内容 |
|------|------|---------|
| **合并单元格** | 40% | 合并单元格数量、比例、复杂度 |
| **表头层级** | 30% | 单级/多级表头检测 |
| **数据结构** | 20% | 公式、超链接、富文本 |
| **表格规模** | 10% | 行列数规模 |

**评分规则：**
- **0-30分**：简单表格 → 推荐Markdown（易读易编辑）
- **31-60分**：中等复杂 → 推荐Markdown（提示可能有损失）
- **61-100分**：复杂表格 → 强制HTML（完整保留结构）

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
  │   └─ 4维度评分算法
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

### 带选项解析

```python
result = parser.parse(
    "data.xlsx",
    output_format="html",
    chunk_rows=512,  # HTML分块大小
    clean_illegal_chars=True,  # 清理非法字符
    preserve_styles=False,  # 保留样式（暂未实现）
    include_empty_rows=False  # 包含空行
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

## 🆚 与竞品对比

| 特性 | TableParser | RAGFlow | Dify | MinerU |
|-----|------------|---------|------|--------|
| 复杂度分析 | ✅ 智能评分 | ❌ | ❌ | ❌ |
| 自适应输出 | ✅ | ❌ | ❌ | ❌ |
| 多格式支持 | ✅ MD/HTML | ✅ | ❌ | ✅ MD |
| 合并单元格 | ✅ 完整支持 | ✅ | ⚠️ 展开 | ⚠️ |
| MCP支持 | ✅ | ❌ | ❌ | ❌ |
| 轻量级 | ✅ 最小依赖 | ⚠️ 重 | ✅ | ⚠️ 依赖MS |

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

### v1.1.0 (2025-11-18)
- ✨ **智能自动保存**：默认保存到Excel同目录，节省90%-99% token
- 📁 自定义保存路径支持
- 💾 自动根据复杂度选择扩展名（.html/.md）
- 🏷️ 返回 `auto_generated` 标记

### v1.0.0 (2025-11-17)
- 🎉 初始版本发布
- 🧠 智能复杂度分析
- 🎯 自适应格式输出
- 💡 MCP工具支持

## 📞 联系方式

- 项目主页：[GitHub Repository]
- 问题反馈：[GitHub Issues]
- 技术方案：查看 `技术方案.md`

---

**TableParser v1.1** - 让表格解析更智能、更简单！ 🚀

