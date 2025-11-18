# TableParser 测试

本目录包含TableParser的测试文件和测试脚本。

## 📁 文件说明

- `test_data.xlsx` - 测试用Excel文件
- `test_mcp_client.py` - MCP HTTP客户端测试脚本

## 🧪 运行测试

### 前提条件

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 启动MCP服务器（HTTP模式）：
```bash
# 在项目根目录执行
python start_mcp_server.py --http --port 8765
```

### 运行MCP客户端测试

在**新的终端窗口**中运行：

```bash
cd /Users/damon/Desktop/品质AI智能客服/TableParser
python tests/test_mcp_client.py
```

**调试模式**（显示详细错误信息）：
```bash
python tests/test_mcp_client.py --debug
```

### 重要说明

本测试脚本使用 **FastMCP Client**，符合 **MCP协议标准**：
- ✅ 使用标准的MCP/JSON-RPC协议
- ✅ 与Claude Desktop使用相同的协议
- ✅ 无需额外的REST API包装层
- ⚠️ 需要异步支持（已封装为同步接口）

## 🔍 测试内容

测试脚本会依次测试所有4个MCP工具：

### 1. 服务器连接测试
检查MCP服务器是否正常运行。

### 2. get_preview - 快速预览
预览表格的前几行数据，不完整解析。

```python
client.get_preview(file_path="test_data.xlsx", max_rows=5)
```

### 3. analyze_complexity - 复杂度分析
分析表格的复杂度并给出建议。

```python
client.analyze_complexity(file_path="test_data.xlsx")
```

### 4. parse_table (文件路径) - 解析表格
使用文件路径解析表格。

```python
client.parse_table(file_path="test_data.xlsx", output_format="auto")
```

### 5. parse_table (Base64) - 解析表格
使用Base64内容解析表格。

```python
with open("test_data.xlsx", "rb") as f:
    content_b64 = base64.b64encode(f.read()).decode()
client.parse_table(file_content_base64=content_b64, output_format="markdown")
```

### 6. batch_parse - 批量解析
批量解析tests目录下的所有Excel文件。

```python
client.batch_parse(
    file_paths=["test_data.xlsx"],
    output_format="auto",
    output_dir="batch_output"
)
```

## 📊 测试输出

测试会生成以下输出文件：

- `output_from_filepath.md` 或 `.html` - 文件路径方式解析结果
- `output_from_base64.md` - Base64方式解析结果
- `batch_output/` - 批量解析输出目录

## ✅ 预期结果

所有测试通过时，会显示：

```
✅ 所有测试通过！(6/6)
```

## 🐛 故障排查

### 问题1：无法连接到服务器

**错误**：`❌ 无法连接到服务器: http://localhost:8765`

**解决**：
1. 确认MCP服务器已启动：`python start_mcp_server.py --http --port 8765`
2. 检查端口是否被占用：`lsof -i :8765`

### 问题2：requests库未安装

**错误**：`❌ 需要安装 requests 库`

**解决**：
```bash
pip install requests
```

### 问题3：测试文件不存在

**错误**：`❌ 测试文件不存在: test_data.xlsx`

**解决**：
确保 `test_data.xlsx` 文件在tests目录下。

## 🔧 自定义测试

您可以修改 `test_mcp_client.py` 中的配置：

```python
# 修改服务器地址
server_url = "http://localhost:8765"  # 改为您的服务器地址

# 修改测试文件
test_file = Path("your_test_file.xlsx")
```

## 📝 使用MCPClient类

您可以在自己的代码中使用 `MCPClient` 类：

```python
from tests.test_mcp_client import MCPClient

# 创建客户端（使用FastMCP Client，符合MCP协议）
client = MCPClient(base_url="http://localhost:8765")

# 解析表格
result = client.parse_table(file_path="data.xlsx", output_format="auto")
if result.get("success"):
    print(f"格式: {result['output_format']}")
    print(f"复杂度: {result['complexity_score']['level']}")

# 分析复杂度
analysis = client.analyze_complexity(file_path="data.xlsx")
if analysis.get("success"):
    print(analysis['recommendation'])

# 快速预览
preview = client.get_preview(file_path="data.xlsx", max_rows=5)
for sheet in preview['sheets']:
    print(f"Sheet: {sheet['name']}, 行数: {sheet['total_rows']}")

# 批量处理
result = client.batch_parse(
    file_paths=["file1.xlsx", "file2.xlsx"],
    output_format="auto",
    output_dir="./output"
)
print(f"成功: {result['succeeded']}/{result['total']}")
```

**注意**：`MCPClient` 内部使用 FastMCP Client，符合MCP协议标准。所有调用都是同步的（内部处理异步）。

## 🚀 持续集成

您可以将此测试脚本集成到CI/CD流程中：

```yaml
# .github/workflows/test.yml 示例
- name: Start MCP Server
  run: python start_mcp_server.py --http --port 8765 &
  
- name: Wait for server
  run: sleep 5
  
- name: Run tests
  run: python tests/test_mcp_client.py
```

