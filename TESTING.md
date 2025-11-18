# TableParser 测试指南

## 🔧 问题排查步骤

如果遇到HTTP测试失败，请按以下步骤排查：

### 步骤1：确认MCP服务器已启动

**终端1**（启动服务器）：
```bash
cd /Users/damon/Desktop/品质AI智能客服/TableParser
python start_mcp_server.py --http --port 8765
```

应该看到类似输出：
```
🚀 启动TableParser MCP服务器...
============================================================
模式: HTTP
地址: http://0.0.0.0:8765
============================================================
```

### 步骤2：运行调试脚本

**终端2**（调试HTTP接口）：
```bash
cd /Users/damon/Desktop/品质AI智能客服/TableParser
python tests/debug_http.py
```

这个脚本会：
- ✅ 测试各个HTTP端点
- ✅ 显示响应格式
- ✅ 帮助诊断问题

### 步骤3：运行完整测试

如果调试脚本通过，运行完整测试：
```bash
python tests/test_mcp_client.py
```

## 🐛 常见问题

### 问题1：连接失败

**现象**：
```
❌ 无法连接到服务器: http://localhost:8765
```

**解决方案**：
1. 确认服务器已启动（步骤1）
2. 检查端口是否被占用：
   ```bash
   lsof -i :8765
   ```
3. 尝试使用 `127.0.0.1` 而不是 `localhost`

### 问题2：404错误

**现象**：
```
❌ 请求失败: 404 Client Error: Not Found
```

**可能原因**：
- FastMCP的HTTP端点格式不对
- 运行 `debug_http.py` 查看正确的端点格式

### 问题3：工具调用失败

**现象**：工具返回错误

**检查**：
1. 确认测试文件存在：
   ```bash
   ls -lh tests/test_data.xlsx
   ```
2. 确认文件路径正确（使用绝对路径）

## 💡 替代测试方案

### 方案A：使用Python直接测试（不需要HTTP）

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path.cwd()))

from table_parser import TableParser

# 创建解析器
parser = TableParser()

# 测试文件
test_file = "tests/test_data.xlsx"

# 1. 预览
print("1. 预览测试:")
preview = parser.preview(test_file, max_rows=5)
print(f"✅ Sheet数: {preview['metadata']['sheets_count']}")

# 2. 分析复杂度
print("\n2. 复杂度分析:")
score = parser.analyze_only(test_file)
print(f"✅ 复杂度: {score.level} (得分: {score.total_score:.1f})")

# 3. 解析表格
print("\n3. 解析表格:")
result = parser.parse(test_file, output_format="auto")
if result.success:
    print(f"✅ 解析成功，格式: {result.output_format}")
else:
    print(f"❌ 解析失败: {result.error}")
```

### 方案B：使用stdio模式（推荐用于Claude Desktop）

stdio模式不需要HTTP，直接通过标准输入输出通信：

```bash
python start_mcp_server.py
# 不加 --http 参数，默认使用stdio模式
```

然后在Claude Desktop配置文件中使用：
```json
{
  "mcpServers": {
    "table-parser": {
      "command": "python",
      "args": [
        "-u",
        "/Users/damon/Desktop/品质AI智能客服/TableParser/start_mcp_server.py"
      ]
    }
  }
}
```

## 📋 完整测试清单

- [ ] Python库导入测试（无需服务器）
  ```bash
  python -c "from table_parser import TableParser; print('✅ 导入成功')"
  ```

- [ ] 基础解析测试（无需服务器）
  ```bash
  python -c "from table_parser import TableParser; p = TableParser(); r = p.parse('tests/test_data.xlsx'); print(f'✅ 解析成功: {r.output_format}')"
  ```

- [ ] HTTP服务器启动测试
  ```bash
  python start_mcp_server.py --http --port 8765
  ```

- [ ] HTTP端点调试测试
  ```bash
  python tests/debug_http.py
  ```

- [ ] 完整HTTP客户端测试
  ```bash
  python tests/test_mcp_client.py
  ```

## 🎯 推荐测试顺序

1. **最简单**：Python直接测试（方案A）
2. **Claude集成**：stdio模式（方案B）
3. **HTTP服务**：HTTP模式（需要debug）

如果只是要验证TableParser功能是否正常，建议使用方案A！

## 📞 获取帮助

如果所有方案都失败，请提供以下信息：
- Python版本：`python --version`
- FastMCP版本：`pip show fastmcp`
- 错误信息的完整输出
- `debug_http.py` 的输出

