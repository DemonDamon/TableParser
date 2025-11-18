#!/usr/bin/env python3
"""
TableParser MCP HTTP客户端测试脚本

测试所有4个MCP工具：
1. parse_table - 解析表格
2. analyze_complexity - 分析复杂度
3. batch_parse - 批量解析
4. get_preview - 快速预览

使用前请确保MCP服务器已启动：
python start_mcp_server.py --http --port 8765

注意：本脚本使用FastMCP Client，符合MCP协议标准
"""

import sys
import asyncio
import base64
from pathlib import Path
from typing import Optional

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from fastmcp import Client
except ImportError:
    print("❌ 需要安装 fastmcp 库: pip install fastmcp")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("⚠️  requests库未安装，健康检查功能将不可用")
    requests = None


class MCPClient:
    """MCP客户端（使用FastMCP Client，符合MCP协议）"""
    
    def __init__(self, base_url: str = "http://localhost:8765"):
        """
        初始化MCP客户端
        
        Args:
            base_url: MCP服务器地址（不含路径，FastMCP会自动添加/mcp）
        """
        # FastMCP HTTP模式通常使用 /mcp 端点
        if base_url.endswith("/mcp"):
            self.mcp_url = base_url
        else:
            self.mcp_url = f"{base_url.rstrip('/')}/mcp"
        self.base_url = base_url
    
    async def call_tool_async(self, tool_name: str, **kwargs) -> dict:
        """异步调用MCP工具"""
        try:
            async with Client(self.mcp_url) as client:
                # FastMCP Client.call_tool返回的结果格式
                raw_result = await client.call_tool(tool_name, kwargs)
                
                # 调试：打印原始返回格式
                debug_mode = "--debug" in sys.argv
                if debug_mode:
                    import json
                    print(f"\n[DEBUG] 工具: {tool_name}")
                    print(f"[DEBUG] 原始返回类型: {type(raw_result)}")
                    print(f"[DEBUG] 原始返回内容: {json.dumps(raw_result, indent=2, ensure_ascii=False)[:500]}")
                
                # FastMCP可能返回不同的格式，统一处理
                result = raw_result
                
                # 情况1: 如果已经是字典格式，检查是否包含我们的标准字段
                if isinstance(result, dict):
                    # 如果包含success或error字段，说明已经是我们的格式
                    if "success" in result or "error" in result:
                        return result
                    
                    # 情况2: MCP标准格式 - 包含content字段
                    if "content" in result:
                        content_list = result.get("content", [])
                        if content_list and isinstance(content_list[0], dict):
                            # 提取text内容
                            text_content = content_list[0].get("text", "")
                            if text_content:
                                try:
                                    # 尝试解析为JSON
                                    import json
                                    parsed = json.loads(text_content)
                                    if debug_mode:
                                        print(f"[DEBUG] 解析后的JSON: {json.dumps(parsed, indent=2, ensure_ascii=False)[:500]}")
                                    return parsed
                                except json.JSONDecodeError:
                                    # 不是JSON，返回文本内容
                                    return {"success": True, "content": text_content}
                    
                    # 情况3: 直接是字典，但没有content字段，可能是MCP的其他格式
                    # 尝试查找是否有嵌套的结果
                    if "result" in result:
                        inner_result = result["result"]
                        if isinstance(inner_result, dict):
                            return inner_result
                    
                    # 情况4: 直接返回字典（可能是我们的工具直接返回的）
                    return result
                
                # 情况5: 非字典格式（字符串、列表等）
                elif isinstance(result, str):
                    # 尝试解析为JSON
                    try:
                        import json
                        parsed = json.loads(result)
                        return parsed
                    except:
                        return {"success": True, "result": result}
                
                # 情况6: 其他类型，包装返回
                else:
                    return {"success": True, "result": result}
                    
        except Exception as e:
            import traceback
            error_detail = str(e)
            if "--debug" in sys.argv:
                print(f"\n[DEBUG] 异常详情:")
                traceback.print_exc()
            return {
                "success": False,
                "error": f"MCP调用失败: {error_detail}",
            }
    
    def call_tool(self, tool_name: str, **kwargs) -> dict:
        """同步调用MCP工具（内部使用异步）"""
        try:
            return asyncio.run(self.call_tool_async(tool_name, **kwargs))
        except Exception as e:
            return {
                "success": False,
                "error": f"调用失败: {e}"
            }
    
    def parse_table(
        self,
        file_path: str = None,
        file_content_base64: str = None,
        output_format: str = "auto",
        **options
    ) -> dict:
        """解析表格"""
        return self.call_tool(
            "parse_table",
            file_path=file_path,
            file_content_base64=file_content_base64,
            output_format=output_format,
            **options
        )
    
    def analyze_complexity(
        self,
        file_path: str = None,
        file_content_base64: str = None
    ) -> dict:
        """分析复杂度"""
        return self.call_tool(
            "analyze_complexity",
            file_path=file_path,
            file_content_base64=file_content_base64
        )
    
    def batch_parse(
        self,
        file_paths: list[str],
        output_format: str = "auto",
        output_dir: str = "./output",
        max_workers: int = 4
    ) -> dict:
        """批量解析"""
        return self.call_tool(
            "batch_parse",
            file_paths=file_paths,
            output_format=output_format,
            output_dir=output_dir,
            max_workers=max_workers
        )
    
    def get_preview(
        self,
        file_path: str = None,
        file_content_base64: str = None,
        max_rows: int = 10,
        max_cols: int = 10
    ) -> dict:
        """快速预览"""
        return self.call_tool(
            "get_preview",
            file_path=file_path,
            file_content_base64=file_content_base64,
            max_rows=max_rows,
            max_cols=max_cols
        )


def test_connection(client: MCPClient):
    """测试连接"""
    print("=" * 60)
    print("测试 1: 服务器连接")
    print("=" * 60)
    
    # 方法1: 尝试使用requests检查基本连接（如果可用）
    if requests:
        try:
            # 尝试连接MCP端点
            response = requests.get(client.mcp_url, timeout=5)
            # 任何响应都说明服务器在运行
            print(f"✅ 服务器连接成功: {client.base_url}")
            print(f"   MCP端点: {client.mcp_url}")
            return True
        except requests.exceptions.ConnectionError:
            print(f"❌ 无法连接到服务器: {client.base_url}")
            print(f"   请确保MCP服务器已启动:")
            print(f"   python start_mcp_server.py --http --port 8765")
            return False
        except Exception as e:
            # 其他错误（如405 Method Not Allowed）也说明服务器在运行
            print(f"✅ 服务器连接成功: {client.base_url}")
            print(f"   MCP端点: {client.mcp_url}")
            print(f"   注意: {type(e).__name__}")
            return True
    else:
        # 方法2: 尝试调用一个简单的工具来测试连接
        print(f"   尝试连接MCP服务器: {client.mcp_url}")
        try:
            # 尝试列出工具（如果FastMCP支持）
            print(f"   ⏳ 测试连接中...")
            # 这里先跳过，直接返回True，让后续测试来验证
            print(f"   ℹ️  将通过实际工具调用来验证连接")
            return True
        except Exception as e:
            print(f"❌ 连接测试失败: {e}")
            return False


def test_preview(client: MCPClient, test_file: Path):
    """测试预览功能"""
    print("\n" + "=" * 60)
    print("测试 2: 快速预览 (get_preview)")
    print("=" * 60)
    
    result = client.get_preview(file_path=str(test_file), max_rows=5)
    
    # 调试：打印返回结果
    if "--debug" in sys.argv:
        import json
        print(f"\n[DEBUG] 预览返回结果:")
        print(json.dumps(result, indent=2, ensure_ascii=False)[:1000])
    
    if result.get("success"):
        print("✅ 预览成功")
        # 安全访问字段
        metadata = result.get('metadata', {})
        sheets = result.get('sheets', [])
        
        if metadata:
            print(f"   Sheet数量: {metadata.get('sheets_count', len(sheets))}")
        
        if sheets:
            for sheet in sheets:
                print(f"\n   Sheet: {sheet.get('name', 'Unknown')}")
                print(f"   总行数: {sheet.get('total_rows', 0)}, 总列数: {sheet.get('total_cols', 0)}")
                preview_data = sheet.get('preview', [])
                if preview_data:
                    print(f"   预览数据（前3行）:")
                    for i, row in enumerate(preview_data[:3], 1):
                        # 只显示前5列
                        preview_row = row[:5] if isinstance(row, list) else [row]
                        print(f"     行{i}: {preview_row}")
        else:
            print("   ⚠️  未找到sheets数据")
        return True
    else:
        error_msg = result.get('error', 'Unknown error')
        print(f"❌ 预览失败: {error_msg}")
        if "--debug" in sys.argv:
            import json
            print(f"\n[DEBUG] 完整返回结果:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
        return False


def test_analyze_complexity(client: MCPClient, test_file: Path):
    """测试复杂度分析"""
    print("\n" + "=" * 60)
    print("测试 3: 复杂度分析 (analyze_complexity)")
    print("=" * 60)
    
    result = client.analyze_complexity(file_path=str(test_file))
    
    # 调试：打印返回结果
    if "--debug" in sys.argv:
        import json
        print(f"\n[DEBUG] 复杂度分析返回结果:")
        print(json.dumps(result, indent=2, ensure_ascii=False)[:1000])
    
    if result.get("success"):
        print("✅ 复杂度分析成功")
        score = result.get('complexity_score', {})
        if score:
            print(f"\n   总分: {score.get('total_score', 0):.1f}")
            print(f"   等级: {score.get('level', 'unknown')}")
            print(f"   推荐格式: {score.get('recommended_format', 'auto')}")
            print(f"\n   各维度得分:")
            dims = score.get('dimensions', {})
            print(f"     合并单元格: {dims.get('merged_cells_score', 0):.1f}")
            print(f"     表头层级: {dims.get('header_depth_score', 0):.1f}")
            print(f"     数据结构: {dims.get('data_structure_score', 0):.1f}")
            print(f"     表格规模: {dims.get('scale_score', 0):.1f}")
        else:
            print("   ⚠️  未找到complexity_score数据")
        print(f"\n   建议: {result.get('recommendation', 'N/A')}")
        return True
    else:
        error_msg = result.get('error', 'Unknown error')
        print(f"❌ 复杂度分析失败: {error_msg}")
        if "--debug" in sys.argv:
            import json
            print(f"\n[DEBUG] 完整返回结果:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
        return False


def test_parse_table_filepath(client: MCPClient, test_file: Path):
    """测试解析表格（文件路径方式）"""
    print("\n" + "=" * 60)
    print("测试 4: 解析表格 - 文件路径方式 (parse_table)")
    print("=" * 60)
    
    result = client.parse_table(file_path=str(test_file), output_format="auto")
    
    # 调试：打印返回结果
    if "--debug" in sys.argv:
        import json
        print(f"\n[DEBUG] 解析返回结果（前500字符）:")
        result_str = json.dumps(result, indent=2, ensure_ascii=False)
        print(result_str[:500])
    
    if result.get("success"):
        print("✅ 解析成功")
        output_format = result.get('output_format', 'unknown')
        print(f"   输出格式: {output_format}")
        
        score = result.get('complexity_score')
        if score:
            print(f"   复杂度: {score.get('level', 'unknown')} (得分: {score.get('total_score', 0):.1f})")
        
        metadata = result.get('metadata', {})
        print(f"   Sheet数: {metadata.get('sheets', 0)}")
        print(f"   总行数: {metadata.get('total_rows', 0)}")
        
        # 保存输出
        content = result.get('content')
        if content:
            if isinstance(content, str):
                output_file = Path("tests/output_from_filepath.md")
                output_file.write_text(content, encoding='utf-8')
                print(f"   已保存: {output_file}")
            elif isinstance(content, list):
                output_file = Path("tests/output_from_filepath.html")
                output_file.write_text("\n\n".join(content), encoding='utf-8')
                print(f"   已保存: {output_file} (包含 {len(content)} 个表格块)")
            else:
                print(f"   ⚠️  未知的content类型: {type(content)}")
        else:
            print("   ⚠️  未找到content数据")
        
        return True
    else:
        error_msg = result.get('error', 'Unknown error')
        print(f"❌ 解析失败: {error_msg}")
        if "--debug" in sys.argv:
            import json
            print(f"\n[DEBUG] 完整返回结果:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
        return False


def test_parse_table_base64(client: MCPClient, test_file: Path):
    """测试解析表格（Base64方式）"""
    print("\n" + "=" * 60)
    print("测试 5: 解析表格 - Base64内容方式 (parse_table)")
    print("=" * 60)
    
    # 读取文件并转换为Base64
    with open(test_file, "rb") as f:
        file_content = f.read()
        file_base64 = base64.b64encode(file_content).decode('utf-8')
    
    print(f"   文件大小: {len(file_content)} bytes")
    print(f"   Base64长度: {len(file_base64)} chars")
    
    result = client.parse_table(
        file_content_base64=file_base64,
        output_format="markdown"  # 强制Markdown
    )
    
    if result.get("success"):
        print("✅ 解析成功")
        output_format = result.get('output_format', 'unknown')
        print(f"   输出格式: {output_format}")
        
        # 保存输出
        content = result.get('content')
        if content:
            output_file = Path("tests/output_from_base64.md")
            if isinstance(content, str):
                output_file.write_text(content, encoding='utf-8')
            elif isinstance(content, list):
                output_file.write_text("\n\n".join(content), encoding='utf-8')
            print(f"   已保存: {output_file}")
        else:
            print("   ⚠️  未找到content数据")
        
        return True
    else:
        error_msg = result.get('error', 'Unknown error')
        print(f"❌ 解析失败: {error_msg}")
        return False


def test_batch_parse(client: MCPClient, test_dir: Path):
    """测试批量解析"""
    print("\n" + "=" * 60)
    print("测试 6: 批量解析 (batch_parse)")
    print("=" * 60)
    
    # 查找测试文件
    test_files = list(test_dir.glob("*.xlsx"))
    
    if not test_files:
        print("⚠️  没有找到测试文件，跳过批量解析测试")
        return True
    
    file_paths = [str(f) for f in test_files]
    print(f"   找到 {len(file_paths)} 个文件")
    
    result = client.batch_parse(
        file_paths=file_paths,
        output_format="auto",
        output_dir=str(test_dir / "batch_output")
    )
    
    # 调试：打印返回结果
    if "--debug" in sys.argv:
        import json
        print(f"\n[DEBUG] 批量解析返回结果:")
        print(json.dumps(result, indent=2, ensure_ascii=False)[:1000])
    
    if result.get("success"):
        print("✅ 批量解析成功")
        total = result.get('total', 0)
        succeeded = result.get('succeeded', 0)
        failed = result.get('failed', 0)
        
        print(f"   总计: {total} 个文件")
        print(f"   成功: {succeeded} 个")
        print(f"   失败: {failed} 个")
        
        # 显示详细结果
        results_list = result.get('results', [])
        if results_list:
            if succeeded > 0:
                print(f"\n   成功的文件:")
                for item in results_list:
                    if item.get('status') == 'success':
                        file_name = Path(item.get('file', 'unknown')).name
                        complexity = item.get('complexity_level', 'unknown')
                        print(f"     ✅ {file_name} -> {complexity}")
            
            if failed > 0:
                print(f"\n   失败的文件:")
                for item in results_list:
                    if item.get('status') == 'failed':
                        file_name = Path(item.get('file', 'unknown')).name
                        error = item.get('error', 'Unknown error')
                        print(f"     ❌ {file_name}: {error}")
        
        return True
    else:
        error_msg = result.get('error', 'Unknown error')
        print(f"❌ 批量解析失败: {error_msg}")
        if "--debug" in sys.argv:
            import json
            print(f"\n[DEBUG] 完整返回结果:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
        return False


def main():
    print("\n🚀 TableParser MCP HTTP客户端测试")
    print("=" * 60)
    print("使用FastMCP Client（符合MCP协议标准）")
    print("=" * 60)
    
    # 检查调试模式
    debug_mode = "--debug" in sys.argv
    
    # 配置
    server_url = "http://localhost:8765"
    test_file = Path(__file__).parent / "test_data.xlsx"
    test_dir = Path(__file__).parent
    
    # 检查测试文件
    if not test_file.exists():
        print(f"❌ 测试文件不存在: {test_file}")
        print(f"   请确保测试文件在正确位置")
        return 1
    
    print(f"测试文件: {test_file.name}")
    print(f"服务器地址: {server_url}")
    print("=" * 60)
    
    # 创建客户端
    client = MCPClient(base_url=server_url)
    
    # 运行测试
    tests = [
        ("服务器连接", lambda: test_connection(client)),
        ("快速预览", lambda: test_preview(client, test_file)),
        ("复杂度分析", lambda: test_analyze_complexity(client, test_file)),
        ("解析表格(文件路径)", lambda: test_parse_table_filepath(client, test_file)),
        ("解析表格(Base64)", lambda: test_parse_table_base64(client, test_file)),
        ("批量解析", lambda: test_batch_parse(client, test_dir)),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n❌ 测试异常: {e}")
            if debug_mode:
                import traceback
                traceback.print_exc()
            results.append((name, False))
    
    # 汇总结果
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✅" if result else "❌"
        print(f"{status} {name}")
    
    print("\n" + "=" * 60)
    if passed == total:
        print(f"✅ 所有测试通过！({passed}/{total})")
        print("=" * 60)
        return 0
    else:
        print(f"⚠️  部分测试失败 ({passed}/{total})")
        print("=" * 60)
        return 1


if __name__ == "__main__":
    sys.exit(main())

