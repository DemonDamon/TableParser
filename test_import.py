"""
快速测试脚本 - 验证模块导入和基本功能
"""

import sys
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

def test_imports():
    """测试模块导入"""
    print("=" * 60)
    print("测试模块导入...")
    print("=" * 60)
    
    try:
        from table_parser import TableParser
        print("✅ TableParser 导入成功")
    except Exception as e:
        print(f"❌ TableParser 导入失败: {e}")
        return False
    
    try:
        from table_parser import ParseResult, ComplexityScore
        print("✅ 类型定义导入成功")
    except Exception as e:
        print(f"❌ 类型定义导入失败: {e}")
        return False
    
    try:
        from table_parser.loader import FileLoader
        from table_parser.analyzer import ComplexityAnalyzer
        from table_parser.converter import FormatConverter
        print("✅ 核心组件导入成功")
    except Exception as e:
        print(f"❌ 核心组件导入失败: {e}")
        return False
    
    return True


def test_basic_functionality():
    """测试基本功能"""
    print("\n" + "=" * 60)
    print("测试基本功能...")
    print("=" * 60)
    
    try:
        from table_parser import TableParser
        
        # 创建解析器
        parser = TableParser()
        print("✅ TableParser 实例化成功")
        
        # 验证组件
        assert hasattr(parser, 'loader'), "缺少loader组件"
        assert hasattr(parser, 'analyzer'), "缺少analyzer组件"
        assert hasattr(parser, 'converter'), "缺少converter组件"
        print("✅ 所有核心组件存在")
        
        # 验证方法
        assert callable(parser.parse), "parse方法不可调用"
        assert callable(parser.analyze_only), "analyze_only方法不可调用"
        assert callable(parser.preview), "preview方法不可调用"
        print("✅ 所有核心方法可调用")
        
        return True
        
    except Exception as e:
        print(f"❌ 基本功能测试失败: {e}")
        return False


def test_dependencies():
    """测试依赖库"""
    print("\n" + "=" * 60)
    print("测试依赖库...")
    print("=" * 60)
    
    deps = {
        "openpyxl": "Excel解析（主引擎）",
        "pandas": "DataFrame操作",
        "fastmcp": "MCP服务器框架",
    }
    
    optional_deps = {
        "chardet": "编码检测（可选）",
        "xlrd": "旧版.xls支持（可选）",
    }
    
    all_ok = True
    
    # 必需依赖
    for dep, desc in deps.items():
        try:
            __import__(dep)
            print(f"✅ {dep:15s} - {desc}")
        except ImportError:
            print(f"❌ {dep:15s} - {desc} [未安装]")
            all_ok = False
    
    # 可选依赖
    for dep, desc in optional_deps.items():
        try:
            __import__(dep)
            print(f"✅ {dep:15s} - {desc}")
        except ImportError:
            print(f"⚠️  {dep:15s} - {desc} [未安装，可选]")
    
    return all_ok


def main():
    print("\n🚀 TableParser 快速测试")
    print("=" * 60)
    
    # 测试导入
    if not test_imports():
        print("\n❌ 模块导入失败，请检查代码")
        return
    
    # 测试基本功能
    if not test_basic_functionality():
        print("\n❌ 基本功能测试失败")
        return
    
    # 测试依赖
    if not test_dependencies():
        print("\n⚠️  部分必需依赖未安装，请运行: pip install -r requirements.txt")
        return
    
    print("\n" + "=" * 60)
    print("✅ 所有测试通过！")
    print("=" * 60)
    print("\n下一步:")
    print("  1. 运行基本示例: python examples/basic_usage.py")
    print("  2. 运行批量处理: python examples/batch_processing.py")
    print("  3. 启动MCP服务器: python table_parser/mcp_server.py")
    print("=" * 60)


if __name__ == "__main__":
    main()

