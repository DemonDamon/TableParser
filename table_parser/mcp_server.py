"""
TableParser MCP服务器

基于FastMCP实现的MCP工具服务器，支持AI智能体直接调用
"""

import base64
import logging
import sys
from pathlib import Path
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

# 处理相对导入和直接运行的兼容性
if __name__ == "__main__":
    # 直接运行时，添加父目录到路径
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from table_parser.parser import TableParser
    from table_parser.types import ComplexityScore
else:
    # 作为模块导入时，使用相对导入
    from .parser import TableParser
    from .types import ComplexityScore

from fastmcp import FastMCP

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 创建MCP服务器
mcp = FastMCP("TableParser")

# 初始化解析器
parser = TableParser()

# 安全配置
ALLOWED_PATHS = [
    "/data",
    "/reports",
    "/tmp",
    "/Users",  # macOS
    "/home",   # Linux
]
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB


def validate_file_path(file_path: str) -> bool:
    """验证文件路径是否在允许的目录中"""
    try:
        abs_path = Path(file_path).resolve()
        return any(
            str(abs_path).startswith(allowed)
            for allowed in ALLOWED_PATHS
        )
    except Exception:
        return False


def validate_file_size(file_path: str) -> bool:
    """验证文件大小"""
    try:
        return Path(file_path).stat().st_size <= MAX_FILE_SIZE
    except Exception:
        return False


def generate_recommendation(score: ComplexityScore) -> str:
    """生成人类可读的建议"""
    if score.level == "simple":
        return (
            f"这是一个简单表格（得分{score.total_score:.1f}），"
            f"推荐使用Markdown格式，易于阅读和编辑。"
        )
    elif score.level == "medium":
        return (
            f"这是一个中等复杂度表格（得分{score.total_score:.1f}），"
            f"可以使用Markdown，但部分结构可能无法完美保留。"
            f"如需精确还原，请使用HTML格式。"
        )
    else:
        return (
            f"这是一个复杂表格（得分{score.total_score:.1f}），"
            f"包含合并单元格或多级表头，强烈推荐使用HTML格式以保留完整结构。"
        )


@mcp.tool()
def parse_table(
    file_path: Optional[str] = None,
    file_content_base64: Optional[str] = None,
    output_format: str = "auto",
    chunk_rows: int = 256,
    clean_illegal_chars: bool = True
) -> dict:
    """
    解析Excel或CSV表格文件
    
    Args:
        file_path: 文件路径（优先使用）
        file_content_base64: Base64编码的文件内容（file_path不存在时使用）
        output_format: 输出格式 (auto/markdown/html)
        chunk_rows: HTML分块行数
        clean_illegal_chars: 是否清理非法字符
        
    Returns:
        解析结果字典
    
    Examples:
        # 解析本地文件
        result = parse_table(file_path="/path/to/data.xlsx")
        
        # 解析Base64内容
        with open("data.xlsx", "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        result = parse_table(file_content_base64=content_b64)
        
        # 强制HTML输出
        result = parse_table(
            file_path="/path/to/data.xlsx",
            output_format="html"
        )
    """
    try:
        # 确定输入源
        if file_path:
            # 安全验证
            if not validate_file_path(file_path):
                return {
                    "success": False,
                    "error": f"文件路径不在允许的目录中: {file_path}"
                }
            
            if not validate_file_size(file_path):
                return {
                    "success": False,
                    "error": f"文件过大（超过50MB）: {file_path}"
                }
            
            input_data = file_path
            logger.info(f"解析文件: {file_path}")
            
        elif file_content_base64:
            try:
                input_data = base64.b64decode(file_content_base64)
                logger.info(f"解析Base64内容 ({len(input_data)} bytes)")
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Base64解码失败: {e}"
                }
        else:
            return {
                "success": False,
                "error": "必须提供 file_path 或 file_content_base64"
            }
        
        # 执行解析
        result = parser.parse(
            input_data,
            output_format=output_format,
            chunk_rows=chunk_rows,
            clean_illegal_chars=clean_illegal_chars
        )
        
        return result.to_dict()
        
    except Exception as e:
        logger.error(f"解析失败: {e}")
        return {
            "success": False,
            "error": str(e)
        }


@mcp.tool()
def analyze_complexity(
    file_path: Optional[str] = None,
    file_content_base64: Optional[str] = None
) -> dict:
    """
    分析表格复杂度（不生成输出内容，仅评估）
    
    Args:
        file_path: 文件路径
        file_content_base64: Base64编码的文件内容
        
    Returns:
        复杂度分析结果字典
    
    Examples:
        # 在解析前先分析
        analysis = analyze_complexity(file_path="/path/to/data.xlsx")
        if analysis["complexity_score"]["level"] == "complex":
            print("检测到复杂表格，推荐使用HTML格式")
    """
    try:
        # 确定输入源
        if file_path:
            # 安全验证
            if not validate_file_path(file_path):
                return {
                    "success": False,
                    "error": f"文件路径不在允许的目录中: {file_path}"
                }
            
            input_data = file_path
            logger.info(f"分析文件复杂度: {file_path}")
            
        elif file_content_base64:
            try:
                input_data = base64.b64decode(file_content_base64)
                logger.info(f"分析Base64内容复杂度 ({len(input_data)} bytes)")
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Base64解码失败: {e}"
                }
        else:
            return {
                "success": False,
                "error": "必须提供 file_path 或 file_content_base64"
            }
        
        # 分析复杂度
        score = parser.analyze_only(input_data)
        
        # 生成建议
        recommendation = generate_recommendation(score)
        
        return {
            "success": True,
            "complexity_score": score.to_dict(),
            "recommendation": recommendation
        }
        
    except Exception as e:
        logger.error(f"复杂度分析失败: {e}")
        return {
            "success": False,
            "error": str(e)
        }


@mcp.tool()
def batch_parse(
    file_paths: list[str],
    output_format: str = "auto",
    output_dir: str = "./output",
    max_workers: int = 4
) -> dict:
    """
    批量解析多个表格文件
    
    Args:
        file_paths: 文件路径列表
        output_format: 输出格式 (auto/markdown/html)
        output_dir: 输出目录
        max_workers: 最大并发数
        
    Returns:
        批量处理结果字典
    
    Examples:
        result = batch_parse(
            file_paths=[
                "/data/report1.xlsx",
                "/data/report2.csv",
                "/data/table3.xlsx"
            ],
            output_format="auto",
            output_dir="./parsed_tables"
        )
    """
    try:
        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        results = []
        succeeded = 0
        failed = 0
        
        def process_file(file_path):
            try:
                # 安全验证
                if not validate_file_path(file_path):
                    return {
                        "file": file_path,
                        "status": "failed",
                        "error": "文件路径不在允许的目录中"
                    }
                
                # 解析文件
                result = parser.parse(file_path, output_format=output_format)
                
                if not result.success:
                    return {
                        "file": file_path,
                        "status": "failed",
                        "error": result.error
                    }
                
                # 保存输出
                file_stem = Path(file_path).stem
                if result.output_format == "markdown":
                    output_file = output_path / f"{file_stem}.md"
                    output_file.write_text(result.content, encoding='utf-8')
                else:  # HTML
                    output_file = output_path / f"{file_stem}.html"
                    output_file.write_text("\n\n".join(result.content), encoding='utf-8')
                
                return {
                    "file": file_path,
                    "status": "success",
                    "output_file": str(output_file),
                    "complexity_level": result.complexity_score.level if result.complexity_score else "unknown"
                }
                
            except Exception as e:
                return {
                    "file": file_path,
                    "status": "failed",
                    "error": str(e)
                }
        
        # 并行处理
        logger.info(f"开始批量处理 {len(file_paths)} 个文件，并发数: {max_workers}")
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(process_file, fp): fp for fp in file_paths}
            
            for future in as_completed(futures):
                result = future.result()
                results.append(result)
                
                if result["status"] == "success":
                    succeeded += 1
                else:
                    failed += 1
        
        logger.info(f"批量处理完成: 成功 {succeeded}, 失败 {failed}")
        
        return {
            "success": True,
            "total": len(file_paths),
            "succeeded": succeeded,
            "failed": failed,
            "results": results
        }
        
    except Exception as e:
        logger.error(f"批量处理失败: {e}")
        return {
            "success": False,
            "error": str(e)
        }


@mcp.tool()
def get_preview(
    file_path: Optional[str] = None,
    file_content_base64: Optional[str] = None,
    max_rows: int = 10,
    max_cols: int = 10
) -> dict:
    """
    预览表格内容（不完整解析，快速返回）
    
    Args:
        file_path: 文件路径
        file_content_base64: Base64编码的文件内容
        max_rows: 最大预览行数
        max_cols: 最大预览列数
        
    Returns:
        预览信息字典
    
    Examples:
        # 快速预览文件内容
        preview = get_preview(
            file_path="/path/to/data.xlsx",
            max_rows=5
        )
        print(f"文件包含 {preview['metadata']['sheets_count']} 个sheet")
        for sheet in preview['sheets']:
            print(f"Sheet: {sheet['name']}, 行数: {sheet['total_rows']}")
    """
    try:
        # 确定输入源
        if file_path:
            # 安全验证
            if not validate_file_path(file_path):
                return {
                    "success": False,
                    "error": f"文件路径不在允许的目录中: {file_path}"
                }
            
            input_data = file_path
            logger.info(f"预览文件: {file_path}")
            
        elif file_content_base64:
            try:
                input_data = base64.b64decode(file_content_base64)
                logger.info(f"预览Base64内容 ({len(input_data)} bytes)")
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Base64解码失败: {e}"
                }
        else:
            return {
                "success": False,
                "error": "必须提供 file_path 或 file_content_base64"
            }
        
        # 预览
        result = parser.preview(
            input_data,
            max_rows=max_rows,
            max_cols=max_cols
        )
        
        return result
        
    except Exception as e:
        logger.error(f"预览失败: {e}")
        return {
            "success": False,
            "error": str(e)
        }


if __name__ == "__main__":
    # 启动MCP服务器
    logger.info("🚀 启动TableParser MCP服务器...")
    mcp.run(transport="stdio")  # 使用标准输入输出（推荐）
    # 或者使用HTTP
    # mcp.run(transport="http", host="0.0.0.0", port=8765)

