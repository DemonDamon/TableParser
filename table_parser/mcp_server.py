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
    clean_illegal_chars: bool = True,
    output_path: Optional[str] = None,
    extract_images: bool = True,
    images_dir: Optional[str] = None
) -> dict:
    """
    解析Excel或CSV表格文件
    
    Args:
        file_path: 文件路径（优先使用）
        file_content_base64: Base64编码的文件内容（file_path不存在时使用）
        output_format: 输出格式 (auto/markdown/html)
        chunk_rows: HTML分块行数
        clean_illegal_chars: 是否清理非法字符
        output_path: 输出文件路径（可选）
            - 如果提供：保存到指定路径
            - 如果不提供且有file_path：默认保存到Excel同目录（自动节省token）
            - 如果不提供且是Base64输入：返回完整内容
        extract_images: 是否提取Excel中的图片（默认True）
        images_dir: 图片保存目录（可选）
            - 如果提供：保存到指定目录
            - 如果不提供：自动保存到Excel同目录的images文件夹
        
    Returns:
        解析结果字典。保存文件时只返回元数据，不返回完整内容（大幅节省token）
    
    Examples:
        # 示例1：自动保存（推荐，自动节省token）
        # 会保存到 /path/to/data.html 或 data.md（取决于复杂度）
        result = parse_table(file_path="/path/to/data.xlsx")
        
        # 示例2：指定保存路径
        result = parse_table(
            file_path="/path/to/data.xlsx",
            output_path="/another/path/result.html"
        )
        
        # 示例3：Base64输入（返回完整内容）
        with open("data.xlsx", "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
        result = parse_table(file_content_base64=content_b64)
        
        # 示例4：强制HTML格式
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
            clean_illegal_chars=clean_illegal_chars,
            extract_images=extract_images,
            images_dir=images_dir
        )
        
        # 确定输出路径
        # 1. 如果明确提供了 output_path，使用它
        # 2. 如果没有提供 output_path，但有 file_path，默认保存到同目录
        # 3. 如果都没有（Base64输入），则返回完整内容
        actual_output_path = output_path
        
        if not actual_output_path and file_path:
            # 自动生成输出路径：同目录，扩展名改为 .html 或 .md
            source_file = Path(file_path)
            if result.output_format == "markdown":
                actual_output_path = str(source_file.with_suffix('.md'))
            else:  # HTML
                actual_output_path = str(source_file.with_suffix('.html'))
            logger.info(f"未指定输出路径，自动保存到: {actual_output_path}")
        
        # 如果有输出路径（明确指定或自动生成），保存文件并只返回元数据
        if actual_output_path:
            try:
                # 验证输出路径安全性
                if not validate_file_path(actual_output_path):
                    return {
                        "success": False,
                        "error": f"输出路径不在允许的目录中: {actual_output_path}"
                    }
                
                output_file = Path(actual_output_path)
                
                # 确保目录存在
                output_file.parent.mkdir(parents=True, exist_ok=True)
                
                # 根据格式保存文件
                if result.output_format == "markdown":
                    # Markdown格式直接保存
                    output_file.write_text(result.content, encoding="utf-8")
                    logger.info(f"Markdown内容已保存到: {actual_output_path}")
                    
                else:  # HTML格式
                    # 构建完整的HTML文档
                    html_parts = []
                    html_parts.append('<!DOCTYPE html>')
                    html_parts.append('<html lang="zh-CN">')
                    html_parts.append('<head>')
                    html_parts.append('    <meta charset="UTF-8">')
                    html_parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
                    html_parts.append(f'    <title>{Path(file_path).stem if file_path else "表格解析结果"}</title>')
                    html_parts.append('    <style>')
                    html_parts.append('        body { font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }')
                    html_parts.append('        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }')
                    html_parts.append('        h1 { color: #333; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }')
                    html_parts.append('        .metadata { background-color: #f0f7ff; padding: 15px; border-radius: 5px; margin: 20px 0; }')
                    html_parts.append('        table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }')
                    html_parts.append('        th, td { border: 1px solid #ddd; padding: 12px 8px; text-align: left; vertical-align: top; }')
                    html_parts.append('        th { background-color: #4a90e2; color: white; font-weight: bold; }')
                    html_parts.append('        tbody tr:nth-child(even) { background-color: #f9f9f9; }')
                    html_parts.append('        tbody tr:hover { background-color: #e8f4ff; }')
                    html_parts.append('        td[rowspan], td[colspan] { background-color: #fff3cd; font-weight: 500; }')
                    html_parts.append('    </style>')
                    html_parts.append('</head>')
                    html_parts.append('<body>')
                    html_parts.append('    <div class="container">')
                    html_parts.append(f'        <h1>{Path(file_path).stem if file_path else "表格解析结果"}</h1>')
                    
                    # 添加元数据信息
                    if result.metadata:
                        html_parts.append('        <div class="metadata">')
                        html_parts.append('            <h3>📋 文件信息</h3>')
                        html_parts.append('            <ul>')
                        html_parts.append(f'                <li><strong>工作表数量：</strong>{result.metadata.get("sheets", 0)}个</li>')
                        html_parts.append(f'                <li><strong>总行数：</strong>{result.metadata.get("total_rows", 0)}行</li>')
                        html_parts.append(f'                <li><strong>总列数：</strong>{result.metadata.get("total_cols", 0)}列</li>')
                        if result.metadata.get("merged_cells_count"):
                            html_parts.append(f'                <li><strong>合并单元格：</strong>{result.metadata["merged_cells_count"]}个</li>')
                        if result.complexity_score:
                            html_parts.append(f'                <li><strong>复杂度评分：</strong>{result.complexity_score.total_score:.1f}/100（{result.complexity_score.level}）</li>')
                        html_parts.append('            </ul>')
                        html_parts.append('        </div>')
                    
                    # 添加表格内容
                    for i, table_html in enumerate(result.content, 1):
                        if len(result.content) > 1:
                            html_parts.append(f'        <h2>表格 {i}</h2>')
                        html_parts.append(f'        {table_html}')
                    
                    html_parts.append('    </div>')
                    html_parts.append('</body>')
                    html_parts.append('</html>')
                    
                    output_file.write_text('\n'.join(html_parts), encoding="utf-8")
                    logger.info(f"HTML内容已保存到: {actual_output_path}")
                
                # 返回元数据和文件路径，不返回完整内容
                return {
                    "success": True,
                    "output_format": result.output_format,
                    "saved_to": str(actual_output_path),
                    "file_size": output_file.stat().st_size,
                    "complexity_score": result.complexity_score.to_dict() if result.complexity_score else None,
                    "metadata": result.metadata,
                    "message": f"✅ 文件已成功保存到 {actual_output_path}（{output_file.stat().st_size / 1024:.1f} KB）",
                    "auto_generated": output_path is None  # 标记是否为自动生成的路径
                }
                
            except Exception as e:
                logger.error(f"保存文件失败: {e}")
                return {
                    "success": False,
                    "error": f"保存文件失败: {e}"
                }
        
        # 没有提供输出路径，返回完整内容
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
                
                # 解析文件（默认提取图片）
                result = parser.parse(
                    file_path, 
                    output_format=output_format,
                    extract_images=True
                )
                
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

