"""
验证工具模块

提供文件路径、参数等验证功能
"""

from pathlib import Path
from typing import Union
import logging

from ..exceptions import ValidationError
from ..types import OutputFormat

logger = logging.getLogger(__name__)

# 支持的文件扩展名
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}

# 支持的输出格式
SUPPORTED_OUTPUT_FORMATS = {"auto", "markdown", "html"}


def validate_file_path(file_path: Union[str, Path]) -> Path:
    """
    验证文件路径
    
    Args:
        file_path: 文件路径
        
    Returns:
        验证后的Path对象
        
    Raises:
        ValidationError: 验证失败时抛出
    """
    if isinstance(file_path, str):
        file_path = Path(file_path)
    
    # 检查文件是否存在
    if not file_path.exists():
        raise ValidationError(f"文件不存在: {file_path}")
    
    # 检查是否是文件
    if not file_path.is_file():
        raise ValidationError(f"路径不是文件: {file_path}")
    
    # 检查文件扩展名
    extension = file_path.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValidationError(
            f"不支持的文件类型: {extension}，"
            f"支持的类型: {', '.join(SUPPORTED_EXTENSIONS)}"
        )
    
    logger.debug(f"文件路径验证通过: {file_path}")
    return file_path


def validate_output_format(output_format: str) -> OutputFormat:
    """
    验证输出格式
    
    Args:
        output_format: 输出格式字符串
        
    Returns:
        验证后的输出格式
        
    Raises:
        ValidationError: 验证失败时抛出
    """
    if output_format not in SUPPORTED_OUTPUT_FORMATS:
        raise ValidationError(
            f"不支持的输出格式: {output_format}，"
            f"支持的格式: {', '.join(SUPPORTED_OUTPUT_FORMATS)}"
        )
    
    return output_format  # type: ignore


def validate_file_size(file_path: Path, max_size: int = 50 * 1024 * 1024) -> bool:
    """
    验证文件大小
    
    Args:
        file_path: 文件路径
        max_size: 最大文件大小（字节），默认50MB
        
    Returns:
        是否通过验证
        
    Raises:
        ValidationError: 文件过大时抛出
    """
    file_size = file_path.stat().st_size
    
    if file_size > max_size:
        raise ValidationError(
            f"文件过大: {file_size / 1024 / 1024:.2f}MB，"
            f"最大允许: {max_size / 1024 / 1024:.2f}MB"
        )
    
    logger.debug(f"文件大小验证通过: {file_size / 1024 / 1024:.2f}MB")
    return True

