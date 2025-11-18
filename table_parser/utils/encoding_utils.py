"""
编码检测工具模块

提供CSV文件编码自动检测功能
"""

from typing import Optional
import logging

logger = logging.getLogger(__name__)


def detect_encoding(file_content: bytes, default: str = "utf-8") -> str:
    """
    检测文件编码
    
    Args:
        file_content: 文件二进制内容
        default: 默认编码
        
    Returns:
        检测到的编码
    """
    try:
        import chardet
        result = chardet.detect(file_content)
        encoding = result.get("encoding")
        confidence = result.get("confidence", 0)
        
        if encoding and confidence > 0.7:
            logger.debug(f"检测到编码: {encoding} (置信度: {confidence:.2f})")
            return encoding
        else:
            logger.debug(f"编码检测置信度过低，使用默认编码: {default}")
            return default
            
    except ImportError:
        logger.warning("chardet未安装，使用默认编码: {default}")
        return default
    except Exception as e:
        logger.warning(f"编码检测失败: {e}，使用默认编码: {default}")
        return default


def try_decode(file_content: bytes, encodings: Optional[list[str]] = None) -> tuple[str, str]:
    """
    尝试使用多种编码解码文件
    
    Args:
        file_content: 文件二进制内容
        encodings: 尝试的编码列表（None则使用默认列表）
        
    Returns:
        (解码后的文本, 使用的编码)
        
    Raises:
        UnicodeDecodeError: 所有编码都失败时抛出
    """
    if encodings is None:
        encodings = ["utf-8", "gbk", "gb2312", "gb18030", "big5", "latin1"]
    
    for encoding in encodings:
        try:
            text = file_content.decode(encoding)
            logger.debug(f"成功使用编码解码: {encoding}")
            return text, encoding
        except (UnicodeDecodeError, LookupError):
            continue
    
    raise UnicodeDecodeError(
        "unknown", file_content, 0, len(file_content),
        f"无法使用任何编码解码文件，尝试过: {encodings}"
    )

