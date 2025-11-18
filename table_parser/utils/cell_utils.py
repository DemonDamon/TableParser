"""
单元格处理工具模块

提供单元格值格式化、空值判断等工具函数
"""

from typing import Any
import re

# 非法字符正则表达式（来自openpyxl）
ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")


def format_cell_value(value: Any, clean_illegal: bool = True) -> str:
    """
    格式化单元格值
    
    Args:
        value: 单元格值
        clean_illegal: 是否清理非法字符
        
    Returns:
        格式化后的字符串
    """
    if value is None:
        return ""
    
    # 转换为字符串
    result = str(value).strip()
    
    # 清理非法字符
    if clean_illegal:
        result = ILLEGAL_CHARACTERS_RE.sub(" ", result)
    
    return result


def is_empty_cell(value: Any) -> bool:
    """
    判断单元格是否为空
    
    Args:
        value: 单元格值
        
    Returns:
        是否为空
    """
    if value is None:
        return True
    
    if isinstance(value, str) and value.strip() == "":
        return True
    
    return False


def clean_string(s: Any) -> Any:
    """
    清理字符串中的非法字符
    
    Args:
        s: 输入值
        
    Returns:
        清理后的值
    """
    if isinstance(s, str):
        return ILLEGAL_CHARACTERS_RE.sub(" ", s)
    return s

