"""
TableParser自定义异常模块

定义了所有自定义异常类型
"""


class TableParserError(Exception):
    """TableParser基础异常"""
    pass


class FileLoadError(TableParserError):
    """文件加载失败异常"""
    pass


class UnsupportedFileTypeError(TableParserError):
    """不支持的文件类型异常"""
    pass


class ParseError(TableParserError):
    """解析错误异常"""
    pass


class ComplexityAnalysisError(TableParserError):
    """复杂度分析错误异常"""
    pass


class ConversionError(TableParserError):
    """格式转换错误异常"""
    pass


class ValidationError(TableParserError):
    """验证错误异常"""
    pass

