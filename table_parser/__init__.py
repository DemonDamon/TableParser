"""
TableParser - 轻量级表格解析工具

智能表格解析，根据复杂度自动选择最佳输出格式（Markdown或HTML）
"""

from .version import __version__
from .parser import TableParser
from .types import (
    ParseOptions,
    ParseResult,
    ComplexityScore,
    OutputFormat,
    ComplexityLevel,
)
from .exceptions import (
    TableParserError,
    FileLoadError,
    ParseError,
    ComplexityAnalysisError,
    ConversionError,
    ValidationError,
)

__all__ = [
    "__version__",
    "TableParser",
    "ParseOptions",
    "ParseResult",
    "ComplexityScore",
    "OutputFormat",
    "ComplexityLevel",
    "TableParserError",
    "FileLoadError",
    "ParseError",
    "ComplexityAnalysisError",
    "ConversionError",
    "ValidationError",
]
