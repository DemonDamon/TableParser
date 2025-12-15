"""
TableParser类型定义模块

定义了所有核心数据类型和类型别名
"""

from typing import Literal, Optional, Union
from dataclasses import dataclass, field
from pathlib import Path

# 输出格式类型
OutputFormat = Literal["auto", "markdown", "html"]

# 复杂度等级
ComplexityLevel = Literal["simple", "medium", "complex"]


@dataclass
class ParseOptions:
    """解析选项配置"""
    
    output_format: OutputFormat = "auto"
    """输出格式: auto(自动选择) / markdown / html"""
    
    chunk_rows: int = 0
    """HTML分块大小（行数），默认为0（不分块，输出完整表格）。设置为正数可以分块处理大表"""
    
    encoding: Optional[str] = None
    """CSV编码（None表示自动检测）"""
    
    clean_illegal_chars: bool = True
    """是否清理非法字符"""
    
    preserve_styles: bool = False
    """是否保留样式（HTML模式）"""
    
    include_empty_rows: bool = False
    """是否包含空行"""


@dataclass
class ComplexityScore:
    """复杂度评分结构"""
    
    # 各维度评分 (0-100)
    merged_cells_score: float
    """合并单元格复杂度"""
    
    header_depth_score: float
    """表头层级复杂度"""
    
    data_structure_score: float
    """数据结构复杂度"""
    
    scale_score: float
    """规模复杂度"""
    
    # 综合得分
    total_score: float
    """总分（加权平均）"""
    
    # 复杂度等级
    level: ComplexityLevel
    """复杂度等级: simple / medium / complex"""
    
    # 推荐输出格式
    recommended_format: OutputFormat
    """推荐的输出格式"""
    
    # 详细信息
    details: dict = field(default_factory=dict)
    """详细分析信息"""
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            "merged_cells_score": self.merged_cells_score,
            "header_depth_score": self.header_depth_score,
            "data_structure_score": self.data_structure_score,
            "scale_score": self.scale_score,
            "total_score": self.total_score,
            "level": self.level,
            "recommended_format": self.recommended_format,
            "details": self.details
        }


@dataclass
class ParseResult:
    """解析结果"""
    
    success: bool
    """是否成功"""
    
    output_format: str
    """实际使用的格式"""
    
    content: Union[str, list[str]]
    """解析内容（Markdown为str, HTML为list）"""
    
    complexity_score: Optional[ComplexityScore]
    """复杂度评分（如果进行了分析）"""
    
    metadata: dict
    """元数据"""
    
    error: Optional[str] = None
    """错误信息（如果失败）"""
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            "success": self.success,
            "output_format": self.output_format,
            "content": self.content,
            "complexity_score": self.complexity_score.to_dict() if self.complexity_score else None,
            "metadata": self.metadata,
            "error": self.error
        }

