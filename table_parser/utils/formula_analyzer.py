"""
公式依赖关系分析器

解析Excel公式，提取数据依赖关系
"""

import logging
import re
from typing import Dict, List, Set, Tuple, Optional
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


class FormulaAnalyzer:
    """
    Excel公式分析器
    
    支持：
    - 识别常见函数（SUM、AVERAGE、COUNT等）
    - 提取单元格引用
    - 分析依赖关系
    - 识别计算类型（合计、百分比、累计等）
    """
    
    # 常见函数分类
    AGGREGATE_FUNCTIONS = {'SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'MEDIAN'}
    PERCENTAGE_FUNCTIONS = {'PERCENTAGE', 'PERCENTRANK'}
    CUMULATIVE_FUNCTIONS = {'CUMSUM', 'RUNNINGTOTAL'}
    LOGICAL_FUNCTIONS = {'IF', 'AND', 'OR', 'NOT', 'IFERROR'}
    LOOKUP_FUNCTIONS = {'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP'}
    
    def analyze_formula(self, cell: Cell) -> Optional[Dict]:
        """
        分析单元格公式
        
        Args:
            cell: openpyxl Cell对象
            
        Returns:
            公式信息字典：
            - formula: 公式字符串
            - formula_type: 公式类型（aggregate/percentage/logical等）
            - functions: 使用的函数列表
            - cell_refs: 引用的单元格列表
            - is_calculation: 是否为计算公式
        """
        # 检查是否为公式
        if cell.data_type != 'f' or not cell.value:
            return None
        
        formula = str(cell.value)
        
        # 去掉开头的等号
        if formula.startswith('='):
            formula = formula[1:]
        
        # 提取函数名
        functions = self._extract_functions(formula)
        
        # 提取单元格引用
        cell_refs = self._extract_cell_references(formula)
        
        # 判断公式类型
        formula_type = self._classify_formula(functions, formula)
        
        # 判断是否为计算公式
        is_calculation = len(functions) > 0 or len(cell_refs) > 0
        
        return {
            "formula": f"={formula}",
            "formula_type": formula_type,
            "functions": list(functions),
            "cell_refs": cell_refs,
            "is_calculation": is_calculation,
            "description": self._describe_formula(formula_type, functions)
        }
    
    def _extract_functions(self, formula: str) -> Set[str]:
        """
        从公式中提取函数名
        
        Args:
            formula: 公式字符串（不含=）
            
        Returns:
            函数名集合
        """
        # 匹配函数名（字母开头，后跟括号）
        pattern = r'\b([A-Z][A-Z0-9_]*)\s*\('
        matches = re.findall(pattern, formula.upper())
        return set(matches)
    
    def _extract_cell_references(self, formula: str) -> List[str]:
        """
        从公式中提取单元格引用
        
        Args:
            formula: 公式字符串
            
        Returns:
            单元格引用列表（如 ['A1', 'B2:B10', 'Sheet2!C3']）
        """
        refs = []
        
        # 匹配单元格引用模式
        # 支持：A1, $A$1, Sheet1!A1, A1:B10等
        patterns = [
            r'\b([A-Z]+\$?[0-9]+)\b',                    # 简单引用：A1
            r'\b([A-Z]+\$?[0-9]+:[A-Z]+\$?[0-9]+)\b',    # 范围引用：A1:B10
            r'\b(\w+![$]?[A-Z]+[$]?[0-9]+)\b',           # 跨表引用：Sheet1!A1
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, formula, re.IGNORECASE)
            refs.extend(matches)
        
        # 去重并排序
        return sorted(set(refs))
    
    def _classify_formula(self, functions: Set[str], formula: str) -> str:
        """
        分类公式类型
        
        Args:
            functions: 函数集合
            formula: 公式字符串
            
        Returns:
            公式类型（aggregate/percentage/logical/lookup/arithmetic/other）
        """
        # 聚合函数（合计、平均等）
        if functions & self.AGGREGATE_FUNCTIONS:
            return "aggregate"
        
        # 百分比计算
        if '%' in formula or functions & self.PERCENTAGE_FUNCTIONS:
            return "percentage"
        
        # 逻辑判断
        if functions & self.LOGICAL_FUNCTIONS:
            return "logical"
        
        # 查找引用
        if functions & self.LOOKUP_FUNCTIONS:
            return "lookup"
        
        # 算术运算（加减乘除）
        if any(op in formula for op in ['+', '-', '*', '/']):
            return "arithmetic"
        
        return "other"
    
    def _describe_formula(self, formula_type: str, functions: List[str]) -> str:
        """
        生成公式描述
        
        Args:
            formula_type: 公式类型
            functions: 函数列表
            
        Returns:
            人类可读的描述
        """
        descriptions = {
            "aggregate": "聚合计算（合计/平均/计数）",
            "percentage": "百分比计算",
            "logical": "逻辑判断",
            "lookup": "查找引用",
            "arithmetic": "算术运算",
            "other": "其他计算"
        }
        
        desc = descriptions.get(formula_type, "计算公式")
        
        if functions:
            func_str = ', '.join(functions[:3])  # 最多显示3个
            desc += f" - 使用函数: {func_str}"
        
        return desc
    
    def analyze_sheet_dependencies(self, sheet: Worksheet) -> Dict:
        """
        分析整个工作表的数据依赖关系
        
        Args:
            sheet: openpyxl Worksheet对象
            
        Returns:
            依赖关系字典：
            - formulas_count: 公式总数
            - aggregate_cells: 聚合计算单元格列表
            - percentage_cells: 百分比计算单元格列表
            - dependency_graph: 依赖关系图
        """
        result = {
            "formulas_count": 0,
            "aggregate_cells": [],
            "percentage_cells": [],
            "calculation_cells": [],
            "formula_types": {}
        }
        
        try:
            # 遍历所有单元格（采样前100行避免性能问题）
            max_rows = min(100, sheet.max_row)
            
            for row in sheet.iter_rows(max_row=max_rows):
                for cell in row:
                    if cell.data_type == 'f':  # 公式单元格
                        formula_info = self.analyze_formula(cell)
                        
                        if formula_info:
                            result["formulas_count"] += 1
                            
                            # 记录单元格坐标
                            cell_coord = cell.coordinate
                            
                            # 按类型分类
                            if formula_info["formula_type"] == "aggregate":
                                result["aggregate_cells"].append({
                                    "cell": cell_coord,
                                    "formula": formula_info["formula"],
                                    "description": formula_info["description"]
                                })
                            elif formula_info["formula_type"] == "percentage":
                                result["percentage_cells"].append({
                                    "cell": cell_coord,
                                    "formula": formula_info["formula"],
                                    "description": formula_info["description"]
                                })
                            else:
                                result["calculation_cells"].append({
                                    "cell": cell_coord,
                                    "formula": formula_info["formula"],
                                    "type": formula_info["formula_type"],
                                    "description": formula_info["description"]
                                })
                            
                            # 统计公式类型
                            ftype = formula_info["formula_type"]
                            result["formula_types"][ftype] = result["formula_types"].get(ftype, 0) + 1
            
            logger.debug(f"工作表依赖分析完成: {result['formulas_count']}个公式")
            
        except Exception as e:
            logger.warning(f"依赖关系分析失败: {e}")
        
        return result

