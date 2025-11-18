"""
文本格式化工具

处理Unicode上下标字符的转换
"""

import re


class TextFormatter:
    """
    文本格式化器
    
    支持：
    - Unicode上下标字符转换为HTML标签
    - 数学符号识别
    """
    
    # Unicode上标字符映射
    SUPERSCRIPT_MAP = {
        '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
        '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
        'ⁿ': 'n', '⁺': '+', '⁻': '-', '⁼': '=', '⁽': '(', '⁾': ')',
    }
    
    # Unicode下标字符映射
    SUBSCRIPT_MAP = {
        '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4',
        '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9',
        '₊': '+', '₋': '-', '₌': '=', '₍': '(', '₎': ')',
    }
    
    def convert_unicode_scripts_to_html(self, text: str, escape_html: bool = True) -> str:
        """
        将Unicode上下标字符转换为HTML标签（同时进行HTML转义）
        
        Args:
            text: 原始文本（如 "H₂O"、"x²"）
            escape_html: 是否转义HTML特殊字符
            
        Returns:
            HTML格式（如 "H<sub>2</sub>O"、"x<sup>2</sup>"）
        """
        from html import escape as html_escape
        
        if not text or not isinstance(text, str):
            return str(text) if text is not None else ""
        
        result = ""
        i = 0
        
        while i < len(text):
            char = text[i]
            
            # 检查是否为上标字符
            if char in self.SUPERSCRIPT_MAP:
                # 收集连续的上标字符
                sup_text = ""
                while i < len(text) and text[i] in self.SUPERSCRIPT_MAP:
                    sup_text += self.SUPERSCRIPT_MAP[text[i]]
                    i += 1
                result += f"<sup>{sup_text}</sup>"
                continue
            
            # 检查是否为下标字符
            if char in self.SUBSCRIPT_MAP:
                # 收集连续的下标字符
                sub_text = ""
                while i < len(text) and text[i] in self.SUBSCRIPT_MAP:
                    sub_text += self.SUBSCRIPT_MAP[text[i]]
                    i += 1
                result += f"<sub>{sub_text}</sub>"
                continue
            
            # 普通字符（需要转义）
            if escape_html:
                result += html_escape(char)
            else:
                result += char
            i += 1
        
        return result
    
    def has_unicode_scripts(self, text: str) -> bool:
        """
        检测文本是否包含Unicode上下标字符
        
        Args:
            text: 文本内容
            
        Returns:
            是否包含上下标
        """
        if not text or not isinstance(text, str):
            return False
        
        for char in text:
            if char in self.SUPERSCRIPT_MAP or char in self.SUBSCRIPT_MAP:
                return True
        
        return False
    
    def extract_formulas_as_text(self, text: str) -> str:
        """
        从文本中提取数学公式并格式化
        
        Args:
            text: 可能包含公式的文本
            
        Returns:
            格式化后的文本
        """
        # 这里可以添加更复杂的公式识别逻辑
        # 暂时只做基础处理
        return text

