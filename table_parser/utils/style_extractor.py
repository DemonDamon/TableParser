"""
样式提取工具

从Excel单元格中提取样式信息（颜色、字体、上下标等）
"""

import logging
from typing import Dict, Optional, List, Tuple
from openpyxl.cell import Cell
from openpyxl.styles import Font, PatternFill

logger = logging.getLogger(__name__)


class StyleExtractor:
    """
    Excel样式提取器
    
    支持提取单元格的各种样式信息
    """
    
    def extract_cell_style(self, cell: Cell) -> Dict[str, any]:
        """
        提取单元格的所有样式信息
        
        Args:
            cell: openpyxl Cell对象
            
        Returns:
            样式字典，包含：
            - background_color: 背景色（RGB十六进制）
            - font_color: 字体颜色
            - is_bold: 是否粗体
            - is_italic: 是否斜体
            - is_underline: 是否下划线
            - font_size: 字体大小
            - has_highlight: 是否有高亮
            - rich_text_parts: 富文本片段（含上下标信息）
        """
        style_info = {
            "background_color": None,
            "font_color": None,
            "is_bold": False,
            "is_italic": False,
            "is_underline": False,
            "font_size": None,
            "has_highlight": False,
            "rich_text_parts": []
        }
        
        try:
            # 提取字体样式
            if cell.font:
                font = cell.font
                style_info["is_bold"] = font.bold or False
                style_info["is_italic"] = font.italic or False
                style_info["is_underline"] = font.underline is not None
                style_info["font_size"] = font.size
                
                # 提取字体颜色
                if hasattr(font, 'color') and font.color:
                    style_info["font_color"] = self._extract_color(font.color)
                
                # 检测上下标
                if hasattr(font, 'vertAlign'):
                    if font.vertAlign == 'superscript':
                        style_info["is_superscript"] = True
                    elif font.vertAlign == 'subscript':
                        style_info["is_subscript"] = True
            
            # 提取背景色（填充）
            if cell.fill and cell.fill.patternType:
                if cell.fill.patternType != 'none':
                    style_info["background_color"] = self._extract_color(cell.fill.fgColor)
                    
                    # 判断是否为高亮（黄色系背景）
                    bg_color = style_info["background_color"]
                    if bg_color and self._is_highlight_color(bg_color):
                        style_info["has_highlight"] = True
            
            # 提取富文本（含上下标信息）
            if hasattr(cell, 'value') and hasattr(cell.value, '__iter__') and not isinstance(cell.value, str):
                # 富文本格式（InlineFont）
                style_info["rich_text_parts"] = self._extract_rich_text(cell.value)
            
        except Exception as e:
            logger.debug(f"样式提取失败: {e}")
        
        return style_info
    
    def _extract_color(self, color) -> Optional[str]:
        """
        提取颜色值
        
        Args:
            color: openpyxl Color对象
            
        Returns:
            RGB十六进制字符串（如 "#FF0000"）
        """
        try:
            if hasattr(color, 'rgb') and color.rgb:
                # RGB格式：AARRGGBB（前两位是透明度）
                rgb = color.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    # 去掉透明度，只保留RGB
                    return f"#{rgb[-6:]}"
            
            if hasattr(color, 'indexed') and color.indexed is not None:
                # 索引颜色（需要查找调色板，这里简化处理）
                return None
            
            if hasattr(color, 'theme') and color.theme is not None:
                # 主题颜色（需要查找主题，这里简化处理）
                return None
                
        except Exception as e:
            logger.debug(f"颜色提取失败: {e}")
        
        return None
    
    def _is_highlight_color(self, color_hex: str) -> bool:
        """
        判断是否为高亮颜色（黄色系）
        
        Args:
            color_hex: RGB十六进制颜色
            
        Returns:
            是否为高亮色
        """
        if not color_hex or not color_hex.startswith('#'):
            return False
        
        try:
            # 提取RGB值
            r = int(color_hex[1:3], 16)
            g = int(color_hex[3:5], 16)
            b = int(color_hex[5:7], 16)
            
            # 判断黄色系（R和G都比较高，B较低）
            # 常见高亮色：FFFF00（黄色）、FFFF99（浅黄）等
            is_yellow = r > 200 and g > 200 and b < 150
            
            return is_yellow
            
        except Exception:
            return False
    
    def _extract_rich_text(self, rich_text_value) -> List[Dict]:
        """
        提取富文本片段（含上下标信息）
        
        Args:
            rich_text_value: 富文本值
            
        Returns:
            富文本片段列表，每个片段包含：
            - text: 文本内容
            - is_superscript: 是否上标
            - is_subscript: 是否下标
            - is_bold: 是否粗体
        """
        parts = []
        
        try:
            for item in rich_text_value:
                if hasattr(item, 'text'):
                    # 富文本片段
                    part = {
                        "text": item.text,
                        "is_superscript": False,
                        "is_subscript": False,
                        "is_bold": False,
                        "is_italic": False
                    }
                    
                    # 提取格式
                    if hasattr(item, 'font') and item.font:
                        font = item.font
                        part["is_bold"] = font.b or False
                        part["is_italic"] = font.i or False
                        
                        # 上下标
                        if hasattr(font, 'vertAlign'):
                            if font.vertAlign == 'superscript':
                                part["is_superscript"] = True
                            elif font.vertAlign == 'subscript':
                                part["is_subscript"] = True
                    
                    parts.append(part)
                else:
                    # 纯文本
                    parts.append({
                        "text": str(item),
                        "is_superscript": False,
                        "is_subscript": False,
                        "is_bold": False,
                        "is_italic": False
                    })
        except Exception as e:
            logger.debug(f"富文本提取失败: {e}")
        
        return parts
    
    def get_cell_html_style(self, cell: Cell) -> str:
        """
        生成单元格的HTML style属性
        
        Args:
            cell: openpyxl Cell对象
            
        Returns:
            HTML style字符串（如 'background-color: #FFFF00; color: #FF0000;'）
        """
        style_info = self.extract_cell_style(cell)
        styles = []
        
        # 背景色
        if style_info["background_color"]:
            styles.append(f"background-color: {style_info['background_color']}")
        
        # 字体颜色
        if style_info["font_color"]:
            styles.append(f"color: {style_info['font_color']}")
        
        # 字体样式
        if style_info["is_bold"]:
            styles.append("font-weight: bold")
        
        if style_info["is_italic"]:
            styles.append("font-style: italic")
        
        if style_info["is_underline"]:
            styles.append("text-decoration: underline")
        
        if style_info["font_size"]:
            styles.append(f"font-size: {style_info['font_size']}pt")
        
        return '; '.join(styles) if styles else ""
    
    def format_rich_text_to_html(self, rich_text_parts: List[Dict]) -> str:
        """
        将富文本片段转换为HTML
        
        Args:
            rich_text_parts: 富文本片段列表
            
        Returns:
            HTML格式字符串
        """
        html = ""
        
        for part in rich_text_parts:
            text = part["text"]
            
            # 应用格式
            if part["is_bold"]:
                text = f"<strong>{text}</strong>"
            
            if part["is_italic"]:
                text = f"<em>{text}</em>"
            
            if part["is_superscript"]:
                text = f"<sup>{text}</sup>"
            
            if part["is_subscript"]:
                text = f"<sub>{text}</sub>"
            
            html += text
        
        return html
    
    def format_rich_text_to_markdown(self, rich_text_parts: List[Dict]) -> str:
        """
        将富文本片段转换为Markdown
        
        Args:
            rich_text_parts: 富文本片段列表
            
        Returns:
            Markdown格式字符串
        """
        md = ""
        
        for part in rich_text_parts:
            text = part["text"]
            
            # 应用格式
            if part["is_bold"]:
                text = f"**{text}**"
            
            if part["is_italic"]:
                text = f"*{text}*"
            
            if part["is_superscript"]:
                text = f"^{text}^"
            
            if part["is_subscript"]:
                text = f"~{text}~"
            
            md += text
        
        return md

