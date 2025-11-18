"""
富文本XML解析器

直接解析Excel的sharedStrings.xml，提取富文本格式（上下标等）
"""

import logging
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)


class RichTextParser:
    """
    Excel富文本XML解析器
    
    直接解析sharedStrings.xml，提取上下标等富文本格式
    """
    
    NAMESPACES = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    }
    
    def __init__(self):
        """初始化解析器"""
        self.shared_strings_cache = {}
    
    def parse_shared_strings(self, excel_path: str) -> Dict[int, List[Tuple[str, Optional[str]]]]:
        """
        解析sharedStrings.xml，提取所有富文本
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            字典：{索引: [(文本, 格式), ...]}
            格式包括：'superscript'、'subscript'、None
        """
        if excel_path in self.shared_strings_cache:
            return self.shared_strings_cache[excel_path]
        
        shared_strings = {}
        
        try:
            with zipfile.ZipFile(excel_path, 'r') as zip_ref:
                # 读取sharedStrings.xml
                if 'xl/sharedStrings.xml' not in zip_ref.namelist():
                    logger.debug("文件中无sharedStrings.xml")
                    return {}
                
                xml_content = zip_ref.read('xl/sharedStrings.xml')
                root = ET.fromstring(xml_content)
                
                # 遍历所有字符串项
                for idx, si in enumerate(root.findall('main:si', self.NAMESPACES)):
                    parts = []
                    
                    # 检查是否有富文本runs
                    runs = si.findall('main:r', self.NAMESPACES)
                    
                    if runs:
                        # 富文本格式
                        for r in runs:
                            text_elem = r.find('main:t', self.NAMESPACES)
                            text = text_elem.text if text_elem is not None and text_elem.text else ''
                            
                            # 提取格式（上下标）
                            vert_align = r.find('main:rPr/main:vertAlign', self.NAMESPACES)
                            va_val = vert_align.get('val') if vert_align is not None else None
                            
                            parts.append((text, va_val))
                    else:
                        # 普通文本
                        text_elem = si.find('main:t', self.NAMESPACES)
                        text = text_elem.text if text_elem is not None and text_elem.text else ''
                        parts.append((text, None))
                    
                    shared_strings[idx] = parts
                
                logger.info(f"✅ 解析了 {len(shared_strings)} 个共享字符串")
                
        except Exception as e:
            logger.warning(f"sharedStrings.xml 解析失败: {e}")
        
        # 缓存结果
        self.shared_strings_cache[excel_path] = shared_strings
        return shared_strings
    
    def get_cell_rich_text(
        self, 
        excel_path: str, 
        cell_string_index: int
    ) -> List[Tuple[str, Optional[str]]]:
        """
        获取单元格的富文本格式
        
        Args:
            excel_path: Excel文件路径
            cell_string_index: 单元格的字符串索引
            
        Returns:
            [(文本, 格式), ...]
        """
        shared_strings = self.parse_shared_strings(excel_path)
        return shared_strings.get(cell_string_index, [])
    
    def format_rich_text_to_html(self, parts: List[Tuple[str, Optional[str]]]) -> str:
        """
        将富文本格式转换为HTML
        
        Args:
            parts: [(文本, 格式), ...]
            
        Returns:
            HTML字符串
        """
        from html import escape
        
        html = ""
        for text, vert_align in parts:
            escaped_text = escape(text)
            
            if vert_align == 'superscript':
                html += f"<sup>{escaped_text}</sup>"
            elif vert_align == 'subscript':
                html += f"<sub>{escaped_text}</sub>"
            else:
                html += escaped_text
        
        return html
    
    def get_cell_string_index_mapping(self, excel_path: str) -> Dict[str, int]:
        """
        获取单元格坐标到字符串索引的映射
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            字典：{'A1': 0, 'B2': 1, ...}
        """
        mapping = {}
        
        try:
            with zipfile.ZipFile(excel_path, 'r') as zip_ref:
                xml_content = zip_ref.read('xl/worksheets/sheet1.xml')
                root = ET.fromstring(xml_content)
                
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                for c in root.findall('.//main:c', ns):
                    coord = c.get('r')
                    cell_type = c.get('t')
                    
                    # 只处理字符串类型（t='s'）
                    if cell_type == 's':
                        v_elem = c.find('main:v', ns)
                        if v_elem is not None and v_elem.text:
                            string_idx = int(v_elem.text)
                            mapping[coord] = string_idx
                
                logger.debug(f"获取了 {len(mapping)} 个单元格的字符串索引映射")
                
        except Exception as e:
            logger.warning(f"字符串索引映射获取失败: {e}")
        
        return mapping

