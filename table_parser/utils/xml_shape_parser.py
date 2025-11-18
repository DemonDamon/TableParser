"""
XML形状解析器

直接解析Excel的XML结构，提取文本框和形状中的文本
"""

import logging
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Optional

logger = logging.getLogger(__name__)


class XMLShapeParser:
    """
    Excel XML形状解析器
    
    通过解析Excel的drawing.xml文件提取形状对象中的文本
    """
    
    # XML命名空间
    NAMESPACES = {
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    def extract_shapes_from_excel(self, excel_path: str) -> List[Dict]:
        """
        从Excel文件中提取所有形状对象的文本
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            形状文本列表
        """
        shapes_text = []
        
        try:
            # Excel文件本质是zip
            with zipfile.ZipFile(excel_path, 'r') as zip_ref:
                # 列出所有文件
                file_list = zip_ref.namelist()
                
                # 查找drawing文件（xl/drawings/drawing*.xml）
                drawing_files = [f for f in file_list if f.startswith('xl/drawings/drawing') and f.endswith('.xml')]
                
                logger.info(f"找到 {len(drawing_files)} 个绘图文件")
                
                for drawing_file in drawing_files:
                    shapes = self._parse_drawing_xml(zip_ref, drawing_file)
                    shapes_text.extend(shapes)
            
            logger.info(f"✅ 提取了 {len(shapes_text)} 个形状对象的文本")
            
        except Exception as e:
            logger.warning(f"XML形状解析失败: {e}")
        
        return shapes_text
    
    def _parse_drawing_xml(self, zip_ref: zipfile.ZipFile, drawing_file: str) -> List[Dict]:
        """
        解析单个drawing.xml文件
        
        Args:
            zip_ref: ZIP文件引用
            drawing_file: drawing文件路径
            
        Returns:
            形状文本列表
        """
        shapes = []
        
        try:
            # 读取XML内容
            xml_content = zip_ref.read(drawing_file)
            root = ET.fromstring(xml_content)
            
            # 查找所有文本框（txSp）
            for txSp in root.findall('.//xdr:sp', self.NAMESPACES):
                text = self._extract_text_from_shape(txSp)
                if text:
                    shapes.append({
                        'type': 'textbox',
                        'text': text,
                        'source': drawing_file
                    })
            
            # 查找其他形状
            for sp in root.findall('.//xdr:sp', self.NAMESPACES):
                text = self._extract_text_from_shape(sp)
                if text and not any(s['text'] == text for s in shapes):  # 去重
                    shapes.append({
                        'type': 'shape',
                        'text': text,
                        'source': drawing_file
                    })
            
        except Exception as e:
            logger.debug(f"解析{drawing_file}失败: {e}")
        
        return shapes
    
    def _extract_text_from_shape(self, shape_element: ET.Element) -> Optional[str]:
        """
        从形状元素中提取文本
        
        Args:
            shape_element: 形状XML元素
            
        Returns:
            文本内容
        """
        texts = []
        
        try:
            # 查找所有文本元素（a:t）
            for t_elem in shape_element.findall('.//a:t', self.NAMESPACES):
                if t_elem.text:
                    texts.append(t_elem.text)
            
            # 合并文本
            if texts:
                return ''.join(texts)
                
        except Exception as e:
            logger.debug(f"文本提取失败: {e}")
        
        return None

