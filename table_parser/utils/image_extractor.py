"""
图片提取工具

从Excel文件中提取图片并保存到本地
"""

import logging
from pathlib import Path
from typing import Optional, List, Tuple
from io import BytesIO

from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

logger = logging.getLogger(__name__)


class ImageExtractor:
    """
    Excel图片提取器
    
    支持从Excel文件中提取所有图片并保存到指定目录
    """
    
    def extract_images(
        self,
        workbook: Workbook,
        output_dir: Optional[Path] = None,
        file_path: Optional[Path] = None
    ) -> Tuple[int, List[str]]:
        """
        提取工作簿中的所有图片
        
        Args:
            workbook: openpyxl Workbook对象
            output_dir: 输出目录（可选）
            file_path: 原始Excel文件路径（用于自动生成输出目录）
            
        Returns:
            (图片数量, 图片路径列表)
        """
        # 确定输出目录
        if output_dir is None and file_path is not None:
            # 默认：Excel同目录下创建images文件夹
            output_dir = file_path.parent / "images"
        elif output_dir is None:
            # 都没有提供，使用当前目录
            output_dir = Path("./images")
        
        # 创建输出目录
        output_dir.mkdir(parents=True, exist_ok=True)
        
        extracted_images = []
        total_count = 0
        
        try:
            # 遍历所有工作表
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # 检查工作表是否有图片
                if not hasattr(sheet, '_images') or not sheet._images:
                    logger.debug(f"Sheet '{sheet_name}' 无图片")
                    continue
                
                # 提取图片
                for idx, img in enumerate(sheet._images, 1):
                    try:
                        # 生成图片文件名
                        # 格式：sheet名_序号.扩展名
                        ext = self._get_image_extension(img)
                        safe_sheet_name = self._sanitize_filename(sheet_name)
                        image_filename = f"{safe_sheet_name}_{idx}.{ext}"
                        image_path = output_dir / image_filename
                        
                        # 保存图片
                        self._save_image(img, image_path)
                        
                        extracted_images.append(str(image_path))
                        total_count += 1
                        
                        logger.debug(f"✅ 保存图片: {image_path}")
                        
                    except Exception as e:
                        logger.warning(f"保存图片失败 ({sheet_name}[{idx}]): {e}")
                        continue
            
            if total_count > 0:
                logger.info(f"✅ 图片提取完成: 共 {total_count} 张，保存到 {output_dir}")
            else:
                logger.info("📝 未检测到图片")
            
            return total_count, extracted_images
            
        except Exception as e:
            logger.error(f"图片提取失败: {e}")
            return 0, []
    
    def _get_image_extension(self, img: OpenpyxlImage) -> str:
        """
        获取图片扩展名
        
        Args:
            img: openpyxl Image对象
            
        Returns:
            图片扩展名（png/jpg/jpeg/gif等）
        """
        # 尝试从图片对象获取格式
        if hasattr(img, 'format'):
            ext = img.format.lower()
            if ext in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff']:
                return ext
        
        # 尝试从ref属性获取扩展名
        if hasattr(img, 'ref'):
            ref = str(img.ref)
            if '.' in ref:
                ext = ref.split('.')[-1].lower()
                if ext in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff']:
                    return ext
        
        # 默认使用png
        return 'png'
    
    def _save_image(self, img: OpenpyxlImage, output_path: Path):
        """
        保存图片到文件
        
        Args:
            img: openpyxl Image对象
            output_path: 输出文件路径
        """
        # 获取图片数据
        if hasattr(img, '_data'):
            # 图片数据在_data属性中
            image_data = img._data()
        elif hasattr(img, 'ref'):
            # 有些版本的openpyxl使用ref属性
            # 这种情况需要从工作簿的_images字典中获取
            raise NotImplementedError("该openpyxl版本的图片提取方式暂不支持")
        else:
            raise ValueError("无法获取图片数据")
        
        # 写入文件
        output_path.write_bytes(image_data)
    
    def _sanitize_filename(self, name: str) -> str:
        """
        清理文件名，移除不安全字符
        
        Args:
            name: 原始名称
            
        Returns:
            安全的文件名
        """
        # 替换不安全字符
        unsafe_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        safe_name = name
        for char in unsafe_chars:
            safe_name = safe_name.replace(char, '_')
        
        # 限制长度
        if len(safe_name) > 50:
            safe_name = safe_name[:50]
        
        return safe_name
    
    def count_images(self, workbook: Workbook) -> int:
        """
        统计工作簿中的图片数量
        
        Args:
            workbook: openpyxl Workbook对象
            
        Returns:
            图片总数
        """
        total = 0
        try:
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                if hasattr(sheet, '_images') and sheet._images:
                    total += len(sheet._images)
        except Exception as e:
            logger.debug(f"图片统计失败: {e}")
        
        return total

