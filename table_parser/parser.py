"""
表格解析器主控制器模块

协调各个组件完成表格解析任务
"""

import logging
from pathlib import Path
from typing import Union, Optional

from .types import ParseOptions, ParseResult, OutputFormat, ComplexityScore
from .exceptions import ParseError
from .loader import FileLoader
from .analyzer import ComplexityAnalyzer
from .converter import FormatConverter
from .utils.validation import validate_file_path, validate_output_format
from .utils.image_extractor import ImageExtractor
from .utils.xml_shape_parser import XMLShapeParser

logger = logging.getLogger(__name__)


class TableParser:
    """
    表格解析器主控制器
    
    职责：
    1. 统一解析接口
    2. 协调各组件工作
    3. 处理异常和容错
    """
    
    def __init__(self):
        """初始化解析器"""
        self.loader = FileLoader()
        self.analyzer = ComplexityAnalyzer()
        self.converter = FormatConverter()
        self.image_extractor = ImageExtractor()
        self.xml_shape_parser = XMLShapeParser()
        logger.info("TableParser 初始化完成")
    
    def parse(
        self,
        file_path: Union[str, Path, bytes],
        output_format: OutputFormat = "auto",
        **options
    ) -> ParseResult:
        """
        主解析方法
        
        Args:
            file_path: 文件路径或二进制内容
            output_format: 输出格式 (auto/markdown/html)
            **options: 其他选项
                - chunk_rows: HTML分块行数 (默认256)
                - encoding: CSV编码 (默认auto)
                - clean_illegal_chars: 清理非法字符 (默认True)
                - preserve_styles: 保留样式 (默认False)
                - include_empty_rows: 包含空行 (默认False)
                - extract_images: 是否提取图片 (默认True)
                - images_dir: 图片保存目录 (默认None，自动生成)
                
        Returns:
            ParseResult: 解析结果对象
            
        Examples:
            >>> parser = TableParser()
            >>> result = parser.parse("data.xlsx", output_format="auto")
            >>> print(result.output_format, result.complexity_score.level)
        """
        try:
            logger.info(f"开始解析任务，输出格式: {output_format}")
            
            # 验证输出格式
            output_format = validate_output_format(output_format)
            
            # 构建解析选项
            parse_options = ParseOptions(
                output_format=output_format,
                chunk_rows=options.get("chunk_rows", 256),
                encoding=options.get("encoding", None),
                clean_illegal_chars=options.get("clean_illegal_chars", True),
                preserve_styles=options.get("preserve_styles", False),
                include_empty_rows=options.get("include_empty_rows", False),
            )
            
            # 步骤1: 加载文件
            logger.info("步骤 1/6: 加载文件...")
            workbook = self.loader.load(file_path)
            
            # 步骤2: 分析复杂度（如果是auto模式）
            complexity_score: Optional[ComplexityScore] = None
            actual_format = output_format
            
            if output_format == "auto":
                logger.info("步骤 2/6: 分析复杂度...")
                complexity_score = self.analyzer.analyze(workbook)
                actual_format = complexity_score.recommended_format
                logger.info(
                    f"自动选择格式: {actual_format} "
                    f"(复杂度: {complexity_score.level}, 得分: {complexity_score.total_score:.1f})"
                )
            else:
                logger.info(f"步骤 2/6: 跳过（用户指定格式: {output_format}）")
            
            # 步骤3: 提取图片（如果启用）
            extracted_images = []
            images_count = 0
            extract_images = options.get("extract_images", True)
            
            if extract_images:
                logger.info("步骤 3/6: 提取图片...")
                
                # 确定图片输出目录
                images_dir = options.get("images_dir", None)
                if images_dir:
                    images_output_dir = Path(images_dir)
                elif isinstance(file_path, (str, Path)):
                    # 自动生成：Excel同目录下的images文件夹
                    images_output_dir = None  # 由ImageExtractor自动处理
                    source_path = Path(file_path) if isinstance(file_path, str) else file_path
                else:
                    # Base64输入，使用当前目录
                    images_output_dir = Path("./images")
                    source_path = None
                
                images_count, extracted_images = self.image_extractor.extract_images(
                    workbook,
                    output_dir=images_output_dir,
                    file_path=source_path if isinstance(file_path, (str, Path)) else None
                )
                
                if images_count > 0:
                    logger.info(f"✅ 提取了 {images_count} 张图片")
                else:
                    logger.info("📝 未检测到图片")
            else:
                logger.info("步骤 3/6: 跳过图片提取（用户禁用）")
            
            # 步骤3.5: 提取文本框/形状中的文本（XML解析）
            shapes_text = []
            if isinstance(file_path, (str, Path)):
                logger.info("步骤 3.5/6: 提取文本框/形状文本...")
                source_path = Path(file_path) if isinstance(file_path, str) else file_path
                shapes_text = self.xml_shape_parser.extract_shapes_from_excel(str(source_path))
                if shapes_text:
                    logger.info(f"✅ 提取了 {len(shapes_text)} 个形状对象的文本")
            
            # 步骤4: 格式转换
            logger.info(f"步骤 4/6: 转换为 {actual_format.upper()} 格式...")
            if actual_format == "markdown":
                content = self.converter.to_markdown(
                    workbook,
                    include_empty_rows=parse_options.include_empty_rows
                )
            else:  # html
                # 传入文件路径用于富文本解析
                excel_file_path = str(file_path) if isinstance(file_path, (str, Path)) else None
                content = self.converter.to_html(
                    workbook,
                    chunk_rows=parse_options.chunk_rows,
                    preserve_styles=parse_options.preserve_styles,
                    include_empty_rows=parse_options.include_empty_rows,
                    excel_path=excel_file_path
                )
            
            # 步骤5: 构建结果
            logger.info("步骤 5/6: 构建解析结果...")
            metadata = self.converter.get_workbook_metadata(workbook)
            
            # 添加图片信息到元数据
            if extracted_images:
                metadata["extracted_images"] = extracted_images
                metadata["images_count"] = images_count
            
            # 添加形状文本信息到元数据
            if shapes_text:
                metadata["shapes_text"] = shapes_text
                metadata["shapes_count"] = len(shapes_text)
            
            result = ParseResult(
                success=True,
                output_format=actual_format,
                content=content,
                complexity_score=complexity_score,
                metadata=metadata,
                error=None
            )
            
            logger.info(
                f"✅ 解析完成！格式: {actual_format}, "
                f"Sheet数: {metadata['sheets']}, "
                f"总行数: {metadata['total_rows']}"
            )
            
            return result
            
        except Exception as e:
            logger.error(f"❌ 解析失败: {e}")
            return ParseResult(
                success=False,
                output_format=output_format,
                content="",
                complexity_score=None,
                metadata={},
                error=str(e)
            )
    
    def analyze_only(
        self,
        file_path: Union[str, Path, bytes]
    ) -> ComplexityScore:
        """
        仅分析复杂度（不生成输出内容）
        
        Args:
            file_path: 文件路径或二进制内容
            
        Returns:
            ComplexityScore: 复杂度评分对象
            
        Raises:
            ParseError: 分析失败时抛出
            
        Examples:
            >>> parser = TableParser()
            >>> score = parser.analyze_only("data.xlsx")
            >>> print(f"复杂度: {score.level}, 推荐格式: {score.recommended_format}")
        """
        try:
            logger.info("开始复杂度分析（仅分析模式）...")
            
            # 加载文件
            workbook = self.loader.load(file_path)
            
            # 分析复杂度
            complexity_score = self.analyzer.analyze(workbook)
            
            logger.info(
                f"✅ 复杂度分析完成: {complexity_score.level} "
                f"(得分: {complexity_score.total_score:.1f})"
            )
            
            return complexity_score
            
        except Exception as e:
            raise ParseError(f"复杂度分析失败: {e}") from e
    
    def preview(
        self,
        file_path: Union[str, Path, bytes],
        max_rows: int = 10,
        max_cols: int = 10
    ) -> dict:
        """
        预览表格内容（快速返回，不完整解析）
        
        Args:
            file_path: 文件路径或二进制内容
            max_rows: 最大预览行数
            max_cols: 最大预览列数
            
        Returns:
            预览信息字典
            
        Raises:
            ParseError: 预览失败时抛出
            
        Examples:
            >>> parser = TableParser()
            >>> preview = parser.preview("data.xlsx", max_rows=5)
            >>> print(preview['sheets'][0]['preview'])
        """
        try:
            logger.info(f"开始预览表格 (max_rows={max_rows}, max_cols={max_cols})...")
            
            # 加载文件
            workbook = self.loader.load(file_path)
            
            sheets = []
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # 提取预览数据
                preview_data = []
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i >= max_rows:
                        break
                    preview_data.append(list(row[:max_cols]))
                
                sheets.append({
                    "name": sheet_name,
                    "preview": preview_data,
                    "total_rows": sheet.max_row,
                    "total_cols": sheet.max_column
                })
            
            result = {
                "success": True,
                "sheets": sheets,
                "metadata": {
                    "sheets_count": len(workbook.sheetnames)
                }
            }
            
            logger.info(f"✅ 预览完成，包含 {len(sheets)} 个sheet")
            return result
            
        except Exception as e:
            raise ParseError(f"预览失败: {e}") from e

