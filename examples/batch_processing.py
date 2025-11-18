"""
批量处理示例

演示如何批量处理多个Excel/CSV文件
"""

import sys
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from table_parser import TableParser


def process_single_file(parser: TableParser, file_path: Path, output_dir: Path) -> dict:
    """处理单个文件"""
    try:
        print(f"处理: {file_path.name}...")
        
        # 解析文件
        result = parser.parse(str(file_path), output_format="auto")
        
        if not result.success:
            return {
                "file": file_path.name,
                "status": "failed",
                "error": result.error
            }
        
        # 保存输出
        if result.output_format == "markdown":
            output_file = output_dir / f"{file_path.stem}.md"
            output_file.write_text(result.content, encoding='utf-8')
        else:  # HTML
            output_file = output_dir / f"{file_path.stem}.html"
            output_file.write_text("\n\n".join(result.content), encoding='utf-8')
        
        return {
            "file": file_path.name,
            "status": "success",
            "output_format": result.output_format,
            "complexity_level": result.complexity_score.level if result.complexity_score else "unknown",
            "output_file": str(output_file)
        }
        
    except Exception as e:
        return {
            "file": file_path.name,
            "status": "failed",
            "error": str(e)
        }


def main():
    print("=" * 60)
    print("批量处理示例")
    print("=" * 60)
    
    # 配置
    input_dir = Path("./data")  # 输入目录
    output_dir = Path("./output")  # 输出目录
    max_workers = 4  # 并发数
    
    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 创建解析器
    parser = TableParser()
    
    # 查找所有Excel和CSV文件
    patterns = ["*.xlsx", "*.xls", "*.csv"]
    files = []
    for pattern in patterns:
        files.extend(input_dir.glob(pattern))
    
    if not files:
        print(f"\n⚠️  在 {input_dir} 目录下没有找到Excel或CSV文件")
        print("   请将文件放入该目录后重试")
        return
    
    print(f"\n找到 {len(files)} 个文件:")
    for f in files:
        print(f"  - {f.name}")
    
    print(f"\n开始批量处理（并发数: {max_workers}）...")
    print("-" * 60)
    
    # 方法1: 串行处理（简单）
    if len(files) <= 3:
        results = []
        for file_path in files:
            result = process_single_file(parser, file_path, output_dir)
            results.append(result)
            
            # 打印进度
            if result["status"] == "success":
                print(f"  ✅ {result['file']} -> {result['output_format']}")
            else:
                print(f"  ❌ {result['file']}: {result['error']}")
    
    # 方法2: 并行处理（快速）
    else:
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有任务
            futures = {
                executor.submit(process_single_file, parser, fp, output_dir): fp
                for fp in files
            }
            
            # 收集结果
            for future in as_completed(futures):
                result = future.result()
                results.append(result)
                
                # 打印进度
                if result["status"] == "success":
                    print(f"  ✅ {result['file']} -> {result['output_format']}")
                else:
                    print(f"  ❌ {result['file']}: {result['error']}")
    
    # 统计结果
    print("\n" + "-" * 60)
    print("处理结果统计:")
    
    succeeded = sum(1 for r in results if r["status"] == "success")
    failed = sum(1 for r in results if r["status"] == "failed")
    
    print(f"  总计: {len(results)} 个文件")
    print(f"  成功: {succeeded} 个")
    print(f"  失败: {failed} 个")
    
    # 按格式统计
    if succeeded > 0:
        markdown_count = sum(
            1 for r in results
            if r["status"] == "success" and r["output_format"] == "markdown"
        )
        html_count = sum(
            1 for r in results
            if r["status"] == "success" and r["output_format"] == "html"
        )
        
        print(f"\n输出格式分布:")
        print(f"  Markdown: {markdown_count} 个")
        print(f"  HTML: {html_count} 个")
    
    # 按复杂度统计
    if succeeded > 0:
        print(f"\n复杂度分布:")
        for level in ["simple", "medium", "complex"]:
            count = sum(
                1 for r in results
                if r["status"] == "success" and r.get("complexity_level") == level
            )
            if count > 0:
                print(f"  {level}: {count} 个")
    
    # 失败文件列表
    if failed > 0:
        print(f"\n失败文件:")
        for r in results:
            if r["status"] == "failed":
                print(f"  - {r['file']}: {r['error']}")
    
    print("\n" + "=" * 60)
    print(f"所有文件已保存到: {output_dir.absolute()}")
    print("=" * 60)


if __name__ == "__main__":
    main()

