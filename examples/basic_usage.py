"""
基本使用示例

演示TableParser的基本功能
"""

import sys
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from table_parser import TableParser


def main():
    # 创建解析器
    parser = TableParser()
    
    print("=" * 60)
    print("TableParser 基本使用示例")
    print("=" * 60)
    
    # 示例1: 自动模式（推荐）
    print("\n示例1: 自动模式解析")
    print("-" * 60)
    
    # 注意：这里需要一个真实的Excel文件
    # 如果没有，会报错，这是正常的
    test_file = "test_data.xlsx"
    
    try:
        result = parser.parse(test_file, output_format="auto")
        
        if result.success:
            print(f"✅ 解析成功！")
            print(f"   输出格式: {result.output_format}")
            print(f"   复杂度等级: {result.complexity_score.level}")
            print(f"   复杂度得分: {result.complexity_score.total_score:.1f}")
            print(f"   Sheet数: {result.metadata['sheets']}")
            print(f"   总行数: {result.metadata['total_rows']}")
            print("\n内容预览（前500字符）:")
            if isinstance(result.content, str):
                print(result.content[:500])
            else:
                print(result.content[0][:500])
        else:
            print(f"❌ 解析失败: {result.error}")
    
    except FileNotFoundError:
        print(f"⚠️  文件不存在: {test_file}")
        print("   请创建一个测试Excel文件或修改test_file变量")
    
    # 示例2: 仅分析复杂度
    print("\n\n示例2: 仅分析复杂度（不生成输出）")
    print("-" * 60)
    
    try:
        score = parser.analyze_only(test_file)
        
        print(f"复杂度分析结果:")
        print(f"  总分: {score.total_score:.1f}")
        print(f"  等级: {score.level}")
        print(f"  推荐格式: {score.recommended_format}")
        print(f"\n各维度得分:")
        print(f"  合并单元格: {score.merged_cells_score:.1f}")
        print(f"  表头层级: {score.header_depth_score:.1f}")
        print(f"  数据结构: {score.data_structure_score:.1f}")
        print(f"  表格规模: {score.scale_score:.1f}")
        
    except FileNotFoundError:
        print(f"⚠️  文件不存在: {test_file}")
    
    # 示例3: 快速预览
    print("\n\n示例3: 快速预览（前5行）")
    print("-" * 60)
    
    try:
        preview = parser.preview(test_file, max_rows=5, max_cols=10)
        
        if preview['success']:
            print(f"文件包含 {preview['metadata']['sheets_count']} 个sheet\n")
            
            for sheet in preview['sheets']:
                print(f"Sheet: {sheet['name']}")
                print(f"  总行数: {sheet['total_rows']}")
                print(f"  总列数: {sheet['total_cols']}")
                print(f"  预览数据:")
                for i, row in enumerate(sheet['preview'][:3], 1):
                    print(f"    行{i}: {row}")
                print()
    
    except FileNotFoundError:
        print(f"⚠️  文件不存在: {test_file}")
    
    # 示例4: 强制指定格式
    print("\n\n示例4: 强制使用Markdown格式")
    print("-" * 60)
    
    try:
        result = parser.parse(test_file, output_format="markdown")
        
        if result.success:
            print(f"✅ 强制Markdown格式解析成功")
            print(f"   (即使表格很复杂也会使用Markdown)")
            print("\n内容预览（前300字符）:")
            print(result.content[:300])
    
    except FileNotFoundError:
        print(f"⚠️  文件不存在: {test_file}")
    
    print("\n" + "=" * 60)
    print("示例结束")
    print("=" * 60)


if __name__ == "__main__":
    main()

