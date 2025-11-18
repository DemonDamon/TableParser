"""
测试修复后的合并单元格处理
"""

import sys
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

from table_parser import TableParser


def main():
    print("=" * 60)
    print("测试修复后的合并单元格处理")
    print("=" * 60)
    
    # 测试文件路径
    test_file = "/Users/damon/Desktop/品质AI智能客服/20251117材料/附件1：【PZ251201RW】AI智能客服核心能力提升研发任务需求明细表.xlsx"
    
    # 创建解析器
    parser = TableParser()
    
    print(f"\n测试文件: {Path(test_file).name}")
    print("-" * 60)
    
    try:
        # 测试1: 自动模式
        print("\n测试1: 自动模式解析...")
        result = parser.parse(test_file, output_format="auto")
        
        if result.success:
            print("✅ 解析成功！")
            print(f"   输出格式: {result.output_format}")
            if result.complexity_score:
                print(f"   复杂度等级: {result.complexity_score.level}")
                print(f"   复杂度得分: {result.complexity_score.total_score:.1f}")
                print(f"   合并单元格得分: {result.complexity_score.merged_cells_score:.1f}")
                print(f"   表头层级得分: {result.complexity_score.header_depth_score:.1f}")
            print(f"   Sheet数: {result.metadata['sheets']}")
            print(f"   总行数: {result.metadata['total_rows']}")
            print(f"   合并单元格数: {result.metadata['merged_cells_count']}")
            
            # 保存输出
            if result.output_format == "markdown":
                output_file = Path("test_output.md")
                output_file.write_text(result.content, encoding='utf-8')
                print(f"\n✅ Markdown已保存: {output_file}")
            else:  # HTML
                output_file = Path("test_output.html")
                html_content = "\n\n".join(result.content)
                output_file.write_text(html_content, encoding='utf-8')
                print(f"\n✅ HTML已保存: {output_file}")
                print(f"   包含 {len(result.content)} 个表格块")
        else:
            print(f"❌ 解析失败: {result.error}")
            return
        
        # 测试2: 强制HTML模式
        print("\n\n测试2: 强制HTML模式...")
        result_html = parser.parse(test_file, output_format="html")
        
        if result_html.success:
            print("✅ HTML模式解析成功！")
            output_file = Path("test_output_forced_html.html")
            html_content = "\n\n".join(result_html.content)
            output_file.write_text(html_content, encoding='utf-8')
            print(f"   已保存: {output_file}")
        else:
            print(f"❌ HTML模式解析失败: {result_html.error}")
        
        # 测试3: Markdown模式
        print("\n\n测试3: 强制Markdown模式...")
        result_md = parser.parse(test_file, output_format="markdown")
        
        if result_md.success:
            print("✅ Markdown模式解析成功！")
            output_file = Path("test_output_forced_md.md")
            output_file.write_text(result_md.content, encoding='utf-8')
            print(f"   已保存: {output_file}")
        else:
            print(f"❌ Markdown模式解析失败: {result_md.error}")
        
        # 测试4: 预览模式
        print("\n\n测试4: 预览模式...")
        preview = parser.preview(test_file, max_rows=5)
        
        if preview['success']:
            print("✅ 预览成功！")
            print(f"   Sheet数: {preview['metadata']['sheets_count']}")
            for sheet in preview['sheets']:
                print(f"\n   Sheet: {sheet['name']}")
                print(f"   总行数: {sheet['total_rows']}, 总列数: {sheet['total_cols']}")
                print(f"   前3行预览:")
                for i, row in enumerate(sheet['preview'][:3], 1):
                    print(f"     行{i}: {row[:5]}...")  # 只显示前5列
        
    except FileNotFoundError:
        print(f"❌ 文件不存在: {test_file}")
        print("   请检查文件路径")
    except Exception as e:
        print(f"❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)


if __name__ == "__main__":
    main()

