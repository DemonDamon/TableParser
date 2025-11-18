#!/usr/bin/env python3
"""
TableParser MCP服务器启动脚本

推荐使用此脚本启动MCP服务器
"""

import sys
import argparse
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

# 导入并运行MCP服务器
from table_parser.mcp_server import mcp, logger


def main():
    parser = argparse.ArgumentParser(
        description="TableParser MCP服务器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # stdio模式（用于Claude Desktop等）
  python start_mcp_server.py
  
  # HTTP模式（用于独立服务）
  python start_mcp_server.py --http --port 8765
        """
    )
    
    parser.add_argument(
        "--http",
        action="store_true",
        help="使用HTTP模式（默认为stdio模式）"
    )
    
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="HTTP服务器监听地址（默认: 0.0.0.0）"
    )
    
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="HTTP服务器端口（默认: 8765）"
    )
    
    args = parser.parse_args()
    
    logger.info("🚀 启动TableParser MCP服务器...")
    logger.info("=" * 60)
    
    if args.http:
        logger.info(f"模式: HTTP")
        logger.info(f"地址: http://{args.host}:{args.port}")
        logger.info("=" * 60)
        mcp.run(transport="http", host=args.host, port=args.port)
    else:
        logger.info(f"模式: stdio（标准输入输出）")
        logger.info(f"适用于: Claude Desktop, Cline等MCP客户端")
        logger.info("=" * 60)
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

