#!/usr/bin/env python3
"""
TableParser MCP服务器启动脚本

推荐使用此脚本启动MCP服务器，而不是直接运行table_parser/mcp_server.py
"""

import sys
from pathlib import Path

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent))

# 导入并运行MCP服务器
from table_parser.mcp_server import mcp, logger

if __name__ == "__main__":
    logger.info("🚀 启动TableParser MCP服务器...")
    logger.info("=" * 60)
    logger.info("使用方式:")
    logger.info("  - stdio模式（推荐，用于Claude等）: 直接运行本脚本")
    logger.info("  - HTTP模式: 修改代码使用 mcp.run(transport='http', port=8765)")
    logger.info("=" * 60)
    
    # 启动服务器（stdio模式）
    mcp.run(transport="stdio")
    
    # 如果需要HTTP模式，注释上面一行，取消下面注释：
    # mcp.run(transport="http", host="0.0.0.0", port=8765)

