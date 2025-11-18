#!/bin/bash
# TableParser HTTP测试快速启动脚本

echo "🚀 TableParser HTTP测试"
echo "=" | tr -d '\n' | xargs -I {} printf '%60s\n' | tr ' ' '='
echo ""

# 检查MCP服务器是否运行
if ! curl -s http://localhost:8765 > /dev/null 2>&1; then
    echo "⚠️  MCP服务器未运行"
    echo "请在另一个终端启动服务器:"
    echo "  python start_mcp_server.py --http --port 8765"
    echo ""
    read -p "按Enter继续启动服务器，或Ctrl+C取消..."
    
    # 启动服务器（后台）
    cd ..
    python start_mcp_server.py --http --port 8765 &
    SERVER_PID=$!
    echo "✅ MCP服务器已启动 (PID: $SERVER_PID)"
    echo "⏳ 等待服务器启动..."
    sleep 3
    cd tests
else
    echo "✅ MCP服务器已运行"
fi

echo ""
echo "🧪 运行测试..."
echo "=" | tr -d '\n' | xargs -I {} printf '%60s\n' | tr ' ' '='
echo ""

# 运行测试
python test_mcp_client.py

TEST_EXIT_CODE=$?

echo ""
echo "=" | tr -d '\n' | xargs -I {} printf '%60s\n' | tr ' ' '='

if [ $TEST_EXIT_CODE -eq 0 ]; then
    echo "✅ 所有测试通过！"
else
    echo "❌ 部分测试失败"
fi

# 如果是我们启动的服务器，询问是否关闭
if [ ! -z "$SERVER_PID" ]; then
    echo ""
    read -p "是否关闭MCP服务器? (y/N): " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        kill $SERVER_PID
        echo "✅ MCP服务器已关闭"
    else
        echo "ℹ️  MCP服务器仍在运行 (PID: $SERVER_PID)"
        echo "   手动关闭: kill $SERVER_PID"
    fi
fi

exit $TEST_EXIT_CODE

