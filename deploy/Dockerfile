FROM python:3.11-slim

LABEL maintainer="TableParser Team"
LABEL description="TableParser - 轻量级智能表格解析工具"

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制代码
COPY table_parser/ ./table_parser/
COPY examples/ ./examples/

# 创建数据目录
RUN mkdir -p /app/data /app/output

# 暴露端口（HTTP模式）
EXPOSE 8765

# 设置环境变量
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app

# 默认启动MCP服务器（stdio模式）
CMD ["python", "table_parser/mcp_server.py"]

