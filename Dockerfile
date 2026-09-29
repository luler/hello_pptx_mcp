FROM python:3.11-slim

# 设置环境变量：避免交互提示、Python 无缓冲、默认端口和数据路径
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=48000 \
    PPTKIT_WORKSPACE=/app/data \
    PPTKIT_OUTPUT_DIR=/app/data/output \
    PPTKIT_STORE_DIR=/app/data/store

# 安装系统依赖：
# 1. libreoffice-nogui: 用于无头 PPTX 转 PDF
# 2. fonts-noto-cjk, fonts-wqy-zenhei, fonts-wqy-microhei: 中文字体，确保渲染预览排版一致不乱码
# 3. curl: 容器健康检查
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-nogui \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    curl \
    && fc-cache -f \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先拷贝依赖文件并安装，利用 Docker 构建缓存
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 拷贝应用源码
COPY . .

# 创建持久化数据目录
RUN mkdir -p /app/data/output /app/data/store

EXPOSE 48000

# 健康检查
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:${PORT:-48000}/healthz || exit 1

# 启动服务（FastAPI + MCP）
CMD ["sh", "-c", "exec uvicorn app:app --host 0.0.0.0 --port ${PORT:-48000}"]
