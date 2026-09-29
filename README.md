# pptx-studio

基于 **FastAPI + MCP** 的声明式 PPTX 生成服务。输入 JSON Spec 即可生成高颜值商业演示文稿，并自动返回 PPTX 下载链接与逐页高清预览图。

---

## 🚀 快速开始

### 方式一：Docker Compose（推荐，内置渲染与中文字体）

```bash
docker compose up -d --build
```
- 服务端点：`http://127.0.0.1:48000`
- MCP 端点：`http://127.0.0.1:48000/mcp`
- API 文档：`http://127.0.0.1:48000/docs`

### 方式二：本地运行

```bash
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 48000 --reload
```

---

## 🔌 MCP 客户端配置（Cherry Studio / Claude Desktop / Cursor）

在客户端 MCP 配置中添加：

```json
{
  "mcpServers": {
    "pptx-studio": {
      "type": "http",
      "url": "http://<你的服务器IP或域名>:48000/mcp"
    }
  }
}
```

---

## 🌐 Nginx 反向代理配置

若通过 Nginx 反向代理，需开启 **Streamable HTTP / SSE 流式传输** 并关闭缓冲：

```nginx
server {
    listen 80;
    listen [::]:80;
    server_name _; # 允许所有域名与 IP 访问

    client_max_body_size 50m;

    location / {
        proxy_pass http://127.0.0.1:48000;

        # 代理标头传递
        proxy_set_header Host $http_host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_set_header X-Forwarded-Host $http_host;

        # MCP 流式传输核心配置（必须关闭缓冲以防连接挂起）
        proxy_http_version 1.1;
        proxy_set_header Connection "";
        proxy_buffering off;
        proxy_cache off;
        chunked_transfer_encoding on;

        # 超时设置
        proxy_read_timeout 3600s;
        proxy_send_timeout 3600s;
    }
}
```

---

## 🛠️ 核心 MCP 工具

| 工具名 | 说明 |
| --- | --- |
| `create_presentation` | 依据 Spec 构建 PPTX，返回下载链接、逐页截图 URL 及 Markdown 图片 |
| `update_page` | 单页精准定向重构与微调，即时更新该页预览 |
| `change_theme` | 一键全局换肤（支持 9 大内置主题或自定义配色） |
| `insert_page` / `delete_page` | 动态插入或删除指定幻灯片 |
| `render_preview` | 对已有文稿重新渲染各页高清截图 |
| `get_design_guide` | 获取版式库（拓扑图、架构栈、看板、网格等）与设计指引 |
| `list_artifacts` / `read_spec` | 查看与检索已生成的产物与 Spec 结构 |

---

## 📋 环境变量

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `BASE_URL` | 自动识别 | 外部访问 Base URL（留空则自动从请求头提取） |
| `PORT` | `48000` | 服务监听端口 |
| `PPTKIT_WORKSPACE` | `/app/data` | 工作目录（挂载数据持久化） |
| `PPTKIT_OUTPUT_DIR` | `/app/data/output` | PPTX / 预览图保存路径 |
| `PPTKIT_STORE_DIR` | `/app/data/store` | 产物持久化索引路径 |
