# -*- coding: utf-8 -*-
"""FastAPI 应用：同时提供 MCP（streamable HTTP）与 REST 接口。

    uvicorn app:app --host 0.0.0.0 --port 48000

REST 端点（便于 curl / 前端直接调用）：
    GET  /                   服务与端点说明
    GET  /healthz            健康检查
    GET  /api/themes         主题列表
    GET  /api/templates      模板列表
    GET  /api/templates/{n}  模板 spec
    POST /api/validate       校验 spec
    POST /api/generate       生成 PPTX（自适应 Host，支持 preview 与密度体检）
    POST /api/artifacts/{id}/page/{n} 单页精修
    GET  /api/artifacts      产物列表
    GET  /api/artifacts/{id} 产物元信息
    GET  /api/download/{id}  下载产物（自适应 filename / id）
    GET  /api/preview/{id}   预览 PNG（?page=1）
    DELETE /api/artifacts/{id}  删除产物
    /mcp                     MCP Streamable HTTP 端点
"""
from __future__ import annotations

import contextlib
import glob
import os


from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

import mcp_server as M
from pptkit import theme as theme_mod, templates
from pptkit.deck import build_from_spec



class SpecIn(BaseModel):
    spec: dict = Field(..., description="演示文稿定义，需含 pages")
    filename: str = ""
    preview: bool = True
    store: bool = True


class PageUpdateIn(BaseModel):
    page_spec: dict = Field(..., description="更新后的单页 spec 定义（必须含 layout）")


@contextlib.asynccontextmanager
async def lifespan(app: FastAPI):
    # streamable_http_app() 会惰性建立 session_manager，必须在启动前调用
    async with M.server.session_manager.run():
        yield


app = FastAPI(
    title="pptx-studio",
    description="声明式 PPTX 生成服务：FastAPI + MCP",
    version="0.2.0",
    lifespan=lifespan,
)


class BaseUrlMiddleware:
    """纯 ASGI 中间件：提取 HTTP 客户端真实请求的协议与 Host 头，注入上下文供 URL 动态拼接。"""
    def __init__(self, inner_app):
        self.inner_app = inner_app

    async def __call__(self, scope, receive, send):
        if scope.get("type") == "http":
            headers = dict(scope.get("headers", []))
            x_proto = headers.get(b"x-forwarded-proto", b"").decode("latin-1")
            proto = x_proto or scope.get("scheme", "http")
            x_host = headers.get(b"x-forwarded-host", b"").decode("latin-1")
            host_header = headers.get(b"host", b"").decode("latin-1")
            host = x_host or host_header or ""
            if host:
                M.current_request_base_url.set(f"{proto}://{host}")
        await self.inner_app(scope, receive, send)


app.add_middleware(BaseUrlMiddleware)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# MCP 端点挂在 /mcp（Streamable HTTP）。
# 注意：streamable_http_app() 内部默认自带 path="/mcp"，若直接 mount("/mcp", ...)
# 实际端点会变成 /mcp/mcp（对外只留一个 307 跳转，客户端按文档配置会连不上）。
# 因此这里把内层 path 置为 "/"，使真实端点恰好是 /mcp。
_mcp_app = M.server.streamable_http_app(streamable_http_path="/")
app.mount("/mcp", _mcp_app)


@app.get("/")
def index():
    return {
        "service": "pptx-studio",
        "version": "0.2.0",
        "mcp_endpoint": "/mcp",
        "docs": "/docs",
        "endpoints": [
            "GET /healthz", "GET /api/themes", "GET /api/templates",
            "GET /api/templates/{name}", "POST /api/validate",
            "POST /api/generate", "POST /api/artifacts/{id}/page/{page_number}",
            "GET /api/artifacts", "GET /api/artifacts/{id}",
            "GET /api/download/{id}", "GET /api/preview/{id}",
            "DELETE /api/artifacts/{id}", "GET /api/server-info"
        ],
        "mcp_tools": [
            "get_design_guide", "list_themes", "list_templates", "get_template",
            "validate_spec", "create_presentation", "update_page", "change_theme",
            "insert_page", "delete_page", "render_preview", "list_artifacts",
            "read_spec", "delete_artifact", "server_info"
        ],
    }


@app.get("/healthz")
def healthz():
    return {
        "ok": True,
        "preview_available": M.render.available(),
        "artifacts": len(M.STORE.list(limit=10000)),
        "base_url": M.get_base_url(),
    }


@app.get("/api/themes")
def api_themes():
    return {
        "themes": [
            {
                "name": n,
                "label": p["label"],
                "mode": p.get("mode", "light"),
                "primary": p["primary"],
                "accent": p["accent"],
                "accent2": p.get("accent2", p["accent"]),
                "bg": p["bg"],
            }
            for n, p in theme_mod.PALETTES.items()
        ],
        "sizes": list(theme_mod.SLIDE_SIZES),
    }


@app.get("/api/templates")
def api_templates():
    return {"templates": templates.listing()}


@app.get("/api/templates/{name}")
def api_template(name: str):
    try:
        return templates.get(name)
    except KeyError as e:
        raise HTTPException(status_code=404, detail=str(e))


@app.post("/api/validate")
def api_validate(payload: SpecIn):
    errs = M._validate_spec(payload.spec)
    return {"ok": not errs, "errors": errs}


@app.post("/api/generate")
def api_generate(payload: SpecIn):
    errs = M._validate_spec(payload.spec)
    if errs:
        raise HTTPException(status_code=422, detail=errs)

    name = M.make_unique_filename(payload.filename, fallback="deck")
    out_path = os.path.join(M.OUTPUT_DIR, name)
    try:
        deck, path, warnings = build_from_spec(
            {**payload.spec, "output": {"path": out_path}})
    except Exception as e:
        raise HTTPException(status_code=500,
                            detail="%s: %s" % (type(e).__name__, e))
    rec = M.STORE.register(name, "pptx", path,
                           spec=payload.spec) if payload.store else {
        "id": None, "name": name, "path": path}

    base_url = M.get_base_url()
    download_url = f"{base_url}/api/download/{rec['name']}"
    alt_download_url = f"{base_url}/api/download/{rec['id']}.pptx" if rec.get("id") else download_url

    pages = payload.spec.get("pages") or []
    body = {
        "ok": True,
        "id": rec["id"],
        "file": rec["name"],
        "path": rec["path"],
        "pages": len(pages),
        "download_url": download_url,
        "alt_download_url": alt_download_url,
        "markdown_download_link": f"[点击下载演示文稿 (PPTX)]({download_url})",
        "warnings": warnings,
        "preview_urls": [],
        "slides": [],
        "user_display_markdown": "",
    }
    if payload.preview and M.render.available():
        try:
            pv = os.path.join(os.path.dirname(path), "preview",
                              M._slug(os.path.splitext(name)[0], "deck"))
            pdf, pngs = M.render.render_pptx(path, out_dir=pv)
            preview_urls = [
                f"{base_url}/api/preview/{rec['id'] or rec['name']}?page={i + 1}"
                for i in range(len(pngs))
            ]
            body["preview_urls"] = preview_urls
            slides_info = []
            for i in range(len(pngs)):
                p_idx = i + 1
                p_meta = pages[i] if i < len(pages) else {}
                slides_info.append({
                    "page": p_idx,
                    "title": p_meta.get("title", f"第 {p_idx} 页"),
                    "layout": p_meta.get("layout", "bullets"),
                    "preview_url": preview_urls[i]
                })
            body["slides"] = slides_info
            body["user_display_markdown"] = M.make_display_markdown(download_url, rec["name"], slides_info)
        except Exception as e:
            body["preview_error"] = str(e)
    elif payload.preview:
        body["preview_error"] = "未检测到 LibreOffice"
    return body


@app.post("/api/artifacts/{item_id:path}/page/{page_number}")
def api_update_page(item_id: str, page_number: int, payload: PageUpdateIn):
    res = M.update_page(item_id, page_number, payload.page_spec)
    data = M.json.loads(res[0])
    if not data.get("ok"):
        raise HTTPException(status_code=400, detail=data.get("error", "更新失败"))
    return data


@app.get("/api/artifacts")
def api_artifacts(kind: str = "", limit: int = Query(25, ge=1, le=200)):
    return {"items": M.STORE.list(kind=kind or None, limit=limit)}


@app.get("/api/artifacts/{item_id}")
def api_artifact(item_id: str):
    rec = M.STORE.get(item_id)
    if not rec:
        raise HTTPException(status_code=404, detail="产物不存在")
    return rec


@app.get("/api/download/{item_id:path}")
def api_download(item_id: str):
    try:
        path = M.STORE.path_of(item_id)
    except Exception:
        cand = os.path.join(M.OUTPUT_DIR, item_id)
        if not os.path.exists(cand) and not cand.endswith(".pptx"):
            cand += ".pptx"
        if os.path.exists(cand):
            path = cand
        else:
            raise HTTPException(status_code=404, detail=f"文件不存在: {item_id}")
    media_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if path.lower().endswith(".pptx")
        else "application/octet-stream"
    )
    return FileResponse(path, filename=os.path.basename(path), media_type=media_type)


@app.get("/api/preview/{item_id:path}")
def api_preview(item_id: str, page: int = Query(1, ge=1)):
    try:
        path = M.STORE.path_of(item_id)
    except Exception:
        cand = os.path.join(M.OUTPUT_DIR, item_id)
        if not os.path.exists(cand) and not cand.endswith(".pptx"):
            cand += ".pptx"
        if os.path.exists(cand):
            path = cand
        else:
            raise HTTPException(status_code=404, detail=f"文件不存在: {item_id}")

    if path.lower().endswith(".png"):
        return FileResponse(path, media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

    pv_dir = os.path.join(os.path.dirname(path), "preview", M._slug(os.path.splitext(os.path.basename(path))[0], "deck"))
    target_png = os.path.join(pv_dir, "page_%02d.png" % page)

    # 1. 命中已生成的高清预览截图（毫秒级极速直出）
    if os.path.exists(target_png):
        return FileResponse(target_png, media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

    # 2. 如果预览目录已有生成的页面截图，判断请求页码是否超出总页数
    existing_pngs = sorted(glob.glob(os.path.join(pv_dir, "page_*.png")))
    if existing_pngs:
        if page > len(existing_pngs):
            raise HTTPException(status_code=404, detail=f"页码超出范围（当前演示文稿共 {len(existing_pngs)} 页）")
        if os.path.exists(existing_pngs[page - 1]):
            return FileResponse(existing_pngs[page - 1], media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})

    # 3. 仅当从未生成过截图时，才调用无头转换补齐生成
    try:
        _, pngs = M.render.render_pptx(path, out_dir=pv_dir)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    if page > len(pngs):
        raise HTTPException(status_code=404, detail="页码超出范围")

    return FileResponse(pngs[page - 1], media_type="image/png", headers={"Cache-Control": "public, max-age=86400"})



@app.delete("/api/artifacts/{item_id}")
def api_delete(item_id: str):
    if not M.STORE.get(item_id):
        raise HTTPException(status_code=404, detail="产物不存在")
    M.STORE.delete(item_id)
    return {"ok": True, "deleted": item_id}


@app.get("/api/server-info")
def api_server_info():
    return JSONResponse(M.json.loads(M.server_info()))
