# -*- coding: utf-8 -*-
"""PPTX 预览渲染：LibreOffice 转 PDF，再用 PyMuPDF 栅格化为 PNG。

容器/Linux 环境通常没有微软雅黑，LibreOffice 会做字体回退，
因此预览图与 Windows 上的最终效果存在字体差异（版式位置基本一致）。
"""
import glob
import os
import shutil
import subprocess
import tempfile
import threading

# 并发转换信号量：控制最多并发 2 个无头 LibreOffice 进程，避免多客户端并发打满 CPU 和内存颠簸
_RENDER_SEMAPHORE = threading.Semaphore(int(os.environ.get("MAX_CONCURRENT_RENDERS", "2")))

SOFFICE_CANDIDATES = ("soffice", "libreoffice")


def soffice_path():
    for cand in SOFFICE_CANDIDATES:
        p = shutil.which(cand)
        if p:
            return p
    for p in ("/usr/bin/soffice", "/usr/bin/libreoffice",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def available():
    return soffice_path() is not None


def pptx_to_pdf(pptx_path, out_dir=None, timeout=180):
    """把 pptx 转成 pdf，返回 pdf 路径。支持多客户端并发隔离与信号量队列排队。"""
    exe = soffice_path()
    if not exe:
        raise RuntimeError("未找到 LibreOffice（soffice），无法生成预览。"
                           "请安装 libreoffice 或关闭 preview。")
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir or os.path.dirname(pptx_path))
    os.makedirs(out_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    target_pdf = os.path.join(out_dir, base_name + ".pdf")

    # 独立的 UserInstallation 与独立的临时输出目录，彻底杜绝并发调用时的输出文件覆盖与锁冲突
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    tmp_out = tempfile.mkdtemp(prefix="lo_out_")
    try:
        cmd = [exe, "-env:UserInstallation=file://%s" % profile,
               "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", tmp_out, pptx_path]
        
        # 通过信号量进行并发限流，超额客户端平滑排队
        with _RENDER_SEMAPHORE:
            proc = subprocess.run(cmd, stdout=subprocess.PIPE,
                                  stderr=subprocess.STDOUT, timeout=timeout)
        
        # 在独立的 tmp_out 中定位生成的 pdf
        candidates = glob.glob(os.path.join(tmp_out, "*.pdf"))
        if not candidates:
            raise RuntimeError("LibreOffice 转换失败: %s"
                               % proc.stdout.decode("utf-8", "ignore")[-500:])
        
        # 移至最终目标路径
        shutil.move(candidates[0], target_pdf)
        return target_pdf
    finally:
        shutil.rmtree(profile, ignore_errors=True)
        shutil.rmtree(tmp_out, ignore_errors=True)


def pdf_to_pngs(pdf_path, out_dir=None, dpi=110, pages=None):
    """PDF 逐页转 PNG，返回 png 路径列表。"""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz  # PyMuPDF 回退

    out_dir = os.path.abspath(out_dir or os.path.dirname(pdf_path))
    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    out = []
    try:
        indices = range(doc.page_count) if pages is None else pages
        for i in indices:
            if i < 0 or i >= doc.page_count:
                continue
            pm = doc[i].get_pixmap(dpi=dpi)
            png = os.path.join(out_dir, "page_%02d.png" % (i + 1))
            pm.save(png)
            out.append(png)
    finally:
        doc.close()
    return out


def render_pptx(pptx_path, out_dir=None, dpi=110, pages=None):
    """一步到位：pptx -> [png]。"""
    out_dir = os.path.abspath(out_dir or (os.path.dirname(pptx_path) + "/preview"))
    pdf = pptx_to_pdf(pptx_path, out_dir=out_dir)
    pngs = pdf_to_pngs(pdf, out_dir=out_dir, dpi=dpi, pages=pages)
    return pdf, pngs

