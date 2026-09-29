# -*- coding: utf-8 -*-
"""pptk —— 轻量 PPTX 生成工具包。

设计源自历史脚本 build_final.py / build_slide.py 的复用与抽象，
把「一页式成效看板」的画法沉淀为可组合的版式库。
"""
__version__ = "0.1.0"

from .deck import Deck, build_from_spec          # noqa: F401
from . import theme, layouts, store, render      # noqa: F401

__all__ = ["Deck", "build_from_spec", "theme", "layouts", "store", "render"]
