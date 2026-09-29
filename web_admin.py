# -*- coding: utf-8 -*-
"""现代化 PPTX 资产在线管理后台单页应用 (SPA)。

无需任何外部 CDN 依赖，纯内置原生 HTML5 + CSS3 + Vanilla JS，
支持：
- 严格按时间倒序展示已生成 PPT 列表；
- 批量勾选与一键批量删除（含级联预览图清理与二次确认）；
- 全景高清大图预览灯箱（键盘左右键翻页、微缩底栏点选跳转）；
- AUTH_KEY 访问凭证安全验证与本地持久化。
"""

def get_admin_html() -> str:
    return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <title>PPTX Studio · 演示文稿资产管理中心</title>
  <style>
    :root {
      --bg: #0b0f19;
      --surface: #111827;
      --surface-subtle: #1f2937;
      --surface-hover: #283548;
      --border: #374151;
      --border-focus: #3b82f6;
      --primary: #2563eb;
      --primary-hover: #1d4ed8;
      --primary-light: #60a5fa;
      --accent: #0ea5e9;
      --danger: #ef4444;
      --danger-hover: #dc2626;
      --text: #f9fafb;
      --text-muted: #9ca3af;
      --text-dim: #6b7280;
      --badge-bg: #1e293b;
      --radius: 10px;
      --shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
      background-color: var(--bg);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }
    /* 顶部导航 */
    header {
      background: rgba(17, 24, 39, 0.85);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--border);
      position: sticky;
      top: 0;
      z-index: 40;
      padding: 0.85rem 1.75rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 0.75rem;
      text-decoration: none;
      color: inherit;
    }
    .brand-icon {
      width: 2.2rem;
      height: 2.2rem;
      background: linear-gradient(135deg, #2563eb, #0ea5e9);
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
      font-size: 1.15rem;
      color: #fff;
      box-shadow: 0 0 12px rgba(37, 99, 235, 0.4);
    }
    .brand-title {
      font-size: 1.15rem;
      font-weight: 700;
      letter-spacing: -0.02em;
    }
    .brand-subtitle {
      font-size: 0.75rem;
      color: var(--text-muted);
    }
    .header-actions {
      display: flex;
      align-items: center;
      gap: 0.85rem;
    }
    .search-box {
      position: relative;
    }
    .search-input {
      background: var(--surface-subtle);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 0.45rem 0.85rem 0.45rem 2.2rem;
      color: var(--text);
      font-size: 0.85rem;
      width: 220px;
      transition: all 0.2s;
    }
    .search-input:focus {
      outline: none;
      border-color: var(--border-focus);
      width: 280px;
      background: var(--surface);
    }
    .search-icon {
      position: absolute;
      left: 0.75rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-dim);
      font-size: 0.85rem;
      pointer-events: none;
    }
    .btn {
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      padding: 0.45rem 0.95rem;
      font-size: 0.85rem;
      font-weight: 500;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: var(--surface-subtle);
      color: var(--text);
      cursor: pointer;
      transition: all 0.15s ease;
      text-decoration: none;
    }
    .btn:hover {
      background: var(--surface-hover);
      border-color: #4b5563;
    }
    .btn-primary {
      background: var(--primary);
      border-color: var(--primary);
      color: #fff;
    }
    .btn-primary:hover {
      background: var(--primary-hover);
      border-color: var(--primary-hover);
    }
    .btn-danger {
      background: rgba(239, 68, 68, 0.15);
      border-color: rgba(239, 68, 68, 0.35);
      color: #f87171;
    }
    .btn-danger:hover {
      background: var(--danger);
      color: #fff;
      border-color: var(--danger);
    }
    .btn-sm {
      padding: 0.3rem 0.65rem;
      font-size: 0.78rem;
    }

    /* 容器主体 */
    main {
      flex: 1;
      max-width: 1400px;
      width: 100%;
      margin: 0 auto;
      padding: 1.5rem 1.75rem;
    }

    /* 统计大盘与工具栏 */
    .dashboard-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 1rem;
      margin-bottom: 1.25rem;
    }
    .stats-group {
      display: flex;
      align-items: center;
      gap: 1.25rem;
    }
    .stat-chip {
      background: var(--surface);
      border: 1px solid var(--border);
      padding: 0.35rem 0.85rem;
      border-radius: 20px;
      font-size: 0.82rem;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }
    .stat-chip b {
      color: var(--primary-light);
      font-size: 0.95rem;
    }

    /* 批量操作悬浮条 */
    .batch-bar {
      background: linear-gradient(90deg, #1e3a8a, #0c4a6e);
      border: 1px solid #3b82f6;
      border-radius: var(--radius);
      padding: 0.65rem 1.25rem;
      margin-bottom: 1.25rem;
      display: none;
      align-items: center;
      justify-content: space-between;
      animation: slideDown 0.2s ease-out;
    }
    .batch-bar.active { display: flex; }
    @keyframes slideDown {
      from { opacity: 0; transform: translateY(-8px); }
      to { opacity: 1; transform: translateY(0); }
    }

    /* 文件卡片表格 */
    .table-container {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      overflow: hidden;
      box-shadow: var(--shadow);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.88rem;
    }
    thead {
      background: var(--surface-subtle);
      border-bottom: 1px solid var(--border);
      color: var(--text-muted);
      font-size: 0.8rem;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }
    th, td {
      padding: 0.85rem 1rem;
      vertical-align: middle;
    }
    tbody tr {
      border-bottom: 1px solid rgba(55, 65, 81, 0.4);
      transition: background 0.15s ease;
    }
    tbody tr:hover {
      background: var(--surface-hover);
    }
    tbody tr.selected {
      background: rgba(37, 99, 235, 0.12);
    }

    /* 缩略图 */
    .thumb-box {
      width: 110px;
      height: 62px;
      border-radius: 6px;
      background: #000;
      border: 1px solid var(--border);
      overflow: hidden;
      cursor: pointer;
      position: relative;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .thumb-img {
      width: 100%;
      height: 100%;
      object-fit: cover;
      transition: transform 0.2s;
    }
    .thumb-box:hover .thumb-img {
      transform: scale(1.08);
    }
    .thumb-hover-overlay {
      position: absolute;
      inset: 0;
      background: rgba(0, 0, 0, 0.4);
      display: flex;
      align-items: center;
      justify-content: center;
      opacity: 0;
      transition: opacity 0.2s;
      color: #fff;
      font-size: 0.75rem;
      font-weight: 500;
    }
    .thumb-box:hover .thumb-hover-overlay { opacity: 1; }

    /* 文件名与详情 */
    .file-meta-col {
      display: flex;
      flex-direction: column;
      gap: 0.25rem;
    }
    .file-title {
      font-weight: 600;
      color: var(--text);
      font-size: 0.92rem;
      cursor: pointer;
    }
    .file-title:hover {
      color: var(--primary-light);
    }
    .file-name {
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      font-size: 0.75rem;
      color: var(--text-dim);
    }

    /* 徽章 */
    .badge {
      display: inline-flex;
      align-items: center;
      padding: 0.2rem 0.55rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 500;
      background: var(--badge-bg);
      color: var(--text-muted);
      border: 1px solid var(--border);
    }
    .badge-primary {
      background: rgba(37, 99, 235, 0.2);
      color: #93c5fd;
      border-color: rgba(59, 130, 246, 0.35);
    }

    /* 操作列 */
    .action-cell {
      white-space: nowrap;
      display: flex;
      gap: 0.4rem;
      align-items: center;
    }

    /* 空态提示 */
    .empty-state {
      padding: 4rem 2rem;
      text-align: center;
      color: var(--text-muted);
    }
    .empty-icon {
      font-size: 3rem;
      margin-bottom: 0.75rem;
      opacity: 0.6;
    }

    /* 灯箱大图预览 Modal */
    .modal-backdrop {
      position: fixed;
      inset: 0;
      background: rgba(3, 7, 18, 0.88);
      backdrop-filter: blur(8px);
      z-index: 100;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 1.5rem;
    }
    .modal-backdrop.active { display: flex; }
    .lightbox-content {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      width: 95%;
      max-width: 1200px;
      max-height: 92vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7);
    }
    .lightbox-header {
      padding: 0.85rem 1.25rem;
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: var(--surface-subtle);
    }
    .lightbox-body {
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
      background: #020617;
      min-height: 480px;
      padding: 1rem;
    }
    .lightbox-img {
      max-width: 100%;
      max-height: 68vh;
      object-fit: contain;
      border-radius: 6px;
      box-shadow: 0 8px 30px rgba(0, 0, 0, 0.6);
      transition: opacity 0.2s;
    }
    .nav-arrow {
      position: absolute;
      top: 50%;
      transform: translateY(-50%);
      width: 44px;
      height: 44px;
      background: rgba(17, 24, 39, 0.75);
      border: 1px solid var(--border);
      color: #fff;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.25rem;
      cursor: pointer;
      transition: all 0.15s;
    }
    .nav-arrow:hover {
      background: var(--primary);
      border-color: var(--primary);
    }
    .nav-prev { left: 1.25rem; }
    .nav-next { right: 1.25rem; }
    .lightbox-thumbs {
      padding: 0.75rem 1.25rem;
      border-top: 1px solid var(--border);
      display: flex;
      gap: 0.65rem;
      overflow-x: auto;
      background: var(--surface-subtle);
    }
    .lightbox-thumb-item {
      width: 80px;
      height: 45px;
      border-radius: 4px;
      border: 2px solid transparent;
      overflow: hidden;
      cursor: pointer;
      flex-shrink: 0;
      opacity: 0.65;
      transition: all 0.15s;
    }
    .lightbox-thumb-item:hover { opacity: 0.9; }
    .lightbox-thumb-item.active {
      border-color: var(--primary-light);
      opacity: 1;
      box-shadow: 0 0 10px rgba(59, 130, 246, 0.5);
    }
    .lightbox-thumb-item img {
      width: 100%;
      height: 100%;
      object-fit: cover;
    }

    /* 认证弹窗 */
    .auth-modal {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 14px;
      width: 100%;
      max-width: 440px;
      padding: 2rem;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.8);
      text-align: center;
    }
    .auth-icon {
      width: 3.2rem;
      height: 3.2rem;
      margin: 0 auto 1.25rem auto;
      background: rgba(37, 99, 235, 0.15);
      border: 1px solid rgba(59, 130, 246, 0.35);
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.5rem;
      color: var(--primary-light);
    }
    .auth-input {
      width: 100%;
      padding: 0.75rem 1rem;
      background: var(--surface-subtle);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: #fff;
      font-size: 0.95rem;
      margin: 1.25rem 0;
      outline: none;
    }
    .auth-input:focus {
      border-color: var(--border-focus);
    }
    .auth-error {
      color: #f87171;
      font-size: 0.82rem;
      margin-top: -0.5rem;
      margin-bottom: 1rem;
      display: none;
    }

    /* Toast 通知 */
    .toast {
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: var(--surface-subtle);
      border: 1px solid var(--border);
      color: #fff;
      padding: 0.75rem 1.25rem;
      border-radius: 8px;
      box-shadow: var(--shadow);
      z-index: 200;
      display: none;
      align-items: center;
      gap: 0.5rem;
      font-size: 0.88rem;
    }
    /* 分页导航栏 */
    .pagination-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 1rem;
      padding: 0.85rem 1.25rem;
      background: var(--surface-subtle);
      border-top: 1px solid var(--border);
      font-size: 0.85rem;
    }
    .pagination-info {
      color: var(--text-muted);
      font-size: 0.82rem;
    }
    .pagination-info b {
      color: var(--text);
    }
    .pagination-controls {
      display: flex;
      align-items: center;
      gap: 0.45rem;
      flex-wrap: wrap;
    }
    .page-size-select {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 6px;
      color: var(--text);
      padding: 0.25rem 0.5rem;
      font-size: 0.82rem;
      cursor: pointer;
      outline: none;
    }
    .page-size-select:focus {
      border-color: var(--border-focus);
    }
    .pagination-numbers {
      display: flex;
      align-items: center;
      gap: 0.25rem;
    }
    .page-num-btn {
      min-width: 30px;
      height: 28px;
      padding: 0 0.35rem;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border-radius: 6px;
      border: 1px solid var(--border);
      background: var(--surface);
      color: var(--text);
      cursor: pointer;
      font-size: 0.8rem;
      transition: all 0.15s ease;
    }
    .page-num-btn:hover {
      background: var(--surface-hover);
      border-color: #4b5563;
    }
    .page-num-btn.active {
      background: var(--primary);
      border-color: var(--primary);
      color: #fff;
      font-weight: bold;
    }
    .page-num-btn:disabled, .btn:disabled {
      opacity: 0.35;
      cursor: not-allowed;
      pointer-events: none;
    }

    /* 移动端标签行：默认在桌面端隐藏 */
    .mobile-tags-row {
      display: none;
    }

    /* ==================== 移动端适配 (Responsive Mobile CSS) ==================== */
    @media (max-width: 768px) {
      body {
        font-size: 14px;
        -webkit-tap-highlight-color: transparent;
      }

      /* 1. 顶部导航栏收敛适配 */
      header {
        padding: 0.65rem 0.85rem;
        flex-wrap: wrap;
        gap: 0.65rem;
      }
      .brand {
        flex: 1;
        min-width: 140px;
        gap: 0.5rem;
      }
      .brand-icon {
        width: 1.85rem;
        height: 1.85rem;
        font-size: 1rem;
      }
      .brand-title {
        font-size: 1.05rem;
      }
      .brand-subtitle {
        display: none; /* 移动端折叠副标题，节省空间 */
      }
      .header-actions {
        width: 100%;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        justify-content: space-between;
      }
      .search-box {
        flex: 1;
        min-width: 0;
      }
      .search-input {
        width: 100% !important;
        font-size: 0.85rem;
        padding-top: 0.45rem;
        padding-bottom: 0.45rem;
      }
      .search-input:focus {
        width: 100% !important;
      }
      #refreshBtn {
        padding: 0.45rem 0.65rem;
        font-size: 0.8rem;
        flex-shrink: 0;
      }
      #authBtn {
        padding: 0.45rem 0.65rem;
        font-size: 0.8rem;
        flex-shrink: 0;
      }
      #authStatusBadge {
        display: none; /* 移动端隐藏状态标签，简化视图 */
      }

      /* 2. 页面主体 */
      main {
        padding: 0.85rem 0.65rem 5rem 0.65rem; /* 底部预留空间给浮动栏 */
      }

      /* 3. 大盘统计 2x2 网格化 */
      .dashboard-bar {
        margin-bottom: 0.85rem;
      }
      .stats-group {
        width: 100%;
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 0.5rem;
      }
      .stat-chip {
        padding: 0.4rem 0.65rem;
        font-size: 0.76rem;
        justify-content: space-between;
        margin: 0;
      }
      .stat-chip span:last-child {
        font-size: 0.72rem;
      }

      /* 4. 表格平滑重构为现代演示文稿卡片 (Card Layout) */
      .table-container {
        background: transparent;
        border: none;
        box-shadow: none;
        overflow: visible;
      }
      table thead {
        display: none; /* 移动端隐藏传统表头 */
      }
      table, tbody {
        display: block;
        width: 100%;
      }
      tbody tr {
        display: grid;
        grid-template-columns: 105px 1fr;
        grid-template-rows: auto auto;
        gap: 0.65rem;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 0.75rem;
        margin-bottom: 0.75rem;
        position: relative;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
        transition: border-color 0.2s, background-color 0.2s;
      }
      tbody tr.selected {
        border-color: var(--primary);
        background: rgba(37, 99, 235, 0.14);
      }
      td {
        display: block;
        padding: 0 !important;
        border: none !important;
      }

      /* 右上角勾选框（超大触控热区） */
      .col-checkbox {
        position: absolute;
        top: 0.4rem;
        right: 0.4rem;
        z-index: 5;
        width: 36px;
        height: 36px;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .col-checkbox input {
        width: 22px;
        height: 22px;
        accent-color: var(--primary);
        cursor: pointer;
      }

      /* 封面微缩图 */
      .col-thumb {
        grid-column: 1;
        grid-row: 1;
      }
      .thumb-box {
        width: 105px;
        height: 59px;
        border-radius: 6px;
      }

      /* 标题与文件信息 */
      .col-meta {
        grid-column: 2;
        grid-row: 1;
        min-width: 0;
        padding-right: 1.8rem !important; /* 为右上角复选框留出安全间距 */
      }
      .file-meta-col {
        gap: 0.2rem;
      }
      .file-title {
        font-size: 0.88rem;
        line-height: 1.35;
        font-weight: 600;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
      }
      .file-name {
        font-size: 0.72rem;
        color: var(--text-dim);
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      /* 移动端标签展示栏 */
      .mobile-tags-row {
        display: flex;
        align-items: center;
        gap: 0.45rem;
        flex-wrap: wrap;
        margin-top: 0.35rem;
        font-size: 0.72rem;
        color: var(--text-muted);
      }
      .mobile-tags-row .badge {
        padding: 0.15rem 0.45rem;
        font-size: 0.7rem;
      }

      /* 隐藏桌面端独立列 */
      .col-pages, .col-size, .col-date {
        display: none !important;
      }

      /* 卡片底部操作按钮栏（全屏等宽大按钮，极佳触控感） */
      .col-actions {
        grid-column: 1 / -1;
        grid-row: 2;
        width: 100%;
        margin-top: 0.2rem;
        padding-top: 0.55rem !important;
        border-top: 1px solid rgba(55, 65, 81, 0.4) !important;
      }
      .action-cell {
        width: 100%;
        display: flex;
        gap: 0.45rem;
        justify-content: stretch !important;
      }
      .action-cell .btn {
        flex: 1;
        justify-content: center;
        padding: 0.5rem 0.25rem;
        font-size: 0.82rem;
        border-radius: 6px;
      }

      /* 5. 分页组件移动端紧凑化 */
      .pagination-bar {
        flex-direction: column;
        align-items: stretch;
        gap: 0.65rem;
        padding: 0.85rem 0.75rem;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        margin-top: 0.5rem;
      }
      .pagination-info {
        text-align: center;
        font-size: 0.78rem;
      }
      .pagination-controls {
        justify-content: center;
        gap: 0.25rem;
        flex-wrap: wrap;
      }
      .page-size-select {
        padding: 0.2rem 0.4rem;
        font-size: 0.78rem;
      }
      .page-num-btn {
        min-width: 28px;
        height: 28px;
        font-size: 0.78rem;
      }

      /* 6. 批量操作条转为底部吸底操作面板 (Bottom Sheet) */
      .batch-bar {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        margin: 0;
        border-radius: 14px 14px 0 0;
        z-index: 90;
        padding: 0.75rem 1rem calc(env(safe-area-inset-bottom, 0px) + 0.75rem);
        box-shadow: 0 -6px 25px rgba(0, 0, 0, 0.6);
        flex-direction: column;
        gap: 0.5rem;
        backdrop-filter: blur(16px);
      }
      .batch-bar > div:first-child {
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }
      .batch-bar > div:last-child {
        width: 100%;
      }
      #batchDeleteBtn {
        width: 100%;
        justify-content: center;
        padding: 0.6rem;
        font-size: 0.88rem;
      }

      /* 7. 全景预览灯箱全屏沉浸适配 */
      .modal-backdrop {
        padding: 0;
      }
      .lightbox-content {
        width: 100vw;
        height: 100vh;
        max-width: 100vw;
        max-height: 100vh;
        border-radius: 0;
        border: none;
      }
      .lightbox-header {
        padding: 0.65rem 0.85rem;
      }
      .lightbox-header #lbTitle {
        max-width: 150px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        display: inline-block;
        vertical-align: middle;
        font-size: 0.9rem;
      }
      .lightbox-header #lbFilename {
        display: none;
      }
      .lightbox-body {
        padding: 0.25rem;
        min-height: unset;
      }
      .lightbox-img {
        max-height: 62vh;
        border-radius: 4px;
      }
      .nav-arrow {
        width: 38px;
        height: 38px;
        font-size: 1rem;
        background: rgba(17, 24, 39, 0.6);
        backdrop-filter: blur(4px);
      }
      .nav-prev { left: 0.4rem; }
      .nav-next { right: 0.4rem; }
      .lightbox-thumbs {
        padding: 0.45rem 0.65rem;
        gap: 0.35rem;
        -webkit-overflow-scrolling: touch;
      }
      .lightbox-thumb-item {
        width: 58px;
        height: 33px;
      }

      /* 8. 授权弹窗移动端自适应 */
      .auth-card {
        width: 90% !important;
        max-width: 340px !important;
        padding: 1.5rem 1.25rem !important;
      }
    }
  </style>
</head>
<body>

  <!-- 顶部导航栏 -->
  <header>
    <a href="/admin" class="brand">
      <div class="brand-icon">P</div>
      <div>
        <div class="brand-title">pptx-studio</div>
        <div class="brand-subtitle">演示文稿资产管理中心 · Asset Manager</div>
      </div>
    </a>

    <div class="header-actions">
      <div class="search-box">
        <span class="search-icon">🔍</span>
        <input type="text" id="searchInput" class="search-input" placeholder="按文件名或标题过滤...">
      </div>

      <button id="refreshBtn" class="btn" title="刷新文件列表">🔄 刷新</button>

      <span id="authStatusBadge" class="badge">检查授权中...</span>
      <button id="authBtn" class="btn btn-sm" style="display:none;">🔑 设置密钥</button>
    </div>
  </header>

  <!-- 页面主体 -->
  <main>
    <!-- 大盘汇总 -->
    <div class="dashboard-bar">
      <div class="stats-group">
        <div class="stat-chip">
          <span>文稿总数</span>
          <b id="totalCount">0</b>
        </div>
        <div class="stat-chip">
          <span>存储占用</span>
          <b id="totalSize">0 MB</b>
        </div>
        <div class="stat-chip">
          <span>排序方式</span>
          <span style="color:var(--text-muted);">🕒 生成时间倒序</span>
        </div>
      </div>
    </div>

    <!-- 批量操作悬浮条 -->
    <div id="batchBar" class="batch-bar">
      <div style="display: flex; align-items: center; gap: 1rem;">
        <span id="batchCountText" style="font-weight: 600; font-size: 0.9rem;">已选择 0 个文稿</span>
        <button id="cancelSelectBtn" class="btn btn-sm" style="background: rgba(255,255,255,0.1); border-color: rgba(255,255,255,0.2);">取消全选</button>
      </div>
      <div>
        <button id="batchDeleteBtn" class="btn btn-danger">🗑️ 批量彻底删除选中文稿</button>
      </div>
    </div>

    <!-- 文件列表表格 -->
    <div class="table-container">
      <table>
        <thead>
          <tr>
            <th style="width: 40px; text-align: center;">
              <input type="checkbox" id="selectAllCheckbox">
            </th>
            <th style="width: 120px;">封面缩略</th>
            <th>文稿名称与标题</th>
            <th>总页数</th>
            <th>文件大小</th>
            <th>生成时间 (倒序)</th>
            <th style="text-align: right; width: 220px;">管理操作</th>
          </tr>
        </thead>
        <tbody id="filesTableBody">
          <tr>
            <td colspan="7" class="empty-state">
              <div class="empty-icon">⏳</div>
              <div>正在加载文稿列表...</div>
            </td>
          </tr>
        </tbody>
      </table>

      <!-- 分页控制栏 -->
      <div id="paginationBar" class="pagination-bar">
        <div class="pagination-info">
          共 <b id="pgTotalItems">0</b> 条文稿，第 <b id="pgCurrentPage">1</b> / <b id="pgTotalPages">1</b> 页
        </div>
        <div class="pagination-controls">
          <label style="font-size: 0.82rem; color: var(--text-muted); display: flex; align-items: center; gap: 0.4rem;">
            每页:
            <select id="pageSizeSelect" class="page-size-select" onchange="changePageSize(this.value)">
              <option value="15" selected>15 条</option>
              <option value="30">30 条</option>
              <option value="50">50 条</option>
              <option value="100">100 条</option>
            </select>
          </label>
          <button id="pgFirstBtn" class="btn btn-sm" onclick="goToPage(1)">⏮️ 首页</button>
          <button id="pgPrevBtn" class="btn btn-sm" onclick="goToPage(currentPage - 1)">◀ 上一页</button>
          <div id="pgNumButtons" class="pagination-numbers"></div>
          <button id="pgNextBtn" class="btn btn-sm" onclick="goToPage(currentPage + 1)">下一页 ▶</button>
          <button id="pgLastBtn" class="btn btn-sm" onclick="goToPage(totalPages)">末页 ⏭️</button>
        </div>
      </div>
    </div>
  </main>

  <!-- 全景大图预览 Modal (Lightbox) -->
  <div id="lightboxModal" class="modal-backdrop">
    <div class="lightbox-content">
      <div class="lightbox-header">
        <div>
          <span id="lbTitle" style="font-weight: 600; font-size: 1rem; color: #fff;"></span>
          <span id="lbFilename" style="font-size: 0.78rem; color: var(--text-muted); margin-left: 0.5rem; font-family: monospace;"></span>
        </div>
        <div style="display: flex; align-items: center; gap: 0.75rem;">
          <span id="lbPageIndicator" class="badge badge-primary">1 / 1</span>
          <a id="lbDownloadBtn" href="#" class="btn btn-sm btn-primary" download>📥 下载 PPTX</a>
          <button id="lbCloseBtn" class="btn btn-sm" style="font-size: 1.1rem; line-height: 1;">✕</button>
        </div>
      </div>

      <div class="lightbox-body">
        <button id="lbPrevBtn" class="nav-arrow nav-prev">◀</button>
        <img id="lbMainImg" class="lightbox-img" src="" alt="幻灯片预览">
        <button id="lbNextBtn" class="nav-arrow nav-next">▶</button>
      </div>

      <div id="lbThumbs" class="lightbox-thumbs">
        <!-- 底部微缩胶囊 -->
      </div>
    </div>
  </div>

  <!-- 授权登录 Modal -->
  <div id="authModal" class="modal-backdrop">
    <div class="auth-modal">
      <div class="auth-icon">🔒</div>
      <h2 style="font-size: 1.25rem; font-weight: 700; margin-bottom: 0.5rem;">访问授权验证</h2>
      <p style="font-size: 0.85rem; color: var(--text-muted); line-height: 1.4;">
        当前服务器已启用 <code style="color:var(--primary-light);">AUTH_KEY</code> 安全防护机制。<br>请输入授权访问密钥以解锁文稿管理权限。
      </p>

      <input type="password" id="authKeyInput" class="auth-input" placeholder="输入 AUTH_KEY 访问凭证..." autocomplete="current-password">
      <div id="authErrorMsg" class="auth-error">授权密钥错误，请重新输入</div>

      <div style="display: flex; gap: 0.75rem; justify-content: center;">
        <button id="authSubmitBtn" class="btn btn-primary" style="width: 100%; padding: 0.7rem;">🔓 解锁并进入管理台</button>
      </div>
    </div>
  </div>

  <!-- Toast 提示条 -->
  <div id="toast" class="toast">
    <span id="toastIcon">ℹ️</span>
    <span id="toastText">操作成功</span>
  </div>

  <script>
    // 全局状态
    let allFiles = [];
    let selectedIds = new Set();
    let currentLightbox = {
      file: null,
      currentPage: 1,
      totalPages: 1
    };
    let currentPage = 1;
    let pageSize = 15;
    let totalPages = 1;
    let totalCount = 0;
    let searchQuery = "";
    let searchDebounceTimer = null;

    function getAuthKey() {
      return localStorage.getItem("pptx_auth_key") || "";
    }

    function setAuthKey(key) {
      if (key) {
        localStorage.setItem("pptx_auth_key", key);
      } else {
        localStorage.removeItem("pptx_auth_key");
      }
    }

    function getAuthHeaders() {
      const key = getAuthKey();
      const headers = { "Content-Type": "application/json" };
      if (key) {
        headers["Authorization"] = `Bearer ${key}`;
      }
      return headers;
    }

    function withAuthParam(url) {
      const key = getAuthKey();
      if (!key) return url;
      const sep = url.includes("?") ? "&" : "?";
      return `${url}${sep}auth_key=${encodeURIComponent(key)}`;
    }

    function showToast(text, isError = false) {
      const toast = document.getElementById("toast");
      document.getElementById("toastIcon").innerText = isError ? "❌" : "✅";
      document.getElementById("toastText").innerText = text;
      toast.classList.add("active");
      setTimeout(() => toast.classList.remove("active"), 3000);
    }

    // 格式化文件字节大小
    function formatBytes(bytes) {
      if (!bytes || bytes <= 0) return "0 B";
      const k = 1024;
      const sizes = ["B", "KB", "MB", "GB"];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
    }

    // 初始化加载（带服务端毫秒级分页与轻量缓存）
    async function loadFiles(page = currentPage) {
      currentPage = page;
      const tbody = document.getElementById("filesTableBody");
      try {
        const params = new URLSearchParams({
          page: currentPage,
          page_size: pageSize,
          search: searchQuery
        });
        const resp = await fetch(`/api/admin/files?${params.toString()}`, {
          headers: getAuthHeaders()
        });

        if (resp.status === 401) {
          showAuthModal(true);
          return;
        }

        if (!resp.ok) {
          throw new Error("加载文稿列表失败 (" + resp.status + ")");
        }

        const data = await resp.json();
        allFiles = data.items || [];
        totalCount = data.total || 0;
        totalPages = data.total_pages || 1;
        currentPage = data.page || 1;

        updateAuthBadge(data.auth_required);
        renderTable(allFiles);
        updateStats(data);
        renderPagination();
      } catch (err) {
        tbody.innerHTML = `
          <tr>
            <td colspan="7" class="empty-state">
              <div class="empty-icon">⚠️</div>
              <div style="color:#f87171;">${escapeHtml(err.message)}</div>
              <button class="btn btn-sm" onclick="loadFiles()" style="margin-top:1rem;">重试</button>
            </td>
          </tr>
        `;
      }
    }

    function updateAuthBadge(authRequired) {
      const badge = document.getElementById("authStatusBadge");
      const authBtn = document.getElementById("authBtn");
      authBtn.style.display = "inline-flex";

      if (authRequired) {
        badge.className = "badge badge-primary";
        badge.innerText = "🔒 密钥防护已开启";
        authBtn.innerText = "🔑 切换密钥";
      } else {
        badge.className = "badge";
        badge.innerText = "🔓 免密公开模式";
        authBtn.innerText = "🔑 配置密钥";
      }
    }

    function updateStats(data) {
      document.getElementById("totalCount").innerText = data.total || 0;
      document.getElementById("totalSize").innerText = formatBytes(data.total_disk_bytes || 0);
    }

    function renderPagination() {
      const bar = document.getElementById("paginationBar");
      if (!bar) return;
      document.getElementById("pgTotalItems").innerText = totalCount;
      document.getElementById("pgCurrentPage").innerText = currentPage;
      document.getElementById("pgTotalPages").innerText = totalPages;

      const firstBtn = document.getElementById("pgFirstBtn");
      const prevBtn = document.getElementById("pgPrevBtn");
      const nextBtn = document.getElementById("pgNextBtn");
      const lastBtn = document.getElementById("pgLastBtn");

      firstBtn.disabled = currentPage <= 1;
      prevBtn.disabled = currentPage <= 1;
      nextBtn.disabled = currentPage >= totalPages;
      lastBtn.disabled = currentPage >= totalPages;

      const numContainer = document.getElementById("pgNumButtons");
      numContainer.innerHTML = "";

      let startP = Math.max(1, currentPage - 2);
      let endP = Math.min(totalPages, startP + 4);
      if (endP - startP < 4) {
        startP = Math.max(1, endP - 4);
      }

      for (let p = startP; p <= endP; p++) {
        const btn = document.createElement("button");
        btn.className = `page-num-btn ${p === currentPage ? 'active' : ''}`;
        btn.innerText = p;
        btn.onclick = () => goToPage(p);
        numContainer.appendChild(btn);
      }
    }

    function goToPage(p) {
      if (p < 1 || p > totalPages || p === currentPage) return;
      loadFiles(p);
    }

    function changePageSize(sz) {
      pageSize = parseInt(sz, 10) || 15;
      loadFiles(1);
    }

    // 渲染表格
    function renderTable(files) {
      const tbody = document.getElementById("filesTableBody");
      updateBatchBar();

      // 判断本页是否全部选中
      const allCurrentSelected = files.length > 0 && files.every(f => selectedIds.has(f.id));
      document.getElementById("selectAllCheckbox").checked = allCurrentSelected;

      if (!files.length) {
        tbody.innerHTML = `
          <tr>
            <td colspan="7" class="empty-state">
              <div class="empty-icon">📂</div>
              <div>暂无已生成的 PPTX 文稿文件</div>
              <div style="font-size:0.75rem; color:var(--text-dim); margin-top:0.35rem;">
                通过 MCP 工具或 REST API 生成后将自动在此列出
              </div>
            </td>
          </tr>
        `;
        return;
      }

      tbody.innerHTML = files.map(f => {
        const thumbUrl = withAuthParam(f.thumbnail_url || `/api/preview/${f.id}?page=1`);
        const dlUrl = withAuthParam(f.download_url || `/api/download/${f.filename}`);
        const pagesCount = f.pages_count || 1;
        const isSelected = selectedIds.has(f.id);

        return `
          <tr id="row-${f.id}" class="${isSelected ? 'selected' : ''}">
            <td class="col-checkbox" style="text-align: center;">
              <input type="checkbox" class="row-checkbox" data-id="${f.id}" ${isSelected ? 'checked' : ''} onchange="toggleSelect('${f.id}')">
            </td>
            <td class="col-thumb">
              <div class="thumb-box" onclick="openLightbox('${f.id}')">
                <img src="${thumbUrl}" class="thumb-img" alt="封面" loading="lazy" onerror="this.src='data:image/svg+xml;utf8,<svg xmlns=\\'http://www.w3.org/2000/svg\\' width=\\'110\\' height=\\'62\\'><rect width=\\'110\\' height=\\'62\\' fill=\\'%231f2937\\'/><text x=\\'55\\' y=\\'35\\' font-size=\\'12\\' fill=\\'%239ca3af\\' text-anchor=\\'middle\\'>无预览</text></svg>'">
                <div class="thumb-hover-overlay">👁️ 查看</div>
              </div>
            </td>
            <td class="col-meta">
              <div class="file-meta-col">
                <div class="file-title" onclick="openLightbox('${f.id}')">${escapeHtml(f.title || f.filename)}</div>
                <div class="file-name">${escapeHtml(f.filename)}</div>
                <div class="mobile-tags-row">
                  <span class="badge badge-primary">${pagesCount} 页</span>
                  <span>💾 ${formatBytes(f.size_bytes)}</span>
                  <span>🕒 ${f.created_at ? f.created_at.slice(5) : '-'}</span>
                </div>
              </div>
            </td>
            <td class="col-pages">
              <span class="badge ${pagesCount > 1 ? 'badge-primary' : ''}">${pagesCount} 页</span>
            </td>
            <td class="col-size" style="color:var(--text-muted); font-size:0.82rem;">${formatBytes(f.size_bytes)}</td>
            <td class="col-date" style="color:var(--text-muted); font-size:0.82rem;">${f.created_at || '-'}</td>
            <td class="col-actions" style="text-align: right;">
              <div class="action-cell">
                <button class="btn btn-sm btn-primary" onclick="openLightbox('${f.id}')" title="查看所有页高清预览">👁️ 预览</button>
                <a href="${dlUrl}" class="btn btn-sm" download title="直接下载 PPTX 演示文稿">📥 下载</a>
                <button class="btn btn-sm btn-danger" onclick="deleteSingle('${f.id}', '${escapeHtml(f.filename)}')" title="删除此文稿">🗑️</button>
              </div>
            </td>
          </tr>
        `;
      }).join("");
    }

    function escapeHtml(str) {
      if (!str) return "";
      return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    // 选中与批量操作
    function toggleSelect(id) {
      if (selectedIds.has(id)) {
        selectedIds.delete(id);
      } else {
        selectedIds.add(id);
      }
      const row = document.getElementById("row-" + id);
      if (row) {
        if (selectedIds.has(id)) row.classList.add("selected");
        else row.classList.remove("selected");
      }
      updateBatchBar();
    }

    document.getElementById("selectAllCheckbox").addEventListener("change", function(e) {
      const checked = e.target.checked;
      const visibleCheckboxes = document.querySelectorAll(".row-checkbox");
      visibleCheckboxes.forEach(cb => {
        const id = cb.getAttribute("data-id");
        cb.checked = checked;
        if (checked) {
          selectedIds.add(id);
          document.getElementById("row-" + id)?.classList.add("selected");
        } else {
          selectedIds.delete(id);
          document.getElementById("row-" + id)?.classList.remove("selected");
        }
      });
      updateBatchBar();
    });

    document.getElementById("cancelSelectBtn").addEventListener("click", () => {
      selectedIds.clear();
      document.querySelectorAll(".row-checkbox").forEach(cb => cb.checked = false);
      document.querySelectorAll("tbody tr").forEach(tr => tr.classList.remove("selected"));
      document.getElementById("selectAllCheckbox").checked = false;
      updateBatchBar();
    });

    function updateBatchBar() {
      const bar = document.getElementById("batchBar");
      const text = document.getElementById("batchCountText");
      if (selectedIds.size > 0) {
        bar.classList.add("active");
        text.innerText = `已勾选 ${selectedIds.size} 个演示文稿`;
      } else {
        bar.classList.remove("active");
      }
    }

    // 删除单项
    async function deleteSingle(id, filename) {
      if (!confirm(`确定彻底删除演示文稿「${filename}」吗？\n删除后关联的高清预览图也会同步清理。`)) {
        return;
      }
      try {
        const resp = await fetch(`/api/artifacts/${encodeURIComponent(id)}`, {
          method: "DELETE",
          headers: getAuthHeaders()
        });
        if (!resp.ok) throw new Error("删除失败");
        showToast("已成功删除 " + filename);
        loadFiles();
      } catch (err) {
        showToast(err.message, true);
      }
    }

    // 批量删除
    document.getElementById("batchDeleteBtn").addEventListener("click", async () => {
      const count = selectedIds.size;
      if (count === 0) return;
      if (!confirm(`⚠️ 风险确认：\n确定批量彻底删除选中的 ${count} 个演示文稿吗？\n此操作不可逆！`)) {
        return;
      }
      try {
        const resp = await fetch("/api/admin/delete", {
          method: "POST",
          headers: getAuthHeaders(),
          body: JSON.stringify({ ids: Array.from(selectedIds) })
        });
        if (!resp.ok) throw new Error("批量删除失败 (" + resp.status + ")");
        const res = await resp.json();
        showToast(`已批量删除 ${res.deleted_count} 个文稿`);
        selectedIds.clear();
        loadFiles();
      } catch (err) {
        showToast(err.message, true);
      }
    });

    // 实时搜索过滤（300ms 防抖，服务端极速过滤）
    document.getElementById("searchInput").addEventListener("input", function(e) {
      clearTimeout(searchDebounceTimer);
      searchQuery = e.target.value.trim();
      searchDebounceTimer = setTimeout(() => {
        loadFiles(1);
      }, 300);
    });

    // 刷新按钮
    document.getElementById("refreshBtn").addEventListener("click", () => loadFiles(1));

    // ==================== 全景预览灯箱 (Lightbox) ====================
    function openLightbox(fileId) {
      const f = allFiles.find(item => item.id === fileId);
      if (!f) return;

      currentLightbox.file = f;
      currentLightbox.currentPage = 1;
      currentLightbox.totalPages = f.pages_count || 1;

      document.getElementById("lbTitle").innerText = f.title || f.filename;
      document.getElementById("lbFilename").innerText = f.filename;
      document.getElementById("lbDownloadBtn").href = withAuthParam(f.download_url || `/api/download/${f.filename}`);

      // 渲染微缩图底栏
      const thumbsContainer = document.getElementById("lbThumbs");
      thumbsContainer.innerHTML = "";
      for (let p = 1; p <= currentLightbox.totalPages; p++) {
        const thumbDiv = document.createElement("div");
        thumbDiv.className = `lightbox-thumb-item ${p === 1 ? 'active' : ''}`;
        thumbDiv.setAttribute("data-page", p);
        const pUrl = withAuthParam(`/api/preview/${f.id}?page=${p}`);
        thumbDiv.innerHTML = `<img src="${pUrl}" alt="第${p}页">`;
        thumbDiv.onclick = () => showLightboxPage(p);
        thumbsContainer.appendChild(thumbDiv);
      }

      showLightboxPage(1);
      document.getElementById("lightboxModal").classList.add("active");
    }

    function showLightboxPage(page) {
      currentLightbox.currentPage = page;
      document.getElementById("lbPageIndicator").innerText = `第 ${page} / ${currentLightbox.totalPages} 页`;
      
      const mainImg = document.getElementById("lbMainImg");
      mainImg.style.opacity = "0.4";
      const pUrl = withAuthParam(`/api/preview/${currentLightbox.file.id}?page=${page}`);
      mainImg.src = pUrl;
      mainImg.onload = () => { mainImg.style.opacity = "1"; };

      // 更新微缩底栏选中状态
      document.querySelectorAll(".lightbox-thumb-item").forEach(item => {
        const itemPage = parseInt(item.getAttribute("data-page"), 10);
        if (itemPage === page) {
          item.classList.add("active");
          item.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
        } else {
          item.classList.remove("active");
        }
      });
    }

    function prevLightboxPage() {
      if (currentLightbox.currentPage > 1) {
        showLightboxPage(currentLightbox.currentPage - 1);
      } else {
        showLightboxPage(currentLightbox.totalPages);
      }
    }

    function nextLightboxPage() {
      if (currentLightbox.currentPage < currentLightbox.totalPages) {
        showLightboxPage(currentLightbox.currentPage + 1);
      } else {
        showLightboxPage(1);
      }
    }

    document.getElementById("lbPrevBtn").addEventListener("click", prevLightboxPage);
    document.getElementById("lbNextBtn").addEventListener("click", nextLightboxPage);
    document.getElementById("lbCloseBtn").addEventListener("click", closeLightbox);
    document.getElementById("lightboxModal").addEventListener("click", (e) => {
      if (e.target.id === "lightboxModal") closeLightbox();
    });

    function closeLightbox() {
      document.getElementById("lightboxModal").classList.remove("active");
    }

    // 键盘快捷键监听
    document.addEventListener("keydown", (e) => {
      if (document.getElementById("lightboxModal").classList.contains("active")) {
        if (e.key === "ArrowLeft") prevLightboxPage();
        else if (e.key === "ArrowRight") nextLightboxPage();
        else if (e.key === "Escape") closeLightbox();
      }
      if (e.key === "Escape" && document.getElementById("authModal").classList.contains("active")) {
        // do nothing
      }
    });

    // 移动端手势左右滑动翻页支持 (Touch Swipe Gestures)
    let touchStartX = 0;
    let touchStartY = 0;
    const lbModal = document.getElementById("lightboxModal");
    lbModal.addEventListener("touchstart", (e) => {
      if (e.changedTouches && e.changedTouches.length > 0) {
        touchStartX = e.changedTouches[0].screenX;
        touchStartY = e.changedTouches[0].screenY;
      }
    }, { passive: true });

    lbModal.addEventListener("touchend", (e) => {
      if (!lbModal.classList.contains("active")) return;
      if (e.changedTouches && e.changedTouches.length > 0) {
        const deltaX = e.changedTouches[0].screenX - touchStartX;
        const deltaY = e.changedTouches[0].screenY - touchStartY;
        // 水平滑动距离超过 40px 且水平滑动位移大于垂直位移 1.2 倍时判定为翻页手势
        if (Math.abs(deltaX) > 40 && Math.abs(deltaX) > Math.abs(deltaY) * 1.2) {
          if (deltaX < 0) {
            nextLightboxPage(); // 向左滑 -> 查看下一页
          } else {
            prevLightboxPage(); // 向右滑 -> 查看上一页
          }
        }
      }
    }, { passive: true });

    // ==================== 认证弹窗 (Auth Modal) ====================
    function showAuthModal(hasError = false) {
      const modal = document.getElementById("authModal");
      const err = document.getElementById("authErrorMsg");
      const input = document.getElementById("authKeyInput");
      input.value = getAuthKey();
      err.style.display = hasError ? "block" : "none";
      modal.classList.add("active");
      setTimeout(() => input.focus(), 100);
    }

    function closeAuthModal() {
      document.getElementById("authModal").classList.remove("active");
    }

    document.getElementById("authBtn").addEventListener("click", () => showAuthModal(false));

    document.getElementById("authSubmitBtn").addEventListener("click", async () => {
      const key = document.getElementById("authKeyInput").value.trim();
      setAuthKey(key);
      document.getElementById("authErrorMsg").style.display = "none";
      try {
        const resp = await fetch("/api/admin/verify", {
          method: "POST",
          headers: getAuthHeaders(),
          body: JSON.stringify({ key: key })
        });
        if (resp.status === 401) {
          document.getElementById("authErrorMsg").style.display = "block";
          return;
        }
        closeAuthModal();
        showToast("授权验证成功");
        loadFiles();
      } catch (err) {
        document.getElementById("authErrorMsg").style.display = "block";
      }
    });

    document.getElementById("authKeyInput").addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        document.getElementById("authSubmitBtn").click();
      }
    });

    // 页面加载启动
    loadFiles();
  </script>
</body>
</html>
"""
