#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档 (.docx) 交互式 Web 前端应用
==============================================================================
Project / Author : https://github.com/endangered-xiezhi
License          : MIT License
==============================================================================

功能：
1. 本地浏览器交互界面，支持拖拽上传 Word (.docx) 文件。
2. 自动检测并可视化呈现当前所有隐藏属性（作者、上次保存者、编辑时长、修订痕迹人员等）。
3. 提供“一键填入合理推荐值”功能，快速消除异常过大的编辑时长或过高修订版本破绽。
4. 在线配置作者、修订人、批注人及自然递增仿真时间戳。
5. 一键导出并直接下载修改后的 Word 文档。

启动方法：
    python3 docx_web_ui.py
运行后会自动在默认浏览器中打开交互界面 (http://127.0.0.1:5050)
"""

import os
import sys
import tempfile
import webbrowser
from threading import Timer
from flask import Flask, request, jsonify, send_file, render_template_string

# 导入核心处理模块
import docx_metadata_editor

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

WATERMARK_URL = "https://github.com/endangered-xiezhi"

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Word 元数据与修订痕迹清洗工具 | endangered-xiezhi</title>
    <style>
        :root {
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --bg: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --danger: #ef4444;
            --danger-bg: #fef2f2;
            --success: #10b981;
            --success-bg: #ecfdf5;
            --warning-bg: #fffbeb;
            --warning-text: #b45309;
            --radius: 12px;
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.07), 0 2px 4px -2px rgb(0 0 0 / 0.05);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
        }

        body {
            background-color: var(--bg);
            color: var(--text-main);
            line-height: 1.5;
            padding: 24px;
        }

        .container {
            max-width: 1100px;
            margin: 0 auto;
        }

        /* 顶部导航 */
        header {
            text-align: center;
            margin-bottom: 24px;
            position: relative;
        }

        .github-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: #1e293b;
            color: #ffffff;
            font-size: 12px;
            font-weight: 500;
            padding: 4px 12px;
            border-radius: 20px;
            text-decoration: none;
            margin-bottom: 12px;
            transition: all 0.2s ease;
        }

        .github-badge:hover {
            background: #334155;
            transform: translateY(-1px);
        }

        header h1 {
            font-size: 26px;
            font-weight: 700;
            color: #1e293b;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }

        header p {
            color: var(--text-muted);
            font-size: 14px;
            margin-top: 6px;
        }

        /* 上传区域 */
        .dropzone {
            background: var(--card-bg);
            border: 2px dashed #cbd5e1;
            border-radius: var(--radius);
            padding: 32px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s ease;
            box-shadow: var(--shadow);
            margin-bottom: 24px;
        }

        .dropzone:hover, .dropzone.dragover {
            border-color: var(--primary);
            background: #eff6ff;
        }

        .dropzone-icon {
            font-size: 40px;
            margin-bottom: 8px;
        }

        .dropzone-text {
            font-size: 16px;
            font-weight: 600;
            color: #334155;
        }

        .dropzone-subtext {
            font-size: 13px;
            color: var(--text-muted);
            margin-top: 4px;
        }

        /* 主体分栏 */
        .main-grid {
            display: grid;
            grid-template-columns: 1fr 1.35fr;
            gap: 20px;
            display: none;
        }

        .card {
            background: var(--card-bg);
            border-radius: var(--radius);
            border: 1px solid var(--border);
            box-shadow: var(--shadow);
            padding: 20px;
        }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 12px;
            margin-bottom: 16px;
        }

        .card-title {
            font-size: 16px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        /* 左侧：检测属性列表 */
        .inspect-item {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            padding: 8px 0;
            border-bottom: 1px solid #f1f5f9;
            font-size: 13px;
        }

        .inspect-item:last-child {
            border-bottom: none;
        }

        .inspect-label {
            color: var(--text-muted);
            flex: 0 0 110px;
        }

        .inspect-value {
            font-weight: 500;
            text-align: right;
            word-break: break-all;
            color: #1e293b;
        }

        .badge-warning {
            background-color: var(--warning-bg);
            color: var(--warning-text);
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            margin-left: 4px;
        }

        .badge-danger {
            background-color: var(--danger-bg);
            color: var(--danger);
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            margin-left: 4px;
        }

        /* 右侧：表单配置 */
        .form-section {
            margin-bottom: 16px;
        }

        .form-section-title {
            font-size: 13px;
            font-weight: 700;
            color: #475569;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-bottom: 10px;
        }

        .form-group {
            display: flex;
            flex-direction: column;
            gap: 4px;
            margin-bottom: 10px;
        }

        .form-group label {
            font-size: 12px;
            font-weight: 600;
            color: #475569;
        }

        .form-group input, .form-group select {
            padding: 8px 12px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 13px;
            outline: none;
            transition: border-color 0.15s;
        }

        .form-group input:focus, .form-group select:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1);
        }

        /* 快捷填充栏 */
        .preset-bar {
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            border-radius: 8px;
            padding: 10px 14px;
            margin-bottom: 16px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .preset-desc {
            font-size: 12px;
            color: #166534;
        }

        .btn-preset {
            background: #16a34a;
            color: #fff;
            border: none;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.15s;
        }

        .btn-preset:hover {
            background: #15803d;
        }

        /* 底部提交按钮 */
        .btn-submit {
            background: var(--primary);
            color: #ffffff;
            border: none;
            width: 100%;
            padding: 12px;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.3);
            transition: all 0.15s;
            margin-top: 10px;
        }

        .btn-submit:hover {
            background: var(--primary-hover);
            transform: translateY(-1px);
        }

        .btn-submit:disabled {
            background: #94a3b8;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        .loading {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 0.8s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .toast {
            position: fixed;
            bottom: 24px;
            right: 24px;
            background: #1e293b;
            color: #fff;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 14px;
            box-shadow: 0 10px 15px -3px rgba(0,0,0,0.3);
            display: none;
            z-index: 100;
        }

        footer {
            text-align: center;
            margin-top: 36px;
            padding-top: 16px;
            border-top: 1px solid var(--border);
            font-size: 13px;
            color: var(--text-muted);
        }

        footer a {
            color: var(--primary);
            text-decoration: none;
            font-weight: 500;
        }

        footer a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>

<div class="container">
    <header>
        <a href="https://github.com/endangered-xiezhi" target="_blank" class="github-badge">
            <svg height="14" width="14" viewBox="0 0 16 16" fill="currentColor">
                <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
            </svg>
            GitHub: endangered-xiezhi
        </a>
        <h1>📝 Word 元数据与修订痕迹清洗工具</h1>
        <p>一键侦测隐藏作者、保存者、异常编辑时长与修订人，并在本地完成脱敏清洗与平滑仿真</p>
    </header>

    <!-- 拖拽上传区 -->
    <div class="dropzone" id="dropzone">
        <div class="dropzone-icon">📄</div>
        <div class="dropzone-text" id="dropzone-text">点击选择 或 直接将 Word 文档 (.docx) 拖入此处</div>
        <div class="dropzone-subtext">所有处理均在本地内存与临时目录完成，安全不外泄</div>
        <input type="file" id="file-input" accept=".docx" style="display: none;">
    </div>

    <!-- 主体：分析与配置面板 -->
    <div class="main-grid" id="main-grid">
        
        <!-- 左侧：侦测报告 -->
        <div class="card">
            <div class="card-header">
                <span class="card-title">🔍 原始隐藏属性检测</span>
                <span id="file-tag" style="font-size: 12px; color: var(--primary); font-weight: 600;"></span>
            </div>

            <div id="inspect-list">
                <!-- 动态填充 -->
            </div>
        </div>

        <!-- 右侧：修改配置表单 -->
        <div class="card">
            <div class="card-header">
                <span class="card-title">✏️ 目标修改值配置</span>
                <span style="font-size: 12px; color: var(--text-muted);">留空项将保持原样不改动</span>
            </div>

            <!-- 一键消除破绽预设 -->
            <div class="preset-bar">
                <div class="preset-desc">
                    <strong>💡 快捷优化：</strong>一键消除异常偏大的编辑时长、过高修订版本，统一保存者
                </div>
                <button type="button" class="btn-preset" id="btn-quick-fill">⚡ 填入合理推荐值</button>
            </div>

            <form id="modify-form">
                <!-- 分组 1: 基本元数据 -->
                <div class="form-section">
                    <div class="form-section-title">📌 文档核心属性</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>文档创建作者 (Creator)</label>
                            <input type="text" id="input-creator" placeholder="例如：张三 / AuthorName">
                        </div>
                        <div class="form-group">
                            <label>上次保存者 (Last Modified By)</label>
                            <input type="text" id="input-modifier" placeholder="例如：张三 / AuthorName">
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>修订版本号 (Revision)</label>
                            <input type="number" id="input-rev-num" placeholder="例如：2 或 3">
                        </div>
                        <div class="form-group">
                            <label>所属单位/公司 (Company)</label>
                            <input type="text" id="input-company" placeholder="例如：某某单位 / 某某大学">
                        </div>
                    </div>
                </div>

                <!-- 分组 2: 统计与时间 -->
                <div class="form-section">
                    <div class="form-section-title">⏱️ 统计信息与时间</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>编辑时间总计 (分钟数)</label>
                            <input type="number" id="input-total-time" placeholder="例如：120 (平抑异常偏大数值)">
                        </div>
                        <div class="form-group">
                            <label>文档标题 (Title)</label>
                            <input type="text" id="input-title" placeholder="如需重命名请填写">
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>创建时间 (Created Time)</label>
                            <input type="datetime-local" id="input-created">
                        </div>
                        <div class="form-group">
                            <label>最后修改时间 (Modified Time)</label>
                            <input type="datetime-local" id="input-modified">
                        </div>
                    </div>
                </div>

                <!-- 分组 3: 修订痕迹与批注 -->
                <div class="form-section">
                    <div class="form-section-title">🖋️ 修订痕迹 (Track Changes) 与批注</div>
                    <div class="form-row">
                        <div class="form-group">
                            <label>修订痕迹新作者 (Track Changes)</label>
                            <input type="text" id="input-rev-author" placeholder="统一替换所有正文修订人">
                        </div>
                        <div class="form-group">
                            <label>批注者姓名 (Comment Author)</label>
                            <input type="text" id="input-comment-author" placeholder="例如：审阅人">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>修订时间处理模式</label>
                        <select id="select-rev-time-mode">
                            <option value="keep">保持现有修订时间不变</option>
                            <option value="fixed">统一设为固定时间</option>
                            <option value="range" selected>自然仿真：在时间区间内平滑递增分布 (推荐)</option>
                        </select>
                    </div>

                    <!-- 固定时间输入框 -->
                    <div class="form-group" id="group-rev-fixed" style="display: none;">
                        <label>统一修订时间</label>
                        <input type="datetime-local" id="input-rev-fixed">
                    </div>

                    <!-- 区间递增输入框 -->
                    <div class="form-row" id="group-rev-range">
                        <div class="form-group">
                            <label>修订起始时间</label>
                            <input type="datetime-local" id="input-rev-start">
                        </div>
                        <div class="form-group">
                            <label>修订截止时间</label>
                            <input type="datetime-local" id="input-rev-end">
                        </div>
                    </div>
                </div>

                <button type="button" class="btn-submit" id="btn-submit">
                    <span id="btn-text">🚀 立即导出并下载修改后的 Word 文档</span>
                </button>
            </form>
        </div>

    </div>

    <footer>
        Word Document Metadata & Track Changes Cleaner · 
        <a href="https://github.com/endangered-xiezhi" target="_blank">github.com/endangered-xiezhi</a> · 
        开源免费 · 纯本地运行
    </footer>
</div>

<div class="toast" id="toast"></div>

<script>
    let currentFile = null;
    let currentInfo = null;

    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');
    const dropzoneText = document.getElementById('dropzone-text');
    const mainGrid = document.getElementById('main-grid');
    const inspectList = document.getElementById('inspect-list');
    const fileTag = document.getElementById('file-tag');
    const toast = document.getElementById('toast');

    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    function toLocalInputFormat(w3cStr) {
        if (!w3cStr) return '';
        try {
            const d = new Date(w3cStr);
            if (isNaN(d.getTime())) return '';
            const pad = (n) => String(n).padStart(2, '0');
            return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
        } catch {
            return '';
        }
    }

    function showToast(msg) {
        toast.innerText = msg;
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 3000);
    }

    async function handleFileSelect(file) {
        if (!file.name.endsWith('.docx')) {
            alert('请上传 .docx 格式的 Word 文档！');
            return;
        }
        currentFile = file;
        dropzoneText.innerHTML = `已选择: <strong>${file.name}</strong> (${(file.size / 1024).toFixed(1)} KB)`;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const resp = await fetch('/api/inspect', {
                method: 'POST',
                body: formData
            });
            const data = await resp.json();
            if (!data.success) {
                alert('解析失败: ' + data.error);
                return;
            }
            currentInfo = data.info;
            renderInspect(currentInfo);
            initFormDefaults(currentInfo);
            mainGrid.style.display = 'grid';
            mainGrid.scrollIntoView({ behavior: 'smooth' });
        } catch (err) {
            alert('请求出错: ' + err.message);
        }
    }

    function renderInspect(info) {
        fileTag.innerText = info.filename;
        let html = '';

        const appendItem = (label, val, badge = '') => {
            html += `
                <div class="inspect-item">
                    <span class="inspect-label">${label}</span>
                    <span class="inspect-value">${val || '<span style="color:#94a3b8">(未设置)</span>'}${badge}</span>
                </div>
            `;
        };

        appendItem('文档标题', info.title);
        appendItem('创建作者', info.creator);

        let modBadge = (info.last_modified_by && info.creator && info.last_modified_by !== info.creator) 
            ? `<span class="badge-warning">与作者不一致</span>` : '';
        appendItem('上次保存者', info.last_modified_by, modBadge);

        let revBadge = (parseInt(info.revision) > 30) ? `<span class="badge-danger">版本过高(${info.revision})</span>` : '';
        appendItem('修订版本号', info.revision, revBadge);

        let totalTimeNum = parseInt(info.total_time);
        let timeBadge = (totalTimeNum > 500) ? `<span class="badge-danger">异常偏大(${totalTimeNum}分钟)</span>` : '';
        appendItem('编辑时间总计', info.total_time ? `${info.total_time} 分钟` : null, timeBadge);

        appendItem('创建时间', info.created);
        appendItem('修改时间', info.modified);
        appendItem('单位 / 公司', info.company);

        if (info.rev_authors && info.rev_authors.length > 0) {
            appendItem('修订人列表', info.rev_authors.join(', '), `<span class="badge-warning">${info.rev_count} 处修订</span>`);
            appendItem('最早修订时间', info.rev_earliest);
            appendItem('最晚修订时间', info.rev_latest);
        } else {
            appendItem('修订人列表', '未检测到正文修订痕迹');
        }

        if (info.comment_authors && info.comment_authors.length > 0) {
            appendItem('批注人列表', info.comment_authors.join(', '));
        }

        if (info.people && info.people.length > 0) {
            appendItem('Office365协作人', info.people.join(', '));
        }

        inspectList.innerHTML = html;
    }

    function initFormDefaults(info) {
        document.getElementById('input-creator').value = info.creator || '';
        document.getElementById('input-modifier').value = info.creator || info.last_modified_by || '';
        document.getElementById('input-rev-num').value = '3';
        document.getElementById('input-company').value = info.company || '';
        document.getElementById('input-total-time').value = '120';
        document.getElementById('input-title').value = info.title || '';

        document.getElementById('input-created').value = toLocalInputFormat(info.created);
        document.getElementById('input-modified').value = toLocalInputFormat(info.modified);

        if (info.rev_authors && info.rev_authors.length > 0) {
            document.getElementById('input-rev-author').value = info.creator || info.rev_authors[0];
        } else {
            document.getElementById('input-rev-author').value = info.creator || '';
        }

        if (info.comment_authors && info.comment_authors.length > 0) {
            document.getElementById('input-comment-author').value = info.comment_authors[0];
        }

        if (info.modified) {
            const modDate = new Date(info.modified);
            const startDate = new Date(modDate.getTime() - 2.5 * 3600 * 1000);
            document.getElementById('input-rev-start').value = toLocalInputFormat(startDate.toISOString());
            document.getElementById('input-rev-end').value = toLocalInputFormat(modDate.toISOString());
        }
    }

    const modeSelect = document.getElementById('select-rev-time-mode');
    const groupFixed = document.getElementById('group-rev-fixed');
    const groupRange = document.getElementById('group-rev-range');
    modeSelect.addEventListener('change', () => {
        const val = modeSelect.value;
        groupFixed.style.display = (val === 'fixed') ? 'block' : 'none';
        groupRange.style.display = (val === 'range') ? 'grid' : 'none';
    });

    document.getElementById('btn-quick-fill').addEventListener('click', () => {
        if (!currentInfo) return;
        const targetName = document.getElementById('input-creator').value.trim() || currentInfo.creator || '原作者';
        document.getElementById('input-creator').value = targetName;
        document.getElementById('input-modifier').value = targetName;
        document.getElementById('input-rev-author').value = targetName;
        document.getElementById('input-comment-author').value = targetName;
        document.getElementById('input-total-time').value = '135';
        document.getElementById('input-rev-num').value = '3';
        modeSelect.value = 'range';
        groupFixed.style.display = 'none';
        groupRange.style.display = 'grid';
        showToast('已一键填入合理去破绽推荐值！');
    });

    document.getElementById('btn-submit').addEventListener('click', async () => {
        if (!currentFile) {
            alert('请先选择文件！');
            return;
        }

        const btn = document.getElementById('btn-submit');
        const btnText = document.getElementById('btn-text');
        btn.disabled = true;
        btnText.innerHTML = '<span class="loading"></span> 正在处理并打包...';

        const formData = new FormData();
        formData.append('file', currentFile);
        formData.append('creator', document.getElementById('input-creator').value);
        formData.append('last_modified_by', document.getElementById('input-modifier').value);
        formData.append('revision_num', document.getElementById('input-rev-num').value);
        formData.append('total_time', document.getElementById('input-total-time').value);
        formData.append('company', document.getElementById('input-company').value);
        formData.append('title', document.getElementById('input-title').value);
        formData.append('created_time', document.getElementById('input-created').value);
        formData.append('modified_time', document.getElementById('input-modified').value);
        formData.append('new_rev_author', document.getElementById('input-rev-author').value);
        formData.append('new_comment_author', document.getElementById('input-comment-author').value);
        
        const mode = modeSelect.value;
        formData.append('rev_time_mode', mode);
        if (mode === 'fixed') {
            formData.append('fixed_rev_date', document.getElementById('input-rev-fixed').value);
        } else if (mode === 'range') {
            formData.append('rev_date_start', document.getElementById('input-rev-start').value);
            formData.append('rev_date_end', document.getElementById('input-rev-end').value);
        }

        try {
            const resp = await fetch('/api/modify', {
                method: 'POST',
                body: formData
            });

            if (!resp.ok) {
                const errData = await resp.json();
                alert('处理失败: ' + (errData.error || '未知错误'));
                return;
            }

            const blob = await resp.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            const originalName = currentFile.name.replace(/\.docx$/i, '');
            a.download = `${originalName}_modified.docx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            showToast('🎉 修改完成，已开始下载！');
        } catch (err) {
            alert('导出出错: ' + err.message);
        } finally {
            btn.disabled = false;
            btnText.innerHTML = '🚀 立即导出并下载修改后的 Word 文档';
        }
    });
</script>

</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/inspect', methods=['POST'])
def api_inspect():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "未接收到文件"}), 400
    
    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({"success": False, "error": "仅支持 .docx 格式"}), 400

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name
        file.save(tmp_path)

    try:
        info = docx_metadata_editor.get_docx_info(tmp_path)
        info["filename"] = file.filename
        return jsonify({"success": True, "info": info})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

@app.route('/api/modify', methods=['POST'])
def api_modify():
    if 'file' not in request.files:
        return jsonify({"error": "未接收到文件"}), 400

    file = request.files['file']
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
        in_path = tmp_in.name
        file.save(in_path)

    out_path = in_path + "_out.docx"

    try:
        def get_val(key):
            v = request.form.get(key, '').strip()
            return v if v else None

        def get_int(key):
            v = get_val(key)
            return int(v) if v and v.isdigit() else None

        creator = get_val('creator')
        last_modified_by = get_val('last_modified_by')
        revision_num = get_int('revision_num')
        total_time = get_int('total_time')
        company = get_val('company')
        title = get_val('title')
        created_time = get_val('created_time')
        modified_time = get_val('modified_time')
        new_rev_author = get_val('new_rev_author')
        new_comment_author = get_val('new_comment_author')

        rev_time_mode = request.form.get('rev_time_mode', 'keep')
        fixed_rev_date = None
        rev_date_start = None
        rev_date_end = None

        if rev_time_mode == 'fixed':
            fixed_rev_date = get_val('fixed_rev_date')
        elif rev_time_mode == 'range':
            rev_date_start = get_val('rev_date_start')
            rev_date_end = get_val('rev_date_end')

        docx_metadata_editor.modify_docx_metadata(
            input_path=in_path,
            output_path=out_path,
            creator=creator,
            last_modified_by=last_modified_by,
            revision_num=revision_num,
            total_time_minutes=total_time,
            company=company,
            title=title,
            created_time=created_time,
            modified_time=modified_time,
            new_rev_author=new_rev_author,
            new_comment_author=new_comment_author,
            fixed_rev_date=fixed_rev_date,
            rev_date_start=rev_date_start,
            rev_date_end=rev_date_end,
            sync_fs_time=False
        )

        original_basename = os.path.splitext(file.filename)[0]
        download_name = f"{original_basename}_modified.docx"

        return send_file(out_path, as_attachment=True, download_name=download_name)

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(in_path):
            os.remove(in_path)

def open_browser(port):
    webbrowser.open(f"http://127.0.0.1:{port}")

if __name__ == '__main__':
    PORT = 5050
    print("=" * 65)
    print("🚀 Word 交互式元数据与修订痕迹清洗工具正在启动...")
    print(f"🔗 Project / Author : {WATERMARK_URL}")
    print(f"🌐 本地访问地址      : http://127.0.0.1:{PORT}")
    print("💡 正在为您自动打开默认浏览器，按 Ctrl + C 可退出服务")
    print("=" * 65)
    Timer(1.2, open_browser, args=[PORT]).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)
