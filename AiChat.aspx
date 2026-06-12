<%@ Page Language="C#" AutoEventWireup="true" Inherits="AiChat" CodeFile="AiChat.aspx.cs" %>
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>AI 助理</title>
    <style>
        /* ===== Theme variables ===== */
        :root {
            color-scheme: dark;
            --bg: #0b1220;
            --bg-gradient: radial-gradient(1200px 600px at 20% 0%, #152a52 0%, #0b1220 60%);
            --panel: #0f1b33;
            --panel-elevated: #14233f;
            --text: #f3f7ff;
            --muted: #c6d0e6;
            --border: rgba(255,255,255,0.12);
            --row-hover: rgba(99,179,237,0.10);
            --chip: rgba(99,179,237,0.18);
            --chip-active-bg: rgba(99,179,237,0.22);
            --chip-active-border: rgba(99,179,237,0.55);
            --accent: #63b3ed;
            --accent-strong: #2563eb;
            --tint-med: rgba(255,255,255,0.06);
            --tint-high: rgba(255,255,255,0.10);
            --input-bg: rgba(5,10,20,0.30);
            --warn: #fbbf24;
            --warn-bg: rgba(251,191,36,0.10);
            --warn-border: rgba(251,191,36,0.40);
            --danger: #fca5a5;
        }
        * { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans TC", Arial, sans-serif;
            background: var(--bg-gradient);
            color: var(--text);
            min-height: 100vh;
        }
        .demo-wrap {
            max-width: 900px;
            margin: 0 auto;
            padding: 32px 24px;
        }
        .demo-wrap h1 { font-size: 22px; margin: 0 0 8px; }
        .demo-wrap p  { color: var(--muted); line-height: 1.6; }
        code { background: var(--tint-med); padding: 2px 6px; border-radius: 4px; font-size: 13px; }

        /* ===== Image zoom overlay (shared) ===== */
        .img-overlay {
            position: fixed; inset: 0;
            background: rgba(0,0,0,0.85);
            z-index: 9999;
            display: flex; align-items: center; justify-content: center;
            cursor: zoom-out;
        }
        .img-overlay img { border-radius: 6px; }

        /* ===== AI bubble + panel ===== */
        #aiBubble {
            position: fixed;
            right: 22px;
            bottom: 22px;
            width: 56px;
            height: 56px;
            border-radius: 50%;
            border: none;
            cursor: pointer;
            background: linear-gradient(135deg, var(--accent) 0%, var(--accent-strong) 100%);
            color: #fff;
            font-weight: 700;
            font-size: 16px;
            letter-spacing: 0.5px;
            box-shadow: 0 8px 24px rgba(37, 99, 235, 0.45), 0 2px 6px rgba(0,0,0,0.25);
            z-index: 9999;
            transition: transform .15s ease, box-shadow .15s ease;
        }
        #aiBubble:hover { transform: translateY(-2px); box-shadow: 0 12px 28px rgba(37, 99, 235, 0.55), 0 3px 8px rgba(0,0,0,0.3); }
        #aiBubble.open { transform: scale(0.9); }
        #aiPanel {
            position: fixed;
            right: 22px;
            bottom: 90px;
            width: 800px;
            height: 760px;
            max-width: calc(100vw - 44px);
            max-height: calc(100vh - 120px);
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: 14px;
            box-shadow: 0 20px 50px rgba(0,0,0,0.45), 0 4px 12px rgba(0,0,0,0.2);
            display: flex;
            flex-direction: row;
            overflow: hidden;
            z-index: 9999;
        }
        #aiPanel[hidden] { display: none; }
        .ai-sidebar {
            width: 160px;
            flex: none;
            display: flex;
            flex-direction: column;
            background: var(--panel-elevated);
            border-right: 1px solid var(--border);
        }
        .ai-new-btn {
            margin: 8px;
            padding: 6px 10px;
            border-radius: 8px;
            background: var(--accent);
            color: #fff;
            border: none;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
        }
        .ai-new-btn:hover { background: var(--accent-strong); }
        .ai-sess-list { flex: 1; overflow-y: auto; padding: 0 6px 6px; display: flex; flex-direction: column; gap: 3px; }
        .ai-sess { position: relative; padding: 6px 8px; border-radius: 8px; cursor: pointer; border: 1px solid transparent; }
        .ai-sess:hover { background: var(--row-hover); }
        .ai-sess.active { background: var(--tint-high); border-color: var(--chip-active-border); }
        .ai-sess-title { font-size: 12px; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding-right: 16px; }
        .ai-sess-tags { display: flex; flex-wrap: wrap; gap: 3px; margin-top: 4px; }
        .ai-sess-tags .tag { font-size: 10px; padding: 1px 6px; border-radius: 999px; background: var(--chip); color: var(--muted); }
        .ai-sess-del {
            position: absolute; top: 4px; right: 4px;
            background: transparent; border: none; color: var(--muted);
            cursor: pointer; font-size: 14px; line-height: 1; padding: 0 4px;
            border-radius: 4px; opacity: 0; transition: opacity .12s;
        }
        .ai-sess:hover .ai-sess-del, .ai-sess.active .ai-sess-del { opacity: 1; }
        .ai-sess-del:hover { background: var(--warn-bg); color: var(--danger); }
        .ai-main { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }
        .ai-head {
            display: flex; align-items: center; justify-content: space-between;
            padding: 10px 14px;
            background: linear-gradient(135deg, rgba(99,179,237,0.15), rgba(37,99,235,0.10));
            border-bottom: 1px solid var(--border);
        }
        .ai-title-input {
            flex: 1; min-width: 0;
            background: transparent; border: 1px solid transparent;
            color: var(--text); font-weight: 600; font-size: 14px;
            outline: none; padding: 4px 8px; border-radius: 6px;
        }
        .ai-title-input:hover { background: var(--tint-med); }
        .ai-title-input:focus { background: var(--tint-med); border-color: var(--accent); }
        #aiClose {
            background: transparent; border: none; color: var(--muted);
            font-size: 22px; line-height: 1; cursor: pointer; padding: 2px 6px;
            border-radius: 6px;
        }
        #aiClose:hover { background: var(--tint-med); color: var(--text); }
        .ai-tags-row {
            display: flex; flex-wrap: wrap; align-items: center; gap: 4px;
            padding: 6px 12px;
            border-bottom: 1px solid var(--border);
            background: var(--panel-elevated);
            min-height: 30px;
        }
        .ai-tag-chip {
            display: inline-flex; align-items: center; gap: 4px;
            background: var(--chip); color: var(--text);
            border: 1px solid var(--chip-active-border);
            border-radius: 999px; padding: 2px 4px 2px 8px; font-size: 11px;
        }
        .ai-tag-chip .x { background: transparent; border: none; color: var(--muted); cursor: pointer; padding: 0 4px; line-height: 1; border-radius: 4px; }
        .ai-tag-chip .x:hover { background: var(--warn-bg); color: var(--danger); }
        .ai-tag-add { background: transparent; border: 1px dashed var(--border); color: var(--muted); border-radius: 999px; padding: 2px 10px; font-size: 11px; cursor: pointer; }
        .ai-tag-add:hover { color: var(--text); border-color: var(--accent); }
        .ai-msgs { flex: 1; overflow-y: auto; padding: 14px; display: flex; flex-direction: column; gap: 10px; }
        .ai-msg { max-width: 86%; padding: 8px 12px; border-radius: 12px; font-size: 13px; line-height: 1.5; white-space: pre-wrap; word-wrap: break-word; }
        .ai-msg.user { align-self: flex-end; background: var(--accent-strong); color: #fff; border-bottom-right-radius: 4px; }
        .ai-msg.assistant { align-self: flex-start; background: var(--tint-med); color: var(--text); border-bottom-left-radius: 4px; }
        .ai-msg.error { align-self: stretch; background: var(--warn-bg); border: 1px solid var(--warn-border); color: var(--warn); font-size: 12px; }
        .ai-msg.typing { align-self: flex-start; background: var(--tint-med); color: var(--muted); }
        .ai-msg.typing .dot { display: inline-block; width: 6px; height: 6px; border-radius: 50%; background: var(--muted); margin: 0 2px; animation: ai-blink 1.2s infinite; }
        .ai-msg.typing .dot:nth-child(2) { animation-delay: .2s; }
        .ai-msg.typing .dot:nth-child(3) { animation-delay: .4s; }
        @keyframes ai-blink { 0%, 80%, 100% { opacity: 0.25; } 40% { opacity: 1; } }
        .ai-msg-img { margin-top: 6px; max-width: 220px; max-height: 160px; border-radius: 8px; display: block; border: 1px solid rgba(255,255,255,0.15); cursor: zoom-in; }
        .ai-input-wrap {
            border-top: 1px solid var(--border);
            padding: 10px;
            display: flex; gap: 8px;
            background: var(--panel-elevated);
        }
        .ai-icon-btn {
            background: var(--tint-med); border: 1px solid var(--border); color: var(--text);
            border-radius: 8px; width: 38px; height: 38px;
            display: inline-flex; align-items: center; justify-content: center;
            cursor: pointer; flex: none;
        }
        .ai-icon-btn:hover { background: var(--tint-high); border-color: var(--accent); }
        .ai-preview {
            display: flex; align-items: center; gap: 8px;
            padding: 8px 10px;
            background: var(--panel-elevated);
            border-top: 1px solid var(--border);
        }
        .ai-preview[hidden] { display: none; }
        .ai-preview img { max-height: 56px; max-width: 90px; border-radius: 6px; border: 1px solid var(--border); object-fit: contain; background: #000; cursor: zoom-in; }
        .ai-preview .ai-prev-meta { color: var(--muted); font-size: 12px; flex: 1; }
        .ai-preview .ai-prev-remove { background: var(--tint-med); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 4px 10px; cursor: pointer; font-size: 12px; }
        .ai-preview .ai-prev-remove:hover { background: var(--warn-bg); color: var(--warn); border-color: var(--warn-border); }
        #aiInput {
            flex: 1; min-height: 38px; max-height: 120px;
            resize: none; padding: 8px 10px; border-radius: 8px;
            border: 1px solid var(--border);
            background: var(--input-bg);
            color: var(--text);
            font-family: inherit; font-size: 13px;
            outline: none;
        }
        #aiInput:focus { border-color: var(--accent); }
        #aiSend {
            white-space: nowrap;
            padding: 8px 16px; border-radius: 8px;
            background: var(--accent); color: #fff; border: 1px solid var(--accent);
            cursor: pointer; font-weight: 600;
        }
        #aiSend:hover { background: var(--accent-strong); border-color: var(--accent-strong); }
        #aiSend:disabled { opacity: 0.5; cursor: not-allowed; }
    </style>
</head>
<body>
    <!-- ============================================================
         Demo page chrome. In a real integration this whole .demo-wrap
         block is unnecessary — only the bubble + panel + script below
         need to live in your host page.
         ============================================================ -->
    <div class="demo-wrap">
        <h1>AI 助理</h1>
        <p>這是公版 AI 問答視窗 — 純前端 + 一個 ASP.NET 後端 proxy,複製整個 <code>ai-chat-widget</code>
        資料夾到任何 IIS 站台、把 <code>web.config</code> 的 <code>AiApiKey</code> 填上,就可以用了。</p>
        <p>右下角的圓鈕點開,就是聊天視窗。多 session、附圖、Markdown、貼上都支援。</p>
    </div>

    <!-- ============================================================
         AI assistant widget — bubble + panel.
         Copy this block (plus the <style> above and <script> below)
         into your own page to embed the widget there.
         ============================================================ -->
    <button id="aiBubble" type="button" title="AI 助理">AI</button>
    <div id="aiPanel" hidden>
        <div class="ai-sidebar">
            <button type="button" id="aiNewChat" class="ai-new-btn">+ 新聊天</button>
            <div id="aiSessionList" class="ai-sess-list"></div>
        </div>
        <div class="ai-main">
            <div class="ai-head">
                <input type="text" id="aiTitle" class="ai-title-input" placeholder="對話標題" />
                <button type="button" id="aiClose" title="關閉">×</button>
            </div>
            <div class="ai-tags-row" id="aiTagsRow"></div>
            <div id="aiMessages" class="ai-msgs"></div>
            <div id="aiPreview" class="ai-preview" hidden>
                <img id="aiPreviewImg" alt="附加圖片" />
                <span class="ai-prev-meta" id="aiPreviewMeta"></span>
                <button type="button" class="ai-prev-remove" id="aiPreviewRemove" title="移除附圖">移除</button>
            </div>
            <div class="ai-input-wrap">
                <button type="button" id="aiAttach" class="ai-icon-btn" title="附加圖片(可貼上)">
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
                    </svg>
                </button>
                <input type="file" id="aiFile" accept="image/*" hidden />
                <textarea id="aiInput" placeholder="輸入問題,可附圖。Enter 送出,Shift+Enter 換行" rows="2"></textarea>
                <button type="button" id="aiSend">送出</button>
            </div>
        </div>
    </div>

    <script>
        // ============================================================
        //  AI Chat Widget — standalone, no host-page dependencies.
        //  Embed: include the <style>, the bubble/panel markup, and
        //  this <script> into any HTML page. Configure proxyUrl below
        //  if your proxy lives at a different path.
        // ============================================================
        (function aiChatWidget() {
            const CONFIG = (window.AiChatConfig || {});
            const PROXY_URL = CONFIG.proxyUrl || 'AiChat.aspx?op=chat';
            const STORAGE_KEY = CONFIG.storageKey || 'aiChat.sessions.v1';
            const ACTIVE_KEY  = CONFIG.activeKey  || 'aiChat.activeId.v1';

            // ---------- DOM ----------
            const bubble = document.getElementById('aiBubble');
            const panel = document.getElementById('aiPanel');
            const closeBtn = document.getElementById('aiClose');
            const msgs = document.getElementById('aiMessages');
            const input = document.getElementById('aiInput');
            const sendBtn = document.getElementById('aiSend');
            const attachBtn = document.getElementById('aiAttach');
            const fileInput = document.getElementById('aiFile');
            const previewBox = document.getElementById('aiPreview');
            const previewImg = document.getElementById('aiPreviewImg');
            const previewMeta = document.getElementById('aiPreviewMeta');
            const previewRemove = document.getElementById('aiPreviewRemove');
            const newChatBtn = document.getElementById('aiNewChat');
            const sessionListEl = document.getElementById('aiSessionList');
            const titleInput = document.getElementById('aiTitle');
            const tagsRow = document.getElementById('aiTagsRow');

            let sessions = [];
            let activeId = null;
            let busy = false;
            let attachedImage = null;

            // ---------- Image zoom overlay ----------
            let overlay = null;
            function openOverlay(src) {
                closeOverlay();
                overlay = document.createElement('div');
                overlay.className = 'img-overlay';
                const img = document.createElement('img');
                const sizeIt = () => {
                    const nw = img.naturalWidth, nh = img.naturalHeight;
                    if (!nw || !nh) return;
                    const ZOOM = 2;
                    const scale = Math.min(ZOOM, (innerWidth * 0.95) / nw, (innerHeight * 0.95) / nh);
                    img.style.width = (nw * scale) + 'px';
                    img.style.height = (nh * scale) + 'px';
                };
                img.onload = sizeIt;
                img.src = src;
                if (img.complete) sizeIt();
                overlay.appendChild(img);
                overlay.addEventListener('click', closeOverlay);
                document.body.appendChild(overlay);
            }
            function closeOverlay() {
                if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
                overlay = null;
            }
            document.addEventListener('keydown', e => { if (e.key === 'Escape') closeOverlay(); });

            // ---------- Sessions (localStorage) ----------
            function uid() {
                return 'sess-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6);
            }
            function loadSessions() {
                try {
                    sessions = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
                    activeId = localStorage.getItem(ACTIVE_KEY) || null;
                } catch (e) { sessions = []; activeId = null; }
                if (!Array.isArray(sessions)) sessions = [];
                if (sessions.length === 0) createSession(false);
                else if (!sessions.find(s => s.id === activeId)) activeId = sessions[0].id;
            }
            function saveSessions() {
                try {
                    localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
                    if (activeId) localStorage.setItem(ACTIVE_KEY, activeId);
                } catch (e) { /* quota */ }
            }
            function getActive() { return sessions.find(s => s.id === activeId); }
            function createSession(rerender) {
                const s = { id: uid(), title: '新對話', tags: [], history: [], createdAt: Date.now(), updatedAt: Date.now() };
                sessions.unshift(s);
                activeId = s.id;
                saveSessions();
                if (rerender !== false) {
                    renderSessionList();
                    renderActiveSession();
                    clearAttachment();
                    input.focus();
                }
            }
            function switchSession(id) {
                if (busy) return;
                activeId = id;
                saveSessions();
                renderSessionList();
                renderActiveSession();
                clearAttachment();
            }
            function deleteSession(id) {
                if (busy) return;
                const s = sessions.find(x => x.id === id);
                if (!s) return;
                if (!confirm('刪除這個對話?\n「' + (s.title || '新對話') + '」')) return;
                sessions = sessions.filter(x => x.id !== id);
                if (activeId === id) activeId = sessions.length > 0 ? sessions[0].id : null;
                if (sessions.length === 0) createSession(false);
                saveSessions();
                renderSessionList();
                renderActiveSession();
            }
            function maybeAutoTitle(s, text) {
                if (s.title === '新對話' && text) {
                    const t = text.replace(/\s+/g, ' ').trim();
                    s.title = t.length > 26 ? t.slice(0, 26) + '...' : t;
                    titleInput.value = s.title;
                }
            }

            // ---------- Rendering ----------
            function renderSessionList() {
                sessionListEl.innerHTML = '';
                sessions.forEach(s => {
                    const item = document.createElement('div');
                    item.className = 'ai-sess' + (s.id === activeId ? ' active' : '');
                    const title = document.createElement('div');
                    title.className = 'ai-sess-title';
                    title.textContent = s.title || '新對話';
                    item.appendChild(title);
                    if (s.tags && s.tags.length > 0) {
                        const tagsEl = document.createElement('div');
                        tagsEl.className = 'ai-sess-tags';
                        s.tags.forEach(t => {
                            const tag = document.createElement('span');
                            tag.className = 'tag';
                            tag.textContent = t;
                            tagsEl.appendChild(tag);
                        });
                        item.appendChild(tagsEl);
                    }
                    const del = document.createElement('button');
                    del.type = 'button';
                    del.className = 'ai-sess-del';
                    del.title = '刪除此對話';
                    del.textContent = '×';
                    del.addEventListener('click', e => { e.stopPropagation(); deleteSession(s.id); });
                    item.appendChild(del);
                    item.addEventListener('click', () => switchSession(s.id));
                    sessionListEl.appendChild(item);
                });
            }
            function renderActiveSession() {
                const s = getActive();
                msgs.innerHTML = '';
                if (!s) return;
                titleInput.value = s.title || '';
                renderActiveTags();
                if (s.history.length === 0) {
                    appendMessage('assistant', '你好,有什麼可以幫你的?');
                } else {
                    s.history.forEach(m => {
                        if (m.role === 'user') {
                            if (Array.isArray(m.content)) {
                                const t = (m.content.find(p => p.type === 'text') || {}).text || '';
                                const u = (m.content.find(p => p.type === 'image_url') || {}).image_url;
                                appendUserMessage(t, u ? u.url : null);
                            } else {
                                appendUserMessage(m.content, null);
                            }
                        } else if (m.role === 'assistant') {
                            appendMessage('assistant', m.content);
                        }
                    });
                }
            }
            function renderActiveTags() {
                tagsRow.innerHTML = '';
                const s = getActive();
                if (!s) return;
                (s.tags || []).forEach(t => {
                    const chip = document.createElement('span');
                    chip.className = 'ai-tag-chip';
                    chip.textContent = t;
                    const x = document.createElement('button');
                    x.type = 'button';
                    x.className = 'x';
                    x.title = '移除';
                    x.textContent = '×';
                    x.addEventListener('click', () => {
                        s.tags = s.tags.filter(v => v !== t);
                        s.updatedAt = Date.now();
                        saveSessions();
                        renderActiveTags();
                        renderSessionList();
                    });
                    chip.appendChild(x);
                    tagsRow.appendChild(chip);
                });
                const add = document.createElement('button');
                add.type = 'button';
                add.className = 'ai-tag-add';
                add.textContent = '+ 標籤';
                add.addEventListener('click', () => {
                    const raw = prompt('輸入標籤(可用逗號分隔多個):');
                    if (!raw) return;
                    const parts = raw.split(/[,,;;]/).map(x => x.trim()).filter(Boolean);
                    if (parts.length === 0) return;
                    s.tags = s.tags || [];
                    parts.forEach(p => { if (!s.tags.includes(p)) s.tags.push(p); });
                    s.updatedAt = Date.now();
                    saveSessions();
                    renderActiveTags();
                    renderSessionList();
                });
                tagsRow.appendChild(add);
            }
            function appendMessage(role, text, cls) {
                const div = document.createElement('div');
                div.className = 'ai-msg ' + (cls || role);
                div.textContent = text;
                msgs.appendChild(div);
                msgs.scrollTop = msgs.scrollHeight;
                return div;
            }
            function appendUserMessage(text, imgUrl) {
                const div = document.createElement('div');
                div.className = 'ai-msg user';
                if (text) {
                    const t = document.createElement('div');
                    t.textContent = text;
                    div.appendChild(t);
                }
                if (imgUrl) {
                    const im = document.createElement('img');
                    im.className = 'ai-msg-img';
                    im.src = imgUrl;
                    im.title = '點擊放大';
                    im.addEventListener('click', () => openOverlay(imgUrl));
                    div.appendChild(im);
                }
                msgs.appendChild(div);
                msgs.scrollTop = msgs.scrollHeight;
                return div;
            }
            function appendTyping() {
                const div = document.createElement('div');
                div.className = 'ai-msg typing';
                div.innerHTML = '<span class="dot"></span><span class="dot"></span><span class="dot"></span>';
                msgs.appendChild(div);
                msgs.scrollTop = msgs.scrollHeight;
                return div;
            }

            // ---------- Title rename ----------
            titleInput.addEventListener('blur', () => {
                const s = getActive();
                if (!s) return;
                const v = titleInput.value.trim() || '新對話';
                if (v !== s.title) {
                    s.title = v;
                    s.updatedAt = Date.now();
                    saveSessions();
                    renderSessionList();
                }
            });
            titleInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); titleInput.blur(); }
            });

            // ---------- Open / close panel ----------
            function open() {
                panel.hidden = false;
                bubble.classList.add('open');
                renderSessionList();
                renderActiveSession();
                setTimeout(() => input.focus(), 0);
            }
            function close() {
                panel.hidden = true;
                bubble.classList.remove('open');
            }
            bubble.addEventListener('click', () => panel.hidden ? open() : close());
            closeBtn.addEventListener('click', close);

            // ---------- New chat ----------
            newChatBtn.addEventListener('click', () => {
                if (busy) return;
                createSession(true);
            });

            // ---------- Image attach ----------
            function downsizeImage(dataUrl) {
                return new Promise((resolve) => {
                    const img = new Image();
                    img.onload = () => {
                        const MAX = 1024;
                        const longest = Math.max(img.width, img.height);
                        if (longest <= MAX) return resolve(dataUrl);
                        const ratio = MAX / longest;
                        const w = Math.round(img.width * ratio);
                        const h = Math.round(img.height * ratio);
                        const c = document.createElement('canvas');
                        c.width = w; c.height = h;
                        c.getContext('2d').drawImage(img, 0, 0, w, h);
                        try { resolve(c.toDataURL('image/jpeg', 0.85)); }
                        catch (e) { resolve(dataUrl); }
                    };
                    img.onerror = () => resolve(dataUrl);
                    img.src = dataUrl;
                });
            }
            function approxKB(dataUrl) {
                const i = dataUrl.indexOf(',');
                const b64 = i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
                return Math.round((b64.length * 3 / 4) / 1024);
            }
            function setAttachedFromFile(file) {
                if (!file || !file.type || file.type.indexOf('image/') !== 0) return;
                if (file.size > 20 * 1024 * 1024) { alert('圖片太大(>20MB),請挑小一點的'); return; }
                const reader = new FileReader();
                reader.onload = async () => {
                    const small = await downsizeImage(reader.result);
                    attachedImage = small;
                    previewImg.src = small;
                    previewMeta.textContent = '已附圖 (' + approxKB(small) + ' KB)';
                    previewBox.hidden = false;
                    input.focus();
                };
                reader.readAsDataURL(file);
            }
            function clearAttachment() {
                attachedImage = null;
                previewBox.hidden = true;
                previewImg.removeAttribute('src');
                previewMeta.textContent = '';
            }
            attachBtn.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', () => {
                const f = fileInput.files && fileInput.files[0];
                fileInput.value = '';
                if (f) setAttachedFromFile(f);
            });
            previewRemove.addEventListener('click', clearAttachment);
            previewImg.addEventListener('click', () => { if (previewImg.src) openOverlay(previewImg.src); });
            input.addEventListener('paste', (ev) => {
                const items = (ev.clipboardData && ev.clipboardData.items) || [];
                for (const it of items) {
                    if (it.kind === 'file' && it.type && it.type.indexOf('image/') === 0) {
                        const blob = it.getAsFile();
                        if (blob) { ev.preventDefault(); setAttachedFromFile(blob); return; }
                    }
                }
            });

            // ---------- Send ----------
            async function send() {
                if (busy) return;
                const text = input.value.trim();
                const img = attachedImage;
                if (!text && !img) return;
                const s = getActive();
                if (!s) return;

                let userContent;
                let displayText = text;
                if (img) {
                    if (!displayText) displayText = '請看這張圖';
                    userContent = [
                        { type: 'text', text: displayText },
                        { type: 'image_url', image_url: { url: img } }
                    ];
                } else {
                    userContent = text;
                }

                input.value = '';
                clearAttachment();
                if (s.history.length === 0) maybeAutoTitle(s, displayText);
                appendUserMessage(displayText, img);
                s.history.push({ role: 'user', content: userContent });
                s.updatedAt = Date.now();
                saveSessions();
                renderSessionList();
                busy = true;
                sendBtn.disabled = true;
                const typing = appendTyping();
                try {
                    const res = await fetch(PROXY_URL, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json; charset=utf-8' },
                        body: JSON.stringify({ messages: s.history })
                    });
                    const data = await res.json();
                    typing.remove();
                    if (!res.ok || data.ok === false) {
                        const err = (data && (data.error || data.detail)) || ('HTTP ' + res.status);
                        appendMessage('assistant', '錯誤: ' + err, 'error');
                        s.history.pop();
                        saveSessions();
                        return;
                    }
                    const reply = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '(空回應)';
                    appendMessage('assistant', reply);
                    s.history.push({ role: 'assistant', content: reply });
                    s.updatedAt = Date.now();
                    saveSessions();
                } catch (e) {
                    typing.remove();
                    appendMessage('assistant', '網路錯誤: ' + e.message, 'error');
                    s.history.pop();
                    saveSessions();
                } finally {
                    busy = false;
                    sendBtn.disabled = false;
                    input.focus();
                }
            }
            sendBtn.addEventListener('click', send);
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
            });

            // ---------- Boot ----------
            loadSessions();
            renderSessionList();
        })();
    </script>
</body>
</html>
