// Bridge: embed AiChat widget into ASM Daily and allow AI to analyze current page data.
// This script contains the AiChat widget JS (from AiChat.aspx) plus a small
// injection that adds an "分析目前資料" button.

(function aiChatWidget() {
  const CONFIG = (window.AiChatConfig || {});
  const PROXY_URL = CONFIG.proxyUrl || 'AiChat.aspx?op=chat';
  const STORAGE_KEY = CONFIG.storageKey || 'aiChat.sessions.v1';
  const ACTIVE_KEY = CONFIG.activeKey || 'aiChat.activeId.v1';

  // ---------- DOM ----------
  const bubble = document.getElementById('aiBubble');
  const panel = document.getElementById('aiPanel');
  const closeBtn = document.getElementById('aiClose');
  const headEl = panel ? panel.querySelector('.ai-head') : null;
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
      input && input.focus();
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
      if (titleInput) titleInput.value = s.title;
    }
  }

  // ---------- Rendering ----------
  function renderSessionList() {
    if (!sessionListEl) return;
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
    if (!msgs) return;
    msgs.innerHTML = '';
    if (!s) return;
    if (titleInput) titleInput.value = s.title || '';
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
    if (!tagsRow) return;
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

    // Bridge button: analyze current ASM table data
    const analyze = document.createElement('button');
    analyze.type = 'button';
    analyze.className = 'ai-tag-add';
    analyze.textContent = '分析目前資料';
    analyze.title = '請 AI 根據目前頁面即時資料做重點分析';
    analyze.addEventListener('click', () => {
      // Page data is now auto-attached on every send, so this only needs to
      // ask the question; no need to dump the whole table into the input box.
      const snapshot = window.getAsmDailySnapshot ? window.getAsmDailySnapshot() : null;
      if (!snapshot) {
        alert('目前沒有可分析的資料（請先等 Excel 讀取完成）。');
        return;
      }
      if (input) {
        input.value = '請根據目前頁面的 ASM Daily 即時資料,條列需要優先關注的 chamber/EQPID 與原因(請引用具體欄位與數值),並針對每個異常給建議 action。';
      }
      open();
      if (input) input.focus();
    });
    tagsRow.appendChild(analyze);
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
  if (titleInput) {
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
  }

  // ---------- Fullscreen + click-outside close ----------
  const FS_KEY = (CONFIG.fullscreenKey || 'aiChat.fullscreen.v1');
  let isFullscreen = false;

  function setFullscreen(on) {
    isFullscreen = !!on;
    try { localStorage.setItem(FS_KEY, isFullscreen ? '1' : '0'); } catch (e) { }
    if (!panel) return;

    if (isFullscreen) {
      panel.classList.add('ai-fullscreen');
    } else {
      panel.classList.remove('ai-fullscreen');
    }

    // update button label
    const btn = document.getElementById('aiFullscreen');
    if (btn) btn.textContent = isFullscreen ? '🗗' : '🗖';
  }

  function toggleFullscreen() { setFullscreen(!isFullscreen); }

  // inject fullscreen button into header
  (function ensureFullscreenButton() {
    if (!headEl) return;
    if (document.getElementById('aiFullscreen')) return;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.id = 'aiFullscreen';
    btn.title = '全螢幕';
    btn.textContent = '🗖';
    btn.style.background = 'transparent';
    btn.style.border = 'none';
    btn.style.color = 'var(--muted)';
    btn.style.fontSize = '18px';
    btn.style.cursor = 'pointer';
    btn.style.padding = '2px 6px';
    btn.style.borderRadius = '6px';
    btn.addEventListener('mouseenter', () => { btn.style.background = 'var(--tint-med)'; btn.style.color = 'var(--text)'; });
    btn.addEventListener('mouseleave', () => { btn.style.background = 'transparent'; btn.style.color = 'var(--muted)'; });
    btn.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); toggleFullscreen(); });

    // place it just before close button
    headEl.insertBefore(btn, closeBtn);

    // restore persisted fullscreen state
    try { isFullscreen = (localStorage.getItem(FS_KEY) === '1'); } catch (e) { isFullscreen = false; }
    setFullscreen(isFullscreen);
  })();

  // Click outside panel auto close
  function handleDocPointerDown(ev) {
    if (!panel || panel.hidden) return;
    const t = ev.target;
    // click on bubble toggles itself, don't auto-close here
    if (bubble && (t === bubble || bubble.contains(t))) return;
    if (panel.contains(t)) return;
    close();
  }
  document.addEventListener('pointerdown', handleDocPointerDown, true);

  // ESC: close overlay handled earlier; here close panel / exit fullscreen
  document.addEventListener('keydown', (e) => {
    if (!panel || panel.hidden) return;
    if (e.key === 'Escape') {
      if (isFullscreen) setFullscreen(false);
      else close();
    }
  });

  // ---------- Open / close panel ----------
  function open() {
    if (!panel) return;
    panel.hidden = false;
    if (bubble) bubble.classList.add('open');
    renderSessionList();
    renderActiveSession();
    setTimeout(() => input && input.focus(), 0);
  }
  function close() {
    if (!panel) return;
    panel.hidden = true;
    if (bubble) bubble.classList.remove('open');
  }
  if (bubble) bubble.addEventListener('click', () => panel.hidden ? open() : close());
  if (closeBtn) closeBtn.addEventListener('click', close);

  // ---------- New chat ----------
  if (newChatBtn) newChatBtn.addEventListener('click', () => {
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
      if (previewImg) previewImg.src = small;
      if (previewMeta) previewMeta.textContent = '已附圖 (' + approxKB(small) + ' KB)';
      if (previewBox) previewBox.hidden = false;
      input && input.focus();
    };
    reader.readAsDataURL(file);
  }
  function clearAttachment() {
    attachedImage = null;
    if (previewBox) previewBox.hidden = true;
    if (previewImg) previewImg.removeAttribute('src');
    if (previewMeta) previewMeta.textContent = '';
  }
  if (attachBtn) attachBtn.addEventListener('click', () => fileInput && fileInput.click());
  if (fileInput) fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    fileInput.value = '';
    if (f) setAttachedFromFile(f);
  });
  if (previewRemove) previewRemove.addEventListener('click', clearAttachment);
  if (previewImg) previewImg.addEventListener('click', () => { if (previewImg.src) openOverlay(previewImg.src); });
  if (input) input.addEventListener('paste', (ev) => {
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
    const text = input ? input.value.trim() : '';
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

    if (input) input.value = '';
    clearAttachment();
    if (s.history.length === 0) maybeAutoTitle(s, displayText);
    appendUserMessage(displayText, img);
    s.history.push({ role: 'user', content: userContent });
    s.updatedAt = Date.now();
    saveSessions();
    renderSessionList();
    busy = true;
    if (sendBtn) sendBtn.disabled = true;
    const typing = appendTyping();
    try {
      const res = await fetch(PROXY_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ messages: buildOutboundMessages(s) })
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
      if (sendBtn) sendBtn.disabled = false;
      input && input.focus();
    }
  }
  if (sendBtn) sendBtn.addEventListener('click', send);
  if (input) input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
  });

  // ---------- Boot ----------
  loadSessions();
  renderSessionList();

  // Expose helpers for bridge
  window.__aiChatOpen = open;
  window.__aiChatSetInputAndOpen = (text) => { if (input) input.value = text || ''; open(); };

  // ===================================================================
  //  Auto-context: every outbound request carries (1) ASM TOOL domain
  //  knowledge and (2) a fresh snapshot of the page's live data, so the
  //  AI can answer questions about the dashboard without the user having
  //  to paste anything. Built fresh on each send and NOT stored in the
  //  saved history (keeps sessions small, always reflects latest data).
  // ===================================================================

  // Static domain knowledge derived from the dashboard's own logic
  // (column meanings + the exact alert thresholds used in ASM Daily.js).
  function buildAsmKnowledge() {
    return [
      '【ASM TOOL 知識庫】(回答時請依此解讀欄位與門檻)',
      '本頁是 ASM 機台(chamber)每日監控儀表板。EQPID 形如 HKG-A01B,末碼 A/B/C/D 對應 chamber 1/2/3/4。',
      '欄位語意與警示門檻:',
      '- Cycle: coating 累積 cycle 數;顯示為「cycle (cycle - UD_PM)」。',
      '- PM: 上次 PM 日期 (MM/DD)。',
      '- Coat: S5 coating 剩餘量;< 3.1 視為偏低(紅字),需留意換 coating。',
      '- State: 機台狀態。SB=standby、PD=production(生產中)、UD=unscheduled down(非預期停機)、UD_M/UD_MO=維修中、SD=scheduled down、SD_M、NS_OFF=關機、EG=工程。',
      '- HPC+/M22/R22/Ehv/ULP: 該 chamber 對各製程的資格旗標,有資格為「V」。HPC+/M22/R22 合計為「HPC+/M22/R22 Chamber」數;Ehv/ULP 合計為「eHV/ULP Chamber」數。',
      '- M22 Inhib / HPC Inhib: inhibit 狀態(灰底代表 inhibit=2)。',
      '- S6_Min: S6 最近一筆量測;> 141 偏高(藍字)。',
      '- W/C: 晶圓計數 (wafer count);> 5000 偏高(藍字),接近需保養。',
      '- Oxide: 最新 oxide 厚度值。目標約 9.55。9.48–9.70=偏低(金)、≥9.70=高(綠)、>10.20=過高(藍)。',
      '- S7 Remain: S7 vessel 剩餘百分比;< 15% 嚴重不足(橘底紅字);< 30% 且距上次秤重 > 60 天(粉紅底)需補。',
      '- S7 Temp: S7 溫度;> 175.9 偏高(紅字)。',
      '- UD_PM: 預估/實際 PM 對應 cycle。階差: Base 分頁結果;> 0.9 需注意。',
      '- Delta: oxide 最新與前一筆差值;> 0.2 上升(藍)、< -0.2 下降(紅);|Δ| > 0.42 變動過大(有底色)。',
      '- Last PA: 最近 3 筆 PA mean;單值 > 5 偏高(紅字)。CH 欄可點開 PA 趨勢圖,EQPID 可點開 OXIDE 趨勢圖。',
      '- 製程欄淺綠/淺黃底色: 代表該 chamber 目前條件(State 為 PD/SB、Oxide 落在該製程區間、無 V、Delta 在 ±0.419 內)建議可做該 coating。',
      '判讀原則: 優先關注 State 異常(UD/UD_M)、Coat<3.1、S7 Remain 偏低、Oxide 超標、Delta 過大、W/C 偏高者;回答請引用具體 EQPID 與數值。'
    ].join('\n');
  }

  // Compact live snapshot of the page (summary + tile states + full table).
  // PA/OXIDE time-series are omitted here to keep每則請求精簡; ask the user to
  // open the chart for deep trend analysis when needed.
  function buildLiveDataContext(snapshot) {
    if (!snapshot) return '【目前頁面資料】(尚無資料,Excel 可能還沒讀取完成)';
    const lines = ['【目前頁面即時資料】'];

    const sm = snapshot.summary || {};
    if (sm.today) lines.push(sm.today);
    if (sm.hmrSummary) lines.push(sm.hmrSummary);
    if (sm.euSummary) lines.push(sm.euSummary);
    if (sm.fileInfo) lines.push(sm.fileInfo);

    const tiles = snapshot.tiles || [];
    if (tiles.length) {
      lines.push('');
      lines.push('Tile 狀態 (eqpid=state): ' +
        tiles.map(t => `${t.eqpid}=${t.class || 'NONE'}`).join(', '));
    }

    const rows = snapshot.rows || [];
    lines.push('');
    lines.push('表格 (' + rows.length + ' 台):');
    if (rows.length) {
      const cols = Object.keys(rows[0]);
      lines.push('欄位: ' + cols.join(' | '));
      rows.forEach(r => {
        lines.push(cols.map(k => {
          const v = r ? r[k] : '';
          return (v == null) ? '' : String(v).trim().replace(/\s+/g, ' ');
        }).join(' | '));
      });
    }
    return lines.join('\n');
  }

  // Build the messages array actually sent to the proxy: a fresh context
  // system message (knowledge + live data) followed by the chat history.
  function buildOutboundMessages(s) {
    let context = buildAsmKnowledge();
    try {
      const snap = window.getAsmDailySnapshot ? window.getAsmDailySnapshot() : null;
      context += '\n\n' + buildLiveDataContext(snap);
    } catch (e) {
      context += '\n\n【目前頁面資料】(讀取失敗: ' + e.message + ')';
    }
    const ctxMsg = { role: 'system', content: context };
    return [ctxMsg].concat((s && s.history) ? s.history : []);
  }

  function buildAnalysisPrompt(snapshot) {
    // Keep prompt compact; include key columns only.
    const lines = [];
    lines.push('你是資深 ASM 製程/設備工程師助理。');
    lines.push('請根據我提供的 ASM Daily 快照做分析：');
    lines.push('1) 先用條列列出需要優先關注的 chamber/EQPID 與原因');
    lines.push('2) 針對異常給建議 action（例如：確認哪個指標/圖/欄位）');
    lines.push('3) 如果資料不足，請直接列出你要我補的欄位/截圖');
    lines.push('');

    lines.push('【摘要】');
    if (snapshot.summary) {
      if (snapshot.summary.today) lines.push('- Today: ' + snapshot.summary.today);
      if (snapshot.summary.hmrSummary) lines.push('- ' + snapshot.summary.hmrSummary);
      if (snapshot.summary.euSummary) lines.push('- ' + snapshot.summary.euSummary);
      if (snapshot.summary.fileInfo) lines.push('- ' + snapshot.summary.fileInfo);
      if (snapshot.summary.status) lines.push('- 狀態: ' + snapshot.summary.status);
    }

    lines.push('');
    lines.push('【Tile 區塊顏色】(eqpid -> class)');
    const tiles = snapshot.tiles || [];
    if (tiles.length === 0) {
      lines.push('(無 tile 資料)');
    } else {
      tiles.forEach(t => {
        lines.push(`${t.eqpid} | ${t.class || ''}`);
      });
    }

    lines.push('');
    lines.push('【PA 原始資料】(paDataByProcessUnit)');
    // For prompt size, include only tools that exist in table snapshot
    const eqSet = new Set((snapshot.rows || []).map(r => (r && r.eqpid) ? String(r.eqpid).trim() : '').filter(Boolean));
    const paAll = snapshot.paDataByProcessUnit || {};
    const paKeys = Object.keys(paAll).filter(k => eqSet.has(k));
    if (paKeys.length === 0) {
      lines.push('(無 PA 資料)');
    } else {
      paKeys.forEach(eqpid => {
        const list = paAll[eqpid] || [];
        // cap points per eqpid to avoid huge prompts
        const tail = list.slice(-60);
        lines.push(`- ${eqpid}: ${tail.length} pts (show last ${tail.length})`);
        tail.forEach(p => {
          lines.push(`  ${p.time || ''} | ${p.mean}`);
        });
      });
    }

    lines.push('');
    lines.push('【OXIDE 原始資料】(oxideDataByEqpid)');
    const oxAll = snapshot.oxideDataByEqpid || {};
    const oxKeys = Object.keys(oxAll).filter(k => eqSet.has(k));
    if (oxKeys.length === 0) {
      lines.push('(無 OXIDE 資料)');
    } else {
      oxKeys.forEach(eqpid => {
        const list = oxAll[eqpid] || [];
        const tail = list.slice(-60);
        lines.push(`- ${eqpid}: ${tail.length} pts (show last ${tail.length})`);
        tail.forEach(p => {
          lines.push(`  ${p.time || ''} | ${p.mean}`);
        });
      });
    }

    lines.push('');
    lines.push('【表格資料】(提供全部欄位)');

    const rows = (snapshot.rows || []).slice(0, 200); // avoid huge prompt

    // Derive column order from the first row keys for stability.
    const colKeys = (rows[0] && typeof rows[0] === 'object') ? Object.keys(rows[0]) : [];
    if (colKeys.length === 0) {
      lines.push('(沒有列資料)');
      lines.push('');
      lines.push('請直接開始分析。');
      return lines.join('\n');
    }

    lines.push('欄位: ' + colKeys.join(' | '));
    rows.forEach(r => {
      lines.push(colKeys.map(k => {
        const v = r ? r[k] : '';
        return (v == null) ? '' : String(v).trim().replace(/\s+/g, ' ');
      }).join(' | '));
    });

    lines.push('');
    lines.push('請直接開始分析。');
    return lines.join('\n');
  }
})();
