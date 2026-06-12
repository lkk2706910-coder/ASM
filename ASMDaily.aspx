<%@ Page Language="C#" CodeFile="ASMDaily.aspx.cs" Inherits="ASMDaily" %>
<!DOCTYPE html>
<html lang="zh-Hant">
<head runat="server">
  <meta charset="UTF-8" />
  <title>ASM Daily</title>

  <!-- 引入 SheetJS 用於讀取 Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- 引入 Chart.js 用於畫圖 -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>

  <style>
  /* ===== ASM Daily.css (inlined) ===== */
/* ASM Daily.css */
body{margin:0;font-family:Arial,"Microsoft JhengHei",sans-serif;background-color:#f5f7fa;}
.header{background:linear-gradient(#dfe9f4,#f5f7fa);padding:10px 30px;font-size:24px;font-weight:700;color:#24496b;display:flex;justify-content:space-between;align-items:center;}
.header-right{font-size:12px;font-weight:400;}
.container{padding:10px 30px 40px 30px;}
.upload-section{margin-bottom:15px;font-size:13px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;}
.upload-left{display:flex;align-items:center;gap:6px;}
.upload-right{display:flex;flex-direction:column;gap:2px;font-size:12px;}
.status{margin-top:0;font-size:12px;color:#555;}
.status.error{color:#b00020;}
.status.success{color:#0b7a0b;}
.file-info{margin-top:0;font-size:12px;color:#333;}
.file-info .stale{color:#b00020;font-weight:700;}
.table-wrapper{overflow-x:auto;}
table{border-collapse:collapse;background-color:#fffff5;border:1px solid #b5b58c;white-space:nowrap;}
th,td{border:1px solid #b5b58c;padding:2px 6px;font-size:12px;text-align:center;height:20px;white-space:nowrap;color:#000;}
th{background-color:#e6f7d9;font-weight:700;cursor:pointer;user-select:none;}
td:first-child,th:first-child{text-align:left;padding-left:8px;}
th.col-eqpid,td.col-eqpid{width:50px;}
th.col-cycle,td.col-cycle{width:50px;}
th.col-pmdate,td.col-pmdate{width:40px;}
th.col-coat,td.col-coat{width:50px;}
th.col-state,td.col-state{width:60px;}
th.col-hpc,td.col-hpc{width:40px;}
th.col-m22,td.col-m22{width:40px;}
th.col-r22,td.col-r22{width:40px;}
th.col-ehv,td.col-ehv{width:40px;}
th.col-ulp,td.col-ulp{width:40px;}
th.col-m22inhib,td.col-m22inhib{width:40px;}
th.col-hpcinhib,td.col-hpcinhib{width:40px;}
th.col-measure,td.col-measure{width:70px;}
th.col-wc,td.col-wc{width:60px;}
th.col-oxide,td.col-oxide{width:60px;}
th.col-s7,td.col-s7{width:70px;}
th.col-s7temp,td.col-s7temp{width:40px;}
th.col-chart,td.col-chart{width:30px;}
th.col-pm,td.col-pm{width:30px;}
th.col-delta,td.col-delta{width:60px;}
th.col-ch,td.col-ch{width:50px;}
th.col-lastpa,td.col-lastpa{width:70px;}
th.sort-asc::after{content:" ▲";font-size:10px;}
th.sort-desc::after{content:" ▼";font-size:10px;}
td.col-ch{cursor:pointer;color:#0066cc;text-decoration:underline;}
.modal-overlay{position:fixed;inset:0;background-color:rgba(0,0,0,0.4);display:none;align-items:center;justify-content:center;z-index:999;}
.modal-overlay.active{display:flex;}
.modal-content{background-color:#fff;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,0.25);padding:16px 20px 20px;min-width:600px;max-width:900px;}
.modal-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;font-size:14px;font-weight:700;}
.modal-close{cursor:pointer;font-size:18px;font-weight:700;}
.modal-body{padding-top:4px;}
.chart-container{width:100%;height:380px;}
.header-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end;}
.nav-btn{font-size:12px;padding:6px 10px;border:1px solid #9db3c7;background:#fff;color:#24496b;border-radius:4px;cursor:pointer;}
.nav-btn:hover{background:#eef5fb;}
.tiles-wrap{display:flex;flex-wrap:wrap;gap:2px;margin:6px 0 8px 0;}
.tile{width:92px;height:18px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;border:1px solid #fff;box-sizing:border-box;user-select:none;cursor:default;color:#000;}
.tile.SB{background:#006400;color:#fff;}
.tile.PD{background:#39e75f;}
.tile.UD_M{background:#ff4da6;}
.tile.UD{background:#ff0000;}
.tile.SD_M{background:#d9ecff;}
.tile.SD{background:#0000ff;color:#fff;}
.tile.NS_OFF{background:#9e9e9e;}
.tile.EG{background:#fff2cc;}
.tile.NONE{background:#fff;color:#000;}

  /* ===== AiChat widget styles (inlined) ===== */
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

    /* fullscreen mode */
    #aiPanel.ai-fullscreen {
      right: 10px;
      bottom: 10px;
      width: calc(100vw - 20px);
      height: calc(100vh - 20px);
      max-width: none;
      max-height: none;
      border-radius: 12px;
    }

    /* Ensure children can expand in fullscreen (fix flex overflow sizing) */
    #aiPanel.ai-fullscreen .ai-main {
      flex: 1;
      min-width: 0;
      min-height: 0;
    }
    #aiPanel.ai-fullscreen .ai-msgs {
      flex: 1;
      min-height: 0;
    }

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
<form id="form1" runat="server">

  <div class="header">
    <div>ASM Daily</div>
    <div class="header-right"><span id="todayText"></span></div>
  </div>

  <div class="container">
    <div class="upload-section">
      <div class="upload-left">
        <label style="font-size:12px; color:#24496b;">
          上傳 Excel：
          <input id="excelFileInput" type="file" accept=".xlsm" />
        </label>
      </div>

      <div class="upload-right">
        <div id="fileInfo" class="file-info"></div>

        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <div id="status" class="status"></div>

          <div class="header-actions" id="quickLinks">
            <button class="nav-btn" data-open="ASM PM.html" type="button">ASM PM</button>
            <button class="nav-btn" data-open="ASM NPW.html" type="button">ASM NPW</button>
            <button class="nav-btn" data-open="ASM SUS.html" type="button">ASM SUS</button>
            <button class="nav-btn" data-open="ASM A-Grade.html" type="button">ASM A-Grade</button>
            <button class="nav-btn" data-open="HKG_GH Chamber List.html" type="button">HKG_GH Table</button>
            <button class="nav-btn" data-open="Oxide Chart.html" type="button">HKG_GH Oxide</button>
            <button class="nav-btn" data-open="HKG MAP.html" type="button">HKG MAP</button>
            <button class="nav-btn" data-open="ALDOX MAP.html" type="button">ALDOX MAP</button>
            <button class="nav-btn" data-open="Defectnotice.ASPX" type="button">ASM Defect Notice</button>
          </div>
        </div>
      </div>
    </div>

    <div id="summaryLines" style="margin-top:0px; font-size:16px; color:#24496b; line-height:0.6; font-weight:700;">
      <span id="hmrSummary">HPC+/M22/R22 Chamber : 0 Chs</span>
      <span style="margin-left:24px;" id="euSummary">eHV/ULP Chamber : 0 Chs</span><br><br>
    </div>

    <div class="table-wrapper">
      <table id="asmTable">
        <thead>
        <tr>
          <th class="col-eqpid">EQPID</th>
          <th class="col-cycle">Cycle</th>
          <th class="col-pmdate">PM</th>
          <th class="col-coat">Coat</th>
          <th class="col-state">State</th>
          <th class="col-hpc">HPC+</th>
          <th class="col-m22">M22</th>
          <th class="col-r22">R22</th>
          <th class="col-ehv">Ehv</th>
          <th class="col-ulp">ULP</th>
          <th class="col-m22inhib">M22 Inhib</th>
          <th class="col-hpcinhib">HPC Inhib</th>
          <th class="col-measure">S6_Min</th>
          <th class="col-wc">W/C</th>
          <th class="col-oxide">Oxide</th>
          <th class="col-s7">S7 Remain</th>
          <th class="col-s7temp">S7 Temp</th>
          <th class="col-chart">UD_PM</th>
          <th class="col-pm">階差</th>
          <th class="col-delta">Delta</th>
          <th class="col-ch">CH</th>
          <th class="col-lastpa">Last PA</th>
        </tr>
        </thead>
        <tbody>
          <!-- 與原本列表一致；每列在 S6_Min 後面插入 <td class="col-wc"></td> -->

          <tr>
            <td class="col-eqpid">HKG-A01B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A01B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A01D</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A01D</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A02A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A02A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A02B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A02B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A02C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A02C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B01B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B01B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B01C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B01C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B02A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B02A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B02B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B02B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B02C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B02C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B03B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B03B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B03C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B03C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B03D</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B03D</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B04A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B04A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B04B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B04B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B04C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B04C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B04D</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B04D</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B05A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B05A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B05B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B05B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B05C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B05C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B06A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B06A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B06B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B06B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B06C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B06C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B07B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B07B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B07C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B07C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B08A</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B08A</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B08B</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B08B</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-B08C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">B08C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A01C</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A01C</td>
            <td class="col-lastpa"></td>
          </tr>

          <tr>
            <td class="col-eqpid">HKG-A02D</td>
            <td class="col-cycle"></td><td class="col-pmdate"></td><td class="col-coat"></td><td class="col-state"></td>
            <td class="col-hpc"></td><td class="col-m22"></td><td class="col-r22"></td>
            <td class="col-ehv"></td><td class="col-ulp"></td><td class="col-m22inhib"></td><td class="col-hpcinhib"></td><td class="col-measure"></td>
            <td class="col-wc"></td>
            <td class="col-oxide"></td><td class="col-s7"></td><td class="col-s7temp"></td><td class="col-chart"></td>
            <td class="col-pm"></td><td class="col-delta"></td><td class="col-ch">A02D</td>
            <td class="col-lastpa"></td>
          </tr>

        </tbody>
      </table>
    </div>
  </div>

  <div id="chamberTiles" class="tiles-wrap"></div>

  <div id="paChartModal" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <div id="paChartTitle">PA Chart</div>
        <div id="paChartClose" class="modal-close">&times;</div>
      </div>
      <div class="modal-body"><div class="chart-container"><canvas id="paChartCanvas"></canvas></div></div>
    </div>
  </div>

  <div id="oxideChartModal" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <div id="oxideChartTitle">OXIDE Chart</div>
        <div id="oxideChartClose" class="modal-close">&times;</div>
      </div>
      <div class="modal-body"><div class="chart-container"><canvas id="oxideChartCanvas"></canvas></div></div>
    </div>
  </div>


  <!-- AiChat widget markup (embedded from AiChat.aspx) -->
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

</form>

<!-- ===== ASM Daily.js (inlined) ===== -->
<script>
// ASM Daily (ASPX 版本) - JS

// 固定 Excel URL
const EXCEL_URL = "http://umcesidb02/PoC/TF/TF2/EQ1/ASM/HKG_GH.xlsm";

// W/C 來源：本頁 ASPX WebMethod
const WC_ENDPOINT = window.location.pathname + "/GetWaferCounts";

// 主分頁欄位名稱對應表格欄位 class
const columnMapping = {
  "EQPID": "col-eqpid",
  "Cycle": "col-cycle",
  "Chart ID": "col-chart",
  "Delta": "col-delta",
  "CH": "col-ch",
  "Last PA": "col-lastpa"
};

// 全域資料
window.paDataByProcessUnit = {};
window.oxideDataByEqpid = {};
window.waferCountByEqpid = {}; // W/C

let paChartInstance = null;
let oxideChartInstance = null;

(function showToday() {
  const el = document.getElementById("todayText");
  if (!el) return;
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  el.textContent = "Today: " + y + "-" + m + "-" + day;
})();

function renderFileInfo(fileName, lastModifiedText) {
  const fileInfoEl = document.getElementById("fileInfo");
  if (!fileInfoEl) return;

  const text = lastModifiedText || "未知";

  // 解析 "YYYY-MM-DD HH:mm"
  let stale = false;
  const m = String(text).match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);
  if (m) {
    const dt = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], 0);
    if (!isNaN(dt.getTime())) {
      const diffMin = Math.abs(Date.now() - dt.getTime()) / 60000;
      stale = diffMin > 30;
    }
  }

  const readAt = (document.getElementById("status") && document.getElementById("status").dataset)
    ? (document.getElementById("status").dataset.readAt || "")
    : "";

  fileInfoEl.innerHTML =
    `檔案：${fileName}　檔案更新日期：` +
    `<span class="${stale ? "stale" : ""}">${text}</span>` +
    (readAt ? `　讀取時間：<span>${readAt}</span>` : "");
}

function buildChamberTilesFromTable() {
  const wrap = document.getElementById("chamberTiles");
  const tbody = document.querySelector("#asmTable tbody");
  if (!wrap || !tbody) return;
  if (wrap.dataset.built === "1") return;
  wrap.dataset.built = "1";

  wrap.innerHTML = "";
  tbody.querySelectorAll("tr").forEach(tr => {
    const eqTd = tr.querySelector("td.col-eqpid");
    if (!eqTd) return;
    const eqpid = (eqTd.textContent || "").trim();
    if (!eqpid) return;

    const div = document.createElement("div");
    div.className = "tile NONE";
    div.textContent = eqpid;
    div.dataset.eqpid = eqpid;
    div.title = eqpid;
    wrap.appendChild(div);
  });
}

function stateToTileClass(stateText) {
  const s = (stateText || "").trim();
  if (s === "NS_OFF") return "NS_OFF";
  if (s.startsWith("UD_M")) return "UD_M";
  if (s.startsWith("UD")) return "UD";
  if (s.startsWith("SD_M")) return "SD_M";
  if (s.startsWith("SD")) return "SD";
  if (s.startsWith("SB")) return "SB";
  if (s.startsWith("PD")) return "PD";
  if (s.startsWith("EG")) return "EG";
  return "NONE";
}

function updateChamberTilesColorsFromTable() {
  const wrap = document.getElementById("chamberTiles");
  const tbody = document.querySelector("#asmTable tbody");
  if (!wrap || !tbody) return;

  const stateByEqpid = {};
  tbody.querySelectorAll("tr").forEach(tr => {
    const eqTd = tr.querySelector("td.col-eqpid");
    const stTd = tr.querySelector("td.col-state");
    if (!eqTd) return;
    const eqpid = (eqTd.textContent || "").trim();
    if (!eqpid) return;

    const st = stTd ? (stTd.textContent || "").trim() : "";
    stateByEqpid[eqpid] = st;
  });

  wrap.querySelectorAll(".tile").forEach(tile => {
    const eqpid = tile.dataset.eqpid;
    const st = stateByEqpid[eqpid] || "";
    tile.classList.remove("SB", "PD", "UD_M", "UD", "SD_M", "SD", "NS_OFF", "EG", "NONE");
    const cls = stateToTileClass(st);
    tile.classList.add(cls);
    tile.title = `${eqpid}\nSTATE: ${st || "(空)"}`;
  });
}

// ========= Base 分頁：階差（RESULT）依 EQPID =========
function parseBaseDeltaByEqpid(workbook) {
  const out = {};
  const sheetName = workbook.SheetNames.find(n => n.toUpperCase() === "BASE") || "Base";
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return out;

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (rows.length < 2) return out;

  const header = rows[0].map(h => String(h || "").trim());
  const colTool = header.findIndex(h => h.toUpperCase() === "TOOL_NAME");
  const colCh = header.findIndex(h => h.toUpperCase() === "CHAMBER_NAME");
  const colRes = header.findIndex(h => h.toUpperCase() === "RESULT");
  if (colTool < 0 || colCh < 0 || colRes < 0) return out;

  const map = { "1": "A", "2": "B", "3": "C", "4": "D" };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const tool = String(r[colTool] ?? "").trim();
    const chName = String(r[colCh] ?? "").trim();
    if (!tool || !chName) continue;

    const m = chName.match(/(\d)(?!.*\d)/);
    if (!m) continue;
    const letter = map[m[1]];
    if (!letter) continue;

    const eqpid = `${tool}${letter}`;
    if (out[eqpid] === undefined) out[eqpid] = String(r[colRes] ?? "").trim();
  }
  return out;
}

// ========= W/C：從 DB 拉取，回填表格 =========
async function fetchWaferCountsFromDb() {
  const res = await fetch(WC_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: "{}",
    cache: "no-store"
  });

  if (!res.ok) throw new Error(`W/C API HTTP ${res.status}`);

  // ASP.NET WebMethod 會回 { d: ... }
  const json = await res.json();
  return (json && json.d) ? json.d : {};
}

function applyWaferCountsToTable(map) {
  const tbody = document.querySelector("#asmTable tbody");
  if (!tbody) return;

  tbody.querySelectorAll("tr").forEach(tr => {
    const eqTd = tr.querySelector("td.col-eqpid");
    const wcTd = tr.querySelector("td.col-wc");
    if (!eqTd || !wcTd) return;

    wcTd.style.color = "";
    wcTd.style.fontWeight = "";

    const eqpid = (eqTd.textContent || "").trim();
    const val = map ? map[eqpid] : undefined;

    if (val === undefined || val === null || String(val).trim() === "") {
      wcTd.textContent = "";
      return;
    }

    const n = Number(val);
    wcTd.textContent = isNaN(n) ? String(val).trim() : String(Math.round(n));

    if (!isNaN(n) && n > 5000) {
      wcTd.style.color = "blue";
      wcTd.style.fontWeight = "bold";
    }
  });
}

async function refreshWaferCountColumn() {
  const statusEl = document.getElementById("status");
  try {
    const map = await fetchWaferCountsFromDb();
    window.waferCountByEqpid = map;
    applyWaferCountsToTable(map);
  } catch (err) {
    console.error(err);
    if (statusEl) {
      const old = statusEl.textContent;
      statusEl.textContent = old ? (old + " / W/C 讀取失敗: " + err.message) : ("W/C 讀取失敗: " + err.message);
      statusEl.classList.remove("success");
      statusEl.classList.add("error");
    }
  }
}

// ========= Excel 主流程 =========
function loadExcelFromArrayBuffer(arrayBuffer, lastModifiedText) {
  const statusEl = document.getElementById("status");
  statusEl.textContent = "";
  statusEl.className = "status";

  try {
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: "array" });

    const basePmByEqpid = parseBaseDeltaByEqpid(workbook);
    renderFileInfo("HKG_GH.xlsm", lastModifiedText);

    // ========= 主分頁 =========
    const mainSheetName = workbook.SheetNames[0];
    const mainSheet = workbook.Sheets[mainSheetName];
    const mainData = XLSX.utils.sheet_to_json(mainSheet, { header: 1 });

    if (mainData.length === 0) {
      statusEl.textContent = "Excel 內容為空。";
      statusEl.classList.add("error");
      return;
    }

    const headerRow = mainData[0].map(h => (h || "").toString().trim());
    const dataRows = mainData.slice(1);

    const headerIndexMap = {};
    headerRow.forEach((name, index) => { if (name) headerIndexMap[name] = index; });

    if (headerIndexMap["EQPID"] === undefined) {
      const eqHeader = headerRow.find(h => h.toUpperCase() === "EQPID");
      if (eqHeader) headerIndexMap["EQPID"] = headerRow.indexOf(eqHeader);
    }

    if (headerIndexMap["EQPID"] === undefined) {
      statusEl.textContent = "Excel 找不到「EQPID」欄位，請確認主分頁標題列名稱。";
      statusEl.classList.add("error");
      return;
    }

    const mainDataByEqpid = {};
    dataRows.forEach(row => {
      if (!row || row.length === 0) return;
      const eqpidCell = row[headerIndexMap["EQPID"]];
      const eqpid = eqpidCell ? eqpidCell.toString().trim() : "";
      if (!eqpid) return;

      const rowObj = {};
      headerRow.forEach((colName, idx) => {
        if (!colName) return;
        rowObj[colName] = row[idx];
      });
      mainDataByEqpid[eqpid] = rowObj;
    });

    // ========= STATUS 分頁 =========
    const statusSheetName = workbook.SheetNames.find(n => n.toUpperCase() === "STATUS") || "STATUS";
    const statusSheet2 = workbook.Sheets[statusSheetName];
    let statusDataByEqpid = {};

    if (!statusSheet2) {
      statusEl.textContent = "找不到 STATUS 分頁，僅更新主分頁資料。";
      statusEl.classList.add("error");
    } else {
      const statusData = XLSX.utils.sheet_to_json(statusSheet2, { header: 1 });
      if (statusData.length > 0) {
        const statusHeaderRow = statusData[0].map(h => (h || "").toString().trim());
        const statusRows = statusData.slice(1);

        const statusHeaderIndex = {};
        statusHeaderRow.forEach((name, index) => { if (name) statusHeaderIndex[name] = index; });

        if (statusHeaderIndex["EQPID"] === undefined) {
          const eqHeader2 = statusHeaderRow.find(h => h.toUpperCase() === "EQPID");
          if (eqHeader2) statusHeaderIndex["EQPID"] = statusHeaderRow.indexOf(eqHeader2);
        }

        let stateCol = statusHeaderIndex["state"]; // 原本用小寫
        if (stateCol === undefined) {
          const alt = statusHeaderRow.find(h => h.toUpperCase() === "STATE");
          if (alt) stateCol = statusHeaderRow.indexOf(alt);
        }

        const remainingCol = statusHeaderIndex["Remaining"] !== undefined
          ? statusHeaderIndex["Remaining"]
          : statusHeaderRow.findIndex(h => h.toUpperCase() === "REMAINING");

        const oxideValCol = statusHeaderIndex["OXIDE Value"] !== undefined
          ? statusHeaderIndex["OXIDE Value"]
          : statusHeaderRow.findIndex(h => h.toUpperCase() === "OXIDE VALUE");

        // Inhibit 欄位（容錯）
        let m22InhibCol = statusHeaderIndex["M22 Inhibit"];
        if (m22InhibCol === undefined) {
          const alt = statusHeaderRow.find(h => h.toUpperCase() === "M22 INHIBIT");
          if (alt) m22InhibCol = statusHeaderRow.indexOf(alt);
        }

        let hpcInhibCol = statusHeaderIndex["HPC+ InHibit"];
        if (hpcInhibCol === undefined) hpcInhibCol = statusHeaderIndex["HPC+ Inhibit"];
        if (hpcInhibCol === undefined) {
          const alt = statusHeaderRow.find(h => h.toUpperCase() === "HPC+ INHIBIT");
          if (alt) hpcInhibCol = statusHeaderRow.indexOf(alt);
        }

        const hasEq = statusHeaderIndex["EQPID"] !== undefined;

        statusRows.forEach(row => {
          if (!row || row.length === 0) return;

          const eqpidCell = hasEq ? row[statusHeaderIndex["EQPID"]] : undefined;
          const eqpid = eqpidCell ? eqpidCell.toString().trim() : "";
          if (!eqpid) return;

          const stateVal = (stateCol !== undefined && stateCol >= 0) ? row[stateCol] : "";
          const remainingVal = (remainingCol !== undefined && remainingCol >= 0) ? row[remainingCol] : "";
          const oxideVal = (oxideValCol !== undefined && oxideValCol >= 0) ? row[oxideValCol] : "";

          const m22InhibVal = (m22InhibCol !== undefined) ? row[m22InhibCol] : "";
          const hpcInhibVal = (hpcInhibCol !== undefined) ? row[hpcInhibCol] : "";

          statusDataByEqpid[eqpid] = {
            state: stateVal,
            remaining: remainingVal,
            oxideValue: oxideVal,
            m22Inhibit: m22InhibVal,
            hpcInhibit: hpcInhibVal
          };
        });
      }
    }

    // ========= UD_PM 分頁 =========
    let udPmByEqpid = {};
    let pmDateByEqpid = {};

    const udSheet = workbook.Sheets["UD_PM"];
    if (udSheet) {
      const rows = XLSX.utils.sheet_to_json(udSheet, { header: 1, defval: "" });
      if (rows.length >= 2) {
        const header = rows[0].map(h => String(h).trim());
        const colTool = header.indexOf("TOOL_NAME");
        const colCh = header.indexOf("CHAMBER_NAME");
        const colRes = header.indexOf("RESULT");
        const colStart = header.indexOf("START_TIME");

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          const tool = String(r[colTool] ?? "").trim();
          const ch = String(r[colCh] ?? "").trim();
          if (!tool || !ch) continue;

          const m = ch.match(/(\d)(?!.*\d)/);
          if (!m) continue;
          const map = { "1": "A", "2": "B", "3": "C", "4": "D" };
          const letter = map[m[1]];
          if (!letter) continue;

          const eqpid = `${tool}${letter}`;

          if (udPmByEqpid[eqpid] === undefined) {
            udPmByEqpid[eqpid] = String(r[colRes] ?? "").trim();

            let mmdd = "";
            const tRaw = (colStart >= 0) ? r[colStart] : "";
            if (tRaw !== "" && tRaw !== null && tRaw !== undefined) {
              let dVal = null;
              if (typeof tRaw === "number") {
                const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(tRaw) : null;
                if (parsed) {
                  const { y, m, d, H = 0, M = 0, S = 0 } = parsed;
                  dVal = new Date(y, m - 1, d, H, M, S);
                } else {
                  dVal = new Date(tRaw);
                }
              } else {
                dVal = new Date(tRaw);
              }

              if (dVal instanceof Date && !isNaN(dVal.getTime())) {
                const mm = String(dVal.getMonth() + 1).padStart(2, "0");
                const dd = String(dVal.getDate()).padStart(2, "0");
                mmdd = `${mm}/${dd}`;
              }
            }
            pmDateByEqpid[eqpid] = mmdd;
          }
        }
      }
    }

    // ========= S6_Min 分頁 =========
    let s6MinByEqpid = {};
    const s6SheetName = workbook.SheetNames.find(n => n.toUpperCase() === "S6_MIN") || "S6_Min";
    const s6Sheet = workbook.Sheets[s6SheetName];
    if (s6Sheet) {
      const s6Data = XLSX.utils.sheet_to_json(s6Sheet, { header: 1 });
      if (s6Data.length > 0) {
        const s6HeaderRow = s6Data[0].map(h => (h || "").toString().trim());
        const s6Rows = s6Data.slice(1);

        const s6HeaderIndex = {};
        s6HeaderRow.forEach((name, index) => { if (name) s6HeaderIndex[name] = index; });

        let toolCol = s6HeaderIndex["TOOL_NAME"];
        if (toolCol === undefined) {
          const alt = s6HeaderRow.find(h => h.toUpperCase() === "TOOL_NAME");
          if (alt) toolCol = s6HeaderRow.indexOf(alt);
        }

        let chamCol = s6HeaderIndex["CHAMBER_NAME"];
        if (chamCol === undefined) {
          const alt = s6HeaderRow.find(h => h.toUpperCase() === "CHAMBER_NAME");
          if (alt) chamCol = s6HeaderRow.indexOf(alt);
        }

        let startTimeCol = s6HeaderIndex["START_TIME"];
        if (startTimeCol === undefined) {
          const alt = s6HeaderRow.find(h => h.toUpperCase().includes("START") && h.toUpperCase().includes("TIME"));
          if (alt) startTimeCol = s6HeaderRow.indexOf(alt);
        }

        let resultCol = s6HeaderIndex["RESULT"];
        if (resultCol === undefined) {
          const alt = s6HeaderRow.find(h => h.toUpperCase() === "RESULT");
          if (alt) resultCol = s6HeaderRow.indexOf(alt);
        }

        if (toolCol !== undefined && chamCol !== undefined && startTimeCol !== undefined && resultCol !== undefined) {
          const tmpByEqpid = {};
          s6Rows.forEach(row => {
            if (!row || row.length === 0) return;
            const tool = row[toolCol] ? row[toolCol].toString().trim() : "";
            const chamber = row[chamCol] ? row[chamCol].toString().trim() : "";
            if (!tool || !chamber) return;

            const lastChar = chamber.slice(-1);
            const digit = parseInt(lastChar, 10);
            let suffix = "";
            if (digit === 1) suffix = "A";
            else if (digit === 2) suffix = "B";
            else if (digit === 3) suffix = "C";
            else if (digit === 4) suffix = "D";
            else return;

            const eqpid = tool + suffix;

            const timeRaw = row[startTimeCol];
            if (!timeRaw && timeRaw !== 0) return;
            let timeVal;
            if (typeof timeRaw === "number") {
              const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(timeRaw) : null;
              if (parsed) {
                const { y, m, d, H = 0, M = 0, S = 0 } = parsed;
                timeVal = new Date(y, m - 1, d, H, M, S);
              } else {
                timeVal = new Date(timeRaw);
              }
            } else {
              timeVal = new Date(timeRaw);
            }
            if (!(timeVal instanceof Date) || isNaN(timeVal.getTime())) return;

            const result = row[resultCol];
            if (result === undefined || result === null || result === "") return;

            if (!tmpByEqpid[eqpid]) tmpByEqpid[eqpid] = [];
            tmpByEqpid[eqpid].push({ time: timeVal, result });
          });

          Object.keys(tmpByEqpid).forEach(eqpid => {
            const list = tmpByEqpid[eqpid];
            list.sort((a, b) => b.time - a.time);
            s6MinByEqpid[eqpid] = list[0].result;
          });
        }
      }
    }

    // ========= S7 Vessel 分頁 =========
    let s7RemainByEqpid = {};
    let s7WeighDateByEqpid = {};

    const s7SheetName =
      workbook.SheetNames.find(n => n.toUpperCase().includes("S7") && n.toUpperCase().includes("VESSEL")) ||
      workbook.SheetNames.find(n => n.toUpperCase() === "S7 VESSEL") ||
      "S7 Vessel";

    const s7Sheet = workbook.Sheets[s7SheetName];

    if (s7Sheet) {
      const s7Data = XLSX.utils.sheet_to_json(s7Sheet, { header: 1 });
      if (s7Data.length > 0) {
        const s7HeaderRow = s7Data[0].map(h => (h || "").toString().trim());
        const s7Rows = s7Data.slice(1);

        const s7HeaderIndex = {};
        s7HeaderRow.forEach((name, index) => { if (name) s7HeaderIndex[name] = index; });

        let eqpidCol = s7HeaderIndex["EQPID"];
        if (eqpidCol === undefined) {
          const altEq = s7HeaderRow.find(h => h.toUpperCase() === "EQPID");
          if (altEq) eqpidCol = s7HeaderRow.indexOf(altEq);
        }
        if (eqpidCol === undefined) eqpidCol = 0;

        const getCol = (colName) => {
          let c = s7HeaderIndex[colName];
          if (c === undefined) {
            const alt = s7HeaderRow.find(h => h.toUpperCase() === colName.toUpperCase());
            if (alt) c = s7HeaderRow.indexOf(alt);
          }
          return c;
        };

        const col33 = getCol("Column33");
        const col13 = getCol("Column13");
        const col16 = getCol("Column16");
        const col12 = getCol("Column12");
        const col11 = getCol("Column11");

        if (col33 !== undefined && col13 !== undefined && col16 !== undefined && col12 !== undefined) {
          s7Rows.forEach(row => {
            if (!row || row.length === 0) return;

            const eqpid = row[eqpidCol] ? row[eqpidCol].toString().trim() : "";
            if (!eqpid) return;

            const v33 = row[col33];
            let v13 = row[col13];
            let v16 = row[col16];
            let v12 = row[col12];

            if (v33 === "" || v33 == null || v12 === "" || v12 == null) return;
            if (v13 === "" || v13 == null) v13 = 0;

            if (typeof v16 === "string" && v16.trim() === "未秤重") v16 = 0.0000017;
            else if (v16 === "" || v16 == null) v16 = 0.0000017;

            const n33 = Number(v33), n13 = Number(v13), n16 = Number(v16), n12 = Number(v12);
            if ([n33, n13, n16, n12].some(x => isNaN(x))) return;

            const diff = n33 - n13;
            const value = n12 - diff * n16;
            s7RemainByEqpid[eqpid] = (value * 100).toFixed(1) + "%";

            if (col11 !== undefined) {
              const dRaw = row[col11];
              if (dRaw || dRaw === 0) {
                let dVal = null;
                if (typeof dRaw === "number") {
                  const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(dRaw) : null;
                  if (parsed) dVal = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
                  else dVal = new Date(dRaw);
                } else {
                  dVal = new Date(dRaw);
                }
                if (dVal instanceof Date && !isNaN(dVal.getTime())) s7WeighDateByEqpid[eqpid] = dVal;
              }
            }
          });
        }
      }
    }

    // ========= S7 Temp 分頁 =========
    let s7TempByEqpid = {};
    const s7TempSheetName = workbook.SheetNames.find(n => n.toUpperCase() === "S7 TEMP") || "S7 Temp";
    const s7TempSheet = workbook.Sheets[s7TempSheetName];

    if (s7TempSheet) {
      const data = XLSX.utils.sheet_to_json(s7TempSheet, { header: 1, defval: "" });
      if (data.length > 1) {
        const header = data[0].map(h => String(h || "").trim());
        const rows = data.slice(1);

        const idx = {};
        header.forEach((name, i) => { if (name) idx[name] = i; });

        const findCol = (name) => {
          let c = idx[name];
          if (c === undefined) {
            const alt = header.find(h => h.toUpperCase() === name.toUpperCase());
            if (alt) c = header.indexOf(alt);
          }
          return c;
        };

        const toolCol = findCol("TOOL_NAME");
        const chamberCol = findCol("CHAMBER_NAME");
        const resultCol = findCol("RESULT");

        if (toolCol !== undefined && chamberCol !== undefined && resultCol !== undefined) {
          const map = { "1": "A", "2": "B", "3": "C", "4": "D" };
          for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const tool = String(r[toolCol] ?? "").trim();
            const chamber = String(r[chamberCol] ?? "").trim();
            if (!tool || !chamber) continue;

            const m = chamber.match(/(\d)(?!.*\d)/);
            if (!m) continue;

            const letter = map[m[1]];
            if (!letter) continue;

            const eqpid = `${tool}${letter}`;
            if (s7TempByEqpid[eqpid] === undefined) {
              s7TempByEqpid[eqpid] = String(r[resultCol] ?? "").trim();
            }
          }
        }
      }
    }

    // ========= Coating 分頁（Cycle） =========
    let cycleDataByEqpid = {};
    const coatingSheetName = workbook.SheetNames.find(n => n.toUpperCase() === "COATING") || "Coating";
    const coatingSheet = workbook.Sheets[coatingSheetName];
    if (coatingSheet) {
      const coatingData = XLSX.utils.sheet_to_json(coatingSheet, { header: 1 });
      if (coatingData.length > 0) {
        const headerRow2 = coatingData[0].map(h => (h || "").toString().trim());
        const rows2 = coatingData.slice(1);
        const idx = {};
        headerRow2.forEach((name, i) => { if (name) idx[name] = i; });

        let chamberIdCol = idx["CHAMBERID"];
        if (chamberIdCol === undefined) {
          const alt = headerRow2.find(h => h.toUpperCase() === "CHAMBERID");
          if (alt) chamberIdCol = headerRow2.indexOf(alt);
        }

        let coatCCol = idx["Coat_c"];
        if (coatCCol === undefined) coatCCol = idx["COAT_C"];
        if (coatCCol === undefined) {
          const alt = headerRow2.find(h => h.toUpperCase() === "COAT_C");
          if (alt) coatCCol = headerRow2.indexOf(alt);
        }

        if (chamberIdCol !== undefined && coatCCol !== undefined) {
          rows2.forEach(r => {
            if (!r || r.length === 0) return;
            const chamberId = r[chamberIdCol] ? r[chamberIdCol].toString().trim() : "";
            if (!chamberId) return;
            const coatC = r[coatCCol];
            if (coatC === undefined || coatC === null || coatC === "") return;
            cycleDataByEqpid[chamberId] = coatC;
          });
        }
      }
    }

    // ========= PA 分頁 =========
    let lastPaByEqpid = {};
    const paSheetName = workbook.SheetNames.find(n => n.toUpperCase() === "PA") || "PA";
    const paSheet = workbook.Sheets[paSheetName];
    window.paDataByProcessUnit = {};

    if (paSheet) {
      const paData = XLSX.utils.sheet_to_json(paSheet, { header: 1 });
      if (paData.length > 0) {
        const paHeaderRow = paData[0].map(h => (h || "").toString().trim());
        const paRows = paData.slice(1);

        const paHeaderIndex = {};
        paHeaderRow.forEach((name, index) => { if (name) paHeaderIndex[name] = index; });

        let unitCol = paHeaderIndex["PROCESSUNIT"];
        if (unitCol === undefined) {
          const alt = paHeaderRow.find(h => h.toUpperCase() === "PROCESSUNIT");
          if (alt) unitCol = paHeaderRow.indexOf(alt);
        }

        let updateTimeCol = paHeaderIndex["UPDATE_TIME"];
        if (updateTimeCol === undefined) {
          const alt = paHeaderRow.find(h => h.toUpperCase().includes("UPDATE") && h.toUpperCase().includes("TIME"));
          if (alt) updateTimeCol = paHeaderRow.indexOf(alt);
        }

        let meanValueCol = paHeaderIndex["MEAN_VALUE"];
        if (meanValueCol === undefined) {
          const alt = paHeaderRow.find(h => h.toUpperCase().includes("MEAN") && h.toUpperCase().includes("VALUE"));
          if (alt) meanValueCol = paHeaderRow.indexOf(alt);
        }

        if (unitCol !== undefined && updateTimeCol !== undefined && meanValueCol !== undefined) {
          const paByUnit = {};
          const fullPaByUnit = {};

          paRows.forEach(row => {
            if (!row || row.length === 0) return;
            const unit = row[unitCol] ? row[unitCol].toString().trim() : "";
            if (!unit) return;

            const timeRaw = row[updateTimeCol];
            if (!timeRaw && timeRaw !== 0) return;

            let timeVal;
            if (typeof timeRaw === "number") {
              const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(timeRaw) : null;
              if (parsed) {
                const { y, m, d, H = 0, M = 0, S = 0 } = parsed;
                timeVal = new Date(y, m - 1, d, H, M, S);
              } else {
                timeVal = new Date(timeRaw);
              }
            } else {
              timeVal = new Date(timeRaw);
            }
            if (!(timeVal instanceof Date) || isNaN(timeVal.getTime())) return;

            const mvRaw = row[meanValueCol];
            if (mvRaw === undefined || mvRaw === null || mvRaw === "") return;
            const mvNum = Number(mvRaw);
            const mv = isNaN(mvNum) ? mvRaw.toString().trim() : mvNum;

            if (!paByUnit[unit]) paByUnit[unit] = [];
            paByUnit[unit].push({ time: timeVal, mean: mv });

            if (!fullPaByUnit[unit]) fullPaByUnit[unit] = [];
            fullPaByUnit[unit].push({ time: timeVal, mean: mv });
          });

          Object.keys(paByUnit).forEach(unit => {
            const list = paByUnit[unit];
            list.sort((a, b) => b.time - a.time);
            const top3 = list.slice(0, 3);
            top3.reverse();

            const plainValues = [];
            const htmlValues = [];
            top3.forEach(item => {
              let val = item.mean;
              if (typeof val === "number") {
                plainValues.push(val);
                htmlValues.push(val > 5 ? `<span style=\"color:red;\">${val}</span>` : String(val));
              } else {
                plainValues.push(val);
                htmlValues.push(String(val));
              }
            });

            lastPaByEqpid[unit] = {
              text: "(" + plainValues.join(", ") + ")",
              html: "(" + htmlValues.join(", ") + ")"
            };
          });

          Object.keys(fullPaByUnit).forEach(unit => fullPaByUnit[unit].sort((a, b) => a.time - b.time));
          window.paDataByProcessUnit = fullPaByUnit;
        }
      }
    }

    // ========= OXIDE 分頁 =========
    let deltaByEqpid = {};
    let oxideLatestTimeByEqpid = {};
    window.oxideDataByEqpid = {};

    const oxideSheetName = workbook.SheetNames.find(n => n.toUpperCase() === "OXIDE") || "OXIDE";
    const oxideSheet = workbook.Sheets[oxideSheetName];
    if (oxideSheet) {
      const oxData = XLSX.utils.sheet_to_json(oxideSheet, { header: 1 });
      if (oxData.length > 0) {
        const oxHeaderRow = oxData[0].map(h => (h || "").toString().trim());
        const oxRows = oxData.slice(1);

        const oxHeaderIndex = {};
        oxHeaderRow.forEach((name, index) => { if (name) oxHeaderIndex[name] = index; });

        let puCol = oxHeaderIndex["PROCESSUNIT"];
        if (puCol === undefined) {
          const alt = oxHeaderRow.find(h => h.toUpperCase() === "PROCESSUNIT");
          if (alt) puCol = oxHeaderRow.indexOf(alt);
        }

        let recipeCol = oxHeaderIndex["RECIPE"];
        if (recipeCol === undefined) {
          const alt = oxHeaderRow.find(h => h.toUpperCase() === "RECIPE");
          if (alt) recipeCol = oxHeaderRow.indexOf(alt);
        }

        let oxUpdateTimeCol = oxHeaderIndex["UPDATE_TIME"];
        if (oxUpdateTimeCol === undefined) {
          const alt = oxHeaderRow.find(h => h.toUpperCase().includes("UPDATE") && h.toUpperCase().includes("TIME"));
          if (alt) oxUpdateTimeCol = oxHeaderRow.indexOf(alt);
        }

        let oxMeanValueCol = oxHeaderIndex["MEAN_VALUE"];
        if (oxMeanValueCol === undefined) {
          const alt = oxHeaderRow.find(h => h.toUpperCase().includes("MEAN") && h.toUpperCase().includes("VALUE"));
          if (alt) oxMeanValueCol = oxHeaderRow.indexOf(alt);
        }

        if (puCol !== undefined && recipeCol !== undefined && oxUpdateTimeCol !== undefined && oxMeanValueCol !== undefined) {
          const fullOxideByEqpid = {};
          const mapEqpidToRecords = {};

          oxRows.forEach(row => {
            if (!row || row.length === 0) return;
            const pu = row[puCol] ? row[puCol].toString().trim() : "";
            const recipe = row[recipeCol] ? row[recipeCol].toString().trim() : "";
            if (!pu || !recipe) return;

            const lastChar = recipe.slice(-1);
            const digit = parseInt(lastChar, 10);
            let suffix = "";
            if (digit === 1) suffix = "A";
            else if (digit === 2) suffix = "B";
            else if (digit === 3) suffix = "C";
            else if (digit === 4) suffix = "D";
            else return;

            const eqpid = pu + suffix;

            const timeRaw = row[oxUpdateTimeCol];
            if (!timeRaw && timeRaw !== 0) return;

            let timeVal;
            if (typeof timeRaw === "number") {
              const parsed = XLSX.SSF ? XLSX.SSF.parse_date_code(timeRaw) : null;
              if (parsed) {
                const { y, m, d, H = 0, M = 0, S = 0 } = parsed;
                timeVal = new Date(y, m - 1, d, H, M, S);
              } else {
                timeVal = new Date(timeRaw);
              }
            } else {
              timeVal = new Date(timeRaw);
            }
            if (!(timeVal instanceof Date) || isNaN(timeVal.getTime())) return;

            const mvRaw = row[oxMeanValueCol];
            if (mvRaw === undefined || mvRaw === null || mvRaw === "") return;
            const mvNum = Number(mvRaw);
            if (isNaN(mvNum)) return;

            if (!fullOxideByEqpid[eqpid]) fullOxideByEqpid[eqpid] = [];
            fullOxideByEqpid[eqpid].push({ time: timeVal, mean: mvNum });

            if (!mapEqpidToRecords[eqpid]) mapEqpidToRecords[eqpid] = [];
            mapEqpidToRecords[eqpid].push({ time: timeVal, mean: mvNum });
          });

          Object.keys(fullOxideByEqpid).forEach(eqpid => fullOxideByEqpid[eqpid].sort((a, b) => a.time - b.time));
          window.oxideDataByEqpid = fullOxideByEqpid;

          Object.keys(mapEqpidToRecords).forEach(eqpid => {
            const list = mapEqpidToRecords[eqpid];
            list.sort((a, b) => b.time - a.time);
            const latest = list[0];
            oxideLatestTimeByEqpid[eqpid] = latest.time;
            if (list.length < 2) return;
            const second = list[1];
            deltaByEqpid[eqpid] = latest.mean - second.mean;
          });
        }
      }
    }

    // ========= 更新 HTML 表格 =========
    const table = document.getElementById("asmTable");
    const tbody = table.querySelector("tbody");
    const rows = tbody.querySelectorAll("tr");
    let updateCount = 0;

    rows.forEach(tr => {
      const eqpidCell = tr.querySelector("td.col-eqpid");
      if (!eqpidCell) return;
      const eqpid = eqpidCell.textContent.trim();

      const mainRow = mainDataByEqpid[eqpid];
      if (mainRow) {
        for (const excelColName in columnMapping) {
          const tdClass = columnMapping[excelColName];
          const targetTd = tr.querySelector("td." + tdClass);
          if (!targetTd) continue;
          const value = mainRow[excelColName];
          if (value !== undefined && value !== null) targetTd.textContent = value;
        }

        const hpcTd = tr.querySelector("td.col-hpc");
        if (hpcTd) hpcTd.textContent = (mainRow["HPC+ Flag"] == 1) ? "V" : "";
        const m22Td = tr.querySelector("td.col-m22");
        if (m22Td) m22Td.textContent = (mainRow["L22 Flag"] == 1) ? "V" : "";
        const r22Td = tr.querySelector("td.col-r22");
        if (r22Td) r22Td.textContent = (mainRow["R22 Flag"] == 1) ? "V" : "";
        const ehvTd = tr.querySelector("td.col-ehv");
        if (ehvTd) ehvTd.textContent = (mainRow["eHVL22 Flag"] == 1) ? "V" : "";
        const ulpTd = tr.querySelector("td.col-ulp");
        if (ulpTd) ulpTd.textContent = (mainRow["ULP Flag"] == 1) ? "V" : "";
      }

      const stRow = statusDataByEqpid[eqpid];
      if (stRow) {
        const stateTd = tr.querySelector("td.col-state");
        const coatTd = tr.querySelector("td.col-coat");
        const oxideTd = tr.querySelector("td.col-oxide");

        // Inhibit
        const m22InhibTd = tr.querySelector("td.col-m22inhib");
        const hpcInhibTd = tr.querySelector("td.col-hpcinhib");
        function paintInhib(td, rawVal) {
          if (!td) return;
          td.textContent = "";
          td.style.backgroundColor = "";
          const n = Number(rawVal);
          if (!isNaN(n) && n === 2) td.style.backgroundColor = "#f2f2f2";
        }
        paintInhib(m22InhibTd, stRow.m22Inhibit);
        paintInhib(hpcInhibTd, stRow.hpcInhibit);

        if (stateTd) {
          let stateText = (stRow.state !== undefined && stRow.state !== null) ? String(stRow.state).trim() : "";
          stateTd.textContent = stateText;
          stateTd.style.color = "";
          if (stateText.startsWith("UD_MO")) stateTd.style.color = "#ff4da6";
          else if (stateText.startsWith("UD_W") || stateText.startsWith("WD_D")) stateTd.style.color = "red";
          else if (stateText.startsWith("SD")) stateTd.style.color = "blue";
          else if (stateText.startsWith("PD")) stateTd.style.color = "#00cc66";
          else if (stateText.startsWith("SB")) stateTd.style.color = "#006633";
        }

        if (coatTd) {
          let coatText = "";
          let coatNumber = null;
          if (stRow.remaining !== undefined && stRow.remaining !== null && stRow.remaining !== "") {
            const num = Number(stRow.remaining);
            if (!isNaN(num)) { coatNumber = num; coatText = num.toFixed(1); }
            else coatText = stRow.remaining.toString();
          }
          coatTd.textContent = coatText;
          coatTd.style.color = "";
          coatTd.style.fontWeight = "";
          if (coatNumber !== null && coatNumber < 3.1) {
            coatTd.style.color = "red";
            coatTd.style.fontWeight = "bold";
          }
        }

        if (oxideTd) {
          let oxideText = "";
          let oxideNum = null;
          if (stRow.oxideValue !== undefined && stRow.oxideValue !== null && stRow.oxideValue !== "") {
            const num2 = Number(stRow.oxideValue);
            if (!isNaN(num2)) { oxideNum = num2; oxideText = num2.toFixed(3); }
            else oxideText = stRow.oxideValue.toString();
          }
          oxideTd.textContent = oxideText;
          oxideTd.style.color = "";
          oxideTd.style.fontWeight = "";

          const t = oxideLatestTimeByEqpid[eqpid];
          if (t instanceof Date && !isNaN(t.getTime())) {
            const y = t.getFullYear();
            const m = String(t.getMonth() + 1).padStart(2, "0");
            const d = String(t.getDate()).padStart(2, "0");
            const hh = String(t.getHours()).padStart(2, "0");
            const mm = String(t.getMinutes()).padStart(2, "0");
            const ss = String(t.getSeconds()).padStart(2, "0");
            oxideTd.title = `UPDATE_TIME : ${y}-${m}-${d} ${hh}:${mm}:${ss}`;
          } else {
            oxideTd.title = "";
          }

          if (oxideNum !== null) {
            if (oxideNum >= 9.48 && oxideNum < 9.70) {
              oxideTd.style.color = "#DAA520";
              oxideTd.style.fontWeight = "bold";
            } else if (oxideNum > 10.20) {
              oxideTd.style.color = "blue";
              oxideTd.style.fontWeight = "bold";
            } else if (oxideNum >= 9.70) {
              oxideTd.style.color = "#006633";
              oxideTd.style.fontWeight = "bold";
            }
          }
        }
      }

      const pmTd = tr.querySelector("td.col-pm");
      const udpmTd = tr.querySelector("td.col-chart");
      const deltaTd = tr.querySelector("td.col-delta");
      const lastPaTd = tr.querySelector("td.col-lastpa");
      const s6Td = tr.querySelector("td.col-measure");
      const s7Td = tr.querySelector("td.col-s7");
      const s7TempTd = tr.querySelector("td.col-s7temp");
      const wcTd = tr.querySelector("td.col-wc");
      const pmDateTd = tr.querySelector("td.col-pmdate");
      const cycleTd = tr.querySelector("td.col-cycle");

      // ======= HPC+~ULP 顏色管理（你說不見的就是這段）：Oxide 範圍 + (各欄位都沒有 V) + (State 以 PD/SB 開頭) → 底色提示 =======
      (function applyHintsByOxideAndNoVAndState() {
        const hpcTd = tr.querySelector("td.col-hpc");
        const m22Td = tr.querySelector("td.col-m22");
        const r22Td = tr.querySelector("td.col-r22");
        const ehvTd = tr.querySelector("td.col-ehv");
        const ulpTd = tr.querySelector("td.col-ulp");
        const oxideTd = tr.querySelector("td.col-oxide");
        const stateTd = tr.querySelector("td.col-state");
        if (!hpcTd || !m22Td || !r22Td || !ehvTd || !ulpTd || !oxideTd || !stateTd) return;

        // 先清除舊底色
        [hpcTd, m22Td, r22Td, ehvTd, ulpTd].forEach(td => td.style.backgroundColor = "");

        // State 必須 PD* 或 SB*
        const stateText = stateTd.textContent.trim();
        const okState = stateText.startsWith("PD") || stateText.startsWith("SB");
        if (!okState) return;

        const oxideNum = Number(oxideTd.textContent.trim());
        if (isNaN(oxideNum)) return;

        const isV = (td) => (td.textContent || "").trim() === "V";

        // HPC+/M22/R22 都沒有 V => 淺綠；但 Coat < 3.0 改淺黃
        const hasAnyV_HMR = isV(hpcTd) || isV(m22Td) || isV(r22Td);

        const coatTd = tr.querySelector("td.col-coat");
        const coatNum = coatTd ? Number((coatTd.textContent || "").trim()) : NaN;

        // 用「資料來源」判斷 Delta（不要用 td 文字，最穩）
        const deltaRaw = (typeof deltaByEqpid !== "undefined") ? deltaByEqpid[eqpid] : undefined;
        const deltaNum = Number(deltaRaw);
        const blockHMRColor = !isNaN(deltaNum) && (deltaNum > 0.419 || deltaNum < -0.419);

        const lightGreen = "#e6f7d9";
        const lightYellow = "#fff2cc";
        const color = (!isNaN(coatNum) && coatNum < 3.0) ? lightYellow : lightGreen;

        if (!hasAnyV_HMR && !blockHMRColor) {
          if (oxideNum >= 9.48 && oxideNum <= 10.20) hpcTd.style.backgroundColor = color;
          if (oxideNum >= 9.55 && oxideNum <= 10.49) m22Td.style.backgroundColor = color;
          if (oxideNum >= 9.61 && oxideNum <= 10.32) r22Td.style.backgroundColor = color;
        }

        // Oxide 9.20~9.88 且 HPC+/M22/R22/ULP 都沒有 V → ULP 淺藍
        const hasAnyV_HMRU = isV(hpcTd) || isV(m22Td) || isV(r22Td) || isV(ulpTd);
        if (!hasAnyV_HMRU && oxideNum >= 9.20 && oxideNum <= 9.88) {
          ulpTd.style.backgroundColor = "#d9ecff";
        }

        // Oxide 8.71~9.75 且 HPC+/M22/R22/ULP/Ehv 都沒有 V → Ehv 淺藍
        const hasAnyV_HMRUE = isV(hpcTd) || isV(m22Td) || isV(r22Td) || isV(ulpTd) || isV(ehvTd);
        if (!hasAnyV_HMRUE && oxideNum >= 8.71 && oxideNum <= 9.75) {
          ehvTd.style.backgroundColor = "#d9ecff";
        }
      })();

      // ===================== [手動固定淺灰名單區] =====================
      (function applyGrayLockRules() {
        const grayLockRules = {
          //"HKG-A02C": ["col-hpc"],
          "HKG-B04D": ["col-m22", "col-r22", "col-ehv", "col-ulp"],
          "HKG-B05A": ["col-m22", "col-r22", "col-ehv", "col-ulp"],
          "HKG-B07B": ["col-ehv"],
          "HKG-B04B": ["col-ehv"],
          "HKG-B07C": ["col-ehv"]
        };

        const cols = grayLockRules[eqpid];
        if (!cols || !cols.length) return;

        cols.forEach(cls => {
          const td = tr.querySelector("td." + cls);
          if (!td) return;
          td.style.backgroundColor = "#e0e0e0";
        });
      })();

      // ===== S7 Remain：含顏色規則 =====
      if (s7Td) {
        const text = s7RemainByEqpid[eqpid];
        s7Td.textContent = text ? text : "";

        s7Td.style.backgroundColor = "";
        s7Td.style.color = "";
        s7Td.style.fontWeight = "";

        if (text) {
          const numStr = String(text).replace("%", "").trim();
          const num = Number(numStr);

          if (!isNaN(num)) {
            const weighDate = s7WeighDateByEqpid ? s7WeighDateByEqpid[eqpid] : null;
            let over60Days = false;

            if (weighDate instanceof Date && !isNaN(weighDate.getTime())) {
              const diffDays = (Date.now() - weighDate.getTime()) / (1000 * 60 * 60 * 24);
              over60Days = diffDays > 60;
            }

            if (num < 30 && over60Days) {
              s7Td.style.backgroundColor = "#ffd1dc";
              s7Td.style.color = "#b00020";
              s7Td.style.fontWeight = "bold";
            }
            else if (num < 15) {
              s7Td.style.backgroundColor = "#ffe0b2";
              s7Td.style.color = "red";
              s7Td.style.fontWeight = "bold";
            }
          }
        }
      }

      // ===== S7 Temp：> 175.9 紅字 =====
      if (s7TempTd) {
        const raw = (s7TempByEqpid[eqpid] !== undefined && s7TempByEqpid[eqpid] !== null)
          ? String(s7TempByEqpid[eqpid]).trim()
          : "";

        s7TempTd.textContent = raw;
        s7TempTd.style.color = "";
        s7TempTd.style.fontWeight = "";

        const num = Number(raw);
        if (!isNaN(num) && num > 175.9) {
          s7TempTd.style.color = "red";
          s7TempTd.style.fontWeight = "bold";
        }
      }

      // 階差（Base）
      if (pmTd) {
        const val = basePmByEqpid[eqpid];
        const txt = (val === undefined || val === null) ? "" : String(val).trim();
        pmTd.textContent = txt;
        pmTd.style.color = "";
        pmTd.style.fontWeight = "";
        const num = Number(txt);
        if (!isNaN(num) && num > 0.9) {
          pmTd.style.color = "black";
          pmTd.style.fontWeight = "bold";
        }
      }

      // UD_PM
      if (udpmTd) {
        const v = udPmByEqpid[eqpid];
        const txt = (v === undefined || v === null) ? "" : String(v).trim();
        udpmTd.textContent = (txt === "" || txt === "0") ? "" : txt;
      }

      // PM 日期
      if (pmDateTd) pmDateTd.textContent = (pmDateByEqpid && pmDateByEqpid[eqpid]) ? pmDateByEqpid[eqpid] : "";

      // S6_Min
      if (s6Td && s6MinByEqpid[eqpid] !== undefined) {
        const raw = s6MinByEqpid[eqpid];
        const txt = (raw === undefined || raw === null) ? "" : String(raw).trim();
        s6Td.textContent = txt;
        s6Td.style.color = "";
        s6Td.style.fontWeight = "";
        const num = Number(txt);
        if (!isNaN(num) && num > 141) {
          s6Td.style.color = "#2f80ed";
          s6Td.style.fontWeight = "bold";
        }
      }

      // W/C：先用快取 map
      if (wcTd) {
        wcTd.style.color = "";
        wcTd.style.fontWeight = "";

        const val = window.waferCountByEqpid ? window.waferCountByEqpid[eqpid] : undefined;

        if (val === undefined || val === null || String(val).trim() === "") {
          wcTd.textContent = "";
        } else {
          const n = Number(val);
          wcTd.textContent = isNaN(n) ? String(val).trim() : String(Math.round(n));

          if (!isNaN(n) && n > 5000) {
            wcTd.style.color = "blue";
            wcTd.style.fontWeight = "bold";
          }
        }
      }

      // Cycle 顯示
      if (cycleTd && cycleDataByEqpid[eqpid] !== undefined) {
        const cycleRaw = cycleDataByEqpid[eqpid];
        const cycleNum = Number(cycleRaw);
        cycleTd.textContent = cycleRaw;
        if (udpmTd) {
          const udpmText = udpmTd.textContent.trim();
          const hasUdpmValue = udpmText !== "" && udpmText !== "-" && udpmText !== "0";
          if (hasUdpmValue) {
            const udpmNum = Number(udpmText);
            if (!isNaN(cycleNum) && !isNaN(udpmNum)) {
              const diff = cycleNum - udpmNum;
              cycleTd.textContent = `${cycleNum} (${diff})`;
            }
          }
        }
      }

      // Delta (OXIDE)
      if (deltaTd && deltaByEqpid[eqpid] !== undefined) {
        const dVal = deltaByEqpid[eqpid];
        let txt = "";
        let num = null;
        const n = Number(dVal);
        if (!isNaN(n)) { num = n; txt = n.toFixed(3); }
        else txt = String(dVal);

        deltaTd.textContent = txt;
        deltaTd.style.color = "";
        deltaTd.style.fontWeight = "";
        deltaTd.style.backgroundColor = "";

        if (num !== null) {
          if (num > 0.2) { deltaTd.style.color = "blue"; deltaTd.style.fontWeight = "bold"; }
          else if (num < -0.2) { deltaTd.style.color = "red"; deltaTd.style.fontWeight = "bold"; }

          if (num > 0.42) deltaTd.style.backgroundColor = "#cce5ff";
          else if (num < -0.42) deltaTd.style.backgroundColor = "#ffd6d6";
        }
      }

      // Last PA
      if (lastPaTd) {
        const info = lastPaByEqpid[eqpid];
        lastPaTd.innerHTML = info ? info.html : "";
      }

      updateCount++;
    });

    statusEl.textContent = "資料更新完成，共更新 " + updateCount + " 筆設備。";
    statusEl.classList.add("success");

    buildChamberTilesFromTable();
    updateChamberTilesColorsFromTable();

    // summary
    (function updateSummaryLines() {
      const tbody2 = document.querySelector("#asmTable tbody");
      if (!tbody2) return;

      let countHMR = 0;
      let countEU = 0;
      const isV = (td) => td && (td.textContent || "").trim().toUpperCase() === "V";

      tbody2.querySelectorAll("tr").forEach(tr => {
        const hpcTd = tr.querySelector("td.col-hpc");
        const m22Td = tr.querySelector("td.col-m22");
        const r22Td = tr.querySelector("td.col-r22");
        const ehvTd = tr.querySelector("td.col-ehv");
        const ulpTd = tr.querySelector("td.col-ulp");

        if (isV(hpcTd)) countHMR++;
        if (isV(m22Td)) countHMR++;
        if (isV(r22Td)) countHMR++;
        if (isV(ehvTd)) countEU++;
        if (isV(ulpTd)) countEU++;
      });

      const hmrEl = document.getElementById("hmrSummary");
      const euEl = document.getElementById("euSummary");
      if (hmrEl) hmrEl.textContent = `HPC+/M22/R22 Chamber : ${countHMR} Chs`;
      if (euEl) euEl.textContent = `eHV/ULP Chamber : ${countEU} Chs`;
    })();

    enableTableSorting("asmTable");
    enableChClickForPA();
    enableEqpidClickForOxide();

    // 最後補一次 W/C（從 DB 拉取覆蓋）
    refreshWaferCountColumn();

  } catch (err) {
    console.error(err);
    const statusEl2 = document.getElementById("status");
    statusEl2.textContent = "讀取 Excel 時發生錯誤：" + err.message;
    statusEl2.className = "status error";
  }
}

// ===== Auto refresh config =====
const EXCEL_AUTO_REFRESH_MS = 5 * 60 * 1000; // 5 minutes
let lastExcelLastModifiedHeader = ""; // raw Last-Modified header value

function formatYmdHm(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${y}-${m}-${day} ${hh}:${mm}`;
}

function fetchAndLoadExcel(forceReload) {
  const statusEl = document.getElementById("status");
  if (statusEl) {
    statusEl.textContent = "正在從伺服器下載 Excel 檔案...";
    statusEl.className = "status";
  }

  fetch(EXCEL_URL, { cache: "no-store" })
    .then(response => {
      if (!response.ok) throw new Error("HTTP 狀態 " + response.status);

      // Server-provided file modification time.
      const lastModified = response.headers.get("Last-Modified") || "";

      // If only want reload on change
      if (!forceReload && lastExcelLastModifiedHeader && lastModified && lastModified === lastExcelLastModifiedHeader) {
        // No change, just update "read time" display
        const readText = formatYmdHm(new Date());
        const st = document.getElementById("status");
        if (st) st.dataset.readAt = readText;
        // refresh file info line to show the new read time (keep same lastModified text)
        renderFileInfo("HKG_GH.xlsm", (document.querySelector('#fileInfo .stale, #fileInfo span')?.textContent || "").trim());
        if (statusEl) statusEl.textContent = "Excel 無更新，略過重載（Last-Modified 未變）";
        return null;
      }

      lastExcelLastModifiedHeader = lastModified || lastExcelLastModifiedHeader;

      let lastModifiedText = "";
      if (lastModified) {
        const d = new Date(lastModified);
        lastModifiedText = (!isNaN(d.getTime())) ? formatYmdHm(d) : lastModified;
      }

      return response.arrayBuffer().then(buf => ({ buf, lastModifiedText }));
    })
    .then(payload => {
      if (!payload) return;
      const { buf, lastModifiedText } = payload;

      // record read time BEFORE rendering file info
      const st = document.getElementById("status");
      if (st) st.dataset.readAt = formatYmdHm(new Date());

      loadExcelFromArrayBuffer(buf, lastModifiedText);
    })
    .catch(err => {
      console.error(err);
      if (statusEl) {
        statusEl.textContent = "下載或解析 Excel 時發生錯誤：" + err.message;
        statusEl.className = "status error";
      }
    });
}

function initExcelAutoRefresh() {
  // First run already happens on DOMContentLoaded
  setInterval(() => {
    // Only reload when Last-Modified changed
    fetchAndLoadExcel(false);
  }, EXCEL_AUTO_REFRESH_MS);
}


function enableTableSorting(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const thead = table.querySelector("thead");
  if (!thead) return;
  const headers = thead.querySelectorAll("th");
  const tbody = table.querySelector("tbody");
  if (!tbody) return;

  let sortState = {};

  headers.forEach((th, colIndex) => {
    th.addEventListener("click", () => {
      const key = colIndex;
      let current = sortState[key] || "none";
      let next;
      if (current === "none") next = "asc";
      else if (current === "asc") next = "desc";
      else next = "none";

      headers.forEach(h => h.classList.remove("sort-asc", "sort-desc"));

      if (next === "none") {
        sortState[key] = "none";
        return;
      }

      sortState[key] = next;
      if (next === "asc") th.classList.add("sort-asc");
      else th.classList.add("sort-desc");

      const rows = Array.from(tbody.querySelectorAll("tr"));
      rows.sort((a, b) => {
        const tdA = a.children[colIndex];
        const tdB = b.children[colIndex];
        const textA = tdA ? tdA.textContent.trim() : "";
        const textB = tdB ? tdB.textContent.trim() : "";

        const numA = parseFloat(textA.replace(/,/g, ""));
        const numB = parseFloat(textB.replace(/,/g, ""));
        const isNumA = !isNaN(numA);
        const isNumB = !isNaN(numB);

        let cmp;
        if (isNumA && isNumB) cmp = numA - numB;
        else cmp = textA.localeCompare(textB, "zh-Hant");

        return next === "asc" ? cmp : -cmp;
      });

      rows.forEach(tr => tbody.appendChild(tr));
    });
  });
}

// === PA Chart ===
function showPAChartForEqpid(eqpid) {
  const paDataAll = window.paDataByProcessUnit || {};
  const unit = eqpid;
  const records = paDataAll[unit];
  if (!records || !records.length) {
    alert("找不到 " + eqpid + " 的 PA 資料。");
    return;
  }

  const labels = records.map(r => {
    const d = r.time;
    if (!(d instanceof Date) || isNaN(d.getTime())) return "";
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${y}-${m}-${day} ${hh}:${mm}`;
  });
  const values = records.map(r => (typeof r.mean === "number" ? r.mean : NaN));

  const modal = document.getElementById("paChartModal");
  const titleEl = document.getElementById("paChartTitle");
  const canvas = document.getElementById("paChartCanvas");
  titleEl.textContent = "PA Chart - " + eqpid;

  if (paChartInstance) { paChartInstance.destroy(); paChartInstance = null; }

  paChartInstance = new Chart(canvas.getContext("2d"), {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: "Mean Value",
        data: values,
        borderColor: "rgba(54, 162, 235, 1)",
        backgroundColor: "rgba(54, 162, 235, 0.1)",
        borderWidth: 2,
        pointRadius: 2,
        fill: true,
        tension: 0.1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { maxRotation: 45, minRotation: 0, autoSkip: true, maxTicksLimit: 15 } },
        y: { beginAtZero: true, min: 0, max: 15 }
      },
      plugins: {
        legend: { display: true },
        tooltip: { mode: "index", intersect: false }
      },
      interaction: { mode: "index", intersect: false }
    }
  });

  modal.classList.add("active");
}

// === OXIDE Chart ===
function showOxideChartForEqpid(eqpid) {
  const oxDataAll = window.oxideDataByEqpid || {};
  const records = oxDataAll[eqpid];
  if (!records || !records.length) {
    alert("找不到 " + eqpid + " 的 OXIDE 資料。");
    return;
  }

  const labels = records.map(r => {
    const d = r.time;
    if (!(d instanceof Date) || isNaN(d.getTime())) return "";
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${y}-${m}-${day} ${hh}:${mm}`;
  });
  const values = records.map(r => r.mean);

  const modal = document.getElementById("oxideChartModal");
  const titleEl = document.getElementById("oxideChartTitle");
  const canvas = document.getElementById("oxideChartCanvas");
  titleEl.textContent = "OXIDE Chart - " + eqpid;

  if (oxideChartInstance) { oxideChartInstance.destroy(); oxideChartInstance = null; }

  const targetLine = values.map(() => 9.55);

  oxideChartInstance = new Chart(canvas.getContext("2d"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "MEAN_VALUE",
          data: values,
          borderColor: "rgba(255, 99, 132, 1)",
          backgroundColor: "rgba(255, 99, 132, 0.1)",
          borderWidth: 2,
          pointRadius: 2,
          fill: true,
          tension: 0.1
        },
        {
          label: "Target 9.55",
          data: targetLine,
          borderColor: "green",
          borderWidth: 1.5,
          pointRadius: 0,
          fill: false,
          borderDash: [6, 4]
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { maxRotation: 45, minRotation: 0, autoSkip: true, maxTicksLimit: 15 } },
        y: { min: 8.8, max: 11.0 }
      },
      plugins: {
        legend: { display: true },
        tooltip: { mode: "index", intersect: false }
      },
      interaction: { mode: "index", intersect: false }
    }
  });

  modal.classList.add("active");
}

function enableChClickForPA() {
  const table = document.getElementById("asmTable");
  if (!table) return;
  const tbody = table.querySelector("tbody");
  if (!tbody) return;

  tbody.querySelectorAll("td.col-ch").forEach(td => {
    td.addEventListener("click", () => {
      const tr = td.closest("tr");
      if (!tr) return;
      const eqpidCell = tr.querySelector("td.col-eqpid");
      if (!eqpidCell) return;
      const eqpid = eqpidCell.textContent.trim();
      if (!eqpid) return;
      showPAChartForEqpid(eqpid);
    });
  });
}

function enableEqpidClickForOxide() {
  const table = document.getElementById("asmTable");
  if (!table) return;
  const tbody = table.querySelector("tbody");
  if (!tbody) return;

  tbody.querySelectorAll("td.col-eqpid").forEach(td => {
    td.style.cursor = "pointer";
    td.style.color = "#cc6600";
    td.style.textDecoration = "underline";

    td.addEventListener("click", () => {
      const eqpid = td.textContent.trim();
      if (!eqpid) return;
      showOxideChartForEqpid(eqpid);
    });
  });
}

(function initModalClose() {
  const modal = document.getElementById("paChartModal");
  const closeBtn = document.getElementById("paChartClose");
  if (!modal || !closeBtn) return;

  closeBtn.addEventListener("click", () => modal.classList.remove("active"));
  modal.addEventListener("click", e => { if (e.target === modal) modal.classList.remove("active"); });
})();

(function initOxideModalClose() {
  const modal = document.getElementById("oxideChartModal");
  const closeBtn = document.getElementById("oxideChartClose");
  if (!modal || !closeBtn) return;

  closeBtn.addEventListener("click", () => modal.classList.remove("active"));
  modal.addEventListener("click", e => { if (e.target === modal) modal.classList.remove("active"); });
})();

function initExcelUpload() {
  const input = document.getElementById("excelFileInput");
  const statusEl = document.getElementById("status");
  const fileInfoEl = document.getElementById("fileInfo");
  if (!input) return;

  if (input.dataset.bound === "1") return;
  input.dataset.bound = "1";

  input.addEventListener("change", () => {
    const file = input.files && input.files[0];
    if (!file) return;

    statusEl.textContent = "正在讀取上傳的檔案...";
    statusEl.className = "status";

    const lm = file.lastModified ? new Date(file.lastModified) : null;
    let lastModifiedText = "";
    if (lm && !isNaN(lm.getTime())) {
      const y = lm.getFullYear();
      const m = String(lm.getMonth() + 1).padStart(2, "0");
      const d = String(lm.getDate()).padStart(2, "0");
      const hh = String(lm.getHours()).padStart(2, "0");
      const mm = String(lm.getMinutes()).padStart(2, "0");
      lastModifiedText = `${y}-${m}-${d} ${hh}:${mm}`;
    } else {
      lastModifiedText = "未知";
    }

    if (fileInfoEl) renderFileInfo(file.name, lastModifiedText);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buf = e.target.result;
        loadExcelFromArrayBuffer(buf, lastModifiedText);
      } catch (err) {
        console.error(err);
        statusEl.textContent = "讀取上傳檔案時發生錯誤：" + err.message;
        statusEl.className = "status error";
      }
    };
    reader.onerror = () => {
      statusEl.textContent = "讀取上傳檔案失敗。";
      statusEl.className = "status error";
    };
    reader.readAsArrayBuffer(file);
  });
}

function initQuickLinkButtons() {
  const wrap = document.getElementById("quickLinks");
  if (!wrap) return;

  if (wrap.dataset.bound === "1") return;
  wrap.dataset.bound = "1";

  const BASE = "http://umcesidb02/PoC/TF/TF2/EQ1/ASM/";

  wrap.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-open]");
    if (!btn) return;

    e.preventDefault();
    e.stopPropagation();

    const file = btn.getAttribute("data-open") || "";
    const url = /^https?:\/\//i.test(file) ? file : (BASE + file);
    window.open(url, "_blank", "noopener");
  });
}

function getAsmDailySnapshot() {
  const tbody = document.querySelector("#asmTable tbody");
  const rows = [];
  if (tbody) {
    tbody.querySelectorAll("tr").forEach(tr => {
      const get = (cls) => {
        const td = tr.querySelector("td." + cls);
        return td ? (td.textContent || "").trim() : "";
      };
      rows.push({
        eqpid: get("col-eqpid"),
        cycle: get("col-cycle"),
        pmdate: get("col-pmdate"),
        coat: get("col-coat"),
        state: get("col-state"),
        hpc: get("col-hpc"),
        m22: get("col-m22"),
        r22: get("col-r22"),
        ehv: get("col-ehv"),
        ulp: get("col-ulp"),
        m22inhib: get("col-m22inhib"),
        hpcinhib: get("col-hpcinhib"),
        s6: get("col-measure"),
        wc: get("col-wc"),
        oxide: get("col-oxide"),
        s7: get("col-s7"),
        s7temp: get("col-s7temp"),
        udpm: get("col-chart"),
        pm: get("col-pm"),
        delta: get("col-delta"),
        ch: get("col-ch"),
        lastpa: get("col-lastpa")
      });
    });
  }

  const fileInfoEl = document.getElementById("fileInfo");
  const statusEl = document.getElementById("status");

  // tiles snapshot
  const tilesWrap = document.getElementById("chamberTiles");
  const tiles = [];
  if (tilesWrap) {
    tilesWrap.querySelectorAll(".tile").forEach(el => {
      const eqpid = (el.dataset.eqpid || el.textContent || "").trim();
      if (!eqpid) return;
      // tile class list includes: tile + state-class
      const classes = Array.from(el.classList);
      const stateClass = classes.find(c => c !== "tile") || "";
      tiles.push({ eqpid, class: stateClass, classes });
    });
  }

  // PA/OXIDE raw (convert Date -> ISO string to be JSON/prompt safe)
  const toIso = (d) => {
    try {
      if (!d) return null;
      const dt = (d instanceof Date) ? d : new Date(d);
      if (!(dt instanceof Date) || isNaN(dt.getTime())) return null;
      return dt.toISOString();
    } catch { return null; }
  };

  const paRaw = {};
  const paAll = window.paDataByProcessUnit || {};
  Object.keys(paAll).forEach(eqpid => {
    const list = paAll[eqpid] || [];
    paRaw[eqpid] = list.map(r => ({ time: toIso(r.time), mean: r.mean }));
  });

  const oxideRaw = {};
  const oxAll = window.oxideDataByEqpid || {};
  Object.keys(oxAll).forEach(eqpid => {
    const list = oxAll[eqpid] || [];
    oxideRaw[eqpid] = list.map(r => ({ time: toIso(r.time), mean: r.mean }));
  });

  return {
    summary: {
      today: (document.getElementById("todayText")?.textContent || "").trim(),
      hmrSummary: (document.getElementById("hmrSummary")?.textContent || "").trim(),
      euSummary: (document.getElementById("euSummary")?.textContent || "").trim(),
      fileInfo: (fileInfoEl?.textContent || "").trim(),
      status: (statusEl?.textContent || "").trim()
    },
    rows,
    tiles,
    paDataByProcessUnit: paRaw,
    oxideDataByEqpid: oxideRaw
  };
}

// expose for AiDailyBridge.js
window.getAsmDailySnapshot = getAsmDailySnapshot;

document.addEventListener("DOMContentLoaded", () => {
  fetchAndLoadExcel(true);
  initExcelAutoRefresh();
  initQuickLinkButtons();
  initExcelUpload();

  // 也先嘗試抓一次 W/C
  refreshWaferCountColumn();
});
</script>

<!-- ===== AiDailyBridge.js (inlined) ===== -->
<script>
// Bridge: embed AiChat widget into ASM Daily and allow AI to analyze current page data.
// This script contains the AiChat widget JS (from AiChat.aspx) plus a small
// injection that adds an "分析目前資料" button.

(function aiChatWidget() {
  const CONFIG = (window.AiChatConfig || {});
  const PROXY_URL = CONFIG.proxyUrl || (window.location.pathname + '?op=chat');
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
    analyze.title = '把目前 ASM Daily 表格資料整理後丟給 AI';
    analyze.addEventListener('click', () => {
      try {
        const snapshot = window.getAsmDailySnapshot ? window.getAsmDailySnapshot() : null;
        if (!snapshot) {
          alert('目前沒有可分析的資料（請先等 Excel 讀取完成）。');
          return;
        }
        const prompt = buildAnalysisPrompt(snapshot);
        if (input) input.value = prompt;
        open();
        if (input) input.focus();
      } catch (e) {
        alert('產生分析資料失敗: ' + e.message);
      }
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
</script>

</body>
</html>
