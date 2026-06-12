<%@ Page Language="C#" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="System.Web.Script.Services" %>

<script runat="server">
  // W/C API
  // Server: UMCESIDB02
  // Table:  [GPTDB_EAS].[dbo].[XSITEUSAGEMETER]
  // 連線字串請放在 web.config 的 <connectionStrings>

  [WebMethod]
  [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
  public static Dictionary<string, object> GetWaferCounts()
  {
    var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

    // 從 web.config 讀取連線字串（這行註解要獨立一行）
    var connStrSettings = System.Configuration.ConfigurationManager.ConnectionStrings["GPTDB_EAS"];

    if (connStrSettings == null || string.IsNullOrWhiteSpace(connStrSettings.ConnectionString))
      throw new Exception("Missing connection string: GPTDB_EAS");

    var connStr = connStrSettings.ConnectionString;

    var sql = @"
      SELECT EQPID, DATA_VAL
      FROM [dbo].[XSITEUSAGEMETER]
      WHERE METERTYPE = 'CH_WAFER_COUNT'
    ";

    using (var conn = new SqlConnection(connStr))
    using (var cmd = new SqlCommand(sql, conn))
    {
      conn.Open();
      using (var rdr = cmd.ExecuteReader())
      {
        while (rdr.Read())
        {
          var eqpid = (rdr["EQPID"] == DBNull.Value) ? "" : rdr["EQPID"].ToString().Trim();
          if (string.IsNullOrEmpty(eqpid)) continue;

          object val = (rdr["DATA_VAL"] == DBNull.Value) ? null : rdr["DATA_VAL"];
          result[eqpid] = val;
        }
      }
    }

    return result;
  }
</script>

<!DOCTYPE html>
<html lang="zh-Hant">
<head runat="server">
  <meta charset="UTF-8" />
  <title>ASM Daily</title>

  <!-- 引入 SheetJS 用於讀取 Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- 引入 Chart.js 用於畫圖 -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>

  <link rel="stylesheet" href="ASM Daily.css" />

  <!-- AiChat widget styles (embedded from AiChat.aspx) -->
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

<script src="ASM Daily.js"></script>
<script src="AiDailyBridge.js"></script>
</body>
</html>
