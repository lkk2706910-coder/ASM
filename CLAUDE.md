# ASM

## 用途
ASM 機台 Daily Dashboard，包含：（1）讀 SQL Server `XSITEUSAGEMETER` 顯示 wafer counts；（2）內嵌可獨立部署的 AI Chat widget proxy。

## 主要檔案
- `ASM Daily.aspx` / `ASM Daily.css` / `ASM Daily.js` — Dashboard 前端 + WebMethod 後端
- `AiChat.aspx` / `AiChat.aspx.cs` — AI Chat widget（標準化 LLM 代理，可獨立 drop-in）
- `AiDailyBridge.js` — Daily 介面與 AiChat 的橋接
- `web.config` — 含 `AiGatewayUrl` / `AiApiKey` / `AiUserId` / `AiSystemPrompt`

## 對外可複用功能
- `AiChat.aspx.cs:HandleChat()` — 把 client 的 messages forward 給上游 LLM，並在 server 端強制注入 system prompt（client 無法覆寫）
- `ASM Daily.aspx:GetWaferCounts()` — SqlClient WebMethod 讀 `[GPTDB_EAS].[dbo].[XSITEUSAGEMETER]`，回傳 EQPID → DATA_VAL 字典
- 整套可複製到任何 IIS 站台當作 AI 助理範本

## 依賴
.NET Framework 4.x / IIS、SheetJS（CDN）、Chart.js（CDN）

## 開發備註
- ⚠️ **`web.config` 內含真實 AI API key（UUID 格式）**，請優先處理：rotate key、從 git history 移除
- `AiSystemPrompt` 可直接從 web.config 覆寫，不需動程式碼
- LLM gateway URL 指向內網（`<內部 server>` GPT 服務）
- ConnectionString `GPTDB_EAS` 從 web.config 讀取（未在 repo 內寫死，做法正確）
