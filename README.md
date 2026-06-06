# AI Report Agent

原 `AI-`(customtkinter 桌面工具)的**前端重構版**:把混亂的多分頁介面,改成一條**看得見的報告流水線**。後端沿用既有報告引擎,前端改用 React 視覺化呈現。

> 📖 **完整使用說明(手動操作 + Agent 操作,逐步圖解)→ [`docs/USAGE.md`](docs/USAGE.md)**
> 第一次使用、或想知道某個功能怎麼點/怎麼用中文叫 Agent 做,請看那份。本 README 只談架構與啟動。

## 為什麼重構

舊版把「製作報告」拆在 6 個並列分頁,使用者看不出先做什麼、在哪一步。新版把核心流程變成 5 步精靈:

```
① 設定來源 → ② 對應標籤 → ③ 驗證 → ④ 產出 → ⑤ 完成/審查
```

每一步都有即時預覽與「現在在做什麼/下一步」的引導。

## 架構

```
AI_Report_Agent/
├─ backend/              FastAPI 本機服務(port 8756)
│  ├─ app/               移植自 AI- 的報告引擎(generator / excel_template /
│  │                     excel_transfer / filename / settings / agent 子系統)
│  ├─ headless_context.py  無 Tk 版 AppContext(agent 與 REST 共用狀態)
│  ├─ native_dialog.py   原生 OS 檔案選取對話框
│  ├─ server.py          API + WebSocket + 靜態前端服務
│  └─ static/            前端 build 輸出(由 FastAPI 直接服務)
├─ frontend/             Vite + React + TS + Tailwind
└─ launch.bat            一鍵啟動(開後端 + 瀏覽器)
```

- **後端**:重用原 `app/` 邏輯,零修改。狀態存在 `~/.auto_report/settings.json`(與舊桌面版共用)。
- **前端**:單頁 SPA,build 後由後端同源服務,不需額外 Node 伺服器。

## 功能(與舊版對等)

| 畫面 | 內容 |
|---|---|
| 報告精靈 | Word / Excel 兩種報告;視覺化標籤編輯器(點欄位→點位置插入)+ AI 建議對應;驗證;SSE 即時產出進度;審查結果 |
| Excel 搬移 | 欄位對應(自動對應同名)、三種寫入模式 |
| AI 助手 | 自然語言操作(WebSocket 串流);沿用原 agent 工具與 orchestrator |
| AI 引擎 | Gemini / Ollama 模型選擇、審查設定、呼叫預算 |
| 設定 | 預設值、圖片寬度、熱鍵橋接目標 |

## 啟動

### 一鍵(正式版)
```
launch.bat
```
會用 `..\AI-\.venv` 的 Python 跑後端,並開瀏覽器到 http://127.0.0.1:8756

### 開發模式(前端熱更新)
```
:: 終端機 1 — 後端
..\AI-\.venv\Scripts\python.exe backend\server.py
:: 終端機 2 — 前端 dev server(會 proxy /api 到後端)
cd frontend
npm run dev
```
然後開 http://127.0.0.1:5275

### 改前端後重新 build
```
cd frontend
npm run build      :: 輸出到 backend/static
```

## 相依

後端:見 `backend/requirements.txt`(已安裝於 `AI-/.venv`)。
前端:見 `frontend/package.json`。

## 注意

- **熱鍵橋接(Ctrl+Shift+M)** 需要本機 Office(COM),為進階功能;在「對應標籤」步驟用滑鼠即可完成相同的事。
- **AI 審查 / docx 預覽** 需要本機安裝 Word(透過 COM 轉 PDF)。
- API key 放 `AI-/.env` 的 `GEMINI_API_KEY`,不進版控。
