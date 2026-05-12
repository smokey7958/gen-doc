# Gen Doc

統一筆記 / 文書 / 表格 / 簡報，內建 AI 編輯助手。一個 `.gd` 容器內可以同時放 Markdown、Word、Excel、PowerPoint 四種文件，右側固定的 AI panel 透過 tool calling 對當前文件做結構化編輯。

完整產品 spec 見 [docs/MVP-SPEC.md](docs/MVP-SPEC.md)。

## 本次落地的實作範圍

這個 commit 是依照 spec 的 **Phase 1–4 vertical slice**：以 Markdown 為主軸跑通完整的「使用者下指令 → AI tool call → diff preview → apply → undo → 存檔」迴圈。

### 已實作（runnable）

- Electron + React 18 + Vite + TypeScript 骨架
- Tab 系統（新增 / 重命名 / 拖曳排序 / 關閉 / dirty 標記）
- Markdown tab 完整編輯（CodeMirror 6 + markdown 語法 + selection-aware）
- **Excel tab 編輯**（SheetJS 路徑 B MVP）：多 sheet、cell 文字／數值編輯、樣式 / 圖表 / 公式快取值 round-trip 保留
- **Word tab 編輯**（mammoth + docx 路徑 B MVP）：段落 / 標題 / 列表結構編輯，inline 格式（粗體 / 斜體 / 圖片 / 表格）round-trip 會被簡化為純文字
- **PowerPoint tab 編輯**（JSZip + OOXML 路徑 B MVP）：可編既有 pptx 的 `<a:t>` 文字 run，layout / 圖片 / 樣式完整保留；新增 / 刪除投影片尚未支援
- `.gd` 容器讀寫（JSZip）— manifest schema、atomic save、自動備份輪轉
- 暫存目錄管理 + 啟動時清理 7 天前殘留
- AI Panel：streaming chat、selection context badge、模型切換、tool toggle
- Anthropic SDK 整合（main process）+ prompt caching markers
- Tool calling loop（spec §6.2 step [1]–[11]）
- 18 個 tool 的 schema 定義；markdown / `excel_set_cell` / `excel_set_range` / `word_replace_paragraph` / `word_insert_paragraph` / `pptx_replace_text` / `read_tab_content` / `cross_tab_summarize` 可實際執行；樣式 / 圖表 / 投影片新增等進階操作回 `not_implemented_in_mvp`
- Diff preview（markdown LCS 行級 diff；其它格式顯示 byte-delta + 描述）
- Apply / Reject / Modify 三按鈕
- Undo stack（SQLite，每 workspace 50 步上限）— binary_replace op 也支援
- 設定頁：API key 經 Electron `safeStorage`（OS keystore）加密儲存、Test connection ping、預設模型、temperature、theme
- 全套快捷鍵（Ctrl/Cmd + N/O/S/Shift+S/Z/L/1-9）
- Native menu（檔案 / 編輯 / 檢視 / AI）

### 路徑 B MVP 限制（之後升級到 Univer 才解）

- **xlsx**：無公式 UI（編輯後寫純值）、無新增 / 刪除 sheet UI
- **docx**：inline runs（粗體 / 斜體 / 連結）round-trip 會掉；表格 flatten 成 tab-separated 純文字
- **pptx**：只能改既有文字框；新增 / 刪除投影片、改字體 / 顏色尚未支援；空 pptx tab 需從既有 .pptx 載入

### 明確 stub

- 進階格式工具：`word_apply_style`、`word_insert_heading`、`excel_apply_format`、`excel_insert_row`、`excel_insert_chart`、`pptx_add_slide`、`pptx_add_bullets`、`convert_md_to_docx`
- Electron Builder 安裝包、Phase 5 的 telemetry / first-run onboarding / sample.gd
- Redo（toolbar 按鈕已停用）；spec §5.4.5 的 redo 設計留 v1.5

## 開發

```bash
npm install
npm run dev          # 開發模式（auto-reload）
npm run typecheck    # tsc --noEmit
npm run build        # 三個 bundle (main / preload / renderer)
```

> ℹ️ Postinstall 會自動做兩件事：
> 1. `npm run ensure-electron` — 確保 Electron 平台 binary 存在（被 `--ignore-scripts`
>    跳過時會自動補抓）
> 2. `electron-rebuild` — 把 `better-sqlite3` 對 Electron 的 Node ABI rebuild
>
> 如果曾用 `--ignore-scripts` 安裝後 dev server 出現 `Error: Electron uninstall`
> 或 `openDatabase()` fail，跑 `npm run ensure-electron && npm run rebuild` 即可。
> Windows native 編譯需 Visual Studio Build Tools，macOS 需要 Xcode CLT。

## 第一次使用

1. `npm run dev` → App 啟動
2. 右下「設定 API key」→ 貼 `sk-ant-...` → 儲存 → Test connection
3. tab bar `+` → Markdown，輸入內容
4. 選一段文字 → AI panel 上方會出現 selection badge → 在輸入框打「把這段改成更精煉的版本」
5. AI 發 tool call → 主編輯區出現 diff preview → `Apply`
6. `Ctrl/⌘+S` → 第一次跳檔案對話框，存成 `.gd`
7. `Ctrl/⌘+Z` 可隨時 undo AI 的變更

## 架構備註

詳見 [STACK.md](STACK.md) 與 spec §8。要點：

- **AI engine 在 renderer**，但 **Anthropic SDK 呼叫在 main**（API key 不過 IPC 邊界）
- **File I/O 在 main**：renderer 沒有 fs 權限，全部走 typed IPC（`src/types/ipc.ts`）
- **Doc model 走 Zustand**：每個 tab 一份 state，AI ChangeSet apply 與 UI 共用同一個 store
- **SQLite 在 main**：`better-sqlite3` 是 sync native module，呼叫包成 async IPC handler
- **`.gd` 是 zip**：JSZip 統一用 forward slash 與 UTF-8，跨平台檔名安全

## Roadmap

詳見 spec §12。下一個重點是 v1.5 的 Univer 整合 spike，把三種 Office 格式從 byte-preserving placeholder 升級成完整編輯。
