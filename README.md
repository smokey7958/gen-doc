# Gen Doc

統一筆記 / 文書 / 表格 / 簡報的桌面編輯器，內建 AI 編輯助手。一個 `.gd` 容器內可同時放 **Markdown、HTML、Word、Excel、PowerPoint** 五種文件，右側固定的 AI panel 透過 tool calling 對文件做結構化編輯——所有 AI 變更先以 diff preview 呈現，按 Apply 才落地，隨時可 undo。

介面支援**中 / 英雙語**（依 OS 語系自動偵測，可在設定切換）。

完整產品 spec 見 [docs/MVP-SPEC.md](docs/MVP-SPEC.md)（內含歷史 Phase 紀錄）。

## 功能總覽

### 工作區與檔案

- `.gd` 容器（zip）讀寫：manifest 驗證（重複 id / 路徑穿越防護）、atomic save、自動備份輪轉（保留 3 份）、暫存目錄自動清理
- Tab 系統：新增 / 重命名 / 拖曳排序 / 關閉 / dirty 標記、右鍵選單（關閉其他 / 右側 / 全部）、`Ctrl+Tab` 循環切換、`Ctrl+1..9` 直達
- 檔案總管側欄：瀏覽資料夾、點擊開檔、已開啟 / 使用中標記
- 拖放開檔：`.gd / .md / .txt / .html / .docx / .xlsx / .pptx` 拖進視窗即開
- 最近開啟清單（選單 + 歡迎頁）
- 匯出：單一頁籤（`Ctrl+E`）或批次匯出多個頁籤到資料夾（`Ctrl+Shift+E`），各以原生格式落地

### Markdown（CodeMirror 6）

- 三種檢視：原始碼 / 對照 / 純預覽（GFM 表格、checkbox、刪除線）
- 格式工具列 + 完整快捷鍵（`Ctrl+B/I/K/1-3`、清單、引用…）、smart list 延續、auto-pair
- 大綱側欄（標題跳轉）、尋找取代（`Ctrl+F`，regex / 大小寫 / 全字）、跳至行（`Ctrl+G`）
- 圖片貼上 / 拖入（base64 內嵌）、**匯出 PDF**
- 狀態列字數統計

### HTML（CodeMirror 6 + 沙箱 iframe）

- 三種檢視（原始碼 / 對照 / 預覽），iframe 沙箱不執行 script
- 語法高亮、配對標籤、自動補齊收尾標籤、`Ctrl+F` 尋找取代

### Word（自製 block 編輯器 + mammoth / docx round-trip）

- WYSIWYG 紙張畫布：頁面大小（A4/Letter/…）、直橫向、邊界三檔，**編輯位置與列印檢視完全一致**（區塊控制元件懸浮在頁邊距）
- 樣式下拉（內文 / 標題 1-3 / 項目符號 / 編號清單）
- 選取範圍格式：**粗體 / 斜體 / 底線 / 刪除線**（`Ctrl+B/I/U`、`Ctrl+Shift+X`），per-run 完整 round-trip
- 段落層級：對齊（左/中/右/兩端）、字型、字級、字色、**螢光標示（6 色）**、**行距（1 / 1.15 / 1.5 / 2）**、超連結、**分頁符**
- 表格（新增/刪除列欄、Tab 巡覽）、圖片（插入 / 等比縮放 / 自由拖放定位）
- 「插入 Markdown / HTML 內容」對話框：原始碼解析成 Word 段落 / 標題 / 清單 / 表格
- 導覽窗格（標題大綱跳轉）、尋找取代、跳至段落（`Ctrl+G`）、列印預覽（工具列保持可用）
- 框選多段落整體拖移 / 刪除；編輯器內建 undo / redo（`Ctrl+Z / Ctrl+Y`）

### Excel（自製 grid + SheetJS round-trip）

- **自寫公式引擎**：tokenizer + parser + evaluator，函數 dispatch 至 formulajs（SUM / IF / VLOOKUP / XLOOKUP 等 50+），支援跨表參照、絕對參照、循環偵測（`#CYCLE!`）、fixpoint 重算
- 公式列 + Name Box（`Ctrl+G` 聚焦、輸入 B5 跳格）
- 範圍選取（拖曳 / Shift+方向鍵）、鍵盤巡覽（Enter / Tab / 方向鍵）
- 儲存格樣式：粗體 / 斜體 / 底線、對齊、字色 / 底色、字型 / 字級、數字格式（千分位 / 貨幣 / 百分比 / 日期…）——**存檔重開完整還原**（直接解析 styles.xml）
- 合併儲存格、**凍結首列 / 首欄**（寫入標準 OOXML，Excel 開啟同樣生效）
- **欄寬 / 列高拖曳調整、雙擊自動適寬**，持久化到 xlsx
- 剪貼簿：範圍複製 / 剪下 / 貼上（公式相對參照自動位移）、**`Ctrl+D / Ctrl+R` 向下 / 向右填滿**
- 多工作表：新增 / 重命名 / 複製 / 刪除 / 拖曳排序、右鍵選單、per-sheet 選取與捲動記憶
- 浮動圖片（插入 / 拖移錨定 / 縮放，寫回 OOXML drawing）
- 跨表尋找取代、狀態列彙總（Sum / Avg / Count）、undo / redo

### PowerPoint（自製 slide 畫布 + OOXML byte-preserving round-trip）

- 以既有 `.pptx` 為基底編輯，layout / master / 動畫等未編輯內容 **byte 級保留**
- 投影片：新增（複製目前頁）/ 刪除 / 縮圖拖曳排序、6 種版面配置套用
- WYSIWYG 畫布：文字框 / 形狀拖移、8 向縮放、對齊輔助線（snap guides）、對齊到投影片（左/中/右/上/中/下）
- 文字格式：粗體 / 斜體 / 底線、字色、字級、字型、**段落對齊（左/中/右）**
- 圖形：矩形 / 圓角矩形 / 橢圓 / 三角形 / 箭頭，**填色、外框（顏色 / 粗細 / 無）、移到最前 / 最後（z-order）**
- 圖片插入（選檔 / 拖放至投影片）、Alt+拖曳複製圖形
- 演講者備忘稿編輯、**放映模式**（F5；方向鍵翻頁、N 切換備忘稿、結尾黑幕）
- 投影片大綱側欄、跨投影片尋找取代、跳至投影片（`Ctrl+G`）、undo / redo

### AI Panel（Anthropic）

- Streaming chat、selection-aware context badge（選取文字自動帶入）、模型切換（Sonnet 4.6 / Opus 4.7）、tool 開關
- Tool calling 完整迴圈：tool_use → ChangeSet → **diff preview**（markdown 行級 LCS；binary 格式顯示 byte-delta + 描述）→ Apply / Reject / Modify
- AI 變更 undo / redo stack（SQLite，每 workspace 50 步）
- Token 用量與成本即時顯示（含 prompt cache 命中率）、prompt caching markers
- 18 個 tool schema；可實際執行 12 個：`md_replace_section / md_append / md_insert_at`、`excel_set_cell / excel_set_range`、`word_replace_paragraph / word_insert_paragraph / word_insert_heading`、`pptx_replace_text`、`read_tab_content`、`convert_md_to_docx`、`cross_tab_summarize`；樣式 / 圖表 / 投影片新增等 6 個進階 tool 回 `not_implemented_in_mvp`
- API key 經 Electron `safeStorage`（OS keystore）加密，明文不落地、不過 IPC 邊界

### 安全與效能

- Renderer 無 fs 權限，全部走 typed IPC；`contextIsolation` + sandbox
- IPC 檔案存取防護：目錄列舉限使用者授權過的根目錄、檔案讀取限支援類型
- 啟動 bundle 程式碼分割：編輯器與 AI 工具鏈按需載入（啟動 JS ~1.7 MB）

## 已知限制

- **Word**：頁首頁尾 / 註腳 / 追蹤修訂僅讀取保留，無編輯 UI；表格儲存格無合併 / 底色
- **Excel**：圖表、條件格式、篩選排序、樞紐分析尚未支援；開檔時既有浮動圖片不保留（SheetJS 限制，App 內插入的圖片正常 round-trip）
- **PowerPoint**：空白 pptx 頁籤需從既有 `.pptx` 開始；表格 / SmartArt 顯示為文字、不可結構化編輯；master slide 唯讀
- **Markdown**：KaTeX / Mermaid 預覽尚未支援
- 多人協作、雲端同步、行動裝置不在 v1.0 範圍

## 開發

```bash
npm install
npm run dev          # 開發模式（auto-reload）
npm run typecheck    # tsc --noEmit
npm run lint         # ESLint（flat config）
npm run build        # 三個 bundle（main / preload / renderer）
npm run package      # electron-builder --dir（可直接執行的 win-unpacked）
npm run dist         # 完整安裝包（Windows NSIS / macOS dmg / Linux AppImage）
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

1. `npm run dev` → App 啟動（或執行打包後的 `release/win-unpacked/Gen Doc.exe`）
2. AI panel「設定 API key」→ 貼 `sk-ant-...` → 儲存 → Test connection
3. Tab bar `+` → 選格式（Markdown / HTML / Word / Excel / PowerPoint），開始編輯
4. 選一段文字 → AI panel 出現 selection badge → 輸入「把這段改成更精煉的版本」
5. AI 發 tool call → diff preview → `Apply`
6. `Ctrl/⌘+S` → 第一次跳檔案對話框，存成 `.gd`
7. `Ctrl/⌘+Z` 可隨時 undo AI 的變更

## 架構備註

詳見 [STACK.md](STACK.md) 與 spec §8。要點：

- **AI engine 在 renderer**，但 **Anthropic SDK 呼叫在 main**（API key 不過 IPC 邊界）
- **File I/O 在 main**：renderer 沒有 fs 權限，全部走 typed IPC（`src/types/ipc.ts`）＋授權根目錄驗證
- **三種 Office 格式為自製編輯器**：mammoth/docx（Word）、SheetJS ×2（Excel）、JSZip + OOXML regex 修補（PowerPoint），未編輯的內容以 byte-preserving 策略保留原樣
- **Doc model 走 Zustand**：每個 tab 一份 state，AI ChangeSet apply 與 UI 共用同一個 store
- **SQLite 在 main**：`better-sqlite3` 是 sync native module，呼叫包成 async IPC handler
- **`.gd` 是 zip**：JSZip 統一用 forward slash 與 UTF-8，跨平台檔名安全

## Roadmap

詳見 spec §12。候選方向：Excel 圖表與條件格式、Word 表格進階（合併 / 底色）、pptx 空白簡報與表格編輯、KaTeX / Mermaid 預覽、AI 樣式類 tool 補完。
