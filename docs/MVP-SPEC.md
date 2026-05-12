# Gen Doc — MVP Specification (v1.0)

> Status: Draft / Locked stack
> Owner: Spec Agent (Agent 1)
> Last updated: 2026-05-11
> Audience: 工程團隊內部 spec，PM / Tech Lead / 各 sub-agent 共讀

> **2026-05-11 update**：新增 HTML 分頁類型（介於 Markdown 與 Word 之間，三種檢視模式：原始碼 / 對照 / 預覽）；新增「批次匯出多個頁籤」工作流（單次資料夾選擇器，每個勾選頁籤以原生格式落地）。詳見 §3.1、§4.7、§5.1、§5.2.2、§5.2.5、§7.1。

---

## 1. 產品願景與目標

Gen Doc 是一款桌面端「統一文書編輯器」，把使用者日常會用到的五種文件型別 — Markdown 筆記、HTML 文件、Word 文件、Excel 表格、PowerPoint 簡報 — 整合在同一個 Electron App 內，以分頁形式並列編輯，整本工作集可打包成單一 `.gd` 檔（zip 容器）攜帶或備份。右側固定一個 AI panel，類似 VS Code Copilot Chat / Cursor 的對話介面，使用者用自然語言下指令，AI 透過 tool calling 對當前文件做結構化編輯，所有變更先以 diff preview 呈現，使用者按 apply 才真正落地，按 undo 可隨時退回。

我們不是再做一個 Office 替代品，也不是再做一個 Notion clone。我們要解決的是「為了一個提案，使用者要在 Word、Excel、PowerPoint、Markdown 筆記之間切換 4 個 App，AI 工具又散落在各家網頁」這個碎片化現況。把編輯器與 AI 助手整合在同一個 surface，讓 AI 可以同時看到簡報的某張投影片、Excel 的某個 range、Word 的某段落，並做跨檔操作，這是 Gen Doc 的差異化。

### 1.1 v1.0 五條核心目標

1. **一個 App 能編輯五種主流格式**：md、html、docx、xlsx、pptx 全部能開、能編、能存，不需要外部依賴 Office / WPS / 瀏覽器。
2. **`.gd` 容器格式可攜帶**：使用者一次打包，跨機器開啟內容完全一致，不依賴雲端帳號。
3. **AI panel 可實際改動文件**：不只是 chat suggestion，而是 tool calling → ChangeSet → preview → apply 的完整 loop，使用者改一份提案的時間能比手動操作縮短 30% 以上（內部 dogfooding 量測）。
4. **對使用者透明的 AI 編輯**：每一次 AI 改動都產生 diff preview 與 undo entry，使用者永遠知道 AI 動了什麼，可隨時退回。
5. **效能可接受**：開啟 10MB 以下的 docx / xlsx / pptx 在 3 秒內呈現可編輯狀態，AI 第一個 token 在 1.5 秒內回來（仰賴 prompt caching）。

---

## 2. 目標使用者

我們鎖定三個 persona，按優先序排列。v1.0 主要為 P1 與 P2 設計，P3 在 v1.5 才被認真服務。

### 2.1 Persona 1：獨立顧問 / 個人工作者 Anna（主要 TA）

- 35 歲，財務顧問，每週要為 3-5 個客戶各做一份提案
- 工作流：Excel 跑數字 → Word 寫報告 → PowerPoint 做呈現
- 痛點：三個檔案之間複製貼上、格式跑掉、AI 工具（ChatGPT、Claude）要不停切視窗、貼進貼出
- Gen Doc 對她的價值：一個檔案打包整個提案，AI panel 一次看懂三份內容，幫她「把 Excel 的 Q3 數字摘要進 Word 第二段，再做成 PPT 第 5 頁的 bullet」

### 2.2 Persona 2：技術文件作者 / RD Lead Ben（次要 TA）
- 30 歲，後端工程師兼技術主管
- 寫 RFC / spec / runbook / 月報，主要用 Markdown，但偶爾要轉成 docx 給主管或 pptx 給高層
- 痛點：Markdown 編輯器（Obsidian / Typora）沒法直接編 docx，跨格式轉換要靠 pandoc
- Gen Doc 對他的價值：Markdown 為主、docx / pptx 為輔，AI 幫他把 spec 摘要轉成主管簡報

### 2.3 Persona 3：學生 / 教師 Cathy（v1.5 才認真服務）

- 報告、講義、評分表三種格式並用
- v1.0 不特別針對她設計表單 / 公式向導等功能
- 提及她是為了確認 v1.0 的基本格式相容性夠用即可

---

## 3. MVP 範圍與非範圍

清楚劃線。v1.0 的目的是**證明整合 + AI 編輯這個 thesis**，不是做一個功能對等 Office 的產品。

### 3.1 v1.0 IN-SCOPE

| 領域 | 包含 |
|------|------|
| 平台 | Windows 10/11、macOS 12+（Apple Silicon 與 Intel） |
| 檔案格式 | md、html、docx、xlsx、pptx 開啟與儲存；`.gd` 打包與解包 |
| 編輯 UI | Tab 系統、單頁簽編輯、Tab 重命名 / 拖曳排序 / 關閉 |
| 匯出 | 單一頁籤匯出（Ctrl+E 或 toolbar / per-tab Download）、批次匯出多個頁籤到資料夾（toolbar `FileDown` 按鈕） |
| 跨格式 | DocxEditor 工具列可「插入 Markdown 內容」與「插入 HTML 內容」，原始碼解析後轉成 Word 段落 / 標題 / 清單 / 表格 |
| AI Panel | Chat、selection-aware context、tool calling、diff preview、apply、undo |
| 設定 | API key 設定、模型切換（Sonnet 4.6 / Opus 4.7）、快捷鍵覆寫 |
| Storage | local file、`.gd` 容器、SQLite 存 chat history |

### 3.2 v1.0 OUT-OF-SCOPE（明確留到 v1.5+）

| 不做 | 理由 |
|------|------|
| 多人即時協作 | CRDT / OT 工程量太大，v1.0 先 single-user |
| 雲端同步 / 帳號系統 | 增加後端負擔，先靠 `.gd` 檔案傳遞即可 |
| docx / pptx 的進階版面（複雜表格、SmartArt、動畫） | Univer 支援度有限，v1.0 開能讀能存即可，render fidelity 先求堪用 |
| Excel 完整公式引擎（VBA、複雜陣列函式） | v1.0 支援基本算術 / SUM / AVERAGE / IF / VLOOKUP 等常用 50 個 function |
| 行動裝置 | 桌面為先 |
| 插件系統 | v2.0 願景 |
| 自訂 AI provider 的 UI（API endpoint、headers） | v1.0 內建 Claude，用 env 變數 / 設定檔切；UI 提供 v1.5 |
| 圖片 / 影片內嵌的 AI 生成 | v1.0 文字操作為主 |

### 3.3 灰色地帶（明確標記）

- **基本表格在 docx 內**：能讀 + 不破壞，編輯只支援文字內容變更，不支援新增/刪除欄列（這留在 v1.5）。
- **pptx 的 master slide / theme**：能讀進來顯示正確，使用者不能在 v1.0 編 master，只能編個別 slide 的文字框與基礎形狀。
- **`.gd` 內嵌圖片**：支援，存在 `assets/images/`，但圖片編輯（裁切、濾鏡）不在 v1.0。

---

## 4. 核心使用者流程

至少五條，依使用頻率與重要性排序。

### 4.1 流程 A — 開新專案 → 加入多種 tab → 存成 `.gd`

1. 使用者按 `New Project`，App 建一個空的 in-memory workspace（manifest.json 雛型 + 0 tabs）
2. 使用者點 `+` 加 tab，dropdown 選 `Markdown / HTML / Word / Excel / PowerPoint`
3. 每個 tab 預設一個空白檔案，使用者各自編輯
4. 使用者按 `Cmd/Ctrl+S`，第一次存檔跳檔案對話框，選擇位置與檔名 → 寫入 `.gd`
5. 之後每次 `Cmd/Ctrl+S` 直接覆寫該 `.gd`
6. **Acceptance**：關閉 App 後重開該 `.gd`，所有 tab 的內容、順序、選取狀態完整還原

### 4.2 流程 B — 開啟既有 `.gd` → 編輯 → AI 修改 → 儲存

1. `File → Open` 選一個 `.gd`，App 解開到暫存目錄，讀 manifest.json，依序載入 tabs
2. 使用者點某個 tab（例：xlsx），選一個 cell range
3. 使用者在 AI panel 輸入「把這個 range 的數字格式改成貨幣 NTD」
4. AI 回應 + 發出 tool call `apply_excel_format`，產生 ChangeSet
5. 主編輯區出現 diff preview overlay（變更前 vs 變更後）
6. 使用者按 `Apply`，變更落地；或按 `Reject` 丟棄
7. `Cmd/Ctrl+Z` 可在 apply 後撤回
8. **Acceptance**：每一步都可中斷，使用者隨時知道下一個 action 是什麼

### 4.3 流程 C — 跨檔 AI 操作（Anna 的核心用例）

1. 使用者打開一份 `.gd`，內含 `report.docx` + `data.xlsx` + `slides.pptx`
2. AI panel 對話：「把 data.xlsx 的 Sheet1 Q3 sum 摘要進 report.docx 第二段，並在 slides.pptx 第 5 頁加 3 個 bullet 描述」
3. AI 透過 tool 序列：`read_excel_range` → `summarize` → `apply_word_paragraph_replace` → `apply_pptx_bullet_insert`
4. 三個 tab 各自顯示 diff preview，使用者可一次 apply all 或逐個 apply
5. **Acceptance**：跨檔操作的 ChangeSet 是 atomic 的，apply all 失敗會 rollback 已變更的部分

### 4.4 流程 D — Markdown 快寫 → 轉發 docx

1. 使用者開新專案，只加一個 md tab，純文字記筆記
2. 寫到一半要交給主管，AI panel：「把目前 md 的內容轉成正式 Word 報告格式，加上標題頁」
3. AI 用 `convert_md_to_docx` tool 在同一份 `.gd` 內新增一個 docx tab
4. 使用者檢視 docx，微調，存檔
5. **Acceptance**：md tab 內容不受影響，新增的 docx tab 自動被 AI 命名（例：`report-2026-04-27.docx`）

### 4.5 流程 E — 設定 API key 與模型切換

1. 第一次開啟 App，AI panel 顯示「請先設定 Anthropic API key」+ 連結到設定頁
2. 使用者貼 API key，按 `Test connection`，App 對 Claude 發一個 ping 訊息確認可用
3. 使用者選預設模型（Sonnet 4.6 / Opus 4.7）
4. 設定寫入 `~/.gendoc/config.json`，加密儲存（Electron safeStorage API）
5. **Acceptance**：API key 在設定檔以 OS keystore 加密形式存放，明文不落地

### 4.6 流程 F — Undo 多步 AI 編輯

1. 使用者連續對 AI 下三個指令，每個都 apply 了
2. `Cmd/Ctrl+Z` 三次，依序退回三步
3. 每步退回都顯示「Undo: [動作描述]」的 toast
4. **Acceptance**：手動編輯與 AI 編輯共用同一個 undo stack，順序正確

### 4.7 流程 G — 批次匯出多個頁籤到資料夾

1. 使用者點工具列的 `FileDown` 圖示按鈕（位於單檔 `Download` 按鈕右側），或從原生功能表「檔案 → 批次匯出多個頁籤…」（快捷鍵 `Ctrl/Cmd+Shift+E`）
2. 跳出 `BatchExportDialog`，列出 workspace 內所有頁籤，每列含 type icon + 檔名 + 格式 label + checkbox；預設全部勾選；空的 binary 頁籤（docx/xlsx/pptx byteLength=0）自動 disabled 並顯示「先在編輯器中輸入再匯出」tooltip
3. 使用者可逐個勾選或勾「全選」切換（含 indeterminate 狀態）
4. 按「選擇資料夾並匯出」→ App 內 dialog 收起 → OS 資料夾選擇器跳出
5. 使用者選資料夾後，每個勾選的頁籤以原生格式（.md / .html / .docx / .xlsx / .pptx）寫入該資料夾；檔名衝突（兩個同名頁籤、或目標資料夾已有同名檔）自動加上 `(2)`、`(3)` 後綴（Windows Explorer 慣例）
6. 結束顯示 toast：「已匯出 N 個檔案到資料夾」（全成功）／「匯出 X 個成功，Y 個失敗（檔名列表）」（部分成功）／「匯出失敗：…」（全失敗）
7. **Acceptance**：(a) 取消資料夾選擇器 → 無 toast、無檔案落地、dialog 已關；(b) 部分失敗時，成功的檔案仍保留在資料夾，使用者知道哪些檔需要重試；(c) 中途切換 workspace（Ctrl+O）不會把 toast 打到 NEW workspace 的 StatusBar；(d) `flushEditors()` 先沖刷 docx/pptx/xlsx 的 400ms 編輯 debounce，匯出 bytes 永遠是最新狀態

---

## 5. 功能規格

### 5.1 編輯器分頁系統

- **Tab bar** 位於主編輯區頂端，水平排列，可拖曳排序（react-dnd 或 dnd-kit）
- 每個 tab 顯示：
  - 格式 icon（md / html / docx / xlsx / pptx 各自圖示，HTML 用 `Code2` 玫紅色）
  - 檔名（雙擊重命名，重命名時驗證副檔名一致性）
  - 修改未存檔指示（檔名旁的小圓點）
  - 關閉按鈕（hover 時出現，未存檔時關閉前提示確認）
  - 單檔下載按鈕（hover 時出現，匯出該頁籤為原生格式檔）
- **Tab 切換**：點擊 tab 標頭，或 `Ctrl/Cmd+1..9` 快速切換到第 N 個 tab
- **Tab 限制**：v1.0 單一 `.gd` 內最多 32 個 tab（軟限制，超過警告但不阻擋）
- **狀態保留**：tab 切換時，每個 tab 的 scroll position、cursor position、selection 都保留在記憶體
- **新增 tab**：tab bar 最右端 `+`，dropdown 列五個格式選項（`Markdown / HTML / Word / Excel / PowerPoint`）；EmptyState 歡迎畫面也提供同樣的 QuickStart 卡片，HTML 卡片位於 Markdown 與 Word 之間

### 5.2 各格式編輯能力

#### 5.2.1 Markdown（CodeMirror 6）

| 能力 | v1.0 |
|------|------|
| 基本語法高亮（標題、粗體、清單、連結、code block） | 支援 |
| 即時預覽（split view / preview-only / source-only） | 支援，預設 split |
| GFM（表格、checkbox、刪除線） | 支援 |
| 數學公式（KaTeX） | 支援 |
| Mermaid 圖 | v1.0 不支援，v1.5 補 |
| 圖片內嵌（拖入 → 寫到 `assets/images/`） | 支援 |
| 自動儲存（debounce 2 秒寫入記憶體 buffer，存檔時序列化） | 支援 |
| Frontmatter（YAML） | 支援，會解析但不特別 render |

**MVP build 待優化清單**（市場對標 → 與目前 `MarkdownEditor.tsx` / `MarkdownToolbar.tsx` 的差距）：

對標對象：Typora、Obsidian、Mark Text、HackMD、Notion、iA Writer、VS Code Markdown。下表只列「目前還沒做」的差距，已實作的（toolbar 14 個按鈕、CM6 syntax highlight、行號、line wrap、undo / history、selection→AI context）不重複列。

| # | 能力 | 標竿做法 | Gen Doc 現況 | 優先序 |
|---|------|---------|--------------|--------|
| 1 | 即時預覽（split / preview-only） | Typora 直接 WYSIWYG；HackMD / VS Code 是右側 pane | **缺，與 spec 表寫的「支援，預設 split」不一致**（嚴重 spec drift） | P0 |
| 2 | 鍵盤快捷鍵（Ctrl+B / I / K / `1-6` heading） | 所有同類產品標配 | toolbar 按鈕有了，但 keymap 沒綁；現在只能滑鼠點 | P0 |
| 3 | Smart list 延續 | Enter 在 `- foo` 後自動加 `- `；空 bullet 兩次 Enter 退出 | 沒實作，使用者要自己重打 `- ` | P0 |
| 4 | Auto-pair（`(`、`[`、`"`、`*`、`` ` ``） | CM6 內建 `closeBrackets` extension | 沒接 | P0 |
| 5 | 字數 / 字元統計（status bar） | iA Writer / Bear / Typora 都顯示 | 沒有 | P0 |
| 6 | Slash command（`/` 喚出區塊插入選單） | Notion、HackMD、Obsidian | 沒有 | P1 |
| 7 | Outline / TOC 側欄（heading 跳轉） | Obsidian、Typora、VS Code | 沒有 | P1 |
| 8 | Markdown 預覽要 render KaTeX / Mermaid / GFM 表格 | Mark Text / Obsidian | 上表寫支援但連 preview 都沒，等於 0 | P1（伴隨 #1） |
| 9 | 剪貼簿圖片貼上 / 拖放 → 寫進 `assets/` | Typora / HackMD | 上表寫支援但實作為 0 | P1 |
| 10 | Find / Replace（Ctrl+F / Ctrl+H） | CM6 內建 `search` extension | 沒接 | P1 |
| 11 | Heading fold / 區塊摺疊 | Typora、Obsidian | CM6 有 `foldGutter` 沒接 | P2 |
| 12 | 高亮當前行 | 多數 IDE-style 編輯器 | CM6 有 `highlightActiveLine` 沒接 | P2 |
| 13 | Focus / Typewriter mode | Typora、iA Writer | 沒有 | P2 |
| 14 | Vim mode | Obsidian、VS Code | 沒有 | v1.5 |
| 15 | Frontmatter schema view | Obsidian Properties | 上表寫「不特別 render」— 維持，但 v1.5 可加 | v1.5 |

**Phase A 範圍（這次 build 直接做掉，全部 P0）**：

1. **Live preview 分割面板** — 用既有 `marked` 依賴 render；右側 50/50；可切換 source / split / preview-only 三模式（toolbar 加切換鈕）
2. **Keymap 綁定** — Ctrl+B / Ctrl+I / Ctrl+` / Ctrl+K（連結）/ Ctrl+1..3（heading）/ Ctrl+Shift+L（無序清單）
3. **Smart list continuation** — Enter 在 `- ` / `1. ` / `> ` / `[ ] ` 行尾自動延續；空項退出
4. **Auto-pair brackets** — 接 CM6 的 `closeBrackets` extension
5. **Status bar** — 字數（去除 markdown 標記後）、字元數、目前游標行/欄

剩下 P1 / P2 留下一輪迭代。

#### 5.2.2 HTML（CodeMirror 6 + sandboxed iframe）

定位：介於 Markdown 與 Word 之間的「結構化純文字」格式 — 內容可攜帶、可直接交給瀏覽器或郵件 client、可由 AI 結構化解析。`.gd` 容器內以 UTF-8 文字儲存，與 markdown 同枝。

| 能力 | v1.0 |
|------|------|
| HTML 語法高亮（tag / attribute / text） | 支援，@codemirror/lang-html |
| 三模式檢視 切換（原始碼 / 對照 / 預覽） | 支援，預設「對照」；選擇持久化至 `localStorage.gendoc.htmlViewMode` |
| 配對標籤亮點（`matchClosingTags`） | 支援 |
| 自動補齊收尾標籤（`autoCloseTags`：輸入 `</` 自動完成） | 支援 |
| 行號 + 自動換行 + Tab 縮排 | 支援 |
| Find / Replace | CM6 內建 `search` 面板（Ctrl+F） |
| 預覽 iframe 安全沙箱 | `sandbox=""` — 不執行 script、無 same-origin、無 top navigation、無 form submit；專為「寫內容、不是測 web app」設計 |
| 預覽自動更新 | 120 ms debounce，與 Markdown preview 同步調 |
| Export 為 `.html` 檔 | 支援，Ctrl+E / toolbar / per-tab Download 三條路徑皆可 |
| 拖入 `.html` / `.htm` 開新 tab | 支援，與其他外部檔同走 `openExternalFile` |
| 圖片內嵌（拖入 → data URI） | v1.0 暫不支援，使用者可手動 `<img src="data:image/png;base64,...">` |
| 鍵盤格式快捷鍵（Ctrl+B 包 `<strong>`） | v1.0 暫不支援，v1.5 補 |
| Outline / heading 跳轉側欄 | v1.0 暫不支援 |
| AI 寫入工具（html_replace_section / html_append） | v1.0 暫不支援；AI 可透過 `read_tab_content` 讀全文文字、提示使用者手動修改 |

**檢視模式 UX 細節**：

- 「原始碼」：左右切到單一 CodeMirror，預覽 iframe 仍掛載但 `display: none`（避免重新 parse 與閃爍）
- 「對照」：左右 50 / 50，CodeMirror + iframe 同時可見
- 「預覽」：iframe 滿版，CodeMirror 仍掛載但 `display: none`（保留 EditorView 與 cursor / undo history）
- 三模式切換是純 CSS 顯隱、不會 unmount / remount → CodeMirror 的 history 永遠保留；同步避開大 HTML 切換時的 iframe 重新建立 browsing context 成本

**AI outline 整合**：`[Active workspace]` 區塊中 HTML tab 顯示為 `[html] file.html (id=…)` 並用 regex 抽取 `<h1>…<h6>` 文字輸出為 `# / ## / …` markers（最多 12 個），AI 可比照 markdown 規劃跨 tab 操作。

#### 5.2.3 Word（Univer Doc）

| 能力 | v1.0 |
|------|------|
| 段落、標題（H1-H6） | 支援 |
| 粗體 / 斜體 / 底線 / 顏色 / 字體 | 支援 |
| 清單（ordered / unordered） | 支援 |
| 簡單表格（讀取保留，編輯只改 cell 文字） | 部分支援 |
| 圖片（讀取顯示、拖入新增） | 支援 |
| 頁首 / 頁尾 | 讀取保留，v1.0 不提供 UI 編輯 |
| 註腳 / 尾註 | 讀取保留，v1.0 不編輯 |
| 追蹤修訂 / 註解 | v1.0 不支援，v1.5 |
| 樣式表（Heading 1 樣式定義） | 讀取保留，使用者用既有樣式套用 |

**MVP build 待優化清單**（屬於上表「支援」範圍但目前實作仍簡化）：

- ~~粗體 / 斜體 / 底線目前以「整段」為單位套用~~ — Phase F-2 已完成 per-range 格式化（contentEditable + run-level 寫回）
- **圖片插入**（拖入或選檔）尚未實作；mammoth 讀取會略過圖片，docx 寫回也尚未組 `<w:drawing>`

**跨格式內容插入**：

- **「插入 Markdown 內容」**（toolbar `FileCode` 按鈕）— 開啟 dialog 貼 markdown 原始碼，按 Ctrl+Enter 後 [`markdownToDocxBlocks`](src/lib/markdown-to-docx.ts) 轉成 DocxBlock 序列（heading1-6 / paragraph / bullet / numbered / table），整批以單一 undo 入口插入 active block 之後（或文件末端）
- **「插入 HTML 內容」**（toolbar `Code2` 按鈕，2026-05-11 新增）— 同樣 Dialog UX，DOMParser 解析 HTML，[`htmlToDocxBlocks`](src/lib/html-to-docx.ts) 走 `<h1-6>` / `<p>` / `<ul><li>` / `<ol><li>` / `<table>` / `<pre>` / `<blockquote>` / `<img alt>` / inline `<strong>` `<em>` `<u>` `<a>` `<code>` `<br>`；`<script>` / `<style>` / `<noscript>` / `<template>` / `<iframe>` / `<svg>` / `<canvas>` 列為 markup-only blacklist，文字內容不會落地（避免 CSS rules / JS code 被當段落）；parser 失敗時錯誤訊息 inline 顯示在 Dialog 內、不關 Dialog 讓使用者修

#### 5.2.4 Excel（Univer Sheet）

| 能力 | v1.0 |
|------|------|
| Cell 文字 / 數字 / 日期輸入 | 支援 |
| 基本公式（50 個常用 function 清單見附錄 A.1） | 支援 |
| Cell 格式（數字格式、貨幣、日期、百分比） | 支援 |
| Cell 樣式（顏色、邊框、字體） | 支援 |
| 多 Sheet | 支援 |
| 凍結欄列 | 支援 |
| 圖表（柱狀、折線、圓餅） | 支援基本三種 |
| 樞紐分析表 / Power Query | v1.0 不支援 |
| VBA / Macro | v1.0 不支援，會在開檔時警告並忽略 |
| 條件格式 | v1.0 簡單版本（>、<、區間） |

**MVP build 待優化清單**：

- ~~**插入 / 刪除欄列的 UI** 尚未上線~~ — 已實作 (XlsxEditor toolbar 上 / 下 / 左 / 右插入 + 刪除)
- ~~**合併儲存格**（`merge cells`）尚未實作~~ — Phase G 已加入 (toolbar Merge / Unmerge 按鈕，read/write `!merges`，insert/delete row/col 時自動位移)
- ~~**數值格式 UI**（千分位、貨幣、百分比、日期）尚未實作~~ — 已實作（toolbar 11 種預設 numFmt 下拉，含整數 / 千分位 / 貨幣 / 百分比 / 日期）
- **xlsx-js-style round-trip**：font / alignment / fontColor 寫入 `styles.xml` 正確，重新開啟時 `.s` 解析回填只能還原 fill；toolbar 的 active 狀態在 reload 後可能不全（讀回路徑需深入解析 cellXfs）

#### 5.2.5 PowerPoint（Univer Slide）

| 能力 | v1.0 |
|------|------|
| Slide 新增 / 刪除 / 排序 | 支援 |
| 文字框、清單 | 支援 |
| 基本形狀（矩形、圓、箭頭） | 支援 |
| 圖片插入 | 支援 |
| 主題（簡單套色） | 支援 v1.0 內建 5 套主題 |
| Master slide 編輯 | v1.0 唯讀 |
| 動畫 / 轉場 | v1.0 讀取保留但不編輯，匯出 pptx 時保留原始定義 |
| 簡報模式（F5 投影） | 支援基本播放，無 presenter view |
| Speaker notes | 支援編輯 |

**MVP build 待優化清單**：

- 編輯介面是 textarea 清單（每個 `<a:r>` 一個 row），**缺 WYSIWYG slide canvas** — 使用者看不到投影片真實版面、圖片位置、形狀層級；保真度只在重新打開原檔（PowerPoint / Keynote）時才看得到
- 新增投影片目前是「複製當前」而非從版面挑空白佈局
- **圖片插入** 尚未實作（需新增 `ppt/media/imageN.*` + slide rels + `<p:pic>` shape）
- **基本形狀 / 箭頭** 尚未支援（spec 列為支援，MVP 實作為 0）

#### 5.2.6 市場對標：Office / Google / iWork vs Gen Doc（Word / Excel / PowerPoint）

對標對象：**Microsoft 365**（Word / Excel / PowerPoint）、**Google Workspace**（Docs / Sheets / Slides）、**Apple iWork**（Pages / Numbers / Keynote）。下表只列「v1.0 應該有但目前還沒做」的差距，已實作的項目（5.2.2 / 5.2.3 / 5.2.4 上方表格的「支援」格）不重複列。

跨產品共通而 Gen Doc 缺的能力另外彙整在表 D。

##### 表 A — Word 編輯器差距

| # | 能力 | MS Word | Google Docs | Pages | Gen Doc 現況 | 優先序 |
|---|------|---------|-------------|-------|--------------|--------|
| W1 | 鍵盤快捷鍵（Ctrl+B / I / U） | ✓ | ✓ | ✓ | toolbar 有，鍵盤未綁 | **P0** |
| W2 | Inline（per-range）格式化 | ✓ | ✓ | ✓ | 目前以「整段」為單位（spec 已標註） | **P0** |
| W3 | Hyperlink 插入（Ctrl+K） | ✓ | ✓ | ✓ | 沒有 | **P0** |
| W4 | Find / Replace（Ctrl+F / Ctrl+H） | ✓ | ✓ | ✓ | 沒有 | **P0** |
| W5 | 行距 / 段距控制 | ✓ | ✓ | ✓ | 沒有 | P1 |
| W6 | 縮排（增 / 減；Tab / Shift+Tab） | ✓ | ✓ | ✓ | 沒有 | P1 |
| W7 | 圖片拖入 / 貼上插入 | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | P1 |
| W8 | 多層次清單（無序內含有序、巢狀縮排） | ✓ | ✓ | ✓ | 平面清單 | P1 |
| W9 | Outline / Document Outline 側欄 | ✓ | ✓ | ✓ | 沒有 | P1 |
| W10 | Right-click context menu | ✓ | ✓ | ✓ | 沒有 | P1 |
| W11 | Format painter（複製格式） | ✓ | ✓ | ✓ | 沒有 | P2 |
| W12 | Spell check / 拼字檢查 | ✓ | ✓ | ✓ | 沒有（瀏覽器 native 可能有） | P2 |
| W13 | 註腳 / 尾註 | ✓ | ✓ | ✓ | spec 標 v1.0 不編輯 | P2 |
| W14 | 註解 / 批註 / 追蹤修訂 | ✓ | ✓ | ✓ | spec 標 v1.5 | v1.5 |
| W15 | 即時協作（Presence） | △（O365） | ✓ | △（iCloud） | spec 標 v2 | OUT |
| W16 | Print preview / Export PDF | ✓ | ✓ | ✓ | 沒有 | P2 |
| W17 | 樣式面板（apply / save Style） | ✓ | ✓ | ✓ | 沒有 | v1.5 |
| W18 | TOC 自動產生 | ✓ | ✓ | ✓ | 沒有 | v1.5 |

##### 表 B — Excel 編輯器差距

| # | 能力 | MS Excel | Google Sheets | Numbers | Gen Doc 現況 | 優先序 |
|---|------|----------|---------------|---------|--------------|--------|
| E1 | 鍵盤快捷鍵（Ctrl+B / I / U） | ✓ | ✓ | ✓ | toolbar 有，鍵盤未綁 | **P0** |
| E2 | Range 選取（拖曳 / Shift+方向鍵） | ✓ | ✓ | ✓ | 只有單格選取 | **P0** |
| E3 | 狀態列彙總（Sum / Avg / Count） | ✓ | ✓ | ✓ | 沒有 | **P0** |
| E4 | Find / Replace | ✓ | ✓ | ✓ | 沒有 | **P0** |
| E5 | 方向鍵 / Tab / Enter 移動游標 | ✓ | ✓ | ✓ | 沒有，使用者要點 | **P0** |
| E6 | 範圍 Copy / Paste（保留格式） | ✓ | ✓ | ✓ | 只能單格貼 | P1 |
| E7 | 公式自動完成（鍵入 `=SU` 提示） | ✓ | ✓ | ✓ | 沒有 | P1 |
| E8 | 合併儲存格 | ✓ | ✓ | ✓ | 沒有（spec 已標註） | P1 |
| E9 | AutoFill（拖右下角填充） | ✓ | ✓ | ✓ | 沒有 | P1 |
| E10 | 凍結窗格 | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | P1 |
| E11 | 條件格式化 | ✓ | ✓ | ✓ | spec 標簡單版本，實作為 0 | P1 |
| E12 | 篩選 / 排序 | ✓ | ✓ | ✓ | 沒有 | P1 |
| E13 | 圖表（柱 / 折 / 圓） | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | P1 |
| E14 | 數值格式 UI（千分位 / 貨幣 / %） | ✓ | ✓ | ✓ | UI 為 0（adapter 可寫） | P1 |
| E15 | 公式列（formula bar） | ✓ | ✓ | ✓ | **已有** | — |
| E16 | Right-click context menu | ✓ | ✓ | ✓ | 沒有 | P1 |
| E17 | 樞紐分析表 | ✓ | ✓ | ✓ | spec 標 OUT | OUT |
| E18 | 命名範圍 | ✓ | ✓ | ✓ | 沒有 | v1.5 |
| E19 | 資料驗證 | ✓ | ✓ | ✓ | 沒有 | v1.5 |
| E20 | Print preview / Export PDF | ✓ | ✓ | ✓ | 沒有 | P2 |

##### 表 C — PowerPoint 編輯器差距

| # | 能力 | MS PowerPoint | Google Slides | Keynote | Gen Doc 現況 | 優先序 |
|---|------|--------------|---------------|---------|--------------|--------|
| P1 | 鍵盤快捷鍵（Ctrl+B / I / U） | ✓ | ✓ | ✓ | toolbar 有，鍵盤未綁 | **P0** |
| P2 | 形狀 resize（8 個 handle） | ✓ | ✓ | ✓ | 只能拖移、不能縮放 | **P0** |
| P3 | 對齊輔助線（snap / smart guides） | ✓ | ✓ | ✓ | 沒有 | **P0** |
| P4 | 插入基本形狀（矩形 / 圓 / 線 / 箭頭） | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | **P0** |
| P5 | Z-order（Bring forward / Send back） | ✓ | ✓ | ✓ | 沒有 | P1 |
| P6 | Group / Ungroup | ✓ | ✓ | ✓ | 沒有 | P1 |
| P7 | 旋轉 handle | ✓ | ✓ | ✓ | 沒有 | P1 |
| P8 | 對齊指令（Align left / center / distribute） | ✓ | ✓ | ✓ | 沒有 | P1 |
| P9 | Outline 模式（左側純文字大綱） | ✓ | ✓ | △ | 沒有 | P1 |
| P10 | Speaker notes（演講者備忘） | ✓ | ✓ | ✓ | spec 寫支援、實作為 0（presenter notes 解析有，UI 缺） | P1 |
| P11 | Find / Replace 跨投影片 | ✓ | ✓ | ✓ | 沒有 | P1 |
| P12 | Slide 縮圖排序（拖曳重排） | ✓ | ✓ | ✓ | 已有縮圖、未支援拖曳排序 | P1 |
| P13 | 圖片插入（拖入 / 貼上） | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | P1 |
| P14 | 主題 / 樣式套用 | ✓ | ✓ | ✓ | spec 寫支援、實作為 0 | P1 |
| P15 | 動畫 / 轉場編輯 | ✓ | ✓ | ✓ | spec 標 OUT，僅讀取保留 | OUT |
| P16 | Presenter view（演講者畫面） | ✓ | ✓ | ✓ | spec 標 OUT | OUT |
| P17 | Right-click context menu | ✓ | ✓ | ✓ | 沒有 | P1 |

##### 表 D — 跨編輯器共通缺口

| # | 能力 | 三家標竿都有 | Gen Doc 現況 | 優先序 |
|---|------|--------------|--------------|--------|
| X1 | 結構性編輯的 App-level Undo / Redo | ✓ | 只有 Markdown（CM6 內建）；Word / Excel / PPT 對結構性操作沒有 undo | **P0** |
| X2 | Find / Replace 全文搜尋 | ✓ | 全部沒有 | **P0** |
| X3 | 鍵盤快捷鍵全套 | ✓ | toolbar 都有按鈕但鍵盤多半沒綁 | **P0** |
| X4 | Right-click context menu | ✓ | 全部沒有 | P1 |
| X5 | Print preview / Export PDF | ✓ | 全部沒有 | P2 |
| X6 | 即時協作 | ✓ | spec 標 v2 | OUT |
| X7 | 雲端同步 | ✓ | spec 標 v2 | OUT |
| X8 | Format painter | ✓ | 全部沒有 | P2 |
| X9 | 模板 / Templates 開新檔 | ✓ | 沒有（first-run 給 sample.gd） | P2 |

##### Phase B 範圍（這次 build 直接做掉）

聚焦在「P0、CP/effort 高」的三項，能在一輪迭代收掉：

1. **格式快捷鍵**（W1 / E1 / P1）— ✅ 完成。新增 [src/lib/use-format-shortcuts.ts](../src/lib/use-format-shortcuts.ts)，三個編輯器各加 `data-{kind}-editor-root` scope 屬性，呼叫 hook 時注入既有 `toggleStyle` / `updateStyle`。Word 支援 B/I/U；Excel 支援 B/I/U；PPT 因 `PptxRunStyle` 暫未建模 underline，僅支援 B/I（用 hook `keys` 限制）。
2. **PowerPoint 形狀 resize handles**（P2）— ✅ 完成。`ShapeFrame` 加上 8-handle（4 corners + 4 edges），各帶獨立 cursor（nwse / nesw / ew / ns），pointerdown 後計算每個 handle 對應的 `dx/dy/dcx/dcy` 分量；min size 0.2"；commit 時 clamp 到 slide 邊界後呼叫新增的 `handleResizeShape` → `moveShapeOnSlide(b, idx, shp, {x,y,cx,cy})`。adapter 不變。
3. **Excel range 選取 + 狀態列彙總**（E2 / E3）— ✅ 完成。`Selection` 升級為 `{ r, c, r2, c2 }`（anchor + opposite corner），加 helpers `rangeOf` / `isInRange` / `rangeAddr` / `rangeStats`。Grid 用 `mousedown` + `mouseenter`（搭配 `draggingRef`）做 drag-select；range 內 cell 疊一層 `bg-primary/15` overlay；anchor 維持 outline。`updateStyle` 從單格升級為 range 全套用，所以選 A1:C3 後按 Ctrl+B 整片變粗體。新增 `<StatusBar>` 顯示 範圍位址 / 加總 / 平均 / 計數 / 儲存格數。

剩下的 P0（W2 inline 格式、W3 hyperlink、W4/E4/P11 Find&Replace、E5 鍵盤導航、X1 App undo、P3 snap guides、P4 插入形狀、PPT underline 建模）排入 **Phase C**。

##### Phase C 範圍（這次 build 直接做掉）

延續 Phase B 的「P0、CP/effort 高」過濾，這次處理三項：

1. **Excel 鍵盤導航**（E5）— ✅ 完成。`Grid` 內加 `tableRef` + `focusCell(r, c)` 透過 `data-cell-r/c` querySelector 跳格。CellInput 的 keydown 加 `handleNav`：Enter / Shift+Enter 上下、Tab / Shift+Tab 左右、ArrowUp/Down 永遠跳、ArrowLeft/Right 只在 caret 在 input 邊界時跳（caret 在中間時走原生游標移動）。Tab 用 preventDefault 阻擋瀏覽器原生 focus traversal。
2. **PPT 底線**（P1 補完）— ✅ 完成。`PptxRunStyle` 新增 `underline?: boolean`；adapter 的 `parseRPr` / `mergeRPr` / `buildRPr` 都加上 `u="sng"` 處理（讀任何 non-`none` 值都收為 true）。Toolbar 加 Underline 按鈕，編輯 + 簡報模式的 inline style 加 `textDecoration`。`useFormatShortcuts` 拿掉 `keys: ['bold','italic']` 限制，Ctrl+U 直接生效。
3. **Find & Replace**（W4 / E4 / P11）— ✅ 完成。新增 [src/components/FindReplaceDialog.tsx](../src/components/FindReplaceDialog.tsx)：通用 dialog 吃 `SearchSegment[]`（id + text + label）+ `onUpdateSegment(id, newText)` callback。三個編輯器各自 flatten 自己的 model：Word→blocks 的文字段（暫不含 table cells）、Excel→active sheet 全 cells（label 顯示 A1:N 樣式）、PPT→active slide 的 runs。Ctrl/Cmd+F 開關 dialog，scoped 在各自的 `data-{kind}-editor-root`。支援大小寫 toggle、上一個 / 下一個、單筆 / 全部取代；全部取代時對同 segment 的多個 match 由右往左 splice 避免 offset 失效。Excel / PPT scope 限定 active sheet / slide，避免「按一下取代了 47 個跨表結果」的驚喜，跨範圍搜尋排 Phase D。

剩下的 P0（W2 inline 格式、W3 hyperlink、X1 App undo、P3 snap guides、P4 插入形狀、跨 sheet/slide F&R、Word table cell F&R）排入 **Phase D**。

##### Phase D 範圍（這次 build 直接做掉）

繼續清剩下的 P0，這次處理三項 + 一項架構決策：

1. **跨範圍 F&R（W4 / E4 / P11 補完）+ Word table cell F&R** — ✅ 完成。Excel 的 `findSegments` 改成走全部 sheets，segment id 升級為 `${sheetIdx}:${r}:${c}`；`applyFindReplace` 解析 sheet 索引，跨表寫入時切 `activeSheetIdx` 並 inline mutate target sheet model。PPT 的 `findSegments` 走全部 slides，id 用 `${slideIndex}:${runId}`，因 runId 本身可能含 `-`，所以新增 `splitFindId` 只對第一個 `:` 切。Word 的 `findSegments` 加上 `block.kind === 'table'` 分支，segment id 用 `${blockId}:${r}:${c}`；`applyFindReplace` 偵測三段式 id 後重建該列的 cells 陣列再 `updateBlock`。
2. **PPT 插入形狀**（P4）— ✅ 完成。adapter 加 `addShapeToSlide(bytes, slideIdx, kind, text?)`，`PptxShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightArrow'`，預設填色 `#4472C4` + 白字置中（`anchor="ctr"` + `algn="ctr"`），有 placeholder run 確保 parser 看得見。PptxEditor 在 `LayoutPicker` 旁加 `<ShapePicker>` dropdown，列出五個形狀。
3. **PPT snap guides**（P3）— ✅ 完成。`SlideCanvas` 把所有非當前 group 的 frame 當 sibling 傳給 `<ShapeFrame>`，並追蹤 `guides` state 渲染 dotted alignment line（`bg-fuchsia-500/80`）。`ShapeFrame` 預先 build x/y target 集合（每個 sibling 的 left/center/right + slide 的 0/中心/邊）；drag 時對自己三條邊（left/center/right、top/middle/bottom）找最近 target，threshold 6 px → EMU；resize 時根據 handle 決定哪條邊在動（'e'/'se'/'ne' 動右、'w'/'nw'/'sw' 動左、'n'/'s' 同理），snap 後把 diff 同步加到 dx/dcx 保持對側錨定。每次 pointermove 都 emit 當下 guides，pointerup 清空。

**X1 App-level undo / redo 延後到 Phase E（架構決策）**。三個 binary editor（Docx / Xlsx / Pptx）目前都各自維護 local model state（`useState<DocxModel>` 等）跟 `workspace.tabs[].data`（Uint8Array bytes）並行，編輯時用 `commit*` debounce serialize 回 bytes。要做 App-level undo 需要把 model 提升到 workspace（讓 workspace.undoSnapshot 攔截），但這樣每次 undo / redo 都要強制 editors 重新解析 bytes（或反向）。較合理的後續路線是**讓 editor 自己管 undo stack**（每個 editor 內維護一個 ring buffer 存 model snapshot，Ctrl+Z 走 in-memory revert 而非 re-parse），把 X1 從 App 層問題降為 editor 層問題。本次 build 不動，先把這個結論記在這裡，避免下次直接捅進 workspace。

剩下未做的 P0 / 高 CP 項：W2 inline 格式（Word run-level B/I/U）、W3 hyperlink、X1 editor-local undo。排入 **Phase E**。

##### Phase E 範圍（這次 build 直接做掉）

挑 X1 跟 W3 進場，W2 因為要重做 run-level 模型留到後面。

1. **X1 editor-local undo / redo** — ✅ 完成。新增 [src/lib/use-undoable-state.ts](../src/lib/use-undoable-state.ts)：`useUndoableState<T>(initial)` 是 `useState` 的 drop-in，內部用兩個 ref 維護 `past` / `future`，`set` 時走 500 ms coalesce window（同一 window 內的 setState 只 push 一次 prev 到 past，避免每個 keystroke 一個 undo step）。`canUndo` / `canRedo` 透過 tick state 觸發 re-render；`resetHistory(next)` 清空兩 stack，用在初次 parse 或切 tab。`useUndoShortcuts({ rootSelector, undo, redo })` 跟既有 `useFormatShortcuts` 同 pattern——只在 `document.activeElement` 在指定 root 內才觸發，capture phase + stopPropagation 避免被其他 keymap 攔截，所以 markdown 編輯器（CM6 內建 undo）跟其他 tab 不會受影響。三個 binary editor（Docx / Xlsx / Pptx）都把 `useState<Model>` 換成 `useUndoableState`：初次 parse 走 `resetHistory`，所有編輯走 `setModel`，並各自加 `useEffect([model])` 在 model swap 後 push bytes（這樣 undo / redo 也會自動 flush 到 disk）。XlsxEditor 順便把原本 inline 的 `writeBack(next)` 都拿掉，統一由 effect 處理避免重複 serialize。
2. **W3 Word hyperlink** — ✅ 完成。`DocxBlock` 加 `link?: string`（block-level external hyperlink，整個 block.text 變成單一連結；per-character link 留到有 run 模型再做）。adapter `makeBlockFromElement` 偵測「整個 block 文字都包在 `<a>` 內」時 lift href 到 `block.link`；`buildParagraph` 在有 link 時把 TextRun 包進 `ExternalHyperlink({ link, children: [run] })`，並強制 underline + color `#0563C1`（Word 標準連結藍）覆蓋使用者設定的 color。DocxEditor toolbar 在 Underline 旁加 Link 按鈕，按下走 `window.prompt`（current value 預填，留白清除），自動加 `https://` 前綴給沒帶 scheme 的輸入。BlockRow 渲染時 link block 強制藍 + 底線，textarea `title` 顯示 URL。

剩餘 P0：**W2 inline 格式**（Word run-level B/I/U）。要做這個必須把 `DocxBlock` 從單 text + block-level style 升級到 `runs[]`，動到 parser、serializer、editor UI、F&R 全部，留 **Phase F**。

##### Phase F-1 範圍（這次 build 直接做掉）

W2 不做完整的 selection-based inline 編輯（那要把 textarea 換成 contentEditable，cursor / selection / IME 等都要重新處理，不適合塞在這次掃尾），只先做**round-trip 保真**：原本 docx 裡有 inline 粗體 / 斜體 / 底線的段落，open → save 後仍然保留。實際編輯該段時才會收合（同既有失真行為）。Selection-based inline editing 排在 **Phase F-2**。

1. **DocxRun 模型** — `DocxRun = { text; style?: { bold?, italic?, underline? } }`，`DocxBlock` 加 `runs?: DocxRun[]`。`block.text` 保留為扁平化的純文字（給 textarea 跟 F&R 讀），`runs[]` 是真正的 source of truth 當它存在。
2. **Parser** — `extractRuns(el)` 走 `createTreeWalker(el, NodeFilter.SHOW_TEXT)`，每個 text node 沿 parent 鏈往上收 `<b/strong>` / `<em/i>` / `<u>`，相鄰同 style 的 run 合併（避免一個粗體字產生 5 個 `{bold:true}` run）。`makeBlockFromElement` 只在「runs 數量 > 1 或 任一 run 有 style」時才 attach `runs`，避免每個普通段落都掛一個冗餘陣列。
3. **Serializer** — `buildParagraph` 抽出 `blockShared`（color / fontSize / fontFamily 走 block-level）跟 inline B/I/U（走 run-level，跟 block.style fallback OR 一起算），有 `runs[]` 時 emit 一個 TextRun per run，否則維持單 TextRun 的舊行為。Hyperlink 包裝 `ExternalHyperlink` 也改成接受多個 child run，跟 ExternalHyperlink 的 children API 一致。
4. **Edit 行為** — `DocxEditor.updateBlock` 偵測 patch 裡有 `text` 時，主動 strip `runs`：textarea 編輯路徑只會生純文字，留著舊 runs 會 desync。其他 patch（kind / align / style / color 等）保留 runs。
5. **可視提示** — `BlockRow` 在 `runs.length > 1` 時，textarea 上面疊一個 read-only preview 用 `<span>` 渲染每個 run 的 B/I/U，並把底下 textarea 的 text 設成 `text-transparent caret-foreground`（caret 還在，glyphs 不重疊）。tooltip 寫「本段含 inline 格式…，編輯時會收合為單一格式」這樣使用者編輯前就知道後果。

剩餘工作（**Phase F-2**）：textarea → contentEditable，selection-based B/I/U toolbar action，inline link 也順帶可以做 partial-text。需要單獨一個 session 處理 cursor / IME / 跟現有 useUndoableState 的協作。

##### Phase F-2 範圍（已完成 — selection-based inline B/I/U）

把 textarea 換成 contentEditable，讓 Ctrl/Cmd+B / I / U 真的對「選取範圍」套用，而不是整段。block-level style 仍保留作為 caret-only / 表格儲存格的 fallback。color / fontSize / fontFamily 仍是 block-level（mammoth 沒給 per-run color，做進來會打亂 round-trip）。

1. **`src/lib/rich-text.ts`** — 一組與 contentEditable DOM 對接的 helper：
   - `runsToHtml` / `domToRuns`：runs ↔ HTML（`<strong>` / `<em>` / `<u>`）。`domToRuns` 用 TreeWalker 走 SHOW_TEXT，沿 parent 鏈收 B/I/U，最後 `coalesceRuns` 合併相鄰同 style；ZWSP（用來在空 styled span 內 host caret 用）會被剝掉，模型不會殘留。
   - `getCharRange` / `setCharRange`：以「字元位置」（不是 DOM offset）保存 / 還原 selection。為什麼用字元而不是 (Node, offset)：rewrite innerHTML 之後 Node 被重建，DOM offset 失效；字元 index 跟 `runsToText` 對齊就好還原。
   - `applyStyleToRange(runs, start, end, key, value?)`：先 `splitRunsAt` 在 boundary 切兩刀，找出範圍內第一個 overlap run 的 key 是否已 set 來決定 toggle 方向（這是 Word / Google Docs 的行為——對部分粗體選取按 Ctrl+B 會把整段都變粗），最後對重疊 run set/clear 該 key 並 coalesce。
2. **`src/components/RichBlock.tsx`** — 一個 contentEditable 段落元件，replace BlockRow 裡的 textarea：
   - 用 `useLayoutEffect` 同步 model → DOM：把當前 `runs` serialize 一個短 key（`B_U|text\u0001…`），跟 ref 中 `lastWrittenKey` 比；不一樣才 rewrite innerHTML。一樣（代表變更來自自家 `oninput`）就跳過，避免每次 keystroke 都 stomp caret。rewrite 前後存 / 還原 char range。
   - IME guard：`composingRef` 在 `compositionstart` set true，`compositionend` set false 並 fire change。期間不 fire `onChange`、layout effect 也不 rewrite，否則 CJK 候選字會被 stomp 掉。
   - `dangerouslySetInnerHTML` 不用——直接讓 layout effect 在 mount 後第一次 render 寫 innerHTML（empty key 一定 mismatch），避免 React 跟我們搶子節點。
   - `onPaste` 強制 plain-text（preventDefault + insert text node），擋掉外部 source 的 inline color / font / span 殘渣。
   - placeholder 用 `[data-rich-block][data-empty="1"]:not(:focus)::before { content: attr(data-placeholder); }`，永遠在 contentEditable subtree 之外（pseudo-element），使用者不可能誤打進去。
3. **`updateBlock` 對 runs 的處理** — patch 帶 `runs` 時，`block.text` 從 runs 重新 derive（`map(r => r.text).join('')`）保持 search index / F&R 看到的純文字一致；patch 帶 `text`（F&R 替換、表格儲存格）仍然 strip `runs` 因為單純的 text 不能還原 inline 樣式。
4. **`toggleStyle` 的 inline 模式** — 有 RichBlock 取得 focus 且 selection range 非空時，用 `applyStyleToRange` 對 runs splice 後 `updateBlock(blockId, { runs })`；selection 為空 / focus 不在 RichBlock（如表格）時 fall back 到既有的 block-level toggle。re-render 後 RichBlock layout effect 重寫 innerHTML 再 setCharRange 還原同一 char range，使用者體感是「選取保持、樣式切換」。
5. **既有 read-only 預覽 overlay 移除** — Phase F-1 那個 `whitespace-pre-wrap break-words` 的疊層 + `text-transparent caret-foreground` hack 拿掉，因為 RichBlock 直接 render 真實格式不需要再疊。Yellow banner 改寫成「B/I/U 對選取生效；對齊 / 字色 / 字型仍整段」。

至此 W2 P0 收完，Word 的 inline 編輯能力從「整段切換」升級到「真正的選取套用」，跟商業 Word / Google Docs 行為對齊。

##### Phase G 範圍（已完成 — UI 細節打磨）

Phase F-2 收完之後，再花一輪掃過界面找「使用者實際會踩到」但還沒做的小坑。每一條都是低風險、可獨立 ship 的局部改動，不動資料模型。原則：跟商業 Office / VS Code 對齊使用者已建立的肌肉記憶，避免使用者「咦怎麼這個快捷鍵這裡沒有」。

1. **儲存狀態指示器** — 既有的 save 是 fire-and-forget，使用者按 Ctrl+S 之後得自己賭它有沒有寫進去。改成 `saveState: 'idle' | 'saving' | 'success' | 'error'` state machine：toolbar 的 Save 按鈕 icon 切 spinner / check / alert，status bar 多一行「X 秒前儲存」（每 15 秒重算的相對時間）；error 時 alert + 紅色 icon + tooltip 帶錯誤訊息。也加了 `savingRef` reentrancy guard，防止使用者連按 Ctrl+S 在慢碟上 race。
2. **Find & Replace 增強（DocxEditor / XlsxEditor / PptxEditor 共用 dialog）** — 加上 regex 與 whole-word 兩個 toggle、replace 完成後 flash 「已取代 N 個」/「找不到符合項目」訊息（auto-clear 2.5s）、無效 regex 顯示紅字錯誤行不讓 dialog 整個炸掉。replaceOne / replaceAll 在 regex mode 走 `String.replace(re, repl)` 讓 `$1` 等 capture group reference 可用，跟 VS Code 對齊。
3. **Tab 鍵盤導航** — Ctrl+Tab / Ctrl+Shift+Tab 在 tabs 之間循環、Ctrl+W 關閉當前 tab。原本 Ctrl+1–9 已有，但對 9 個以上的 .gd 沒辦法，且 Tab cycle 是大家肌肉記憶。實作放在 App.tsx 的 window keydown listener，preventDefault 避免被 textarea 吃掉。
4. **Tab 右鍵選單** — 右鍵 tab → 重新命名 / 關閉 / 關閉其他 / 關閉右側 / 全部關閉。bulk close 對 dirty tab 一次性 confirm（顯示「有 N 個未儲存」），不一個個彈窗。menu 用 fixed 定位，window-level mousedown / Esc 關閉，本身 stopPropagation 避免馬上自關。
5. **Markdown F&R via CM6 search panel** — 之前只有 Docx/Xlsx/Pptx 有 F&R，Markdown 沒有。加入 `@codemirror/search` 套件，`search({ top: true })` extension + searchKeymap 進 EditorState；toolbar 也加一個 Search button 呼叫 `openSearchPanel`，跟內建 Ctrl+F 同入口。免費拿到 regex / case / whole-word 三個 toggle，UX 跟 VS Code 同源。
6. **Excel 合併儲存格** — `XlsxSheet.merges: MergeRange[]`，read 從 `ws['!merges']` 還原，write 寫回同欄位（xlsx-js-style 直接 round-trip）。adapter 提供 `mergeRange` / `unmergeAt` / `isMergeCovered` / `mergeAtAnchor`，editor toolbar 加 Merge / Unmerge 兩個按鈕（依選取範圍動態啟用）。Grid 在 render 時對 covered cell 直接 `return null`，anchor cell 帶 `rowSpan` / `colSpan`，視覺上跟 Excel 一致。`insertRowAt` / `deleteRowAt` / `insertColAt` / `deleteColAt` 都會同步位移 / 收縮 / 丟棄受影響的 merge，不會遺留 stale 矩形。點擊 merge 內任一格會自動把 selection 擴張到整個 merge 矩形，所以 toolbar 的 Unmerge / 統計顯示都對。
7. **Excel 公式引擎** — 在 `src/lib/xlsx-formula.ts` 自寫 tokenizer + recursive-descent parser + AST evaluator，function dispatch 接到 `@formulajs/formulajs`（MIT 授權，附錄 A.1 全部 50 個 function 直接 cover）。HyperFormula 以 GPL-3.0-only 在商業關門產品不適用，因此不採用。
   - 文法：cell ref（`A1` / `$A$1` / `$A1` / `A$1`）、跨表 ref（`Sheet1!A1` / `'Sheet With Spaces'!A1`）、range（`A1:B5` / `Sheet1!A1:B5`）、運算子 `+ - * / ^ & = <> < > <= >=` + unary minus + 百分比後綴；字面值含數字（容許 thousands separator）、`"string"`（`""` 跳脫）、`TRUE` / `FALSE` 與 7 種 error literal。
   - 錯誤碼：`#DIV/0!` / `#VALUE!` / `#REF!` / `#NAME?` / `#N/A` / `#NUM!` / `#CYCLE!` / `#ERROR!`，全部走 `class FormulaError extends Error` propagate。
   - Function dispatch 用 `RANGE_FLATTEN_FNS`（SUM/AVERAGE/SUMIF/COUNTIFS 等）/ `RANGE_2D_FNS`（VLOOKUP/HLOOKUP/INDEX/MATCH/XLOOKUP）/ default scalar with implicit-intersection 三套策略，配合 SUMIF / COUNTIF / SUMIFS / COUNTIFS 的 (range, criterion) 序列特殊處理。`normalizeResult` 把 formulajs 各種混合回傳（number / string / Error / Array spill）統一收斂成 `CellValue`。
   - **Recompute strategy**：`recomputeAllFormulas` 走 8-pass fixpoint，每個 pass 重置 cache、重算所有有 `formula` 欄位的儲存格、若沒任何 cell 變動就 break。沒做 dependency DAG（MVP 表小，O(formulas × passes) 跑得動），用 `inProgress: Set<string>` 偵測 cycle 並 emit `#CYCLE!`。函式做成 generic over `T extends FormulaSheet` 讓 XlsxSheet 的 merges / style 等額外欄位 round-trip 不被吃掉。
   - **Adapter round-trip**：`XlsxCell.formula?: string`（含開頭 `=`），read 從 SheetJS `cell.f` 還原（補上 `=`），write 寫回 `cell.f`（脫掉 `=`，符合 SheetJS 約定）；空字串 / undefined 對應 plain literal 路徑。
   - **Editor 整合**：`commitCell` 偵測到輸入以 `=` 開頭時 set `formula = text`，否則 strip formula 欄位；patch 完馬上跑 `recomputeAllFormulas` 讓 dependents 一起更新。`FormulaBar` 與 `CellInput` 在 focus 時切換到 source view（顯示 `=A1+B1`）、blur 時切回 computed value（顯示 `42`），並在 blur 與 source 比對避免假性 commit。`parseXlsx` 與 tab swap 也都先 recompute 一次，讓開檔當下顯示就是我們引擎算出來的值，不是 SheetJS 的快取。
   - 不在範圍：external workbook ref、structured table ref、`{...}` 陣列公式、defined names、動態陣列 spill（UNIQUE / SORT 等的多 cell spill 收斂到第一格）—— 留 v1.5。

### 5.3 `.gd` 檔案格式設計

`.gd` 是一個 zip 容器，副檔名換成 `.gd`，內部結構：

```
mybook.gd
├── manifest.json
├── doc/
│   ├── 01-notes.md
│   ├── 02-landing.html
│   ├── 03-report.docx
│   ├── 04-data.xlsx
│   └── 05-slides.pptx
├── assets/
│   └── images/
│       ├── img-001.png
│       └── img-002.jpg
└── meta/
    ├── chat-history.sqlite      ← 可選，使用者可在設定中關閉內嵌 chat history
    └── thumbnails/              ← tab 預覽圖快取（可重建）
```

#### 5.3.1 manifest.json schema（JSON Schema 風格）

```json
{
  "version": "1.0",
  "title": "Q3 客戶提案",
  "createdAt": "2026-04-27T10:00:00Z",
  "modifiedAt": "2026-04-27T15:23:00Z",
  "tabs": [
    {
      "id": "tab-uuid-1",
      "name": "備忘",
      "type": "markdown",
      "file": "doc/01-notes.md",
      "order": 0
    },
    {
      "id": "tab-uuid-2",
      "name": "報告",
      "type": "docx",
      "file": "doc/02-report.docx",
      "order": 1
    }
  ],
  "settings": {
    "embedChatHistory": true,
    "defaultModel": "claude-sonnet-4-6"
  }
}
```

- `version`：semver，v1.0 寫 `"1.0"`，未來升版時做向後相容
- `tabs[].id`：UUID v4，跨 session 穩定，AI tool call 用此 id 鎖定 target tab
- `tabs[].type`：enum `markdown` / `html` / `docx` / `xlsx` / `pptx`
- `tabs[].file`：相對 `.gd` 根的路徑

#### 5.3.2 開檔流程

1. 使用者選 `.gd` → JSZip 解到一個暫存目錄（`os.tmpdir()/gendoc-{uuid}/`）
2. 讀 `manifest.json`，依 `tabs[].order` 載入每個 tab
3. 各 tab 的 doc engine（Univer / CodeMirror）讀對應檔案，建立 in-memory model
4. 編輯時 model 變更不立即寫回暫存目錄；存檔時才序列化 → 重新打包 zip → 覆寫 `.gd`
5. App 關閉時清掉暫存目錄

#### 5.3.3 衝突處理

v1.0 single-user，沒有真正的衝突。但要處理：
- 開啟時偵測檔案 lock（同一個 `.gd` 在另一個 App instance 已開）→ readonly 提示
- 暫存目錄殘留偵測（上次 crash 沒清掉）→ 啟動時掃描，超過 7 天的暫存目錄自動清除

### 5.4 AI Panel 行為規格

固定在主視窗右側，預設寬度 380px，可拖曳調整 280-600px。

#### 5.4.1 主要區域

- **Chat history**：scrollable 訊息列表，user / assistant / tool result 三種訊息類型
- **Selection context badge**：當使用者在主編輯區選了內容，AI panel 上方出現一個 badge：「Context: B2:D8 in data.xlsx (24 cells)」可移除
- **Input area**：多行文字框，`Enter` 送出，`Shift+Enter` 換行
- **Model picker**：下拉選 Sonnet 4.6 / Opus 4.7，預設 Sonnet
- **Tool toggle**：可暫時關閉 tool calling，只做 chat（用於使用者只想「問問」不想被改動）

#### 5.4.2 Selection-aware context 規則

- 主編輯區當前 selection 自動成為下一次 prompt 的 context
- Context 內容序列化規則：
  - **md**：raw markdown text（最多 4000 chars，超過截斷並提示）
  - **html**：raw HTML source（與 md 同 cap；HtmlEditor 目前未推送 selection 到 workspace.selection — AI 透過 outline + read_tab_content 取得內容）
  - **docx**：plain text + 段落結構（JSON 格式）
  - **xlsx**：cell range 的 2D array + sheet 名 + 範圍位址
  - **pptx**：當前 slide 的所有文字 + slide index
- 沒有 selection 時，預設 context = 「current tab 的整檔大綱」（不是整檔內容，避免爆 token）；HTML tab 的大綱由 buildOutline 用 regex 抽取 `<h1>…<h6>` 文字（與 markdown 的 ATX heading 抽取同字典輸出）

#### 5.4.3 Tool calling 流程

詳見 §6.2。

#### 5.4.4 Diff preview 設計

- AI 發出 tool call 後，App 不立即套用，先建立一個 `pendingChange` object
- 主編輯區 overlay 一個半透明的 diff layer：
  - md：左右分欄，舊版 vs 新版，行級 diff（紅刪綠加）
  - docx：在原文上 inline 顯示，紅色刪除線 / 綠色新增
  - xlsx：被影響的 cell 邊框變綠，hover 顯示 tooltip「舊值 → 新值」
  - pptx：受影響的 slide 縮圖出現「Pending」徽章，點擊進入該 slide 看 inline diff
- AI panel 對應訊息下方出現 `[Apply] [Reject] [Modify]` 三個按鈕
- `Apply`：套用 ChangeSet，加入 undo stack
- `Reject`：丟棄 pendingChange
- `Modify`：把 pendingChange 內容回填到 input box，使用者改後重送

#### 5.4.5 Undo / Redo 規則

- 統一 undo stack，限制 50 步（per-session，存在 SQLite）
- 每個 entry 紀錄：
  - 時間戳
  - 動作來源（manual / ai）
  - 影響的 tab id
  - inverse patch（用於 redo 回去）
- `Cmd/Ctrl+Z` undo、`Cmd/Ctrl+Shift+Z` redo
- 跨 tab 的 AI ChangeSet 視為一個 atomic entry，undo 會一次回退所有受影響 tab

### 5.5 設定與快捷鍵

#### 5.5.1 設定項

| 類別 | 項目 |
|------|------|
| AI | API key、預設模型、temperature、max tokens、prompt cache 開關 |
| 編輯 | 字體、字級、行距、theme（light / dark / system） |
| 檔案 | 自動儲存間隔、最近開啟列表大小 |
| 隱私 | 是否內嵌 chat history 到 `.gd`、telemetry 開關 |
| 進階 | 暫存目錄路徑、log level |

#### 5.5.2 預設快捷鍵（可在設定中覆寫）

| 動作 | Win/Linux | macOS |
|------|-----------|-------|
| 新專案 | Ctrl+N | Cmd+N |
| 開啟 | Ctrl+O | Cmd+O |
| 儲存 | Ctrl+S | Cmd+S |
| 另存新檔 | Ctrl+Shift+S | Cmd+Shift+S |
| 匯出目前頁籤 | Ctrl+E | Cmd+E |
| 批次匯出多個頁籤 | Ctrl+Shift+E | Cmd+Shift+E |
| 關閉 tab | Ctrl+W | Cmd+W |
| 切換 tab | Ctrl+1..9 | Cmd+1..9 |
| Undo | Ctrl+Z | Cmd+Z |
| Redo | Ctrl+Shift+Z | Cmd+Shift+Z |
| 開啟 AI panel focus | Ctrl+L | Cmd+L |
| Apply pending change | Ctrl+Enter | Cmd+Enter |
| Reject pending change | Ctrl+Backspace | Cmd+Backspace |

---

## 6. AI 整合規格

### 6.1 Provider 抽象

雖然 v1.0 預設 Claude，仍要把 provider 抽到一層介面，避免日後綁死。

```ts
// src/ai/provider.ts
export interface AIProvider {
  name: string;
  models: ModelDescriptor[];
  chat(req: ChatRequest): AsyncIterable<ChatChunk>;
  countTokens(text: string): number;
}

export interface ChatRequest {
  model: string;
  messages: Message[];
  tools?: ToolDefinition[];
  system?: string;
  cacheBreakpoints?: number[];   // index of messages to mark as cache_control
  maxTokens: number;
  temperature?: number;
}

export interface ChatChunk {
  type: 'text' | 'tool_use' | 'tool_use_input' | 'message_stop' | 'usage';
  data: unknown;
}
```

實作放在 `src/ai/providers/anthropic.ts`，未來擴充 `openai.ts`、`ollama.ts`。

### 6.2 Tool calling 流程

完整 loop：

```
[1] User prompt + selection context
       │
       ▼
[2] Build messages: system + history + user + tools schema
       │  (system + tools schema 標記 cache_control)
       ▼
[3] anthropic.messages.stream({...})
       │
       ▼
[4] Stream chunks → render text incrementally in panel
       │
       ▼ (when tool_use detected)
[5] Suspend stream, parse tool_use input
       │
       ▼
[6] Validate tool call against schema
       │
       ▼
[7] Build ChangeSet (do not apply yet)
       │
       ▼
[8] Render diff preview in main editor + Apply/Reject UI
       │
       ▼ (user clicks Apply)
[9] Apply ChangeSet to doc model
       │  push to undo stack
       │  serialize tool_result back to messages
       ▼
[10] Continue stream → AI may emit more tool_use or final text
       │
       ▼
[11] Stream end → idle
```

#### 6.2.1 Tool 清單（v1.0 共 18 個）

**Markdown：**
- `md_replace_section` (heading, newContent)
- `md_append` (text)
- `md_insert_at` (line, text)

**Word：**
- `word_replace_paragraph` (paragraphIndex, newText)
- `word_insert_paragraph` (afterIndex, text, style?)
- `word_apply_style` (range, styleName)
- `word_insert_heading` (afterIndex, level, text)

**Excel：**
- `excel_set_cell` (sheet, address, value)
- `excel_set_range` (sheet, range, values2D)
- `excel_apply_format` (sheet, range, formatString)
- `excel_insert_row` (sheet, rowIndex, count)
- `excel_insert_chart` (sheet, range, chartType)

**PowerPoint：**
- `pptx_replace_text` (slideIndex, shapeId, newText)
- `pptx_add_slide` (afterIndex, layout, content)
- `pptx_add_bullets` (slideIndex, shapeId, bullets[])

**跨檔：**
- `read_tab_content` (tabId, options)
- `convert_md_to_docx` (sourceTabId, destTabName)
- `cross_tab_summarize` (sourceTabIds[], destTabId, instructions)

每個 tool 的完整 JSON schema 由 Agent 3 撰寫於 `src/ai/tools/`，本 spec 只規範語意。

### 6.3 對話 context 管理

#### 6.3.1 訊息結構

```ts
type Message =
  | { role: 'user'; content: ContentBlock[] }
  | { role: 'assistant'; content: ContentBlock[] }
  | { role: 'tool_result'; tool_use_id: string; content: ContentBlock[] };

type ContentBlock =
  | { type: 'text'; text: string }
  | { type: 'image'; source: ImageSource }
  | { type: 'tool_use'; id: string; name: string; input: object };
```

#### 6.3.2 Cache 策略

依照 Anthropic prompt caching 設計：

1. **System prompt + tools schema** 永遠標 `cache_control: ephemeral`，命中率最高
2. **首段歷史**（前 N 則 message，N 由 token budget 決定）標 cache breakpoint
3. **當前 selection context** 不快取（每次都不同）
4. **使用者 prompt** 不快取

預期命中率：第二次 turn 開始，input token 的 70-85% 可以走 cache hit，價格降到 1/10。

#### 6.3.3 Context 大小控制

- 單一 turn input token 上限：80k（Sonnet 4.6 為 200k context，留餘裕給 tool result）
- 超過時自動摘要：把最舊的 10 則 message 用一個小 prompt 摘要成 1 則 system note，繼續對話
- 摘要動作對使用者透明，但 chat panel 顯示一個小灰色標籤「Earlier conversation summarized」

#### 6.3.4 跨檔 context

當使用者問題涉及多 tab，AI 不會自動把所有 tab 內容塞進 context。流程是：

1. 第一次對話只給 manifest 列表 + 各 tab 的「outline」（md 標題、xlsx sheet 名、docx 大綱、pptx slide 標題）
2. AI 自己決定要 read 哪幾個 tab，發出 `read_tab_content` tool call
3. App 回 tool_result，把那些 tab 的具體內容塞進 messages
4. 之後 AI 才開始改檔

這套設計避免使用者每次提問都付整本 `.gd` 的 token 費用。

### 6.4 Token 成本估算

以下為一個典型 working session（Anna 改一份提案，30 分鐘 / 20 個 turn）的估算。

| 項目 | 數值 |
|------|------|
| 平均 input token / turn (with cache) | 4k(uncached) + 18k(cached @ 0.1x) |
| 平均 output token / turn | 800 |
| Sonnet 4.6 定價 | $3 / MTok input, $15 / MTok output, cached read $0.30 / MTok |
| 單 turn 成本 | $3 × 4/1000 + $0.30 × 18/1000 + $15 × 0.8/1000 = $0.012 + $0.0054 + $0.012 = $0.0294 |
| 20 turn session 成本 | ≈ $0.59 |
| 月度（10 session/週、4 週） | ≈ $24 |

對 Anna 這種付費意願高的使用者，月成本可接受。我們會在設定頁顯示 token usage 統計，讓使用者有透明度。

Opus 4.7 用於 heavy thinking（例如「全文重寫」「跨檔大改」），單 turn 成本約 4-5 倍，使用者主動切才會用。

---

## 7. 資料模型

### 7.1 manifest.json schema（完整版）

```json
{
  "$schema": "https://gendoc.app/schemas/manifest-v1.0.json",
  "version": "1.0",
  "title": "string",
  "createdAt": "ISO-8601 timestamp",
  "modifiedAt": "ISO-8601 timestamp",
  "tabs": [
    {
      "id": "uuid v4",
      "name": "display name",
      "type": "markdown | html | docx | xlsx | pptx",
      "file": "doc/relative-path",
      "order": "integer >= 0"
    }
  ],
  "settings": {
    "embedChatHistory": "boolean",
    "defaultModel": "string"
  },
  "metadata": {
    "appVersion": "string",
    "lastEditedBy": "string (optional, for v2.0 multi-user prep)"
  }
}
```

### 7.2 ChangeSet 物件結構

ChangeSet 是 AI 與 doc engine 之間的中介層。

```ts
interface ChangeSet {
  id: string;                    // uuid
  origin: 'ai' | 'manual';
  createdAt: string;             // iso 8601
  description: string;           // human-readable, 給 undo toast 用
  ops: ChangeOp[];
}

type ChangeOp =
  | { tabId: string; type: 'md_text'; before: string; after: string; range?: [number, number] }
  | { tabId: string; type: 'word_paragraph'; paraIndex: number; before: ParagraphData; after: ParagraphData }
  | { tabId: string; type: 'excel_cell'; sheet: string; address: string; before: CellValue; after: CellValue }
  | { tabId: string; type: 'pptx_text'; slideIndex: number; shapeId: string; before: string; after: string }
  | { tabId: string; type: 'tab_create'; tab: TabDescriptor; data: Uint8Array }
  | { tabId: string; type: 'tab_delete'; tab: TabDescriptor; data: Uint8Array };
```

每個 op 都帶 `before` 與 `after`，用於 undo 時直接套 inverse。`tab_create` 的 inverse 是 `tab_delete`（資料已存在 ChangeSet 內），反之亦然。

### 7.3 Chat history 儲存（SQLite）

用 `better-sqlite3`，DB 檔放在：
- session-only：`%TEMP%/gendoc-{instance}/chat.sqlite`
- 內嵌入 `.gd`：使用者啟用後，存檔時複製到 `.gd:meta/chat-history.sqlite`

#### 7.3.1 Schema

```sql
CREATE TABLE conversations (
  id TEXT PRIMARY KEY,           -- uuid
  title TEXT,                    -- AI 摘要的簡短標題
  created_at INTEGER NOT NULL,   -- unix epoch ms
  updated_at INTEGER NOT NULL,
  workspace_id TEXT              -- 對應某個 .gd 的 manifest.title hash
);

CREATE TABLE messages (
  id TEXT PRIMARY KEY,
  conversation_id TEXT NOT NULL REFERENCES conversations(id),
  role TEXT NOT NULL,            -- user | assistant | tool_result
  content TEXT NOT NULL,         -- JSON-serialized ContentBlock[]
  tool_use_id TEXT,              -- nullable, 對應 tool_result
  created_at INTEGER NOT NULL,
  token_input INTEGER,           -- 該 turn 的統計（assistant 訊息上）
  token_output INTEGER,
  cache_read INTEGER,
  cache_creation INTEGER
);

CREATE INDEX idx_messages_conv ON messages(conversation_id, created_at);

CREATE TABLE undo_entries (
  id TEXT PRIMARY KEY,
  changeset_json TEXT NOT NULL,  -- 序列化的 ChangeSet
  applied_at INTEGER NOT NULL,
  workspace_id TEXT
);
```

---

## 8. 技術架構圖

```
┌──────────────────────────────────────────────────────────────────────┐
│                          Electron Main Process                       │
│                                                                      │
│  ┌────────────────┐  ┌────────────────┐  ┌────────────────────────┐  │
│  │  File I/O      │  │  Crypto /      │  │  IPC bridge            │  │
│  │  - .gd zip rw  │  │  safeStorage   │  │  (contextBridge to     │  │
│  │  - JSZip       │  │  - API key     │  │   Renderer)            │  │
│  └────────────────┘  └────────────────┘  └────────────────────────┘  │
│           │                  │                       │               │
│           └──────────────────┴───────────────────────┘               │
└────────────────────────────────────────┬─────────────────────────────┘
                                         │
                            ┌────────────▼────────────┐
                            │   IPC channel (typed)   │
                            └────────────┬────────────┘
                                         │
┌────────────────────────────────────────▼─────────────────────────────┐
│                        Electron Renderer Process                     │
│                          (React 18 + Vite)                           │
│                                                                      │
│  ┌────────────────────────────┐    ┌──────────────────────────────┐  │
│  │   Tab Bar / Workspace UI   │    │      AI Panel (right)        │  │
│  │   (shadcn Tabs + dnd-kit)  │    │  - Chat history list         │  │
│  └─────────────┬──────────────┘    │  - Input + model picker      │  │
│                │                   │  - Selection context badge   │  │
│  ┌─────────────▼──────────────┐    │  - Diff preview controls     │  │
│  │    Editor surface          │    └────────────┬─────────────────┘  │
│  │  ┌────────┐ ┌────────────┐ │                 │                    │
│  │  │ MD tab │ │ Univer tab │ │    ┌────────────▼─────────────────┐  │
│  │  │ (CM6)  │ │ (doc/sheet/│ │    │   AI engine (renderer-side)  │  │
│  │  │        │ │  slide)    │ │    │  - Provider abstraction      │  │
│  │  └────────┘ └────────────┘ │    │  - Tool dispatcher           │  │
│  └─────────────┬──────────────┘    │  - ChangeSet builder         │  │
│                │                   └────────────┬─────────────────┘  │
│                │                                │                    │
│  ┌─────────────▼──────────────┐    ┌────────────▼─────────────────┐  │
│  │     Doc model store        │◄───┤   ChangeSet apply / undo     │  │
│  │  (Zustand, per-tab state)  │    │   (transactional)            │  │
│  └─────────────┬──────────────┘    └──────────────────────────────┘  │
│                │                                                     │
│  ┌─────────────▼──────────────┐                                      │
│  │   SQLite (better-sqlite3)  │                                      │
│  │  - chat history            │                                      │
│  │  - undo stack              │                                      │
│  └────────────────────────────┘                                      │
└──────────────────────────────────────────────────────────────────────┘

External: Anthropic API (HTTPS, streaming)
                ▲
                │
                └── via AI engine in Renderer
```

關鍵說明：
- **AI engine 放在 Renderer 而非 Main**：streaming UI 直接更新方便，API key 從 Main 透過 IPC 取出後只在記憶體
- **File I/O 必須在 Main**：Renderer 沒有檔案系統權限，所有 zip / fs 操作走 IPC
- **doc model 走 Zustand**：每個 tab 一個 slice，AI 與 UI 共用同一個 state

---

## 9. 里程碑時程（11 週 MVP）

依 Phase 拆解，每個 phase 結尾要有 demo。

### Phase 1（Week 1-2）— 骨架與 Markdown tab

- 目標：Electron + React + Vite + Tailwind + shadcn 起骨架
- 完成 tab 系統 UI（純樣式，無功能）
- CodeMirror 6 整合，md tab 可開可存單檔（先不打包 .gd）
- 設定頁雛型（API key 欄位）
- **Demo 1**：能開一個空 App，新增一個 md tab，編輯，存成單一 .md 檔

### Phase 2（Week 3-4）— Univer 整合 + .gd 容器

- 目標：docx / xlsx / pptx 三種 tab 透過 Univer 整合
- JSZip 串起 `.gd` 打包與解包邏輯
- manifest.json 讀寫
- Tab 重命名 / 拖曳 / 關閉
- **Demo 2**：能建立含 4 種 tab 的 .gd，存檔、重開、所有內容還原

### Phase 3（Week 5-7）— AI Panel 與 chat 基礎

- 目標：AI panel UI、Anthropic SDK 整合、streaming chat（不含 tool calling）
- Provider 抽象、cache 策略
- Selection context 自動帶入
- SQLite chat history
- **Demo 3**：使用者可在 AI panel chat 問問題，AI 看得懂當前 selection 但還不能改文件

### Phase 4（Week 8-10）— Tool calling 與 ChangeSet

- 目標：完成 18 個 tool 的 schema 與 dispatcher
- ChangeSet 建構與 apply 邏輯
- Diff preview UI（四種格式各一套）
- Undo stack
- **Demo 4**：完整流程跑通，使用者可以「把這段改成正式語氣」AI 真的改，可以 undo

### Phase 5（Week 11）— Polish 與 packaging

- 目標：bug fix、效能優化、Electron Builder 打 Windows / macOS 安裝包
- 設定頁完整、快捷鍵、telemetry opt-in
- 內部 dogfooding 5-10 人
- **Release 1.0**：對外發佈，附 release notes

### 9.1 風險緩衝

- 每個 phase 內留 20% 浮動時間給未預期的 issue
- 若 phase 3 延遲，phase 4 的 tool 數量先砍到 10 個（保留四種格式各 2-3 個核心 tool）
- Univer 的 docx / pptx render fidelity 是最大不確定性，phase 2 第一週做 spike 評估

---

## 10. 風險與緩解

### 10.1 Univer 的 docx/pptx fidelity 不夠

- **風險**：Univer 在 doc / slide 領域比 sheet 年輕，複雜 docx（多欄位、複雜表格、SmartArt）可能 render 走樣
- **緩解**：phase 2 第一週做 fidelity spike，挑 10 份典型 docx / pptx 跑測試，量化「可接受度」；不可接受時降低 v1.0 的編輯範圍至「文字段落 + 簡單格式」，render fidelity 只保證「能開不破壞」
- **Owner**：Agent 2

### 10.2 AI tool calling 改錯文件 / 改錯位置

- **風險**：tool input 中的 paragraphIndex / cellAddress 容易因為文件變動而失效，AI 可能改到錯誤位置
- **緩解**：每個 tool call 在 dispatcher 層做 pre-flight 驗證（target 還存在嗎？範圍合法嗎？），失敗則回 tool_error 讓 AI 重試；diff preview 是最後一道防線，使用者一定看得到才 apply
- **Owner**：Agent 3 + Agent 4

### 10.3 Token 費用爆炸

- **風險**：使用者塞超大檔案進 context 或反覆對話，月度費用失控，使用者抱怨
- **緩解**：(1) 設定頁顯示 token usage 統計與費用估算 (2) cross-tab 操作走 outline-first，AI 自己決定要 read 哪幾個 tab (3) 單 turn input 超過 80k 自動摘要 (4) 設定可開「警告：本月已用 $X」上限
- **Owner**：Agent 4

### 10.4 .gd 檔案損壞

- **風險**：存檔過程 crash，zip 損壞使用者整本內容遺失
- **緩解**：寫檔走 atomic write（先寫 `.gd.tmp`，flush，rename）；自動備份最近 3 版到 `~/.gendoc/backups/{title}-{timestamp}.gd`；File menu 提供「Restore from backup」
- **Owner**：Agent 2

### 10.5 Electron 安全（API key 外洩）

- **風險**：API key 若以明文存 `config.json`，使用者主機被滲透即外洩
- **緩解**：用 Electron `safeStorage` API，背後是 OS keystore（Windows DPAPI / macOS Keychain）；contextIsolation: true、sandbox: true、disable nodeIntegration in renderer；對外請求只能走 IPC 中繼
- **Owner**：Agent 2

### 10.6 跨平台路徑與編碼

- **風險**：Windows / macOS 路徑分隔符、檔名編碼（中文檔名）、行尾差異讓 zip 內檔案在跨平台時錯位
- **緩解**：JSZip 統一用 forward slash 與 UTF-8；md / json 一律 LF 行尾；CI 多平台跑 e2e（Phase 5）
- **Owner**：Agent 2

### 10.7 使用者學習曲線

- **風險**：四格式並列 + AI panel 對非技術使用者太重，使用者打開兩次就放棄
- **緩解**：First-run onboarding（3 頁 walkthrough）+ 預設一個 sample.gd 含每種 tab 的範例 + AI panel 預設帶 5 個常用 prompt 範例（dropdown 可選）
- **Owner**：Agent 2 + 設計

---

## 11. 成功指標

v1.0 發佈後 90 天內量測，每兩週 review。

### 11.1 採用指標

- **Activation rate**：下載並開啟過至少一次 `.gd`，且使用 AI panel 至少 1 次的比例 ≥ 40%
- **D7 retention**：第 1 天活躍使用者中，第 7 天仍開啟 App 的比例 ≥ 25%
- **Average session length**：> 12 分鐘（短於這個代表使用者沒進到真實工作）

### 11.2 品質指標

- **AI apply 率**：使用者按 Apply 對 Reject 的比例 ≥ 70%（衡量 AI 改得準不準）
- **Crash rate**：每千 session crash 次數 < 3
- **Save success rate**：> 99.95%（檔案損壞是底線）
- **Cold start time**：< 2.5 秒（Electron baseline）
- **Open .gd time**：10MB 以下檔案 < 3 秒進入可編輯狀態

### 11.3 經濟指標（內部）

- 平均月度 token 成本 / 活躍使用者：< $30
- Cache hit rate：input token 的 cache read 占比 > 60%

### 11.4 質性訊號

- 內部 dogfooding 5 位 PM/RD 在 Phase 5 結束時的 NPS ≥ 30
- 對外 closed beta 30 位使用者於發佈後 30 天的 NPS ≥ 20

---

## 12. 未來路線圖

**v1.5（發佈後 3-6 個月）**：補完 docx 進階版面（複雜表格編輯、追蹤修訂、註解）、Markdown 加 Mermaid，Excel 補樞紐分析表入門版，AI 增加 image 生成 / 編輯（透過 Claude vision + 外掛圖像 provider），plugin SDK alpha，行動裝置 read-only viewer。

**v2.0（12 個月後）**：多人即時協作（CRDT 為基底，先支援同一份 .gd 的同步，雲端與 self-hosted 兩種 deploy）、自訂 AI provider 的 UI 完整化、開放 plugin marketplace、行動裝置 full editor、企業 SSO 與審計 log。長期願景是讓 Gen Doc 成為「文件世界的 VS Code」：一個可擴充的、AI-native 的文書工作平台。

---

## 附錄 A — 細節清單

### A.1 Excel v1.0 支援的 50 個函式

數學：SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND, ROUNDUP, ROUNDDOWN, ABS, INT, MOD, POWER, SQRT, RAND, RANDBETWEEN

邏輯：IF, AND, OR, NOT, IFERROR, IFS, SWITCH

文字：CONCATENATE, LEFT, RIGHT, MID, LEN, LOWER, UPPER, TRIM, FIND, REPLACE, SUBSTITUTE, TEXT

日期：TODAY, NOW, YEAR, MONTH, DAY, WEEKDAY, DATE, DATEDIF

查詢：VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP

統計：COUNTIF, SUMIF, AVERAGEIF, COUNTIFS, SUMIFS

### A.2 預設 5 套 PPT 主題

1. Minimal Light（白底黑字、無花俏配色，commercial 預設）
2. Minimal Dark
3. Corporate Blue（深藍主題色，標題深色 bar）
4. Warm Cream（米色背景，適合人文 / 學術）
5. Vibrant（多色 accent，適合產品發表）

### A.3 First-run 預設 sample.gd 內容

- `intro.md`：3 頁說明 Gen Doc 是什麼、五種 tab 怎麼用、AI panel 怎麼下指令
- `landing.html`：示範 HTML tab 三模式切換的最小頁面（含 inline `<style>`、`<h1>`、`<p>`、`<ul>`）
- `data.xlsx`：兩 sheet，各 50 列假資料（虛構 Q1-Q4 銷售）
- `report.docx`：對應上述資料的範例報告，3 頁
- `slides.pptx`：5 頁範例簡報，套 Minimal Light 主題
