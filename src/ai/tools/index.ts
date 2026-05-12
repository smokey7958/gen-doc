/**
 * The 18 tool definitions referenced in spec §6.2.1. Markdown + cross-file
 * tools are fully implemented; Word / Excel / PowerPoint write tools return
 * `not_implemented_in_mvp` so the AI knows to fall back gracefully.
 *
 * Schemas live next to the runtime so the dispatcher can validate input
 * with the same source of truth the model sees.
 *
 * R372 — uniform `tabId` description across every tool that exposes it.
 * The dispatcher's `findTab` helper (dispatcher.ts:44-48) used by every
 * `expect*Tab` path applies the same「omit → use active tab」 fallback
 * rule, but the schema previously documented this only on
 * md_replace_section + read_tab_content — the other 14 tabId fields were
 * bare `{ type: 'string' }`. AI's plan-time reaction varied: some tools
 * 「我得先 find active tab id and pass it explicitly」 (defensive but
 * wasteful), some「會自動 fallback 嗎？試試看」 (one round-trip on
 * uncertainty). Stating the same fallback verbatim on every tool
 * eliminates the per-tool inference cost and matches the「same
 * convention across all tools」 invariant findTab actually enforces.
 */

import type { ToolDefinition } from '../../types/ai';

export const TOOLS: ToolDefinition[] = [
  // ── Markdown ────────────────────────────────────────────────────────
  {
    name: 'md_replace_section',
    description:
      '替換 Markdown 文件中指定標題（heading）所屬章節的內容。包含該標題本身與後續所有內容直到下一個同級或更高層級標題前。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        heading: { type: 'string', description: '要替換的章節標題（不含 # 前綴；ATX 風格結尾的 # 標記會自動處理，「## 標題 ##」跟「## 標題」視為等價）' },
        newContent: {
          type: 'string',
          // R374 — describe the empty-newContent semantic. replaceMdSection
          // (dispatcher.ts:1239-1261) builds the result as
          // `[head, middle, tail].filter(s => s.length > 0).join('\n')` —
          // an empty `middle` (when newContent is "") gets filtered out,
          // leaving `head + tail` joined. That's the「刪除整段章節（含
          // 標題）」 semantic. Without this hint AI typing「請刪除『xxx』
          // 這節」 might pass a placeholder string (「(已刪除)」) or
          // ask the user to manually clean up — both worse than the
          // direct empty-string delete the dispatcher already supports.
          // Same description-aligns-with-implementation pattern as R363
          // (cross_tab_summarize instructions purpose) / R365 / R367.
          description: '新的章節內容（請包含原本的標題行，例如 "## 標題\\n..."）。傳空字串 "" 表示「整個刪除這個章節（含標題行）」。',
        },
      },
      required: ['heading', 'newContent'],
    },
  },
  {
    name: 'md_append',
    description: '將文字附加到 Markdown 文件結尾。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        // R380 — align `text` description with R331's actual empty-doc
        // behaviour. Original「會自動在前面加入空行分隔」 was true for
        // non-empty docs but over-promised for empty docs — R331 added a
        // length===0 short-circuit so an empty markdown tab (0-byte .md,
        // Ctrl+A+Delete on default boilerplate, freshly cleared dest tab
        // for cross_tab_summarize) appends `text` directly with no `\n\n`
        // prefix, avoiding the「兩行空白開頭」 visual bug R331 fixed.
        // AI reading the schema previously couldn't predict the empty-
        // doc case without testing; describing it here makes the
        // behaviour discoverable plan-time and matches the R331 fix
        // family's「dispatcher behavior surfaces in description」
        // pattern (R365 / R369 / R374 for sibling clamp / delete /
        // empty-special-case docs).
        text: { type: 'string', description: '要附加的內容；非空文件會自動在前面加入空行分隔；空文件直接從 text 開頭（不會有前綴空行）' },
      },
      required: ['text'],
    },
  },
  {
    name: 'md_insert_at',
    // R365 — describe the silent-clamp behavior. dispatcher.ts:289 does
    // `Math.max(0, Math.min(line - 1, lines.length))` — values > doc length
    // clamp to end-of-doc (append semantics), not a range error. Without
    // this hint, AI passing `line: 999` on a 5-line doc would expect an
    // error and not know to use `md_append` for explicit append-to-end.
    // Mirrors the R353 / R364 pattern: surface limits / fallbacks in the
    // description so AI's plan-time decision matches dispatcher's runtime
    // behavior. The clamp itself is good UX (forgiving for off-by-one AI
    // estimates of doc length); just needs to be documented.
    description:
      '在 Markdown 指定行號（1-based）之前插入文字。line 超過總行數會 clamp 到文件結尾（等同 md_append 的 append 語意），不會回報 out-of-range。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        line: { type: 'integer', description: '1-based 行號；超過文件總行數會 clamp 到結尾' },
        // R378 — add description. md_insert_at uses `lines.splice(idx, 0, text)`
        // where `text` is inserted as ONE array element. Multi-line content
        // (containing `\n`) ends up as a single element that, when joined
        // back with `\n`, produces correct multi-line markdown — AI can
        // pass `"## 標題\n內文段落"` to insert two lines at once. Without
        // this hint AI might assume single-line-only and split into
        // multiple md_insert_at calls, wasting round-trips.
        text: { type: 'string', description: '要插入的文字內容；可含 \\n 換行字元，會被當成多行整體插入到指定行之前' },
      },
      required: ['line', 'text'],
    },
  },

  // ── Word（spec stub）────────────────────────────────────────────────
  {
    name: 'word_replace_paragraph',
    // R364 — description mentions the R359 block-kind guard upfront so AI
    // knows the tool is text-only and table / image blocks need a
    // different recovery path. Before R364, the description claimed
    // 「替換文字」 with no caveat; AI calling on a table block got the
    // typed `block_kind_mismatch` error (R359) and had to discover the
    // limitation via trial-and-error. With the hint here, AI can pre-
    // emptively recommend「圖片 / 表格目前需請使用者在 Word 編輯器內手動
    // 修改」 to the user without burning a round-trip. Same plan-time-
    // vs-runtime alignment pattern as R357 (NOT_IMPLEMENTED messages
    // embed redirects from the description) and R353 (excel_set_cell
    // formula / null / finite hints).
    description:
      '替換 docx 中第 N 段（0-based）的文字。MVP 限制：只改文字，inline 樣式（粗體 / 斜體）不會保留；無法套用於 table / image 區塊（會回 block_kind_mismatch，這兩種區塊目前需請使用者在 Word 編輯器內手動修改）。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        paragraphIndex: { type: 'integer', description: '0-based 段落 index（必須指向文字型區塊：paragraph / heading1-6 / bullet / numbered）' },
        text: { type: 'string', description: '替換後的純文字' },
      },
      required: ['paragraphIndex', 'text'],
    },
  },
  {
    name: 'word_insert_paragraph',
    // R355 — describe heading-level coverage explicitly. The kind enum was
    // previously capped at heading3 even though docx-adapter.ts:39-46 declares
    // `DocxBlockKind` includes heading1–heading6 and the serializer's switch
    // at docx-adapter.ts:545-563 maps every level to docx's `HeadingLevel.
    // HEADING_N`. The schema cap therefore hid h4–h6 from AI for no model-
    // side reason — AI could only emit「請插入第 4 層標題」by failing the
    // word_insert_heading stub (which lists levels 1–6 but always returns
    // NOT_IMPLEMENTED). Aligning the enum surfaces what's already wired.
    //
    // R369 — also describe the silent-clamp behavior for paragraphIndex.
    // dispatcher.ts:927 does `Math.max(0, Math.min(paragraphIndex, model.
    // blocks.length))` — values >= blocks.length clamp to「insert at
    // end」 (equivalent to omitting paragraphIndex). Without this hint AI
    // passing `paragraphIndex: 999` on a 5-block doc would expect a
    // range error and not know to use the「append」semantic of omitting
    // the field. Same R365-shape gap for md_insert_at; siblings now
    // aligned: insert tools document the lax clamp、replace tools (R364
    // word_replace_paragraph) document the strict reject.
    description: '在 docx 指定位置插入新段落（kind 可選一般段落、heading1-6、bullet、numbered）。paragraphIndex 超過總段落數會 clamp 到結尾（等同省略 paragraphIndex 的 append 語意），不會回報 out-of-range。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        paragraphIndex: {
          type: 'integer',
          description: '插入位置（0-based，新段落會出現在此 index）；省略則 append；超過總段落數會 clamp 到結尾',
        },
        // R378 — add description. word_insert_paragraph inserts ONE new
        // paragraph block. Unlike md_insert_at, `\n` inside text doesn't
        // create multiple paragraphs — docx's TextRun stays single-run
        // and `\n` either renders as a literal space or gets dropped
        // depending on Word's interpretation. To insert MULTIPLE
        // paragraphs at once, AI should call word_insert_paragraph
        // multiple times with successive paragraphIndex values.
        // (R316's same-tab-pending guard prevents this within one turn,
        // so AI needs to do it across turns — user Applies each, AI
        // continues.)
        text: { type: 'string', description: '新段落的純文字內容；單一 paragraph block。要插入多段請分多次呼叫（內含 \\n 不會自動拆段）' },
        kind: {
          type: 'string',
          enum: [
            'paragraph',
            'heading1',
            'heading2',
            'heading3',
            'heading4',
            'heading5',
            'heading6',
            'bullet',
            'numbered',
          ],
          // R366 — describe what each kind produces in the output docx.
          // Without this hint AI just sees the enum values; their meaning
          // is mostly obvious (paragraph / heading1-6) but bullet vs.
          // numbered is now a meaningful distinction post-R360 (numbered
          // produces real「1. 2. 3.」 markers via the document's numbering
          // config; pre-R360 both fell back to bullet output). Naming the
          // marker style in the description makes the difference
          // discoverable without AI having to test-emit-and-inspect.
          description: '段落樣式：paragraph（一般段落）/ heading1-6（六級標題）/ bullet（• 項目符號清單）/ numbered（1. 2. 3. 編號清單）；省略則用 paragraph。',
        },
      },
      required: ['text'],
    },
  },
  {
    name: 'word_apply_style',
    description: '套用指定段落樣式。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        range: {
          type: 'array',
          items: { type: 'integer' },
          minItems: 2,
          maxItems: 2,
        },
        styleName: { type: 'string' },
      },
      required: ['range', 'styleName'],
    },
  },
  {
    name: 'word_insert_heading',
    // R355 — point AI at the working sibling. word_insert_heading itself is
    // a NOT_IMPLEMENTED stub, but R355's enum expansion on
    // word_insert_paragraph means `word_insert_paragraph({ paragraphIndex:
    // afterIndex + 1, text, kind: 'headingN' })` covers EVERY case this
    // tool was meant to handle. Without this redirect, AI looking at the
    // schema sees「v1.0 stub」 and either gives up on inserting headings
    // or tries the stub and gets NOT_IMPLEMENTED back — same R354-shape
    // pessimism (description claims a capability is missing when a
    // working sibling tool already covers it). Keep the tool present so
    // its NOT_IMPLEMENTED return path stays a wire-compatible signal,
    // but treat the description as a doc-only redirect.
    description: 'v1.0 stub — 已被 word_insert_paragraph 取代：請改呼叫 word_insert_paragraph({ paragraphIndex: afterIndex + 1, text, kind: "heading1" | … | "heading6" })，效果完全等價且支援全 6 個 heading level。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        afterIndex: { type: 'integer' },
        level: { type: 'integer', minimum: 1, maximum: 6 },
        text: { type: 'string' },
      },
      required: ['afterIndex', 'level', 'text'],
    },
  },

  // ── Excel（spec stub）───────────────────────────────────────────────
  {
    name: 'excel_set_cell',
    // R353 — description aligned with actual dispatcher behaviour after the
    // R297 / R303 / R352 sweep. Original「文字 / 數字 / 布林皆可（傳 string）」
    // was misleading: it suggested AI must coerce everything to string before
    // sending, but R297 made the dispatcher natively accept the four-type
    // union AND R303 added formula support (string starting with `=` is
    // recognized as a formula). The (傳 string) parenthetical was a holdover
    // from the early MVP when xlsx-adapter's inferCell re-typed every cell
    // from text content; today the cell carries the native type when AI
    // passes the right primitive, which avoids a string→type round-trip
    // and any ambiguity in special cases (e.g. value `"TRUE"` string vs.
    // `true` boolean produces the same byte output but explicit primitive
    // is clearer intent).
    //
    // Three additions versus the old description:
    //   • formula support: leading-`=` strings become formula cells
    //   • null support: idiomatic way to clear a cell
    //   • R352 hint: numeric values must be finite (NaN/Infinity rejected)
    description: '寫入單一 cell 的值。value 接受 string / number / boolean / null：number 必須為有限值（NaN / Infinity 會被擋）；字串以「=」開頭視為公式（如 `=SUM(A1:A10)`）；null 等同清空儲存格。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        sheet: { type: 'string', description: '工作表名稱；省略則用第一個 sheet' },
        address: { type: 'string', description: 'A1 notation, e.g. B7' },
        value: { description: '新值；string / number / boolean / null（number 須有限；字串以「=」開頭視為公式）' },
      },
      required: ['address', 'value'],
    },
  },
  {
    name: 'excel_set_range',
    // R351 — pair description with excel_set_cell sibling. Original
    // 「批量寫入 2D values」 was missing two AI-actionable invariants the
    // dispatcher actually enforces:
    //   • `sheet` is optional with the same「省略則用第一個 sheet」 default
    //     that excel_set_cell advertised at its sheet property — both
    //     tools route through `sheet && sheet.length > 0 ? sheet :
    //     model.activeSheet` (R326). AI looking at the two tools side-
    //     by-side saw the sheet hint only on set_cell and assumed
    //     set_range needed it, padding calls with `sheet: ""` (which
    //     R326 catches but still wastes a tokens-worth of typing).
    //   • Jagged-row tolerance: R298 validates per-cell types and R327
    //     bounds-check against EXCEL_MAX_ROWS/COLS, but neither path
    //     rejects rows of unequal width — R328's reduce-based maxRowLen
    //     handles them and R333's description string was just fixed to
    //     reflect the actual writable range. Telling AI「rows 可不等
    //     長」 upfront cuts a class of cautious-AI behavior where it
    //     pre-pads every row to maxRowLen to「play it safe」, doubling
    //     the request payload for no benefit.
    description:
      '從 startAddress 開始批量寫入 2D values（會自動延伸 sheet 範圍）。rows 寬度可以不一致（jagged），每列各自寫到其實際長度。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        sheet: { type: 'string', description: '工作表名稱；省略則用第一個 sheet（跟 excel_set_cell 一致）' },
        startAddress: { type: 'string', description: '左上角 A1 座標，e.g. A1' },
        values: {
          type: 'array',
          description: '2D 陣列；每一列是一 row 的 cell 值（cell 接受 string / number / boolean / null）',
          items: { type: 'array' },
        },
      },
      required: ['startAddress', 'values'],
    },
  },
  {
    name: 'excel_apply_format',
    description: '套用 cell 格式字串。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        sheet: { type: 'string' },
        range: { type: 'string' },
        formatString: { type: 'string' },
      },
      required: ['sheet', 'range', 'formatString'],
    },
  },
  {
    name: 'excel_insert_row',
    description: '插入指定行數。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        sheet: { type: 'string' },
        rowIndex: { type: 'integer' },
        count: { type: 'integer', minimum: 1 },
      },
      required: ['sheet', 'rowIndex', 'count'],
    },
  },
  {
    name: 'excel_insert_chart',
    description: '基於 range 建立圖表。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        sheet: { type: 'string' },
        range: { type: 'string' },
        chartType: { type: 'string', enum: ['bar', 'line', 'pie'] },
      },
      required: ['sheet', 'range', 'chartType'],
    },
  },

  // ── PowerPoint（spec stub）──────────────────────────────────────────
  {
    name: 'pptx_replace_text',
    // R367 — description aligned with R349's foundBy-based behaviour. The
    // dispatcher distinguishes THREE write paths post-R349:
    //   • runIndex provided        → full replace (AI's intent at this run)
    //   • oldText exact match      → full replace
    //   • oldText substring match  → only the matching substring is replaced
    // The OLD description「傳 oldText 做字串配對」 conflated exact +
    // substring into one「字串配對」 term, so AI couldn't tell:
    //   (a) if it wanted to「fully replace a run that contains text X」 →
    //       passing oldText=X with full-run-text would do full replace
    //       (good), but passing only a fragment did substring replace
    //       (which is sometimes what they wanted, sometimes not)
    //   (b) the「prefer runIndex」 hint from R349 — when AI accidentally
    //       provides both, the runtime error from R349 says「provide
    //       valid runIndex OR non-empty oldText」 but didn't surface
    //       the「provide BOTH = runIndex wins, oldText ignored」 rule.
    // pptx_empty hint also surfaced (R356) — name the recovery path so
    // AI can preempt it.
    description:
      '替換 slide 上的文字 run。三種匹配模式：(1) 傳 runIndex 指定第幾個 run（全文替換、最精準）；(2) 傳 oldText 且該值跟某個 run 的全部文字完全相同 → 全文替換；(3) 傳 oldText 且只是某個 run 的子字串 → 只替換該子字串。**請擇一提供 runIndex / oldText**：兩者同傳會以 runIndex 為準、oldText 忽略。對未初始化（0 bytes）的 pptx tab 會回 pptx_empty，請使用者先點該 tab 讓 PptxEditor 自動建空白投影片。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        slideIndex: { type: 'integer', description: '0-based 投影片 index' },
        runIndex: { type: 'integer', description: '0-based run index（在投影片內，跨所有 text shape 編號）；命中時走全文替換' },
        oldText: { type: 'string', description: '舊文字；先試 run.text === oldText 全文配對（全文替換）、失敗才試 run.text.includes(oldText) 子字串配對（只替換該子字串）。與 runIndex 擇一；同傳時被忽略。' },
        newText: { type: 'string', description: '替換後的文字（全文替換模式下成為新 run.text；子字串替換模式下替換掉 oldText 的位置）' },
      },
      required: ['slideIndex', 'newText'],
    },
  },
  {
    name: 'pptx_add_slide',
    description: '新增一張 slide。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        afterIndex: { type: 'integer' },
        layout: { type: 'string' },
        content: { type: 'object' },
      },
      required: ['afterIndex', 'layout'],
    },
  },
  {
    name: 'pptx_add_bullets',
    description: '在 slide 上加 bullets。v1.0 stub。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        slideIndex: { type: 'integer' },
        shapeId: { type: 'string' },
        bullets: { type: 'array', items: { type: 'string' } },
      },
      required: ['slideIndex', 'shapeId', 'bullets'],
    },
  },

  // ── 跨檔（read 全部支援；convert/summarize 部分支援）────────────────
  {
    name: 'read_tab_content',
    // R311 — description aligned with actual dispatcher behavior. Original
    // text claimed「其它格式回大綱與基本 metadata」(returns outline + metadata
    // for non-markdown formats), but doReadTabContent (dispatcher.ts:251-260,
    // post-R279/R281) returns ONLY `[type/name/byteLength + 改用 writeHint]`
    // for docx/xlsx/pptx — no outline. The mismatch made AI plan
    // read-then-write flows on binary tabs, get a "not supported" string
    // back, and either retry (burning round-trips and tokens) or reply
    // 「我無法讀取此檔案」to the user — looking like a system fault rather
    // than a v1.0 design choice. Replace with the truthful contract:
    // markdown → raw text; binary → metadata-only + write-tool pointer.
    // Same R279/R281 family that fixed buildOutline + system-prompt; this
    // closes the schema-side description (which is the canonical surface
    // AI reads when deciding whether to call a tool).
    //
    // R343 — add HTML to the text-extracting branch. Dispatcher's
    // doReadTabContent (post-HTML support) treats markdown AND html
    // symmetrically as text-based formats and returns raw .content for
    // both. The description previously only named markdown, so an AI
    // following the description would skip html tabs (assume they're in
    // the metadata-only docx/xlsx/pptx category) — but the user expects
    // 「請整理 index.html 的內容」 to work, same as for markdown. Same
    // R311-shape mismatch: the schema-side description is what AI plans
    // against; if it doesn't match the dispatcher, AI plans wrong tools.
    // maxChars truncation also applies to html in the dispatcher; the
    // option description's「僅對 markdown 生效」 was likewise wrong.
    description: '讀取 tab 內容。markdown / html 回 raw text；docx/xlsx/pptx **只回 metadata**（type / 檔名 / byteLength），不抽取內文 — 要修改這三種格式請直接呼叫 word_replace_paragraph / excel_set_cell / pptx_replace_text，無須事先 read。',
    input_schema: {
      type: 'object',
      properties: {
        tabId: { type: 'string', description: '目標 tab id；省略時使用當前 active tab' },
        options: {
          type: 'object',
          properties: {
            maxChars: { type: 'integer', description: '截斷上限，預設 8000（對 markdown / html 生效）' },
          },
        },
      },
      required: [],
    },
  },
  {
    name: 'convert_md_to_docx',
    // R354 — description corrected. Was「v1.0 stub：尚未實作 docx 寫入」 which
    // is factually wrong: docx writing IS implemented (word_replace_paragraph
    // / word_insert_paragraph have been wired since R266/R280, and the
    // system prompt's rule 5 explicitly directs AI to those tools). What's
    // missing is the tab-CREATION infrastructure — the dispatcher doesn't
    // construct `tab_create` ops yet (R347 doc-block confirms「tab_create
    // 跟 tab_delete 兩種 op 全 codebase 沒有任何 constructor」), so this
    // tool can't spawn a new docx tab from markdown content even though
    // both the read side (raw markdown) and the write side (docx
    // insertion) work. The old phrasing risked making AI under-use the
    // sibling docx write tools — same R311 / R343 / R353 family bug where
    // schema description and dispatcher capability disagree, but here the
    // description was MORE pessimistic than reality.
    //
    // Fallback hint added per R344's pattern (system prompt rule 5
    // already tells AI「HTML 沒有專用寫入工具…請直接回覆完整 HTML 文字」
    // — convert_md_to_docx is the analogous「沒有新 tab 建立工具」 case;
    // AI's recovery is to ask the user to manually create a new docx tab
    // and then call word_insert_paragraph paragraph by paragraph).
    description: '把 markdown tab 轉成新的 docx tab。v1.0 stub — 尚未實作「新增 docx tab」的 op 建構；現有 docx 寫入工具（word_replace_paragraph / word_insert_paragraph）需要既存的 docx tab 才能作用。請建議使用者先用工具列「+」新增空白 docx tab，再逐段呼叫 word_insert_paragraph 寫入。',
    input_schema: {
      type: 'object',
      properties: {
        // R401 — add `description` to bring this stub's property schema into
        // lockstep with every other tool in TOOLS. Pre-R401 these two fields
        // were bare `{ type: 'string' }` while sibling tools (e.g.,
        // cross_tab_summarize.sourceTabIds / .destTabId immediately below)
        // describe what AI should pass. The tool itself returns NOT_IMPLEMENTED
        // (dispatcher.ts:127-131) so AI never reaches the input-consumption
        // code path, but AI's plan-time reasoning still inspects the schema
        // when deciding which tool to invoke — undocumented fields would
        // make AI either skip this tool entirely (defensive) or guess at
        // semantics. Document them so AI's plan-step is informed even when
        // the runtime path returns the redirect.
        sourceTabId: { type: 'string', description: '要轉成 docx 的來源 markdown tab id（從 [Active workspace] outline 的 id= 值取得）' },
        destTabName: { type: 'string', description: '新建 docx tab 的顯示名稱；省略則由 dispatcher 自動命名（目前未實作，先看 description 的 fallback 流程）' },
      },
      required: ['sourceTabId'],
    },
  },
  {
    name: 'cross_tab_summarize',
    // R363 — three description bugs fixed:
    //
    // 1. Function name + old description「讀多個 tab、產出摘要寫入目標
    //    tab」 implies the dispatcher reads source tabs internally and
    //    auto-summarizes. It does NOT — dispatcher.ts:417-419 only
    //    appends `instructions` to `dest.content`:
    //      const after = `${dest.content}${sep}${instructions}`;
    //    The doc-block above the function explicitly says「We don't
    //    actually do summarization here — that's a model job. We just
    //    collect source content and surface it; AI is expected to call
    //    this after it has already produced the summary text, with
    //    `instructions` containing the summary itself.」AI reading the
    //    OLD description treated this as「magic auto-summarizer」,
    //    omitted the read_tab_content prep step, then either:
    //      (a) passed empty / vague instructions, expecting the tool
    //          to「do the work」 — got literal empty text appended
    //      (b) passed「請摘要這幾份」 as instructions, got literal
    //          那行字 appended to dest (instead of an actual summary)
    //    Either way the user saw their summary tab grow useless text.
    //
    // 2. sourceTabIds purpose was not documented. Dispatcher uses it
    //    only as a sanity guard (non-empty array required at line
    //    317-322); it's never read for content. AI thought passing
    //    fake ids was fine since「dispatcher doesn't validate them」 —
    //    but they're meant as CITATION, telling Gen-Doc which tabs
    //    informed the summary so the audit trail is accurate. Name
    //    them as「citation source ids」 in the description.
    //
    // 3. 「markdown→markdown 完整支援；其它組合 v1.0 stub」 was
    //    half-true. Dest enforcement is real (line 411 `if dest.type
    //    !== 'markdown' return NOT_IMPLEMENTED`), but SOURCE type
    //    doesn't matter at all since sources aren't read — saying
    //    「源類型不限」 is more accurate. Phase H+ can extend dest to
    //    html (sibling text-format); flag that explicitly.
    description:
      '把已經摘要好的文字 (instructions) 附加到一個既有的 markdown tab，並把 source tab ids 記錄為這段摘要的引用來源。**請注意：此工具不會自動讀取或摘要 source tabs** — 呼叫前必須先用 read_tab_content 讀完每個 source，自己合成摘要文字，再放進 instructions 一起傳入。dest 目前限定 markdown（html / docx / xlsx / pptx 是 v1.0 stub）。source 類型不限（這個欄位只做引用記錄）。',
    input_schema: {
      type: 'object',
      properties: {
        sourceTabIds: {
          type: 'array',
          items: { type: 'string' },
          description: '引用來源 tab id 清單；告訴 Gen-Doc「這段摘要是基於哪幾個 tab 寫的」。dispatcher 不讀這幾個 tab 的內容、純引用記錄。',
        },
        destTabId: {
          type: 'string',
          description: '寫入摘要的目標 tab id；目前限定 markdown（html / docx / xlsx / pptx 走 v1.0 stub）。',
        },
        instructions: {
          type: 'string',
          description: '**已經寫好的摘要 markdown 文字本身**（不是「請摘要 …」的指令）。dispatcher 把這段字直接 append 到 destTabId.content 結尾。',
        },
      },
      required: ['sourceTabIds', 'destTabId', 'instructions'],
    },
  },
];
