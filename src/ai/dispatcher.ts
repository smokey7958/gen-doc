/**
 * Tool dispatcher — converts a Claude tool_use into a ChangeSet (or a
 * tool_result error). Pure logic that doesn't mutate state directly; the
 * caller decides whether to enqueue the ChangeSet for diff preview.
 *
 * Spec §6.2 step [5]–[7]: validate → build ChangeSet → defer apply.
 */

import { v4 as uuid } from 'uuid';
import type { ToolUseBlock } from '../types/ai';
import type { ChangeSet, ChangeOp } from '../types/changeset';
import type { Tab, DocxTab, PptxTab, XlsxTab } from '../types/tab';
import { parseXlsx, serializeXlsx, parseA1, colIndexToLetter } from '../lib/xlsx-adapter';
import { isFormulaSource } from '../lib/xlsx-formula';
import { parseDocx, serializeDocx } from '../lib/docx-adapter';
import { parsePptx, serializePptx } from '../lib/pptx-adapter';
import { slicePreview } from '../lib/utils';

export interface DispatchSuccess {
  ok: true;
  changeset: ChangeSet;
  /** Short human-readable summary for diff preview & undo toast. */
  summary: string;
  /** Optional read-only tool_result (e.g. for read_tab_content). */
  inlineResult?: string;
}

export interface DispatchFailure {
  ok: false;
  error: string;
}

export type DispatchResult = DispatchSuccess | DispatchFailure;

export interface DispatchContext {
  tabs: Tab[];
  activeTabId: string | null;
}

// R358 — fully-qualified tool names in the fallback hint. The original
// listed bare `set_cell / replace_paragraph / replace_text` which DO NOT
// EXIST in the tool schema (they are prefixed `excel_set_cell` /
// `word_replace_paragraph` / `pptx_replace_text`). AI reading this hint
// has two options:
//   1. Best case — pattern-match the prefix-less names against the schema,
//      pick the right prefixed sibling, retry. Works but burns a chain-of-
//      thought step.
//   2. Worst case — pass the bare name through, hit `unknown_tool:
//      set_cell` from the dispatcher's default branch, lose another turn
//      figuring out the disambiguation.
// Either way the user-facing latency takes a hit on every stub-tool
// fallback. Fully-qualifying the names removes the guesswork. The「或先
// 在 markdown tab 編寫內容」escape hatch stays — it's a real fallback
// for cases where none of the basic tools fit (e.g., excel_insert_chart
// has no equivalent at all, and AI can recommend「我幫你在 markdown 寫
// 圖表的 caption + 占位描述、之後你手動建 Excel chart」).
const NOT_IMPLEMENTED = (tool: string): DispatchFailure => ({
  ok: false,
  error: `not_implemented_in_mvp: ${tool} — 樣式 / 格式 / 圖表 / 投影片新增等進階操作尚未支援。請改用基本的 excel_set_cell / word_replace_paragraph / pptx_replace_text，或先在 markdown tab 編寫內容。`,
});

function findTab(ctx: DispatchContext, tabId?: string): Tab | null {
  if (tabId) return ctx.tabs.find((t) => t.id === tabId) ?? null;
  if (ctx.activeTabId) return ctx.tabs.find((t) => t.id === ctx.activeTabId) ?? null;
  return null;
}

export async function dispatchTool(
  block: ToolUseBlock,
  ctx: DispatchContext,
): Promise<DispatchResult> {
  // R299 — centralized input shape guard. Every helper below destructures
  // `block.input as { ... }` immediately; if block.input is null / undefined
  // / non-object at runtime (malformed tool_use chunk, future SDK partial-
  // parse path, test mock typo), destructuring throws TypeError "Cannot
  // destructure property 'X' of 'null'". R228 outer catch wraps it as
  // `dispatch_threw: Cannot destructure...` — functional but unhelpful for
  // AI to recover from. Same R291 / R295 / R297 / R298 idiom: convert raw
  // JS TypeError into a typed `bad_tool_input` error that AI can read.
  // Note: `typeof null === 'object'` is the JS gotcha — combine both
  // checks. Arrays also fail since `Array.isArray([])` would treat them as
  // valid objects but no tool helper expects array input; explicit reject.
  if (!block.input || typeof block.input !== 'object' || Array.isArray(block.input)) {
    return { ok: false, error: `bad_tool_input: ${block.name} requires an object input` };
  }
  switch (block.name) {
    case 'md_replace_section':
      return doMdReplaceSection(block, ctx);
    case 'md_append':
      return doMdAppend(block, ctx);
    case 'md_insert_at':
      return doMdInsertAt(block, ctx);
    case 'read_tab_content':
      return doReadTabContent(block, ctx);
    case 'cross_tab_summarize':
      return doCrossTabSummarize(block, ctx);
    case 'excel_set_cell':
      return doExcelSetCell(block, ctx);
    case 'excel_set_range':
      return doExcelSetRange(block, ctx);
    case 'word_replace_paragraph':
      return doWordReplaceParagraph(block, ctx);
    case 'word_insert_paragraph':
      return doWordInsertParagraph(block, ctx);
    case 'pptx_replace_text':
      return doPptxReplaceText(block, ctx);
    // R357 — tool-specific NOT_IMPLEMENTED messages for the two stubs whose
    // tool-description (R354 / R355) already documents a working workaround.
    // The generic NOT_IMPLEMENTED helper's「請改用基本的 set_cell /
    // replace_paragraph / replace_text，或先在 markdown tab 編寫內容」 hint
    // doesn't tell AI WHICH redirect to take for THIS specific stub — AI
    // gets the message at runtime (tool_result), not the plan-time schema
    // description, so the redirect from the description doesn't reach the
    // recovery loop. Embedding the redirect into the runtime error mirrors
    // R356's「accurate recovery hint」pattern for pptx_empty: align what
    // AI sees during execution with what the user-facing fallback path
    // actually is. The other six stubs (word_apply_style /
    // excel_apply_format / excel_insert_row / excel_insert_chart /
    // pptx_add_slide / pptx_add_bullets) have no working sibling, so the
    // generic「樣式 / 格式 / 圖表 / 投影片新增等進階操作尚未支援」 message
    // fits them — keep them on the generic helper.
    case 'word_insert_heading':
      return {
        ok: false,
        error: 'not_implemented_in_mvp: word_insert_heading — 已被 word_insert_paragraph 取代，請改呼叫 word_insert_paragraph({ paragraphIndex: afterIndex + 1, text, kind: "heading1" | … | "heading6" })，效果完全等價且支援全 6 個 heading level。',
      };
    case 'convert_md_to_docx':
      return {
        ok: false,
        error: 'not_implemented_in_mvp: convert_md_to_docx — 尚未實作「新增 docx tab」的 op 建構。現有 docx 寫入工具（word_replace_paragraph / word_insert_paragraph）需要既存 docx tab 才能作用。請告訴使用者先用工具列「+」新增空白 docx tab，再逐段呼叫 word_insert_paragraph 寫入。',
      };
    case 'word_apply_style':
    case 'excel_apply_format':
    case 'excel_insert_row':
    case 'excel_insert_chart':
    case 'pptx_add_slide':
    case 'pptx_add_bullets':
      return NOT_IMPLEMENTED(block.name);
    default:
      return { ok: false, error: `unknown_tool: ${block.name}` };
  }
}

// ── Markdown ops ─────────────────────────────────────────────────────────

type ExpectMdResult =
  | { ok: true; tab: import('../types/tab').MarkdownTab }
  | { ok: false; error: string };

function expectMdTab(ctx: DispatchContext, tabId?: string): ExpectMdResult {
  const tab = findTab(ctx, tabId);
  // R310 — disambiguate `findTab(...) === null`. The helper returns null for
  // TWO distinct reasons: (a) tabId provided but not present in ctx.tabs (AI
  // typo'd the id, or referenced a just-closed tab), (b) no tabId AND no
  // activeTabId (workspace is empty). Collapsing both under 'no_active_tab'
  // misleads AI: case (a) needs "re-check the outline for valid ids", case
  // (b) needs "ask user to open / create a tab". With a typo'd tabId AI
  // previously responded 「請先開啟或新增一個檔案再試」 even though the
  // user had several tabs open — confusing UX. Echoing the bad id in
  // tab_not_found lets AI compare against the outline ([Active workspace]
  // block sent on every turn, see orchestrator.ts:942-979 buildOutline)
  // and self-correct without bouncing back to the user. Same R291/R295/
  // R297/R298 idiom: convert ambiguous error into typed signal AI can
  // recover from.
  if (!tab) {
    return { ok: false, error: tabId ? `tab_not_found: ${tabId}` : 'no_active_tab' };
  }
  // R399 — echo the resolved tab id in tab_type_mismatch errors. Matches
  // R310's `tab_not_found: ${tabId}` echo idiom: when AI emits multiple
  // tool_uses in one turn with different tab references and one of them
  // mis-targets a wrong-type tab, the error message alone tells AI WHICH
  // call failed (cross-reference the echoed id against the [Active
  // workspace] outline on the next turn) without needing the user to
  // describe it. Use `tab.id` (the resolved id) rather than `tabId` (the
  // arg, which may be undefined when active-tab fallback was used) so
  // even the「omit tabId → use active tab → active tab is wrong type」
  // case surfaces a useful id. Sibling expects (expectXlsxTab /
  // expectDocxTab / expectPptxTab) get the same echo for parity.
  if (tab.type !== 'markdown')
    return { ok: false, error: `tab_type_mismatch: expected markdown, got ${tab.type} (tabId: ${tab.id})` };
  return { ok: true, tab };
}

function doMdReplaceSection(
  block: ToolUseBlock,
  ctx: DispatchContext,
): DispatchResult {
  const { tabId, heading, newContent } = block.input as {
    tabId?: string;
    heading: string;
    newContent: string;
  };
  // R291 — strict typeof checks aligned with sibling tools (doMdAppend /
  // doMdInsertAt / doWordReplaceParagraph / doWordInsertParagraph /
  // doPptxReplaceText / doCrossTabSummarize). The original `!heading`
  // loose check let non-string truthy values (number, boolean, array)
  // through; `replaceMdSection` then calls `heading.trim()` which throws
  // TypeError, caught by R228's outer dispatch_threw but with a much
  // less actionable error message to the AI than missing_required_fields.
  if (typeof heading !== 'string' || heading.length === 0 || typeof newContent !== 'string') {
    return { ok: false, error: 'missing_required_fields: heading, newContent' };
  }
  const got = expectMdTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.content;
  const replaced = replaceMdSection(before, heading, newContent);
  if (replaced === null) {
    return { ok: false, error: `heading_not_found: "${heading}"` };
  }
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: `替換章節「${heading}」`,
      ops: [
        { tabId: tab.id, type: 'md_text', before, after: replaced } satisfies ChangeOp,
      ],
    },
    summary: `替換章節「${heading}」`,
  };
}

function doMdAppend(block: ToolUseBlock, ctx: DispatchContext): DispatchResult {
  const { tabId, text } = block.input as { tabId?: string; text: string };
  // R324 — also reject empty `text`. The previous check only enforced type;
  // empty string passed through, produced a no-op-ish changeset that just
  // appended `\n` or `\n\n` to the doc (via the sep logic below), dirtied
  // the workspace, and the user got a「附加 0 字元」 PendingChange card
  // with no visible content change. Same fail-fast idiom as R283
  // (doCrossTabSummarize) / R313 (doPptxReplaceText empty oldText) / R314
  // (doExcelSetRange empty values) — convert ambiguous "semantically empty"
  // input into a typed error so AI can self-correct (e.g., re-issue as
  // md_insert_at if the intent was a paragraph break, or skip the call
  // entirely). Other markdown tools (doMdInsertAt's blank paragraph,
  // doMdReplaceSection's delete-section semantic) legitimately accept
  // empty text and stay unchanged.
  if (typeof text !== 'string' || text.length === 0) {
    return {
      ok: false,
      error: 'missing_required_fields: text (must be non-empty — append of empty content has no meaning; use md_insert_at if you intended a blank paragraph break)',
    };
  }
  const got = expectMdTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  // R331 — also treat an empty document as "no separator needed". Without the
  // length===0 branch, the existing chain falls through to the default `\n\n`
  // case (empty string doesn't endsWith `\n\n` nor `\n`), so md_append on an
  // empty markdown tab produces `'\n\n' + text` — the file opens with two
  // leading blank lines before the first heading. Realistic triggers: user
  // drops a 0-byte .md from disk → openExternalFile decodes "" → AI does its
  // first md_append; user clears NEW_TAB_DEFAULTS's "# 新文件\n\n" boilerplate
  // with Ctrl+A/Delete before asking AI to draft something; or a .gd is saved
  // with an empty markdown tab (the editor permits empty buffers — content
  // length isn't gated anywhere). The two-blank-line preamble looks like a
  // typo to the user and breaks the visual top-of-document expectation. Same
  // fail-fast input-cleanup family as R324 (empty `text` rejection), R313
  // (empty oldText), R314 (empty values).
  const sep =
    tab.content.length === 0 || tab.content.endsWith('\n\n')
      ? ''
      : tab.content.endsWith('\n')
        ? '\n'
        : '\n\n';
  const after = `${tab.content}${sep}${text}`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: '附加內容到 Markdown',
      ops: [{ tabId: tab.id, type: 'md_text', before: tab.content, after }],
    },
    summary: `附加 ${text.length} 字元到 ${tab.name}`,
  };
}

function doMdInsertAt(block: ToolUseBlock, ctx: DispatchContext): DispatchResult {
  const { tabId, line, text } = block.input as {
    tabId?: string;
    line: number;
    text: string;
  };
  // R295 — strict integer check. `typeof line === 'number'` lets NaN /
  // Infinity / fractions through; `lines.splice(NaN, 0, text)` treats NaN
  // as 0, silently inserting at wrong position. Number.isInteger blocks
  // all three pathological cases at once. Negative line is also caught
  // by the same gate (line numbers are 1-based positive integers per the
  // tool schema).
  if (!Number.isInteger(line) || line < 1 || typeof text !== 'string') {
    return { ok: false, error: 'missing_required_fields: line, text (line must be a positive integer)' };
  }
  const got = expectMdTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const lines = tab.content.split('\n');
  const idx = Math.max(0, Math.min(line - 1, lines.length));
  lines.splice(idx, 0, text);
  const after = lines.join('\n');
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: `在第 ${line} 行插入內容`,
      ops: [{ tabId: tab.id, type: 'md_text', before: tab.content, after }],
    },
    summary: `在 ${tab.name} 第 ${line} 行插入`,
  };
}

// ── Cross-file ───────────────────────────────────────────────────────────

function doReadTabContent(block: ToolUseBlock, ctx: DispatchContext): DispatchResult {
  const { tabId, options } = block.input as {
    tabId?: string;
    options?: { maxChars?: number };
  };
  const tab = findTab(ctx, tabId);
  // R310 — same disambiguation as expectMdTab sibling above. doReadTabContent's
  // previous bare 'tab_not_found' was the inverse mistake: it claimed a
  // specific tab was missing even when tabId was omitted (in which case the
  // failure is really「no active tab」). Pair-it-up with the rest of the
  // dispatcher's helpers so AI sees consistent vocabulary across every
  // tool path.
  if (!tab) {
    return { ok: false, error: tabId ? `tab_not_found: ${tabId}` : 'no_active_tab' };
  }
  // R346 — validate `maxChars` before using it as a slice index. Schema
  // says `integer`, but the runtime cast at the destructure is bypassed
  // by whatever JSON the model emits; AI passing pathological values
  // surface as silent garbage rather than typed errors:
  //   • `maxChars: 0` → max=0 → `length > 0` true for any non-empty
  //     content → `slice(0, 0)` returns "" → AI gets just the truncation
  //     marker「[…內容過長，已截斷至 0 字]」 with literally none of the
  //     content. AI then re-plans against an empty read and either calls
  //     read_tab_content again with a different cap (wasting a turn) or
  //     replies「文件內容讀不到」 to the user — looks like a system fault.
  //   • `maxChars: -1` → max=-1 → `length > -1` always true → `slice(0,-1)`
  //     drops the LAST character (`String.slice` treats negative as
  //     length+arg) → AI gets nearly-full content but tagged as
  //     truncated and missing the final char (which could be a closing
  //     `>` / `}` / `)` mid-token — confusing for AI trying to reason
  //     about the syntax tree).
  //   • `maxChars: NaN` / `Infinity` → `length > NaN` false → returns
  //     full content with no truncation marker, despite AI's intent
  //     having been bounded.
  //   • `maxChars: 5.5` (fractional) → `slice(0, 5.5)` floors to 5 with
  //     no signal — minor but inconsistent.
  // Fall back to default 8000 when AI passes anything outside the legal
  // positive-integer range, matching the same R295 idiom for invalid
  // line / paragraphIndex / slideIndex / runIndex inputs across the
  // dispatcher (those reject with missing_required_fields; here we
  // silently default because maxChars is OPTIONAL — the AI calling
  // read_tab_content with a bad cap clearly wants to read, just got
  // the cap wrong, and falling through to the canonical 8000 default
  // lets the call still produce useful output rather than failing).
  const rawMax = options?.maxChars;
  const max =
    typeof rawMax === 'number' && Number.isInteger(rawMax) && rawMax > 0
      ? rawMax
      : 8000;
  let result: string;
  if (tab.type === 'markdown' || tab.type === 'html') {
    // HTML is text-based like markdown: return raw source, capped at maxChars
    // with the same truncation marker. AI can edit via TODO future html_* tools
    // or by reading + advising the user; for now the read path mirrors md.
    //
    // R381 — surrogate-pair-aware truncation. `String.length` counts UTF-16
    // code units, and `slice(0, n)` is code-unit-indexed too — so a doc
    // whose code unit at position `max - 1` is the HIGH surrogate of an
    // emoji (e.g.「🚀」 = `🚀`, two code units) gets sliced
    // mid-pair, leaving an orphan high surrogate at the truncation
    // boundary. The resulting string is technically valid UTF-16 but
    // contains an unpaired surrogate that renders as「�」 / fails clean
    // JSON serialization on some downstream paths (Anthropic's API
    // tolerates it but some intermediate logs / SDK normalisers don't).
    // Realistic trigger: a markdown doc with emoji bullet points that
    // happens to have one straddling the 8000-char boundary — common
    // for emoji-heavy task lists / chat exports / changelogs.
    //
    // `Array.from(str)` iterates by code POINT (Unicode-aware), so the
    // resulting array has whole emoji as single elements. Slicing this
    // array and joining gives a valid UTF-16 string with no orphan
    // surrogates. Cost: `Array.from(10MB)` allocates a 10MB-ish array,
    // which is fine for the typical < 1MB markdown / html source the
    // user-facing tab limit puts us at.
    //
    // Also note: this slightly changes the truncation semantics for
    // emoji-heavy docs —「max=8000」 now means「up to 8000 Unicode
    // code points」 rather than「up to 8000 UTF-16 code units」. For
    // ASCII / CJK docs the two are identical (each char is one code
    // point AND one code unit); only surrogate-pair-using emoji
    // differ. The user / AI mental model of「max chars」 aligns more
    // naturally with code points anyway — counting an emoji as 1
    // matches what the viewer sees, not 2.
    // Fast path: `length` (code units) is an upper bound on code-point
    // count. When `length <= max`, code-point count is also <= max, so no
    // need to Array.from. Only when `length > max` do we actually need to
    // count code points to decide whether the content is over-cap.
    // Without the inner check, an emoji-only doc with 4001 emojis (length
    // = 8002 code units, code points = 4001, max = 8000) would enter the
    // truncation branch via length > max but its real code-point count
    // is UNDER max — appending the truncation marker without actually
    // truncating would lie to AI about「截斷」 happening when it didn't.
    if (tab.content.length > max) {
      const codePoints = Array.from(tab.content);
      if (codePoints.length > max) {
        result = codePoints.slice(0, max).join('') + `\n\n[…內容過長，已截斷至 ${max} 字]`;
      } else {
        // length > max but codePoints <= max — content has surrogates that
        // inflate code-unit count above the cap, but real code-point count
        // is within limit. Return full content WITHOUT the truncation
        // marker, since no truncation actually happened.
        result = tab.content;
      }
    } else {
      result = tab.content;
    }
  } else {
    // R281 — sibling fix to R279 (buildOutline + system-prompt cleanup).
    // Original message conflated "read_tab_content can't extract plain text
    // from this binary format" (true — this function's own limitation) with
    // "v1.0 唯讀" (false — basic write tools are implemented for all three
    // binary formats: word_replace_paragraph / excel_set_cell / pptx_
    // replace_text). The mixed message made AI default to refusing writes
    // on binary tabs even when the requested operation was supported.
    // Reframe the response to describe ONLY this function's limit and
    // point at the available write tools — same three-edge consistency
    // R279 set up (outline says basic-write-available, system prompt says
    // not_implemented_in_mvp is per-operation not per-format, this
    // response now says ditto).
    const writeHint =
      tab.type === 'docx'
        ? 'word_replace_paragraph / word_insert_paragraph'
        : tab.type === 'xlsx'
          ? 'excel_set_cell / excel_set_range'
          : 'pptx_replace_text';
    result = `[${tab.type} tab "${tab.name}"，${tab.data.byteLength} bytes — read_tab_content 尚未支援此格式的純文字抽取。要編輯請直接呼叫 ${writeHint}（不需先讀取就能修改）。]`;
  }
  // Read tools don't produce a ChangeSet — we return inlineResult so the
  // caller can synthesize a tool_result message back to the model.
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: `讀取 ${tab.name}`,
      ops: [],
    },
    summary: `讀取 ${tab.name}`,
    inlineResult: result,
  };
}

function doCrossTabSummarize(block: ToolUseBlock, ctx: DispatchContext): DispatchResult {
  const { sourceTabIds, destTabId, instructions } = block.input as {
    sourceTabIds: string[];
    destTabId: string;
    instructions: string;
  };
  if (!Array.isArray(sourceTabIds) || sourceTabIds.length === 0) {
    return {
      ok: false,
      error: 'sourceTabIds_empty: provide at least one tab id (use the outline\'s id= values) — sourceTabIds documents which tabs informed the summary, even when the summary text itself is precomputed in `instructions`',
    };
  }
  // R403 — also validate each entry is a non-empty string. The destructure
  // cast (`sourceTabIds: string[]`) is TypeScript-only; runtime can hold
  // anything AI emitted. Today the dispatcher doesn't actually consume the
  // contents (R363 doc-block confirms「dispatcher uses it only as a sanity
  // guard」), so a stray number / null / object in the array slips through
  // silently — but that's a dormant trap for any future code path that
  // reads the citation for display (PendingChange detail panel, undo
  // toast, exported audit log). The non-empty-array check at line 468
  // would pass for `[123, null, 'valid-id']`, masking the input
  // corruption until a downstream consumer dereferences `.toLowerCase()`
  // or `.split()` on a non-string. Same fail-fast idiom as R297
  // (excel_set_cell value) / R298 (excel_set_range per-cell) /
  // R325 (destTabId strict-type) — convert ambiguous「technically a
  // string[] per TS, garbage at runtime」 into a typed error AI can
  // self-correct by re-emitting the tool_use with the proper outline
  // id= values. Empty-string entries are also rejected (no legitimate
  // tab has empty id; sneaks past Array.isArray + length checks).
  const badIdx = sourceTabIds.findIndex(
    (id) => typeof id !== 'string' || id.length === 0,
  );
  if (badIdx >= 0) {
    return {
      ok: false,
      error: `bad_sourceTabIds: entry ${badIdx} is not a non-empty string (got ${typeof sourceTabIds[badIdx]} — sourceTabIds should be tab id strings from the outline's id= values)`,
    };
  }
  // R283 — instructions guard, mirrors the sibling tool validation idiom
  // (doMdAppend / doMdInsertAt / doWordReplaceParagraph / doWordInsertParagraph
  // / doPptxReplaceText all check `typeof X !== 'string'` for their string
  // params). The `as { instructions: string }` cast at the destructure is
  // TypeScript-only; runtime gets whatever the AI emitted. An omitted
  // instructions field (despite tool schema marking it required) coerces to
  // `"undefined"` via string concatenation at line ~257, appending the
  // literal word "undefined" to the destination markdown. Empty string is
  // also caught — appending a blank line carries no information and
  // pollutes the doc with no user-visible content.
  if (typeof instructions !== 'string' || instructions.length === 0) {
    return { ok: false, error: 'missing_required_fields: instructions' };
  }
  // R325 — also strict-type destTabId. Without this, AI passing null /
  // undefined / number for destTabId would walk into ctx.tabs.find with
  // a non-string predicate, return undefined, and surface as the
  // misleading `dest_tab_not_found` below (which AI then reads as "the
  // tab was closed" rather than "you passed the wrong type"). Same fix
  // shape as R310's expectMdTab disambiguation.
  if (typeof destTabId !== 'string' || destTabId.length === 0) {
    return { ok: false, error: 'missing_required_fields: destTabId' };
  }
  const dest = ctx.tabs.find((t) => t.id === destTabId);
  // R325 — include destTabId in the error so AI can compare against the
  // outline's id= values without an extra round-trip. R310 already
  // aligned `tab_not_found: ${tabId}` for expectMd/Xlsx/Docx/Pptx +
  // doReadTabContent; this is the lone callsite that side-stepped the
  // findTab helper (uses ctx.tabs.find directly because the destTabId is
  // explicit, not optional/active-tab fallback) and so was missed.
  if (!dest) return { ok: false, error: `dest_tab_not_found: ${destTabId}` };
  // R402 — distinguish「dest type not supported」 from「entire tool not
  // implemented」. The previous NOT_IMPLEMENTED('cross_tab_summarize')
  // returned the generic helper's「樣式 / 格式 / 圖表 / 投影片新增等進
  // 階操作尚未支援。請改用基本的 excel_set_cell / word_replace_paragraph
  // / pptx_replace_text，或先在 markdown tab 編寫內容」 — which is wrong
  // on two axes:
  //   (a) cross_tab_summarize is fully implemented; it just constrains dest
  //       to markdown (R325 + this line). Telling AI「整個工具尚未支援」
  //       implies the feature doesn't exist, when actually a different
  //       destTabId pick would work.
  //   (b) The redirect to「改用基本的 excel_set_cell / ...」 is wrong —
  //       those are unrelated single-cell writers, NOT a cross-tab
  //       summary alternative. AI following the redirect either writes
  //       to a single cell (defeating the purpose of「跨檔摘要」) or
  //       gives up on the task entirely.
  // Concrete failure: AI calls cross_tab_summarize({destTabId: '<xlsx-
  // tab-uuid>', sourceTabIds: [...], instructions: '...'}). Pre-R402
  // dispatcher returns NOT_IMPLEMENTED with the generic message. AI tells
  // user「跨檔摘要尚未支援」 — but the user has a perfectly fine markdown
  // tab they could have used as dest. The friction is misplaced.
  //
  // Echo dest.id matches R310 / R399's tabId-echo idiom — AI can identify
  // which dest pick was wrong (vs scanning the outline blindly). The
  // recovery hint「請改傳 markdown 類型的 destTabId」 names the actual
  // fix concretely.
  if (dest.type !== 'markdown')
    return {
      ok: false,
      error: `dest_type_mismatch: cross_tab_summarize 目前只支援寫入 markdown tab，destTabId 對應的是 ${dest.type} (tabId: ${dest.id})。請改傳一個 markdown 類型的 destTabId（從 [Active workspace] outline 中找 [md] 開頭的項目）。`,
    };

  // We don't actually do summarization here — that's a model job. We just
  // collect source content and surface it; AI is expected to call this
  // after it has already produced the summary text, with `instructions`
  // containing the summary itself.
  //
  // R348 — match doMdAppend's R331 empty-doc handling: skip the separator
  // entirely when dest.content is empty. The original separator chain
  //
  //   dest.content + (endsWith '\n' ? '' : '\n') + '\n' + instructions
  //
  // for an empty string `""` falls through both ternaries the wrong way:
  // `"".endsWith('\n')` is false → adds `'\n'`, then the literal `'\n'`
  // appends another newline. Result: `'\n\n' + instructions` — the
  // summarized doc opens with two blank lines before the first character
  // of the summary. Realistic trigger: AI calls cross_tab_summarize on
  // a freshly added markdown tab where the user cleared the default
  // 「# 新文件\n\n」 boilerplate (or AI was directed at a brand-new
  // markdown tab created specifically as a summary destination — common
  //「為這幾個 tab 產一份摘要到新的 tab」 ask). Same R331-shape bug, same
  // R331 fix (length===0 short-circuit). Other doMd* paths now consistent:
  // R331 covers append, R348 covers cross_tab_summarize; replace_section
  // doesn't need the guard (it fully replaces, no separator-with-prior-
  // content step) and insert_at uses splice (no separator either).
  const sep =
    dest.content.length === 0
      ? ''
      : dest.content.endsWith('\n\n')
        ? ''
        : dest.content.endsWith('\n')
          ? '\n'
          : '\n\n';
  const after = `${dest.content}${sep}${instructions}`;
  // R404 — align changeset.description with summary so the undo toast names
  // the dest tab. Pre-R404 description was the hardcoded「跨檔摘要寫入」 with
  // no dest identifier; the App.tsx undo flow renders「Undo:
  // ${cs.description}」 as a toast (changeset-apply.ts:31), so a user with
  // multiple cross_tab_summarize entries in their undo stack (legitimate
  // for workflow「先做 budget 摘要、再做 Q1 摘要、再做 Q2 摘要」) would
  // see THREE identical「Undo: 跨檔摘要寫入」 toasts as they Ctrl+Z back.
  // No way to tell which dest tab each undo just touched. Every sibling
  // dispatcher tool (doExcelSetCell at line 681-682, doMdReplaceSection
  // at line 206-211, doWordReplaceParagraph at line 1029, doPptxReplaceText
  // at line 1262) uses the same descriptive string for both `description`
  // and `summary` — this was the lone outlier. Using `summary`'s exact
  // wording keeps the「PendingChange card / undo toast」 voice consistent
  // (R332 / R349 family established「summary names the user-visible
  // change」 idiom for the rest of the dispatcher).
  const desc = `跨檔摘要 → ${dest.name}`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: desc,
      ops: [{ tabId: dest.id, type: 'md_text', before: dest.content, after }],
    },
    summary: desc,
  };
}

// ── Excel ops ────────────────────────────────────────────────────────────

function expectXlsxTab(ctx: DispatchContext, tabId?: string): { ok: true; tab: XlsxTab } | { ok: false; error: string } {
  const tab = findTab(ctx, tabId);
  // R310 — see expectMdTab doc-block.
  if (!tab) {
    return { ok: false, error: tabId ? `tab_not_found: ${tabId}` : 'no_active_tab' };
  }
  // R399 — see expectMdTab sibling for the tabId-echo rationale.
  if (tab.type !== 'xlsx')
    return { ok: false, error: `tab_type_mismatch: expected xlsx, got ${tab.type} (tabId: ${tab.id})` };
  return { ok: true, tab };
}

async function doExcelSetCell(block: ToolUseBlock, ctx: DispatchContext): Promise<DispatchResult> {
  const { tabId, sheet, address, value } = block.input as {
    tabId?: string;
    sheet?: string;
    address: string;
    value: string | number | boolean | null;
  };
  // R291 — strict typeof. `parseA1(address.toUpperCase())` throws on non-
  // string; align with sibling tools' validation idiom.
  if (typeof address !== 'string' || address.length === 0) {
    return { ok: false, error: 'missing_required_fields: address' };
  }
  // R297 — value primitive guard. Tool schema declares
  // `value: string | number | boolean | null` as required, but the `as`
  // cast at the destructure is TypeScript-only — runtime sees whatever AI
  // emits. Without this check, an omitted value coerces via
  // `String(undefined) === 'undefined'` and the cell is silently
  // populated with the literal text "undefined" (likewise "[object
  // Object]" for an object, "1,2,3" for an array). Same R283 idiom for
  // doCrossTabSummarize's instructions field. Reject anything that's
  // not in the schema's primitive union.
  const valType = typeof value;
  if (
    value !== null &&
    valType !== 'string' &&
    valType !== 'number' &&
    valType !== 'boolean'
  ) {
    return { ok: false, error: 'missing_required_fields: value (must be string, number, boolean, or null)' };
  }
  // R352 — reject non-finite numeric values (NaN / Infinity / -Infinity).
  // R297 catches non-primitive types, but `typeof NaN === 'number'` and
  // `typeof Infinity === 'number'` so both slip past the union check.
  // Down the pipeline `String(NaN) === "NaN"` and `String(Infinity)
  // === "Infinity"`, so the cell ends up showing the literal text
  // 「NaN」 / 「Infinity」. Excel's native error format is `#NUM!` /
  // `#DIV/0!` — the JS string form looks like a buggy paste rather
  // than a computed error. Realistic trigger: AI runs a divide-by-zero
  // / overflow / NaN-from-parseFloat() in its planning step and writes
  // the result verbatim. Same fail-fast idiom as R297 / R298 / R314 /
  // R346 — convert ambiguous corruption into a typed error so AI can
  // re-plan (e.g., write `null` for missing, or emit a string label
  // like "—" for unrepresentable values).
  if (valType === 'number' && !Number.isFinite(value as number)) {
    return {
      ok: false,
      error: 'bad_cell_value: numeric values must be finite (got NaN / Infinity / -Infinity); use null for empty cells, or a string label like "—" for unrepresentable results',
    };
  }
  const got = expectXlsxTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.data;

  const model = parseXlsx(before);
  // R326 — treat empty `sheet` ("" / null / undefined) as "use the active
  // sheet". The schema describes sheet as optional with「省略則用第一個 sheet」
  // semantics; AI commonly implements「omit」 as `sheet: ""` (empty string is
  // a natural fallback for an optional schema field). `??` doesn't fire on
  // empty string, so the empty string passes through to `.find()` which
  // returns undefined → AI saw the misleading `sheet_not_found:` (empty
  // after the colon). Excel sheet names are 1-31 chars and can't be empty,
  // so treating "" as「未提供」 is safe — no legitimate sheet has an empty
  // name to mask.
  const sheetName = sheet && sheet.length > 0 ? sheet : model.activeSheet;
  const target = model.sheets.find((s) => s.name === sheetName);
  if (!target) return { ok: false, error: `sheet_not_found: ${sheetName}` };
  const a1 = parseA1(address.toUpperCase());
  if (!a1) return { ok: false, error: `bad_a1_address: ${address}` };
  // R301 — include negative-bound check. parseA1 returns `parseInt(row) - 1`,
  // so a `"0"` row in the address (AI mistakenly treating Excel rows as
  // 0-based) gives r=-1. Without the `< 0` guard, `target.cells[-1][0]`
  // reads undefined, the R269 `const { formula } = prevCell` destructure
  // throws TypeError on undefined, R228 wraps it as the cryptic
  // `dispatch_threw: Cannot destructure...`. Same address_out_of_range
  // error for both bound directions keeps AI's recovery path uniform.
  if (
    a1.r < 0 ||
    a1.c < 0 ||
    a1.r >= target.cells.length ||
    a1.c >= (target.cells[0]?.length ?? 0)
  ) {
    return { ok: false, error: `address_out_of_range: ${address}` };
  }
  // R269 — preserve prev cell's style / rawType when overriding text; only
  // clear `formula` since AI is setting a literal value. Same invariant
  // XlsxEditor.writeCell (XlsxEditor.tsx:626-638) enforces on user typing:
  // a bare `{ text }` replacement throws away bold / colour / numberFormat
  // / font / fontSize and any rawType info, so AI's value lands correctly
  // but every prior formatting on that cell vanishes after Apply.
  const prevCell = target.cells[a1.r][a1.c];
  const newText = value === null ? '' : String(value);
  // R303 — mirror XlsxEditor.writeCell's two-branch invariant. R269 only
  // mirrored the non-formula branch (drop `formula` to keep text
  // authoritative). The formula branch (set `formula = text` when the
  // value starts with `=`) was missing — AI calls to set a formula via
  // excel_set_cell produced cells with the formula text as a literal
  // string, not a computable formula. Now AI typing `'=SUM(B1:B10)'`
  // produces the same cell as user-typed `=SUM(B1:B10)`.
  const { formula: _dropF, ...restCell } = prevCell;
  void _dropF;
  target.cells[a1.r][a1.c] = isFormulaSource(newText)
    ? { ...restCell, text: newText, formula: newText }
    : { ...restCell, text: newText };
  const after = serializeXlsx(model, before);
  // R368 — preview-truncate the value in the description. The raw String(value)
  // can be hundreds-to-thousands of characters when AI writes long text into a
  // cell (Excel's per-cell text limit is 32,767 chars), and that string ends
  // up in THREE user-visible / persistent surfaces:
  //   • PendingChange card title in AIPanel — overflows the card width
  //   • Undo toast「Undo: Sheet1!B5 = ...」— bumps toast height
  //   • SQLite undo_entries.changeset_json — `description` is persisted
  //     verbatim, multiplying long-value rows in the undo store
  // The actual cell write (`target.cells[a1.r][a1.c]`) and BinaryReplaceOp's
  // `before` / `after` still carry the full value bytes; this truncation
  // is purely for the human-readable summary string. 60-char + ellipsis
  // mirrors the selection-context preview convention (MarkdownEditor /
  // HtmlEditor's R336 selection badge) so the visual cue「…」 is consistent
  // across the app's preview surfaces.
  // R382 — use slicePreview helper for code-point-aware truncation. The
  // inline `.slice(0, 60)` could split a surrogate pair when value
  // contains emoji at the 60-char boundary, producing orphan high
  // surrogate that gets persisted to SQLite undo_entries.changeset_json
  // as broken UTF-16. See lib/utils.ts:slicePreview doc-block.
  const valuePreview =
    value === null
      ? '(空)'
      : slicePreview(String(value).replace(/\s+/g, ' '), 60);
  const desc = `${target.name}!${address.toUpperCase()} = ${valuePreview}`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: desc,
      ops: [{ tabId: tab.id, type: 'binary_replace', before, after, description: desc }],
    },
    summary: desc,
  };
}

async function doExcelSetRange(block: ToolUseBlock, ctx: DispatchContext): Promise<DispatchResult> {
  const { tabId, sheet, startAddress, values } = block.input as {
    tabId?: string;
    sheet?: string;
    startAddress: string;
    values: Array<Array<string | number | boolean | null>>;
  };
  // R291 — strict typeof for startAddress; same rationale as doExcelSetCell
  // sibling. `parseA1(startAddress.toUpperCase())` throws on non-string.
  if (typeof startAddress !== 'string' || startAddress.length === 0 || !Array.isArray(values))
    return { ok: false, error: 'missing_required_fields: startAddress, values' };
  // R314 — also reject an effectively-empty matrix. The schema check above
  // accepts `values=[]`, `values=[[]]`, `values=[[], [], []]` (any 2D shape
  // where no cell actually exists). The downstream walk then silently
  // produces a no-op changeset: grid-grow `while (cells.length < start.r +
  // 0)` doesn't iterate, write-loop `for (r < 0)` doesn't iterate,
  // serializeXlsx round-trips model unchanged, and dispatcher returns
  // `ok: true` with a `binary_replace` op carrying byte-equivalent (but
  // ref-different) before/after. The user sees a confusing PendingChange
  // card titled「批量設定 Sheet1!A1:A0（0×0）」 — the `A1:A0` reversed
  // range is the cosmetic tell — and Apply still flips `dirty: true` +
  // pushes a row onto the persistent undo stack, wasting an undo slot on
  // an op that has no observable effect. R302 closed the `ops.length === 0`
  // path for read-only tools (collapse to synth tool_result); this is the
  // sibling guard at the input layer for batch writes where ops > 0 but
  // every op is on a zero-cell matrix. Reject upfront with a typed error
  // AI can self-correct from (e.g., recheck whether the planning step
  // forgot to fill the values array).
  const hasAnyCell = values.some((row) => Array.isArray(row) && row.length > 0);
  if (!hasAnyCell) {
    return {
      ok: false,
      error: 'empty_values: 2D matrix has no cells (outer or every inner row is empty)',
    };
  }
  const got = expectXlsxTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.data;

  const model = parseXlsx(before);
  // R326 — empty `sheet` falls back to activeSheet; same rationale as
  // doExcelSetCell sibling above.
  const sheetName = sheet && sheet.length > 0 ? sheet : model.activeSheet;
  const target = model.sheets.find((s) => s.name === sheetName);
  if (!target) return { ok: false, error: `sheet_not_found: ${sheetName}` };
  const start = parseA1(startAddress.toUpperCase());
  if (!start) return { ok: false, error: `bad_a1_address: ${startAddress}` };
  // R301 — same negative-bound guard as doExcelSetCell sibling. `A0` /
  // `B0` etc. pass parseA1 with r=-1; without this check the grid-grow
  // `while (target.cells.length < start.r + values.length)` would
  // iterate while `length < -1+N`, then the per-cell `target.cells[
  // start.r + r][start.c + c]` walks off the negative axis and the R269
  // destructure throws on undefined. Fail-fast with bad_a1_address keeps
  // the error message readable for AI.
  if (start.r < 0 || start.c < 0) {
    return { ok: false, error: `bad_a1_address: ${startAddress} (row must be ≥ 1)` };
  }
  // R327 — upper-bound guard. `doExcelSetRange` is the only dispatcher path
  // that proactively grows the cells grid (see the `while (target.cells.length
  // < neededRows)` loop below); without an upper bound, an AI call with
  // `startAddress: "A1000000"` or a `values` array with 16K+ entries per
  // row drives that loop into millions of `{text: ''}` object allocations.
  // Each empty row of `target.colCount` cells is small (~100 bytes), but
  // 1M rows × even 10 cols is ~50 MB of in-memory model state created
  // BEFORE the actual write loop runs. Pathological combinations (1M rows ×
  // 16K cols) hit hundreds of MB and either freeze the UI for seconds /
  // tens-of-seconds or OOM the renderer outright (Electron auto-restarts
  // the crashed renderer and the user loses every unsaved tab).
  //
  // `doExcelSetCell` is safe — it does fail-fast `address_out_of_range`
  // against the current sheet bounds, never grows. Only the batch path
  // needs this guard.
  //
  // Excel 規範: 1,048,576 rows × 16,384 cols (xlsx max). Beyond these the
  // output would be an invalid xlsx anyway; rejecting at the dispatcher
  // gives AI a typed error string it can self-correct from (most likely
  // path is「AI thought rows are 1-based-positive-integer but typo'd a
  // zero too many」 — bad_a1_address style recovery). Total cell count
  // not separately bounded — within row/col caps the realistic worst-
  // case is ~17B cells, which is its own absurdity; the rows × cols
  // dimensional caps are the meaningful boundary.
  const EXCEL_MAX_ROWS = 1_048_576;
  const EXCEL_MAX_COLS = 16_384;
  // R328 — compute maxRowLen via reduce (not Math.max(...spread)). For very
  // large `values` (e.g., 100K rows × 3 cols, a legal pattern AI might emit
  // for a time-series fill), `Math.max(...values.map((row) => row.length))`
  // spreads N elements as function arguments and hits V8's argument-count
  // limit (~65K) — throws `RangeError: Maximum call stack size exceeded`
  // BEFORE the R327 bounds check below can reject. The throw propagates as
  // `dispatch_threw: Maximum call stack size exceeded`, AI sees an
  // implementation-detail error for what looks like a legal Excel range,
  // can't recover. reduce has no stack-spread, handles any array size.
  // Also reused at the grid-grow loop below to avoid recomputing
  // (previously a second `Math.max(...spread)` site with the same hazard).
  // Defensive `Array.isArray(row)` because R298's per-cell validation runs
  // LATER — a non-array row in `values` would crash on `.length` if we
  // didn't guard; this check turns it into 0 (no contribution to maxRowLen)
  // and R298 catches the actual error with a clean message a few lines down.
  const maxRowLen = values.reduce<number>(
    (m, row) => Math.max(m, Array.isArray(row) ? row.length : 0),
    0,
  );
  const neededRows = start.r + values.length;
  const neededCols = start.c + maxRowLen;
  if (neededRows > EXCEL_MAX_ROWS) {
    return {
      ok: false,
      error: `range_out_of_bounds: would write to row ${neededRows} (Excel max is ${EXCEL_MAX_ROWS}); check startAddress + values.length`,
    };
  }
  if (neededCols > EXCEL_MAX_COLS) {
    return {
      ok: false,
      error: `range_out_of_bounds: would write to column ${neededCols} (Excel max is ${EXCEL_MAX_COLS}); check startAddress + max row length`,
    };
  }

  // R298 — per-cell value primitive guard. Mirror of R297 (doExcelSetCell)
  // for the batch path. Tool schema declares values as `Array<Array<string
  // | number | boolean | null>>` but unsafe `as` cast doesn't enforce at
  // runtime. Without this, an AI call with `undefined` / object / array
  // anywhere in the 2D matrix silently coerces via `String(v)` to literal
  // `"undefined"` / `"[object Object]"` / `"1,2,3"` at the corresponding
  // cell — and `set_range` can blast dozens of such cells per call. Two-
  // pass approach: validate all cells first, fail-fast atomically if any
  // is non-primitive (with row/col coords for AI to fix); only proceed
  // to the mutation loop if every cell is clean. O(N×M) double-walk is
  // negligible (typical 10×10 = 200 checks).
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (!Array.isArray(row)) {
      return { ok: false, error: `bad_values: row ${r} is not an array` };
    }
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      const t = typeof v;
      if (v !== null && t !== 'string' && t !== 'number' && t !== 'boolean') {
        return {
          ok: false,
          error: `bad_cell_value at row ${r}, col ${c}: must be string, number, boolean, or null`,
        };
      }
      // R352 — also reject non-finite numbers per-cell. Mirror of the
      // single-cell guard in doExcelSetCell above; see that doc-block for
      // the full「NaN literal text in cells」rationale. Batch path is
      // even more impactful — one call can blast dozens of NaN cells if
      // AI's calculation produces non-finite results for multiple
      // entries in a derivative formula (e.g., dividing column B by a
      // column C that has zeros).
      if (t === 'number' && !Number.isFinite(v as number)) {
        return {
          ok: false,
          error: `bad_cell_value at row ${r}, col ${c}: numeric values must be finite (got NaN / Infinity / -Infinity); use null for empty cells or a string label like "—" for unrepresentable results`,
        };
      }
    }
  }

  // Grow grid if needed. `neededRows` / `neededCols` computed once at R327
  // above (via reduce, not Math.max-spread — see doc-block there).
  while (target.cells.length < neededRows) {
    target.cells.push(new Array(target.colCount).fill(null).map(() => ({ text: '' })));
  }
  if (neededCols > target.colCount) {
    for (const row of target.cells) {
      while (row.length < neededCols) row.push({ text: '' });
    }
    target.colCount = neededCols;
  }
  target.rowCount = Math.max(target.rowCount, neededRows);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      // R269 — same prev-style preservation as doExcelSetCell. excel_set_range
      // is even more destructive without this: a single AI call can blast
      // formatting on dozens of cells in one click. See sibling at the
      // single-cell path for the full bug shape and writeCell parity.
      const prevCell = target.cells[start.r + r][start.c + c];
      const newText = v === null ? '' : String(v);
      // R303 — same writeCell two-branch parity as doExcelSetCell. Batch
      // writes can include formula cells (e.g., AI generating a column
      // of `=B*1.1` derivative formulas); without this, those cells
      // land as literal strings, not computable formulas.
      const { formula: _dropF, ...restCell } = prevCell;
      void _dropF;
      target.cells[start.r + r][start.c + c] = isFormulaSource(newText)
        ? { ...restCell, text: newText, formula: newText }
        : { ...restCell, text: newText };
    }
  }
  const after = serializeXlsx(model, before);
  // R333 — use `maxRowLen` (computed once at R328 via reduce) for the end-
  // address + dimension summary instead of `values[0].length`. Jagged 2D
  // arrays (rows of differing widths — e.g., `[['a'], ['b','c','d']]`) are
  // legal input: R298's per-cell validation walks `values[r][c]` row-by-
  // row, and the write loop at line 663-682 likewise iterates each row's
  // own length, so cells get written across the union of all column
  // indices — and the grid-grow loop at line 632-642 grew the sheet to
  // `start.c + maxRowLen` columns. The OLD description sampled only
  // `values[0].length`, so a jagged write produced a misleading summary
  // like `Sheet1!A1:A2（2×1）` while the actual write spanned `A1:C2`
  // (3 columns wide because row 1 has 3 entries). The undo toast / diff
  // preview / inflight Apply card all read this string — the user clicks
  // Apply expecting a 2×1 footprint and 2×3 worth of cells change,
  // including the right-most ones in row 0 that the description didn't
  // hint at. Same `colIndexToLetter(start.c + maxRowLen - 1)` end formula
  // the grid-grow path already used; the (rows × cols) pair in parens is
  // now `values.length × maxRowLen` for the same reason. Empty-row
  // safety: if maxRowLen is 0 (every row is empty — already rejected by
  // R298 but defensive), `colIndexToLetter(start.c - 1)` would be ≤ start
  // column; the description would degrade to `A1:` which is harmless
  // since the write loop also produces zero writes.
  const endAddr = `${colIndexToLetter(start.c + Math.max(maxRowLen, 1) - 1)}${start.r + values.length}`;
  const desc = `${target.name}!${startAddress.toUpperCase()}:${endAddr}（${values.length}×${maxRowLen}）`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: `批量設定 ${desc}`,
      ops: [{ tabId: tab.id, type: 'binary_replace', before, after, description: `批量設定 ${desc}` }],
    },
    summary: `批量設定 ${desc}`,
  };
}

// ── Word ops ─────────────────────────────────────────────────────────────

function expectDocxTab(ctx: DispatchContext, tabId?: string): { ok: true; tab: DocxTab } | { ok: false; error: string } {
  const tab = findTab(ctx, tabId);
  // R310 — see expectMdTab doc-block.
  if (!tab) {
    return { ok: false, error: tabId ? `tab_not_found: ${tabId}` : 'no_active_tab' };
  }
  // R399 — see expectMdTab sibling for the tabId-echo rationale.
  if (tab.type !== 'docx')
    return { ok: false, error: `tab_type_mismatch: expected docx, got ${tab.type} (tabId: ${tab.id})` };
  return { ok: true, tab };
}

async function doWordReplaceParagraph(block: ToolUseBlock, ctx: DispatchContext): Promise<DispatchResult> {
  const { tabId, paragraphIndex, text } = block.input as {
    tabId?: string;
    paragraphIndex: number;
    text: string;
  };
  // R295 — strict integer check; see doMdInsertAt sibling.
  // `model.blocks[NaN] = {...}` silently sets a "NaN" property on the
  // array and the serializer ignores it, so AI thinks the replace
  // succeeded while docx content is unchanged.
  if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0 || typeof text !== 'string')
    return { ok: false, error: 'missing_required_fields: paragraphIndex, text (paragraphIndex must be a non-negative integer)' };
  const got = expectDocxTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.data;

  const model = await parseDocx(before);
  if (paragraphIndex < 0 || paragraphIndex >= model.blocks.length)
    return { ok: false, error: `paragraph_out_of_range: ${paragraphIndex}` };
  // R359 — reject text replacement on non-text blocks (table / image). The
  // serializer at docx-adapter.ts:258-263 dispatches by kind:
  //   • kind === 'table'           → buildTable(b) reads b.rows; ignores b.text
  //   • kind === 'image' && image  → buildImageParagraph(b) reads b.image; ignores b.text
  //   • else                       → buildParagraph(b) uses b.text / b.runs
  // So `{...model.blocks[N], text}` on a table block updates the JS
  // object's text field, but the serializer never reads it — the resulting
  // docx is byte-equivalent to before, and the dispatcher's ok:true +
  // BinaryReplaceOp with `before === after` is the same silent-no-op
  // shape R314 (empty values) / R349 (oldText mismatch) closed for other
  // paths. User sees「替換第 N 段」 in the PendingChange card, clicks
  // Apply, nothing visually changes. Worse: workspace.dirty flips true
  // (the BinaryReplaceOp is still recorded), so the next close prompts
  // 「未儲存」on a no-op edit.
  //
  // Fail-fast with a typed error that names what AI can do instead. For
  // table cells the user has to edit in-place via DocxEditor (no AI tool
  // for cells today); for images the only "edit" is replace via the
  // editor's right-click. AI knowing this saves a retry round-trip.
  const targetBlock = model.blocks[paragraphIndex];
  if (targetBlock.kind === 'table') {
    return {
      ok: false,
      error: `block_kind_mismatch: paragraph ${paragraphIndex} is a table block — word_replace_paragraph only works on text blocks (paragraph / heading1-6 / bullet / numbered). 表格 cell 內容目前需請使用者在 Word 編輯器內手動修改。`,
    };
  }
  if (targetBlock.kind === 'image') {
    return {
      ok: false,
      error: `block_kind_mismatch: paragraph ${paragraphIndex} is an image block — word_replace_paragraph only works on text blocks. 圖片區塊目前需請使用者在 Word 編輯器內手動更換。`,
    };
  }
  // R268 — drop runs[] when overriding text. The serializer at
  // docx-adapter.ts:519 picks `runs[]` over `text` when both are present;
  // a naive { ...prev, text } spread leaves OLD runs intact, so the
  // serializer emits OLD text from runs and AI's replacement is silently
  // dropped at the byte layer (next parse round-trips back to OLD text).
  // Same invariant DocxEditor.updateBlock already enforces for its
  // text-only patch path: a plain-text replacement can't carry per-run
  // styling so we clear runs to keep text authoritative.
  const updated = { ...targetBlock, text };
  delete updated.runs;
  model.blocks[paragraphIndex] = updated;
  const after = await serializeDocx(model);
  // Human-readable description: switch the AI's 0-based paragraph index back
  // to 1-based for display, matching the markdown `第 ${line} 行` / pptx
  // `投影片 ${slideIndex + 1}` convention. Otherwise the Apply system message
  // reads "Applied: file.docx：替換第 0 段" which looks like a typo.
  const desc = `${tab.name}：替換第 ${paragraphIndex + 1} 段`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: desc,
      ops: [{ tabId: tab.id, type: 'binary_replace', before, after, description: desc }],
    },
    summary: desc,
  };
}

async function doWordInsertParagraph(block: ToolUseBlock, ctx: DispatchContext): Promise<DispatchResult> {
  // R355 — kind union expanded to heading4-6. The docx model + serializer
  // (docx-adapter.ts:39-46 + 545-563) have always supported heading1-6;
  // the schema enum and this destructure cast were capped at heading3,
  // hiding three valid kinds. The dispatcher uses `kind` directly in
  // `model.blocks.splice(...{ kind: kind ?? 'paragraph' })`, so widening
  // the cast is enough to make AI's `kind: 'heading4'` calls produce
  // proper docx h4 paragraphs. Same destructure pattern as bullet /
  // numbered; no extra branching needed.
  const { tabId, paragraphIndex, text, kind } = block.input as {
    tabId?: string;
    paragraphIndex?: number;
    text: string;
    kind?:
      | 'paragraph'
      | 'heading1'
      | 'heading2'
      | 'heading3'
      | 'heading4'
      | 'heading5'
      | 'heading6'
      | 'bullet'
      | 'numbered';
  };
  if (typeof text !== 'string')
    return { ok: false, error: 'missing_required_fields: text' };
  // R295 — when paragraphIndex is provided, require a non-negative integer.
  // `Math.min(NaN, N) === NaN` → `Math.max(0, NaN) === NaN` → splice treats
  // NaN as 0 and silently inserts at position 0 regardless of AI's intent.
  // Optional path is preserved: undefined means "append at end" per the
  // tool schema, which the line below already handles.
  if (paragraphIndex !== undefined && (!Number.isInteger(paragraphIndex) || paragraphIndex < 0))
    return { ok: false, error: 'paragraphIndex must be a non-negative integer or omitted' };
  // R397 — runtime guard for `kind`. The `as` cast at the destructure is
  // TypeScript-only; AI can emit any string at runtime and it bypasses the
  // type system. If AI passes an unrecognized kind (typo「heading7」, an
  // imagined「subheading」, a future-version「callout」 a downstream tool
  // schema hasn't taught it about), the insert proceeds:
  //   • `model.blocks.splice(idx, 0, { kind: 'heading7', text })` lands a
  //     block with structurally invalid kind
  //   • `buildParagraph`'s switch (docx-adapter.ts:597-632) has no case
  //     for 'heading7' → falls to `default: break;` → opts stays
  //     baseOpts (no heading, no bullet, no numbered) → emits a PLAIN
  //     paragraph
  //   • Word doc shows regular body text where AI intended a heading
  //   • dispatcher's `desc = `在第 ${idx + 1} 段插入${kind ?? '段落'}`` (line
  //     1068) still embeds 'heading7' in the PendingChange card label,
  //     so the user reads「在第 N 段插入 heading7」 but the actual
  //     output is plain text — confusing, and the silent degradation
  //     compounds.
  // Same R291 / R295 / R297 / R298 fail-fast idiom for runtime input
  // validation: reject with a typed error AI can self-correct from
  // (retry with a valid kind from the listed union) rather than commit
  // a silently-degraded write. Listing the valid kinds in the error
  // message matches doExcelSetCell's「must be string, number, boolean,
  // or null」 pattern — AI gets the recovery path inline.
  const VALID_KINDS = new Set([
    'paragraph',
    'heading1',
    'heading2',
    'heading3',
    'heading4',
    'heading5',
    'heading6',
    'bullet',
    'numbered',
  ]);
  if (kind !== undefined && !VALID_KINDS.has(kind)) {
    return {
      ok: false,
      error: `bad_kind: "${String(kind)}" — kind must be one of: ${[...VALID_KINDS].join(', ')}, or omit for plain paragraph`,
    };
  }
  const got = expectDocxTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.data;

  const model = await parseDocx(before);
  const idx = paragraphIndex === undefined ? model.blocks.length : Math.max(0, Math.min(paragraphIndex, model.blocks.length));
  model.blocks.splice(idx, 0, {
    id: `inserted-${Date.now()}`,
    kind: kind ?? 'paragraph',
    text,
  });
  const after = await serializeDocx(model);
  // 1-based for display (see comment in doWordReplaceParagraph). For an
  // append (paragraphIndex === undefined → idx === blocks.length) we still
  // show the human-friendly position the new paragraph will land at.
  const desc = `${tab.name}：在第 ${idx + 1} 段插入${kind ?? '段落'}`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: desc,
      ops: [{ tabId: tab.id, type: 'binary_replace', before, after, description: desc }],
    },
    summary: desc,
  };
}

// ── PowerPoint ops ───────────────────────────────────────────────────────

function expectPptxTab(ctx: DispatchContext, tabId?: string): { ok: true; tab: PptxTab } | { ok: false; error: string } {
  const tab = findTab(ctx, tabId);
  // R310 — see expectMdTab doc-block.
  if (!tab) {
    return { ok: false, error: tabId ? `tab_not_found: ${tabId}` : 'no_active_tab' };
  }
  if (tab.type !== 'pptx')
    // R399 — see expectMdTab sibling for the tabId-echo rationale.
    return { ok: false, error: `tab_type_mismatch: expected pptx, got ${tab.type} (tabId: ${tab.id})` };
  return { ok: true, tab };
}

async function doPptxReplaceText(block: ToolUseBlock, ctx: DispatchContext): Promise<DispatchResult> {
  const { tabId, slideIndex, runIndex, oldText, newText } = block.input as {
    tabId?: string;
    slideIndex: number;
    runIndex?: number;
    oldText?: string;
    newText: string;
  };
  // R295 — strict integer for slideIndex. `model.slides.find((s) => s.index
  // === NaN)` never matches (NaN !== NaN), so the function returns the
  // misleading `slide_not_found` error instead of a clean
  // `missing_required_fields`. Same fix shape as doMdInsertAt /
  // doWordReplaceParagraph siblings.
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || typeof newText !== 'string')
    return { ok: false, error: 'missing_required_fields: slideIndex, newText (slideIndex must be a non-negative integer)' };
  const got = expectPptxTab(ctx, tabId);
  if (!got.ok) return { ok: false, error: got.error };
  const tab = got.tab;
  const before = tab.data;
  if (before.byteLength === 0)
    // R356 — accurate recovery hint. Was「open an existing pptx file first」
    // which implies the user must have a .pptx on disk to load — wrong on
    // two counts:
    //   (a) pptx-adapter.ts:176-220 ships `createBlankPptx()` that
    //       constructs a minimal one-slide deck from scratch, so the user
    //       doesn't need any existing file.
    //   (b) PptxEditor.tsx auto-runs createBlankPptx the moment the tab
    //       becomes active (its mount effect serializes the blank deck
    //       back into tab.data), so the actual user-visible recovery is
    //       just「點一下這個 pptx tab 讓 PptxEditor 載入」 — no file
    //       picker, no upload, no manual create.
    // AI used to repeat the「open an existing pptx file first」line back
    // to the user (since the error became the basis of its「我目前無法操
    // 作這個 tab」reply), sending them on a fake「先去找一個 .pptx 來開」
    // errand. The correct action is to activate the tab — explain that
    // directly. Same shape as R343 / R344 / R353 / R354 / R355 fixes:
    // align the surface AI reads with the actual recovery path.
    return {
      ok: false,
      error: 'pptx_empty: this pptx tab has not been initialised yet (0 bytes). Ask the user to click the tab once so PptxEditor mounts; it auto-builds a blank one-slide deck on first activation. Then retry pptx_replace_text.',
    };

  const model = await parsePptx(before);
  const slide = model.slides.find((s) => s.index === slideIndex);
  if (!slide) return { ok: false, error: `slide_not_found: ${slideIndex}` };

  let target = -1;
  // R349 — track WHICH lookup method located the target so the write step
  // can decide between full-replace and substring-replace consistently.
  // Previous code used `typeof oldText === 'string' && run.text !== oldText`
  // at the write step — but that condition is true even when target was
  // located by `runIndex` AND oldText was supplied as a hint that doesn't
  // match. Concrete failure:
  //   • AI calls pptx_replace_text({runIndex: 2, oldText: "舊文字",
  //     newText: "新文字"}) — both fields provided.
  //   • Lookup uses runIndex first, finds run 2 (whose text happens to
  //     differ from oldText, e.g. AI's planner had stale info).
  //   • Write step: `typeof "舊文字" === 'string' && run.text !== "舊文字"`
  //     → true → substring branch → `run.text.replace("舊文字", "新文字")`
  //     finds no「舊文字」 substring in run.text → returns unchanged.
  //   • model.runs[target].text re-assigned to identical string.
  //   • serializePptx produces byte-equivalent output.
  //   • dispatcher returns ok:true; user Applies「投影片 N 文字替換」card
  //     and nothing visually changes — same silent-no-op shape R314
  //     (empty values) and R313 (empty oldText) closed for their respective
  //     paths.
  // The fix is to use the method that found the target as the source of
  // truth for the write: runIndex-found targets always do full-replace
  // (AI's intent is "replace whole text at this run"), oldText-found
  // targets do substring-replace only if exact-match failed.
  let foundBy: 'runIndex' | 'oldTextExact' | 'oldTextSubstring' | null = null;
  // R295 — strict integer for runIndex. NaN runIndex would `findIndex`
  // every run via `r.runIndex === NaN` (NaN !== NaN) → returns -1 → falls
  // through to "run_not_found". Worse, fractional runIndex (e.g. 0.5)
  // also never matches integer run indices. Same idiom as slideIndex
  // sibling above.
  if (Number.isInteger(runIndex) && (runIndex as number) >= 0) {
    target = slide.runs.findIndex((r) => r.runIndex === runIndex);
    if (target >= 0) foundBy = 'runIndex';
  } else if (typeof oldText === 'string' && oldText.length > 0) {
    // R313 — require oldText to be non-empty. Without the length check, an AI
    // call that omits runIndex AND passes oldText="" (easy regression when
    // AI's planner hallucinates required fields as empty strings, or a tool
    // schema-mocking test) walks into a silent corruption:
    //   1. `r.text === ''` typically misses (runs have content).
    //   2. Fuzzy fallback `r.text.includes('')` is ALWAYS true — empty string
    //      is a substring of every string per the JS spec — so findIndex
    //      returns 0 (first run).
    //   3. The substring-replace branch below runs
    //      `slide.runs[0].text.replace('', newText)`, and
    //      `String.prototype.replace` treats the empty pattern as "first
    //      zero-length match", which is position 0 — i.e., prepends newText.
    //   4. dispatcher returns ok:true; user Applies; the first run's text
    //      silently gains a `newText` prefix the AI never intended to
    //      insert. The Apply DiffPreview summary reads "投影片 1 文字替
    //      換" — looks like a legitimate replace, doesn't surface the
    //      prepend behavior.
    // Reject empty oldText here so the fallback only fires for a real
    // substring search. AI sees `run_not_found` and re-checks outline /
    // tries a different parameter. Same R291/R295 fail-fast idiom for
    // dispatcher input validation.
    target = slide.runs.findIndex((r) => r.text === oldText);
    if (target >= 0) {
      foundBy = 'oldTextExact';
    } else {
      // Fuzzy: find runs containing oldText as substring.
      target = slide.runs.findIndex((r) => r.text.includes(oldText));
      if (target >= 0) foundBy = 'oldTextSubstring';
    }
  }
  if (target < 0)
    return {
      ok: false,
      error: 'run_not_found: provide a valid non-negative runIndex or a non-empty oldText that matches a run on this slide',
    };

  // R349 — write decision is driven by `foundBy`, not by oldText presence:
  //   • runIndex      → full replace (AI's intent at this run)
  //   • oldTextExact  → full replace (run.text === oldText; substring would
  //                     produce an identical result anyway)
  //   • oldTextSubstring → substring replace (the only path where
  //                     `run.text.replace(oldText, newText)` is the actual
  //                     user intent — replace a specific fragment of a
  //                     longer run)
  // The old condition `oldText is string && run.text !== oldText` ALSO
  // fired for the `runIndex` path when oldText was supplied as a hint
  // that didn't match — substring-replace would no-op, producing a silent
  // success.
  if (foundBy === 'oldTextSubstring') {
    slide.runs[target] = {
      ...slide.runs[target],
      // Non-null assert is safe: we only reach this branch if foundBy ===
      // 'oldTextSubstring', which requires the else-if's `typeof oldText
      // === 'string' && oldText.length > 0` to have passed.
      text: slide.runs[target].text.replace(oldText!, newText),
    };
  } else {
    slide.runs[target] = { ...slide.runs[target], text: newText };
  }
  const after = await serializePptx(model, before);
  const desc = `${tab.name}：投影片 ${slideIndex + 1} 文字替換`;
  return {
    ok: true,
    changeset: {
      id: uuid(),
      origin: 'ai',
      createdAt: new Date().toISOString(),
      description: desc,
      ops: [{ tabId: tab.id, type: 'binary_replace', before, after, description: desc }],
    },
    summary: desc,
  };
}

// ── helpers ──────────────────────────────────────────────────────────────

/**
 * Replace a Markdown section starting at the first heading whose text matches
 * `heading` (any level), continuing until the next heading of equal-or-higher
 * level (i.e. fewer or equal `#`s). Returns null if no match.
 */
export function replaceMdSection(
  source: string,
  heading: string,
  newContent: string,
): string | null {
  const lines = source.split('\n');
  let startIdx = -1;
  let level = 0;
  // R329 — mirror `MarkdownEditor.parseOutline`'s regex exactly so ATX
  // closing `#` markers (e.g., `# Title #`, `## Section ##`) are stripped
  // from the captured heading text. Without this the dispatcher saw
  // "Title #" while the editor's outline pane (and the user, who reads
  // it) saw clean "Title"; AI calling `md_replace_section({heading:
  // "Title"})` produced a confusing heading_not_found despite the
  // section visibly existing. CommonMark spec defines heading text as
  // excluding the optional trailing `#` sequence, so this is the
  // canonical interpretation — `(.+?)\s*#*\s*$` is non-greedy + absorbs
  // any number of trailing `#` chars with surrounding whitespace.
  const headingRe = /^(#{1,6})\s+(.+?)\s*#*\s*$/;
  // R387 — accept both backtick (```) and tilde (~~~) fence forms. CommonMark
  // §4.5 allows either char (3+ of one kind) to open a fenced code block;
  // R315's original regex only matched backticks, leaving tilde-fenced docs
  // unprotected. A doc like
  //   ## Examples
  //   ~~~bash
  //   # install dependencies
  //   ~~~
  //   ## Deployment
  // would let the AI call md_replace_section({heading:"install dependencies"})
  // and the dispatcher would happily replace from the fake heading through
  // the rest of the doc — same data-loss shape R315 was supposed to close,
  // just with a different fence flavor. Tilde fences are uncommon but real
  // (Pandoc/Hugo default to them in some configs, and authors mixing
  // languages use ~~~ to avoid backtick-conflict with bash `cmd substitution`).
  // Cross-line: orchestrator.ts:1092 buildOutline and
  // MarkdownEditor.tsx:77 parseOutline get the same extension so all three
  // outline-scanning surfaces stay consistent (the same three-way invariant
  // R342 / R315 already enforce). We treat ANY fence line as a toggle —
  // strict CommonMark only closes a fence with the same char + ≥ same
  // length, but tracking that adds complexity for a degenerate edge case
  // (intentionally mismatched fences in a single doc), and our previous
  // pre-R387 backtick-only behavior was already lossy in the same way.
  const fenceRe = /^\s*(?:```|~~~)/;

  // R315 — track ``` fences so `# comment` inside a fenced code block isn't
  // matched as a real ATX heading. Without this guard, a doc like
  //   ## Examples
  //   ```bash
  //   # install dependencies
  //   ```
  //   ## Deployment
  // would let AI call `md_replace_section({heading:"install dependencies"})`
  // (the AI sees this fake heading via buildOutline's same fence-blind walk,
  // R315 sibling fix in orchestrator.ts) and the dispatcher would happily
  // find the fake heading at the `# install dependencies` line, scan
  // forward for the next level-≤-1 heading (`## Deployment` is level 2,
  // doesn't trigger), reach EOF, and replace EVERYTHING from the fake
  // heading to end-of-doc with newContent — including the closing ``` ` ``,
  // the entire `## Deployment` section, and any trailing content. Severe
  // data loss; Ctrl+Z recovers but only if the user notices before
  // continuing.
  //
  // MarkdownEditor's parseOutline at MarkdownEditor.tsx:70-93 already
  // tracks fences correctly; the AI path was the lone fence-blind surface.
  // Single pass tracks fence state across both phases (start-search and
  // end-search) so a fenced block between startIdx and the next real
  // sibling heading doesn't prematurely terminate the section either.
  let inFence = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (fenceRe.test(line)) {
      inFence = !inFence;
      continue;
    }
    if (inFence) continue;
    const m = line.match(headingRe);
    if (!m) continue;
    if (startIdx === -1) {
      if (m[2].trim() === heading.trim()) {
        startIdx = i;
        level = m[1].length;
      }
    } else if (m[1].length <= level) {
      // First subsequent heading at same-or-shallower level marks the
      // section boundary.
      // (Found; break below sets endIdx.)
      // We assign endIdx here directly and break to avoid restructuring
      // the loop into two halves.
      const head = lines.slice(0, startIdx).join('\n');
      const tail = lines.slice(i).join('\n');
      const middle = newContent.endsWith('\n') ? newContent.slice(0, -1) : newContent;
      return [head, middle, tail].filter((s) => s.length > 0).join('\n');
    }
  }
  if (startIdx === -1) return null;

  // Reached EOF without finding a sibling heading — replace to end of doc.
  const head = lines.slice(0, startIdx).join('\n');
  const middle = newContent.endsWith('\n') ? newContent.slice(0, -1) : newContent;
  return [head, middle].filter((s) => s.length > 0).join('\n');
}
