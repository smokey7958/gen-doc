/**
 * Cross-editor Find & Replace dialog.
 *
 * Each editor flattens its document model into an ordered list of
 * `SearchSegment`s (id + text + optional label) and hands them to this
 * dialog along with an `onUpdateSegment(id, newText)` callback. The
 * dialog walks the segments to compute matches, navigates between them,
 * and rewrites segment text on Replace / Replace All.
 *
 * Match position is tracked as a (segmentIndex, offset) pair into the
 * concatenated stream of segment texts. Replacements are applied per
 * segment so the host editor's existing setters (which themselves drive
 * undo, serialization, etc.) stay in charge.
 *
 * Match modes: literal (default), regex (with capture-group expansion),
 * and whole-word (decorates the pattern with `\b` boundaries). Replacement
 * preserves no formatting — literal/regex substring replace inside the
 * segment's plain text. Per-run formatting preservation is Phase D.
 */

import { useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown, ChevronUp, Replace, Search, X } from 'lucide-react';
import { cn, sliceCodePoints } from '../lib/utils';
import { useT } from '../lib/i18n';

export interface SearchSegment {
  /** Stable id the host editor uses to address this text fragment. */
  id: string;
  /** Plain-text content of the segment. */
  text: string;
  /** Human label for the match list ("A1" / "Slide 2 / Box 1"). Optional. */
  label?: string;
}

interface Match {
  segmentId: string;
  segmentIndex: number;
  /** Char offset into the segment's text where the match starts. */
  offset: number;
  length: number;
  label?: string;
}

interface Props {
  open: boolean;
  onClose: () => void;
  segments: SearchSegment[];
  onUpdateSegment: (id: string, newText: string) => void;
  /** Optional hook called when the highlighted match changes — lets the
   * host scroll the underlying segment into view / focus its input. */
  onLocateSegment?: (id: string) => void;
  title?: string;
  /** Bumped by the host every time Ctrl+F is pressed. Re-focuses + selects
   *  the query field even when `open` was already true — matches VS Code /
   *  Chrome convention where Ctrl+F always brings focus back to the
   *  search box (and Esc closes). Without this, repeat Ctrl+F just
   *  toggled the dialog shut, which surprises users who clicked into
   *  the document to navigate a match and then hit Ctrl+F to come back. */
  focusNonce?: number;
}

/** Match-mode toggles. `regex` and `wholeWord` are mutually orthogonal —
 *  whole-word with regex still applies `\b` boundaries to the user's
 *  pattern, matching VS Code / Word behaviour. */
interface MatchOptions {
  caseSensitive: boolean;
  regex: boolean;
  wholeWord: boolean;
}

/** Escape every RegExp metachar so a plain-string query becomes a literal
 *  regex source. Used when `regex` is off but `wholeWord` is on. */
function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Read the text the user has selected in whatever element currently has
 * focus. Bridges the two distinct selection APIs:
 *   • `<input>` / `<textarea>` expose selection via selectionStart/End on
 *     the element — `window.getSelection()` does NOT cover these in Chromium
 *   • Everything else (contentEditable, plain text nodes) uses
 *     `window.getSelection()`
 *
 * Returns the first line only — the query field is single-line, so feeding
 * a multi-line selection in would look broken; this matches VS Code's
 * behaviour ("find next occurrence of what I just selected"). Capped at
 * 200 chars so an accidental select-all doesn't paste an entire document
 * into the query field.
 */
function readActiveSelectionText(): string {
  const ae = document.activeElement;
  let raw = '';
  if (ae instanceof HTMLInputElement || ae instanceof HTMLTextAreaElement) {
    const start = ae.selectionStart ?? 0;
    const end = ae.selectionEnd ?? 0;
    if (end > start) raw = ae.value.slice(start, end);
  } else {
    const sel = window.getSelection();
    if (sel) raw = sel.toString();
  }
  if (!raw) return '';
  const firstLine = raw.split(/\r?\n/, 1)[0]?.trim() ?? '';
  if (!firstLine) return '';
  // R383 — code-point-aware slice, no ellipsis. This output becomes the
  // dialog's initial search query — any extra suffix would be searched
  // FOR (find `searched-text…` won't match the source emoji that got
  // surrogate-split). Surrogate split is the realistic risk: a user
  // selects a sentence that contains emoji at the 200-code-unit
  // boundary, opens find, and the auto-populated query has an orphan
  // high surrogate that doesn't match anything in the document. Same
  // R382 helper but the no-suffix flavor.
  return sliceCodePoints(firstLine, 200);
}

/**
 * Build a single global RegExp for the given query + options. Returns null
 * when the user has typed an invalid regex so the caller can show an
 * inline error rather than crashing the dialog.
 */
function compilePattern(query: string, opts: MatchOptions): RegExp | null {
  if (!query) return null;
  let source = opts.regex ? query : escapeRegExp(query);
  if (opts.wholeWord) source = `\\b(?:${source})\\b`;
  const flags = opts.caseSensitive ? 'g' : 'gi';
  try {
    return new RegExp(source, flags);
  } catch {
    return null;
  }
}

/**
 * Walk segments, collect every occurrence of `query`. Honors all three
 * match-option toggles. Empty query yields no matches so the user doesn't
 * accidentally "find" every gap.
 *
 * Matches with zero length (e.g., `(?=)`) advance the cursor by one to
 * avoid an infinite loop — same defence as `String.matchAll`.
 */
function collectMatches(
  segments: SearchSegment[],
  query: string,
  opts: MatchOptions,
): Match[] {
  const re = compilePattern(query, opts);
  if (!re) return [];
  const out: Match[] = [];
  for (let i = 0; i < segments.length; i += 1) {
    const seg = segments[i];
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(seg.text)) !== null) {
      out.push({
        segmentId: seg.id,
        segmentIndex: i,
        offset: m.index,
        length: m[0].length,
        label: seg.label,
      });
      if (m[0].length === 0) re.lastIndex += 1;
    }
  }
  return out;
}


/**
 * Last-used Find state, preserved across dialog mount/unmount within the
 * renderer-process session. EditorSurface keys its `<ErrorBoundary
 * key={active.id}>` on the active tab id — every tab switch unmounts the
 * editor subtree, taking its `<FindReplaceDialog />` instance with it.
 * Without this cache, typing a query in one editor and hopping to another
 * tab to verify something wipes the query before you can re-open Find on
 * the new tab. VS Code / Word both keep the last query primed across
 * files; this gives the three FindReplaceDialog hosts (Docx / Pptx / Xlsx)
 * the same continuity. Only fields the user actively configures are
 * cached: query, replacement, and the three match-option toggles.
 * `activeIdx` is intentionally per-document — different docs have
 * different match sets, so reusing an index would point at unrelated text.
 * `statusMsg` is also per-document and per-action, so we skip it. Cleared
 * on app reload, same volatility as the editor view-state Maps.
 */
let lastFindState = {
  query: '',
  replacement: '',
  caseSensitive: false,
  useRegex: false,
  wholeWord: false,
};

export function FindReplaceDialog({
  open,
  onClose,
  segments,
  onUpdateSegment,
  onLocateSegment,
  title,
  focusNonce,
}: Props): JSX.Element | null {
  // R405 — bilingual.
  const t = useT();
  // Seed from the module-level cache so a remount (tab switch) brings back
  // whatever the user last had typed. The selection-prefill effect below
  // still wins when the user has a non-empty selection at the moment they
  // hit Ctrl+F, matching VS Code's "find selected" gesture.
  const [query, setQuery] = useState(lastFindState.query);
  const [replacement, setReplacement] = useState(lastFindState.replacement);
  const [caseSensitive, setCaseSensitive] = useState(lastFindState.caseSensitive);
  const [useRegex, setUseRegex] = useState(lastFindState.useRegex);
  const [wholeWord, setWholeWord] = useState(lastFindState.wholeWord);
  // Persist user-configured fields back to the module cache on every
  // change so the next mount picks them up. Cheap (one object assignment
  // per keystroke) and avoids the alternative of plumbing this state up
  // through workspace store, which would entrench dialog internals in a
  // place where they don't belong.
  useEffect(() => {
    lastFindState = { query, replacement, caseSensitive, useRegex, wholeWord };
  }, [query, replacement, caseSensitive, useRegex, wholeWord]);
  const [activeIdx, setActiveIdx] = useState(0);
  /**
   * One-shot status line shown under the dialog after a Replace / Replace
   * All — "已取代 5 個" / "找不到符合項目". Auto-clears so it doesn't
   * stale into the next operation.
   */
  const [statusMsg, setStatusMsg] = useState<string | null>(null);
  const statusTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flashStatus = (msg: string) => {
    setStatusMsg(msg);
    if (statusTimer.current) clearTimeout(statusTimer.current);
    statusTimer.current = setTimeout(() => setStatusMsg(null), 2500);
  };
  useEffect(
    () => () => {
      if (statusTimer.current) clearTimeout(statusTimer.current);
    },
    [],
  );

  /**
   * Two-step "全部取代" guard. Bulk replace is destructive in this app
   * because undo is per-segment-editor — once we've rewritten N different
   * segments there's no single Ctrl+Z that brings them all back. We split
   * the click for batches of 5+ so a fat finger doesn't blow away dozens
   * of matches across the doc. Single-digit batches just go through, since
   * forcing confirmation on a 2-replacement run would be more annoying
   * than helpful (matches VS Code / Word's no-confirm flow for small ops).
   * Window auto-cancels after 4s; the effect below also cancels on any
   * change to the query / options / segments so a stale "確認取代" never
   * applies to a different match set than the one the user just saw.
   */
  const REPLACE_ALL_CONFIRM_THRESHOLD = 5;
  const [confirmAll, setConfirmAll] = useState(false);
  const confirmTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (confirmTimer.current) clearTimeout(confirmTimer.current);
    },
    [],
  );
  const queryInputRef = useRef<HTMLInputElement>(null);
  // Where to send focus when the dialog closes. Captured at the moment the
  // dialog opens — typically the toolbar Find button or the cell/textbox the
  // user was just editing — so dismissing with Esc/× doesn't strand focus on
  // <body>, where subsequent app shortcuts (Ctrl+W close tab, Ctrl+S save…)
  // technically still work via the document-level keymap but typing into the
  // editor doesn't. If onLocateSegment moved focus into the editor while the
  // dialog was open, that element is *also* what activeElement points at
  // here — restoring is harmless because we only restore when post-close
  // focus has fallen back to <body> (i.e. nothing else claimed it).
  const restoreFocusToRef = useRef<HTMLElement | null>(null);

  const opts = useMemo<MatchOptions>(
    () => ({ caseSensitive, regex: useRegex, wholeWord }),
    [caseSensitive, useRegex, wholeWord],
  );

  // Detect invalid regex up-front so we can show an inline error and skip
  // collecting matches against a broken pattern.
  const regexError = useMemo<string | null>(() => {
    if (!useRegex || !query) return null;
    try {
      // Use the same source we'd compile so wholeWord-decorated patterns
      // also surface errors.
      const src = wholeWord ? `\\b(?:${query})\\b` : query;
      // eslint-disable-next-line no-new
      new RegExp(src, caseSensitive ? 'g' : 'gi');
      return null;
    } catch (e) {
      return (e as Error).message;
    }
  }, [useRegex, query, caseSensitive, wholeWord]);

  // Recompute matches whenever segments / query / option toggles change.
  const matches = useMemo(
    () => (regexError ? [] : collectMatches(segments, query, opts)),
    [segments, query, opts, regexError],
  );

  // Clamp activeIdx into the valid range whenever matches shrink.
  useEffect(() => {
    if (matches.length === 0) {
      setActiveIdx(0);
    } else if (activeIdx >= matches.length) {
      setActiveIdx(matches.length - 1);
    }
  }, [matches.length, activeIdx]);

  // Cancel the "確認取代" pending state whenever the match set could have
  // shifted under the user's feet — query / option toggle / replacement
  // text / segments edited externally. Without this, a user could stage a
  // confirm against 12 matches, change the query, and the second click
  // would silently apply to a different (possibly larger) match set.
  useEffect(() => {
    if (!confirmAll) return;
    setConfirmAll(false);
    if (confirmTimer.current) clearTimeout(confirmTimer.current);
    // confirmAll intentionally omitted — including it would re-trigger
    // the effect on its own state change and blank an in-flight confirm.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, replacement, caseSensitive, useRegex, wholeWord, segments]);

  /**
   * After `replaceOne`, jump activeIdx to the first match that starts
   * at-or-after the byte right past the inserted replacement. Without this,
   * a replacement string that re-introduces the search query (e.g.
   * `TODO` → `TODO(joe)`) makes the recomputed match list include the
   * just-inserted `TODO`; activeIdx stays at 0 pointing into the new
   * insertion, and clicking Replace again expands it to `TODO(joe)(joe)`,
   * then `TODO(joe)(joe)(joe)`, ad infinitum. Setting the ref before
   * `onUpdateSegment` lets this effect fire on the next render once
   * matches have recomputed against the new segment text. Same advance
   * convention as VS Code / Word.
   */
  const advanceAfterReplaceRef = useRef<{ segmentIndex: number; offsetAfter: number } | null>(null);
  useEffect(() => {
    const tgt = advanceAfterReplaceRef.current;
    if (!tgt) return;
    advanceAfterReplaceRef.current = null;
    if (matches.length === 0) return;
    let idx = matches.findIndex(
      (m) =>
        m.segmentIndex > tgt.segmentIndex ||
        (m.segmentIndex === tgt.segmentIndex && m.offset >= tgt.offsetAfter),
    );
    // Wrap to the start when nothing remains after the replacement —
    // continues the iteration cycle so users can keep mashing Replace.
    if (idx < 0) idx = 0;
    setActiveIdx(idx);
  }, [matches]);

  // Auto-focus the query field when the dialog opens — same UX as Word /
  // Sheets pressing Ctrl+F. Also select-all so re-opening with a stale
  // query lets the user overtype immediately. `focusNonce` re-fires this
  // on every Ctrl+F even when the dialog was already open, so users who
  // clicked into the document to navigate a match can hit Ctrl+F to
  // come right back to the query field.
  //
  // Pre-fill from the host's current selection (Adobe / VS Code / Word
  // convention — "find more of what I just selected"). The read MUST happen
  // before the `el.focus()` below: stealing focus to the dialog's input
  // collapses the underlying selection in some hosts, so the read at that
  // point would come back empty. Only overwrites when the host has a
  // non-empty selection — an empty selection preserves whatever query the
  // user already had typed, matching VS Code (re-opening with no selection
  // doesn't blank your last query).
  useEffect(() => {
    if (open) {
      // R300 — reset transient statusMsg on open. The dialog never
      // unmounts (returns null when open=false, but state survives);
      // flashStatus's 2.5s timer can be cut short by close-before-fire,
      // leaving「已取代 12 個」visible the next time the user opens
      // Find for a fresh search. Stale flash misleads — the message
      // claims an action that didn't happen in the current session.
      // Same per-action-state-cleanup-on-reopen idiom as SettingsDialog's
      // confirmClearKey at [SettingsDialog.tsx:101-105]. Also clear the
      // pending timer so the new session's flashStatus calls aren't
      // races against a leftover from the previous open.
      setStatusMsg(null);
      if (statusTimer.current) {
        clearTimeout(statusTimer.current);
        statusTimer.current = null;
      }
      const prefill = readActiveSelectionText();
      queueMicrotask(() => {
        if (prefill) setQuery(prefill);
        const el = queryInputRef.current;
        if (!el) return;
        el.focus();
        el.select();
      });
    }
  }, [open, focusNonce]);

  // Locate hook — fire whenever the active match changes.
  useEffect(() => {
    if (!open) return;
    const cur = matches[activeIdx];
    if (cur && onLocateSegment) onLocateSegment(cur.segmentId);
  }, [activeIdx, matches, open, onLocateSegment]);

  // Capture-on-open / restore-on-close so dismissing the dialog returns
  // focus to whatever element triggered Ctrl+F (toolbar button, editor
  // surface, …) rather than dropping it on <body>. We only restore when
  // post-close `activeElement` has fallen back to body — if anything else
  // (e.g. onLocateSegment landed focus inside the editor) holds focus,
  // pulling it away would yank the user out of where they navigated to.
  useEffect(() => {
    if (open) {
      restoreFocusToRef.current = document.activeElement as HTMLElement | null;
      return;
    }
    const target = restoreFocusToRef.current;
    restoreFocusToRef.current = null;
    if (!target) return;
    if (!document.body.contains(target)) return;
    const ae = document.activeElement;
    if (ae !== null && ae !== document.body) return;
    target.focus();
  }, [open]);

  // Document-level Esc — closes the dialog even when focus has moved into
  // the underlying editor (cell, slide text box, doc block) to navigate a
  // match. Without this, users have to click back into the dialog before
  // Esc takes effect, which contradicts VS Code / Sheets / Chrome where
  // Esc-anywhere always dismisses the find widget. We skip the close when
  // the user is actively editing an input — cell-edit cancel still owns
  // Esc in that case and would lose its "revert to original" behaviour.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      const ae = document.activeElement as HTMLElement | null;
      if (ae) {
        const tag = ae.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA' || ae.isContentEditable) return;
      }
      onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  // F3 = next match, Shift+F3 = previous match. Document-level so the
  // shortcut works even after the user has clicked into the underlying
  // editor (a cell / slide text box / paragraph) to read context around
  // the current hit. Without this, navigating matches required either
  // clicking back into the query field or pressing Ctrl+F again to
  // re-focus + Enter — three keystrokes / a mouse trip for a one-key
  // shortcut. MarkdownEditor already gets F3 / Shift+F3 for free via
  // CM6's built-in search (see MarkdownEditor.tsx:468); this brings
  // XlsxEditor / DocxEditor / PptxEditor up to the same baseline. F3
  // is not a printable character so we don't need to skip when focus
  // is in an input — typing F3 in any field can never be the user's
  // intent. Matches Word / VS Code / Chrome conventions.
  const matchCount = matches.length;
  useEffect(() => {
    if (!open) return;
    if (matchCount === 0) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'F3') return;
      e.preventDefault();
      e.stopPropagation();
      if (e.shiftKey) {
        setActiveIdx((i) => (i - 1 + matchCount) % matchCount);
      } else {
        setActiveIdx((i) => (i + 1) % matchCount);
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, matchCount]);

  // Esc closes; Enter = next; Shift+Enter = previous. Scoped to dialog
  // root so the host editor's keymap doesn't fight us.
  if (!open) return null;

  const total = matches.length;
  const current = matches[activeIdx];

  const next = () => {
    if (total === 0) return;
    setActiveIdx((i) => (i + 1) % total);
  };
  const prev = () => {
    if (total === 0) return;
    setActiveIdx((i) => (i - 1 + total) % total);
  };

  /**
   * Replace the *current* match. In literal mode we splice `replacement`
   * verbatim. In regex mode we run `text.replace(re, replacement)` over
   * just the matched substring so capture-group placeholders ($1, $&,
   * etc.) get expanded — same UX as VS Code.
   */
  const replaceOne = () => {
    if (!current) return;
    const seg = segments[current.segmentIndex];
    if (!seg || seg.id !== current.segmentId) return;
    const matched = seg.text.slice(current.offset, current.offset + current.length);
    const expanded = useRegex
      ? matched.replace(compilePattern(query, opts) ?? /a^/, replacement)
      : replacement;
    const before = seg.text.slice(0, current.offset);
    const after = seg.text.slice(current.offset + current.length);
    // Arm the post-replace advance BEFORE calling onUpdateSegment — the
    // host editor will synchronously update segments, the matches useMemo
    // will recompute, and the ref-watching effect will fire on the next
    // render to skip activeIdx past `expanded`'s end.
    advanceAfterReplaceRef.current = {
      segmentIndex: current.segmentIndex,
      offsetAfter: current.offset + expanded.length,
    };
    onUpdateSegment(seg.id, before + expanded + after);
    flashStatus(t('已取代 1 個', 'Replaced 1'));
  };

  /**
   * Click handler for "全部取代". Routes to `runReplaceAll` either
   * immediately (small batch / already confirmed) or stages a confirm
   * state that turns the button into "確認取代 N 個" for 4 seconds.
   */
  const onReplaceAllClick = () => {
    if (matches.length === 0) {
      flashStatus(query ? t('找不到符合項目', 'No matches found') : '');
      return;
    }
    if (matches.length >= REPLACE_ALL_CONFIRM_THRESHOLD && !confirmAll) {
      setConfirmAll(true);
      if (confirmTimer.current) clearTimeout(confirmTimer.current);
      confirmTimer.current = setTimeout(() => setConfirmAll(false), 4000);
      return;
    }
    if (confirmTimer.current) clearTimeout(confirmTimer.current);
    setConfirmAll(false);
    runReplaceAll();
  };

  /**
   * Replace every match. In regex mode we let `String.replace(re, repl)`
   * do the heavy lifting per segment so capture groups expand correctly.
   * In literal mode we splice from rightmost to leftmost so earlier
   * offsets stay valid.
   */
  const runReplaceAll = () => {
    if (matches.length === 0) {
      flashStatus(query ? t('找不到符合項目', 'No matches found') : '');
      return;
    }
    const re = useRegex ? compilePattern(query, opts) : null;
    const bySegment = new Map<string, Match[]>();
    for (const m of matches) {
      const list = bySegment.get(m.segmentId) ?? [];
      list.push(m);
      bySegment.set(m.segmentId, list);
    }
    let count = 0;
    for (const [segId, segMatches] of bySegment) {
      const seg = segments.find((s) => s.id === segId);
      if (!seg) continue;
      if (re) {
        // Reset lastIndex so .replace starts from byte 0 of THIS segment.
        re.lastIndex = 0;
        onUpdateSegment(seg.id, seg.text.replace(re, replacement));
        count += segMatches.length;
        continue;
      }
      let text = seg.text;
      const sorted = [...segMatches].sort((a, b) => b.offset - a.offset);
      for (const m of sorted) {
        text = text.slice(0, m.offset) + replacement + text.slice(m.offset + m.length);
        count += 1;
      }
      onUpdateSegment(seg.id, text);
    }
    flashStatus(t(`已取代 ${count} 個`, `Replaced ${count}`));
  };

  return (
    <div
      role="dialog"
      // R129 — aria-label language alignment (English → Traditional Chinese,
      // mirroring the visible heading). This was the lone English aria-label
      // in the entire renderer: every other one is Chinese —
      //   App.tsx:1259 / TabBar.tsx:424   `尚未儲存`
      //   AIPanel.tsx:593                 `關閉錯誤訊息`
      //   AIPanel.tsx:758/903/1295        `停用工具呼叫` / `複製訊息` / `送出`
      //   DocxEditor.tsx:3371/3530        `編輯連結` / `插入 Markdown`
      //   FileExplorer.tsx (banner X)     `關閉錯誤訊息`
      //   GoToDialog.tsx:161              `aria-label={label}` — the visible
      //                                    Chinese label like 「跳到第幾段？」
      //   MarkdownEditor.tsx:1067         `插入圖片 / 插入連結`
      //   SettingsDialog.tsx:318          `隱藏 API key / 顯示 API key`
      //   Toaster.tsx:101                 `通知`
      // Whereas this one read English `"Find and Replace"` — the only outlier
      // among 13+ Chinese aria-labels, and the *most prominent* of them
      // because role="dialog" + aria-label is what a screen reader announces
      // the moment focus enters the dialog (Ctrl+F). The visible heading at
      // line 574 already renders `{title ?? '尋找與取代'}`, so AT users heard
      // English while sighted users read Chinese — a worse split than just
      // "stale string" because aria-label *overrides* the visible heading
      // for AT (per ARIA spec — aria-label takes precedence over inner text
      // for accessible-name calculation). Mirror the visible heading's
      // expression so AT and visible UI stay in lockstep across all three
      // hosts (DocxEditor + XlsxEditor → default `尋找與取代`,
      // PptxEditor.tsx:1008 → override `尋找與取代 · 全部投影片`).
      aria-label={title ?? '尋找與取代'}
      // Floating panel pinned to the upper-right of the editor — small
      // enough to not block content, draggable would be Phase D polish.
      className="absolute top-3 right-3 z-30 w-[360px] rounded-md border bg-background shadow-lg p-3 text-sm"
      onKeyDown={(e) => {
        // R233 — IME composition guard, mirrors R231 (TabBar rename) /
        // R232 (XlsxEditor cell edit + sheet rename). CJK users search
        // for Chinese / Japanese / Korean text by typing pinyin /
        // bopomofo / kana / jamo into the find input, then pressing
        // Enter to confirm the IME candidate. Without this guard, the
        // dialog-level Enter handler (which bubbles from the find /
        // replacement <input>) fires preventDefault + next() → IME's
        // confirmation is blocked, the input stays on the raw IME
        // buffer, and the search runs against「ㄒㄐㄚ」instead of
        // 「西甲」. Same isComposing escape hatch the AIPanel input
        // already uses for its send-Enter path.
        if (e.nativeEvent.isComposing) return;
        if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        } else if (e.key === 'Enter') {
          e.preventDefault();
          if (e.shiftKey) prev();
          else next();
        }
      }}
    >
      <div className="flex items-center justify-between mb-2">
        <div className="font-medium flex items-center gap-1.5">
          <Search className="h-3.5 w-3.5" />
          {title ?? '尋找與取代'}
        </div>
        {/* R128 — cross-component close-X scope-disclosure parity. Bare
            「關閉 (Esc)」 was the lone outlier among close-X buttons in
            the codebase after R119/R121 finished sweeping the others:
              TabBar.tsx:467         關閉此頁籤 (Ctrl+W)   ← R119
              AIPanel.tsx:592        關閉錯誤訊息          ← scope-disclosed
              FileExplorer.tsx:325   關閉錯誤訊息          ← R121
              Toaster.tsx:150        關閉通知              ← R121
              FindReplaceDialog:580  關閉 (Esc)            ← was here, bare
            And the self-contradicting cross-reference made it doubly
            wrong: DocxEditor.tsx:3589-3593 explicitly cites this exact
            line as the precedent for "naming what's being cancelled
            rather than echoing the bare label" with 取消插入 (Esc) —
            but the actual string here was bare 關閉, breaking the very
            claim that referenced it (same paradox R121 caught between
            AIPanel.tsx:583 and Toaster.tsx). 關閉尋找與取代 mirrors
            the dialog title's anchor (`尋找與取代`, line 574 default;
            PptxEditor.tsx:1008 overrides to 「尋找與取代 · 全部投影
            片」 but the scope name without the qualifier still reads
            correctly there) so the tooltip stays accurate across all
            three host editors (Docx / Pptx / Xlsx). The (Esc) gate is
            unconditional here because — unlike TabBar's Ctrl+W which
            only targets the active tab — Esc inside this dialog always
            closes it (line 561-563 onKeyDown handler), so the shortcut
            hint is always truthful. */}
        <button
          type="button"
          onClick={onClose}
          className="p-0.5 rounded hover:bg-secondary text-muted-foreground"
          title={t('關閉尋找與取代 (Esc)', 'Close Find & Replace (Esc)')}
          // R152 — icon-only-button accessible-name parity. Same round
          // adds aria-label to the three other close-X buttons (AIPanel
          // 錯誤訊息, TabBar 頁籤 X, FileExplorer 錯誤橫幅, Toaster 通知);
          // this dialog's close-X was the lone holdout. The shortcut
          // parenthetical lives in `title` (visual hover) only, per the
          // AIPanel.tsx:775 convention — aria-label carries the action
          // verb; SR engines announce keyboard shortcuts separately.
          aria-label={t('關閉尋找與取代', 'Close Find & Replace')}
        >
          <X className="h-3.5 w-3.5" />
        </button>
      </div>

      <div className="space-y-2">
        <div className="flex items-center gap-1">
          <input
            ref={queryInputRef}
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            // R155 — explicit accessible name. Placeholder「尋找…」alone is the
            // SR fallback, but it disappears the moment the user types — SR
            // engines that re-announce on focus return then read「edit, blank」
            // with no hint of what this field is for. `aria-label` stays
            // anchored to the role, mirroring the convention used at every
            // other input in this dialog (none currently!) and the broader
            // codebase pattern of icon-only-button aria-label parity (R152).
            aria-label={t('尋找', 'Find')}
            // Select-all when the user manually clicks back into a non-empty
            // query field — common pattern after Replace All when they want
            // to immediately search for something else.
            onFocus={(e) => e.currentTarget.select()}
            // R127 — CJK ellipsis harmonization (`...` → `…`). This file
            // was the lone half-width holdout in the entire src tree:
            // grepping `[\u4e00-\u9fff]+\.\.\.` returns these two
            // placeholders only (here and 「取代為...」 ~80 lines below),
            // while 13 other files — every editor (DocxEditor.tsx:2556
            // 「輸入文字…」, PptxEditor.tsx:2596 「在這裡寫下這張投影片
            // 的備忘稿…」, XlsxEditor.tsx:2402 「輸入內容或 =公式…」),
            // [R144 — PptxEditor cite updated 2582 → 2596 to track
            //  the +14 line shift introduced by R144's anchor block
            //  + tooltip wrap; same same-round-self-drift correction
            //  paradigm as R142.]
            // AIPanel.tsx:1229/1231 (「請先設定 API key…」, 「AI 回應中…」),
            // SettingsDialog/StatusBar/MarkdownToolbar/menu/toast — already
            // use the full-width 「…」 form. Same CJK-typography direction
            // R124 set for `（）` and R126 just set for 「？」: when the
            // surrounding string is Chinese, reach for the CJK glyph
            // instead of the three-period ASCII fallback. Pure visual
            // change — placeholder pixel width drops by ~12px which is
            // strictly nicer in this narrow find input, and JS string
            // length changes (3→1) don't touch the search query (the
            // placeholder is HTML-attribute only, never read as `value`).
            placeholder={t('尋找…', 'Find…')}
            // Red border / ring when the user has typed something but
            // there are zero hits — without this, the only "no match"
            // signal was the 11px "0 個結果" tucked into the status
            // line, which is easy to miss after a typo. Mirror VS Code
            // / Word's convention. We deliberately skip the red state
            // when `regexError` is set: the dedicated red error line
            // below already explains why the pattern matches nothing,
            // and stacking two red signals on the same field reads as
            // panic rather than information (per the existing comment
            // on the regex-error block — "keeps the field outline
            // calm").
            className={cn(
              'flex-1 px-2 py-1 text-xs border rounded bg-background outline-none focus:ring-1',
              query && total === 0 && !regexError
                ? 'border-destructive/60 focus:ring-destructive'
                : 'focus:ring-primary',
            )}
          />
          {/* R89 — disabled-state tooltip flip across all four action buttons
              in this dialog (prev / next / Replace / Replace All). The dialog
              already recognises the no-match state visually: the search input
              at line ~611 flips its border to destructive/red when `query &&
              total === 0`. Mirror that recognition on hover too — previously
              all four buttons kept their action-name tooltip even when greyed
              out, so a user with a typo'd query saw「下一個 (Enter / F3)」on
              a button that did nothing. Wording 「沒有符合項目」matches the
              「N 個結果 / 0 個結果」status-line vocabulary the dialog already
              uses, so the row reads as one voice. */}
          <button
            type="button"
            onClick={prev}
            disabled={total === 0}
            title={total === 0 ? t('沒有符合項目', 'No matches') : t('上一個 (Shift+Enter / Shift+F3)', 'Previous (Shift+Enter / Shift+F3)')}
            // R155 — icon-only-button accessible-name parity, same pattern
            // R152 closed for the four close-X buttons across the codebase.
            // Shortcut parenthetical lives in `title` only, per the
            // AIPanel.tsx:775 convention — aria-label carries the action
            // verb only, since SR engines announce shortcuts via the
            // separate keyboard-state surface.
            aria-label={t('上一個符合項目', 'Previous match')}
            className={cn(
              'p-1 rounded text-muted-foreground hover:text-foreground hover:bg-secondary',
              total === 0 && 'opacity-40 pointer-events-none',
            )}
          >
            <ChevronUp className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onClick={next}
            disabled={total === 0}
            title={total === 0 ? t('沒有符合項目', 'No matches') : t('下一個 (Enter / F3)', 'Next (Enter / F3)')}
            aria-label={t('下一個符合項目', 'Next match')}
            className={cn(
              'p-1 rounded text-muted-foreground hover:text-foreground hover:bg-secondary',
              total === 0 && 'opacity-40 pointer-events-none',
            )}
          >
            <ChevronDown className="h-3.5 w-3.5" />
          </button>
        </div>

        <div className="flex items-center gap-1">
          <input
            type="text"
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            // R155 — same accessible-name treatment as the 尋找 input above.
            aria-label={t('取代為', 'Replace with')}
            // Enter inside the replacement field = replace current match,
            // matching VS Code / Word / Sheets convention. The dialog-level
            // onKeyDown maps Enter → next(); without stopPropagation here,
            // typing Enter after the replacement string would skip to the
            // next match without ever applying the edit. stopPropagation
            // keeps that handler from also firing; replaceOne() itself
            // advances activeIdx via advanceAfterReplaceRef, so the user
            // still moves forward through the document naturally. We only
            // intercept bare Enter — Shift+Enter falls through to prev()
            // for symmetry with the find input's behaviour.
            onKeyDown={(e) => {
              // R233 — same IME guard as the dialog-level handler above.
              // The replacement input is the second CJK typing surface
              // in this dialog (find input is the first, gated by the
              // dialog-level handler's R233 short-circuit). Without
              // this, typing CJK replacement text and pressing Enter to
              // confirm a candidate fires replaceOne() with the raw
              // IME buffer in `replacement`, replacing matches with
              // unconfirmed pinyin / bopomofo / etc.
              if (e.nativeEvent.isComposing) return;
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                e.stopPropagation();
                replaceOne();
              }
            }}
            placeholder={t('取代為…', 'Replace with…')}
            className="flex-1 px-2 py-1 text-xs border rounded bg-background outline-none focus:ring-1 focus:ring-primary"
          />
          <button
            type="button"
            onClick={replaceOne}
            disabled={!current}
            // Surface the Enter-in-replacement-field shortcut the same way
            // the navigation buttons advertise theirs ("下一個 (Enter / F3)" /
            // "上一個 (Shift+Enter / Shift+F3)" — siblings ~12 lines up). The
            // replacement input's onKeyDown (line ~657) does call replaceOne
            // on bare Enter, but with no tooltip hint a user focused in the
            // 取代為 field has no visual cue that Enter applies the edit —
            // they have to mouse to this button or guess. Qualifying with
            // 「於取代欄位」disambiguates from the dialog-level Enter that
            // triggers next() everywhere else (see onKeyDown at line ~564).
            // R89 — disabled-flip per the prev/next batch comment above.
            title={!current ? t('沒有符合項目可取代', 'No matches to replace') : t('取代目前項目（於『取代為』欄位按 Enter）', 'Replace current match (press Enter in the Replace field)')}
            // R155 — sibling parity with the prev/next icon-only buttons
            // above. ReplaceAll right next door has visible text「全部取代」
            // / 「確認取代 N 個」 so it doesn't need aria-label; this Replace
            // button is the only icon-only one in the row.
            aria-label={t('取代目前項目', 'Replace current match')}
            className={cn(
              'p-1 rounded text-muted-foreground hover:text-foreground hover:bg-secondary',
              !current && 'opacity-40 pointer-events-none',
            )}
          >
            <Replace className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onClick={onReplaceAllClick}
            disabled={total === 0}
            // R89 — disabled branch added in front of the existing two-state
            // confirmAll/normal title. Order matters: when total === 0 the
            // confirmAll latch can't have been set anyway (button can't be
            // clicked), so the disabled string takes precedence cleanly.
            title={
              total === 0
                ? t('沒有符合項目可取代', 'No matches to replace')
                : confirmAll
                  ? t('再次點擊以確認取代', 'Click again to confirm replacing')
                  : t('取代所有符合項目', 'Replace all matches')
            }
            className={cn(
              'px-2 py-1 text-[11px] border rounded transition-colors whitespace-nowrap',
              confirmAll
                ? 'bg-amber-500/15 border-amber-500/60 text-amber-700 hover:bg-amber-500/25'
                : 'hover:bg-secondary',
              total === 0 && 'opacity-40 pointer-events-none',
            )}
          >
            {confirmAll ? t(`確認取代 ${total} 個`, `Confirm: replace ${total}`) : t('全部取代', 'Replace All')}
          </button>
        </div>

        <div className="flex items-center justify-between text-[11px] text-muted-foreground">
          <div className="flex items-center gap-3">
            <label className="inline-flex items-center gap-1 cursor-pointer">
              <input
                type="checkbox"
                checked={caseSensitive}
                onChange={(e) => setCaseSensitive(e.target.checked)}
                className="h-3 w-3"
              />
              {t('區分大小寫', 'Match case')}
            </label>
            <label className="inline-flex items-center gap-1 cursor-pointer" title={t('完整單字符合 (\\b 邊界)', 'Match whole words (\\b boundaries)')}>
              <input
                type="checkbox"
                checked={wholeWord}
                onChange={(e) => setWholeWord(e.target.checked)}
                className="h-3 w-3"
              />
              {t('全字符合', 'Whole word')}
            </label>
            <label className="inline-flex items-center gap-1 cursor-pointer" title={t('JavaScript 正規表示式語法', 'JavaScript regex syntax')}>
              <input
                type="checkbox"
                checked={useRegex}
                onChange={(e) => setUseRegex(e.target.checked)}
                className="h-3 w-3"
              />
              .*
            </label>
          </div>
          {/* R154 — Live-region announcement for the match-counter, mirroring
              the role="status" aria-live="polite" pair already in use at
              [Toaster.tsx:108-109] and [DocxEditor.tsx:1561-1562]. The
              counter text changes on every keystroke (search query rebuild)
              and on every Next / Prev / F3 navigation; without aria-live, SR
              users typing into the search field heard nothing — they had no
              feedback for「query found N matches」or「now on match 3 of 17」.
              `aria-atomic="true"` so the entire string is re-read on each
              update (the change between「12 / 17」 and「13 / 17」 is small but
              the whole new context is what the user wants, not a diff). */}
          <span role="status" aria-live="polite" aria-atomic="true">
            {total === 0
              ? query
                ? t('0 個結果', '0 results')
                : ''
              : `${activeIdx + 1} / ${total}${current?.label ? ` · ${current.label}` : ''}`}
          </span>
        </div>

        {/* Regex syntax error — keeps the field outline calm but tells the
            user why their pattern matches nothing. R154 — `role="alert"`
            (implicit `aria-live="assertive"`) so SR users typing an invalid
            pattern (`[` 等未閉合 metacharacter) hear the parser's error
            immediately, instead of just an empty「0 個結果」 with no clue
            why. Mirrors the urgency tier the Toaster uses for error variants
            (Toaster.tsx:109 → `aria-live={t.variant === 'error' ? 'assertive'
            : 'polite'}`). */}
        {regexError && (
          <div role="alert" className="text-[11px] text-destructive truncate" title={regexError}>
            {t('正規表示式錯誤：', 'Regex error: ')}{regexError}
          </div>
        )}

        {/* Flash status from Replace / Replace All — auto-clears after 2.5s
            so it never staleness-overrides the next operation. R154 — same
            polite live-region as the counter above so SR users hear「已取代
            N 個」 after pressing 全部取代; without it, the green flash was
            sighted-only feedback for an action they just triggered. */}
        {statusMsg && (
          <div role="status" aria-live="polite" className="text-[11px] text-emerald-600">{statusMsg}</div>
        )}
      </div>
    </div>
  );
}
