/**
 * Markdown editor — CodeMirror 6 (source) + `marked` rendered preview.
 *
 * Spec §5.2.1 promises split-view live preview, smart shortcuts, and a
 * status bar; this component owns all of those:
 *
 *   - Tri-mode layout (source / split / preview-only) toggled from toolbar
 *   - Toolbar commands also bound to Ctrl+B / I / ` / K / 1-3 / Shift+L keys
 *   - Smart list continuation on Enter (- / 1. / > / [ ])
 *   - Auto-paired brackets / quotes via CM6's `closeBrackets`
 *   - Status bar with word count, character count, and cursor line/col
 *
 * Edits update the workspace tab's `content`; selections feed the AI panel
 * context. External edits (e.g. AI ChangeSet apply) flow back via the
 * `tab.content` effect below.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { EditorView, keymap, lineNumbers, highlightActiveLine } from '@codemirror/view';
import { EditorState } from '@codemirror/state';
import { defaultKeymap, history, historyKeymap, indentWithTab } from '@codemirror/commands';
import { markdown } from '@codemirror/lang-markdown';
import { closeBrackets, closeBracketsKeymap } from '@codemirror/autocomplete';
import { search, searchKeymap, gotoLine } from '@codemirror/search';
import { marked } from 'marked';
import { useWorkspace } from '../store/workspace';
import { notify } from '../store/toast';
import type { MarkdownTab } from '../types/tab';
import { MarkdownToolbar, type MarkdownViewMode } from './MarkdownToolbar';
import {
  applyImage,
  applyLink,
  setLinePrefix,
  smartListContinue,
  wrapSelection,
} from '../lib/markdown-commands';
import { cn, slicePreview } from '../lib/utils';
import { exitExportPdf, tryEnterExportPdf } from '../lib/export-pdf-busy';

interface Props {
  tab: MarkdownTab;
}

interface DocStats {
  words: number;
  chars: number;
  line: number;
  col: number;
}

interface OutlineEntry {
  /** Heading depth: 1 for `#`, 2 for `##`, …, capped at 6. */
  level: number;
  /** Visible heading text — closing trailing `#`s and surrounding whitespace stripped. */
  text: string;
  /** 1-based line number in the source — what CM6 wants for `doc.line()`. */
  lineNumber: number;
  /** 0-based occurrence index across the whole doc — used to pick the
   *  matching `<h1..h6>` in the rendered preview when jumping. */
  occurrence: number;
}

/**
 * Walk the source line-by-line and pull out ATX headings (`# / ## / ...`).
 * We track ``` and ~~~ fences ourselves so a `# comment` inside a code block
 * doesn't pollute the outline. Setext (`===` / `---` underline) headings are
 * ignored — uncommon in our toolbar-driven flow, and disambiguating from
 * horizontal rules requires lookahead we don't need yet.
 */
function parseOutline(text: string): OutlineEntry[] {
  const lines = text.split('\n');
  const out: OutlineEntry[] = [];
  let inFence = false;
  let occurrence = 0;
  // R387 — match BOTH backtick and tilde fences per CommonMark §4.5. The
  // outline pane is the user-facing surface; an unprotected `# comment`
  // inside a `~~~bash` block would otherwise show up as a clickable
  // heading entry that scrolls to the wrong line. See dispatcher.ts:1275
  // for the analogous AI-path fix and the data-loss scenario it closes.
  const fenceRe = /^\s*(?:```|~~~)/;
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (fenceRe.test(line)) {
      inFence = !inFence;
      continue;
    }
    if (inFence) continue;
    const m = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (!m) continue;
    out.push({
      level: m[1].length,
      text: m[2].trim(),
      lineNumber: i + 1,
      occurrence,
    });
    occurrence += 1;
  }
  return out;
}

/**
 * Per-tab scroll + selection memory across remounts within the
 * renderer-process session. `EditorSurface` keys its `<ErrorBoundary
 * key={active.id}>` on the active tab id, so EVERY tab switch unmounts the
 * editor subtree and constructs a fresh `EditorView` for the next active
 * tab — meaning a user reading at line 500 of a long markdown doc, glancing
 * at another tab, and switching back lands at line 1 with no record of
 * where they were. Stashing `scrollDOM.scrollTop` + the main selection
 * range here on unmount and restoring on the next mount recovers the
 * position without persisting anything to the .gd manifest (which would
 * entrench format coupling).
 *
 * Selection (not just scroll) matters because CM6's default initial
 * selection is position 0: a keyboard-only user who Ctrl+Tabs back and
 * presses Down expects to step from line 500, but with scroll-only restore
 * the caret is invisibly at the top — Down moves to line 2 and the screen
 * jumps to the top of the doc. Selection clamps to the current doc length
 * so a shrinking AI changeset between mount and remount can't position
 * the caret out of bounds.
 *
 * Map persists for the renderer process lifetime; cleared on app reload —
 * same volatility as `lastEditAt`. Bounded by the # of markdown tabs the
 * user touches in a session, which is small enough that GC isn't worth
 * the wire-up.
 */
const viewMemory = new Map<string, { scroll: number; selFrom: number; selTo: number }>();
/**
 * Per-tab outline-pane scroll memory — same shape and lifetime as
 * `viewMemory` above, but for the OutlinePanel (left rail). Without it,
 * the user could scroll a long doc's outline to read H30-H40 for
 * context, switch tabs, switch back, and the panel jumps to the top
 * while the editor itself remembered exactly where it was. The existing
 * `scrollIntoView({ block: 'nearest' })` on activeIdx only tracks the
 * caret heading, so it can't carry over a manually-chosen browse
 * position. Mirrors the same fix in DocxEditor (Round 35) and
 * PptxEditor (Round 36) `navScrollMemory`.
 */
const outlineScrollMemory = new Map<string, number>();

/**
 * Pop the OS file picker for an image. Returns null on cancel. Mirrors the
 * `settled` flag pattern already used in XlsxEditor / DocxEditor / PptxEditor —
 * Chromium fires no event when the user dismisses the picker, so we attach a
 * focus-event timeout fallback (next frame, +200 ms grace) and let `change`
 * win if it's actually coming. Without this the promise hangs forever and
 * the next click on the 圖片 button silently no-ops because we never re-arm.
 */
function pickImageFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml';
    let settled = false;
    const finish = (f: File | null) => {
      if (settled) return;
      settled = true;
      resolve(f);
    };
    input.onchange = () => finish(input.files?.[0] ?? null);
    const onFocus = () => {
      window.removeEventListener('focus', onFocus);
      setTimeout(() => finish(null), 200);
    };
    setTimeout(() => window.addEventListener('focus', onFocus), 0);
    input.click();
  });
}

/** Read a File as a base64 `data:` URL. We embed images inline rather than
 *  copying to disk because the markdown tab has no notion of an asset folder
 *  — the doc is one .md file the user can move freely, and a sibling
 *  `images/` directory would silently break on copy. Data URLs keep the
 *  whole document self-contained at the cost of a larger source file. */
function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(typeof fr.result === 'string' ? fr.result : '');
    fr.onerror = () => reject(fr.error ?? new Error('讀取檔案失敗'));
    fr.readAsDataURL(file);
  });
}

/**
 * Strip markdown syntax tokens before counting words. Heuristic-only — we don't
 * need a full parser for a status-bar number, and `marked` would be overkill.
 */
function countWords(text: string): number {
  const stripped = text
    .replace(/```[\s\S]*?```/g, ' ') // fenced code
    .replace(/`[^`]*`/g, ' ') // inline code
    .replace(/!\[[^\]]*\]\([^)]*\)/g, ' ') // images
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1') // links → keep label
    .replace(/[*_~>#-]+/g, ' '); // markdown punctuation
  // Split on whitespace + CJK char boundaries (CJK each char counts as a word).
  const cjk = stripped.match(/[\u4e00-\u9fff\u3040-\u30ff\u3400-\u4dbf]/g)?.length ?? 0;
  const words = stripped
    .replace(/[\u4e00-\u9fff\u3040-\u30ff\u3400-\u4dbf]/g, ' ')
    .split(/\s+/)
    .filter(Boolean).length;
  return cjk + words;
}

export function MarkdownEditor({ tab }: Props): JSX.Element {
  const containerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  // Loop guard for split-view scroll sync — when we drive one pane
  // programmatically the resulting `scroll` event would otherwise feed back
  // into the other pane and ricochet. Cleared on the next animation frame.
  const programmaticScrollRef = useRef(false);
  const patchTab = useWorkspace((s) => s.patchTab);
  const setSelection = useWorkspace((s) => s.setSelection);

  // R275 — markdown analog of docx/pptx/xlsx's `lastWrittenBytesRef` (R253 /
  // R254). When AI Apply / Ctrl+Z / Ctrl+Shift+Z mutates tab.content from
  // outside, the [tab.content] effect below dispatches a doc replacement
  // into CM6, which fires the updateListener with docChanged=true, which
  // echoes the same content back via patchTab — triggering a workspace-
  // wide re-render and bumping lastEditAt (which resets the autosave
  // debounce to "ε after AI Apply" instead of "ε after the user's last
  // real edit"). Track the last value already pushed to / received from
  // the store so the updateListener can short-circuit on programmatic
  // dispatches: pre-arm in [tab.content] before view.dispatch, and clear
  // in updateListener when a real edit lands. Initialize from the mount-
  // time tab.content so the very first user keystroke still triggers a
  // patchTab (newContent never equals tab.content at mount because they
  // differ by at least the typed character).
  const lastPatchedContentRef = useRef<string>(tab.content);

  // Persist view mode the same way `outlineOpen` is — without this, every
  // tab switch unmounts MarkdownEditor (EditorSurface's ErrorBoundary keys
  // off active.id) and resets the choice back to 'split'. Users who prefer
  // preview-only reading or source-only writing had to re-toggle on every
  // bounce between markdown tabs. The choice is global, not per-tab: tabs
  // are interchangeable enough that a single preference matches user intent.
  const [viewMode, setViewMode] = useState<MarkdownViewMode>(() => {
    try {
      const v = localStorage.getItem('gendoc.markdownViewMode');
      if (v === 'source' || v === 'split' || v === 'preview') return v;
    } catch {
      /* private mode / quota — fall through to default */
    }
    return 'split';
  });
  useEffect(() => {
    try {
      localStorage.setItem('gendoc.markdownViewMode', viewMode);
    } catch {
      /* preference just won't persist */
    }
  }, [viewMode]);
  const [stats, setStats] = useState<DocStats>({ words: 0, chars: 0, line: 1, col: 1 });
  const [previewHtml, setPreviewHtml] = useState<string>('');
  // Latched while a PDF export is in flight so the toolbar can grey-out the
  // button — a second click would spawn a second hidden BrowserWindow + save
  // dialog, and the resulting two save dialogs stack confusingly.
  const [exportPdfBusy, setExportPdfBusy] = useState(false);
  /**
   * Synchronous in-flight gate for `handleExportPdf`. R196 — same shape as
   * R162's `sendingRef` in AIPanel: the React state `exportPdfBusy` is the
   * source of truth for the toolbar button's `disabled` prop, but
   * setExportPdfBusy(true) at line 751 only takes effect after React commits
   * the next render. A rapid double-click on 「輸出為 PDF」 in the React-batch
   * window between click 1's setState and the render commit fires both
   * `handleExportPdf` calls with the SAME closure `exportPdfBusy === false`
   * — the gate at line 737 passes for both, two `markdown.exportPdf` IPCs
   * fly, the user gets two stacked OS save dialogs (and presumably two
   * redundant PDF writes if they pick destinations on both). The ref is
   * set synchronously before the await, closing the race even when
   * React hasn't re-rendered yet.
   */
  const exportPdfBusyRef = useRef(false);
  // Outline (TOC) sidebar — Adobe Acrobat's "Bookmarks" pane equivalent.
  // Defaults off so existing users see no layout change; persisted per-app
  // (not per-tab) so the choice sticks across sessions.
  const [outlineOpen, setOutlineOpen] = useState<boolean>(() => {
    try {
      return localStorage.getItem('gendoc.markdownOutlineOpen') === '1';
    } catch {
      return false;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem('gendoc.markdownOutlineOpen', outlineOpen ? '1' : '0');
    } catch {
      /* private mode / quota — preference just won't persist */
    }
  }, [outlineOpen]);

  // Inline link/image dialog state. Replaces the previous `window.prompt`
  // flow inside markdown-commands (modal, themeless, single-field). When
  // opened we snapshot the current selection — `defaultLabel` from the
  // selected text, `defaultUrl` empty — and apply on commit. The keymap
  // (Mod-k) reads `openLinkDialogRef.current` because CM6 closures are
  // captured at editor-init time and we want to call into a fresh dialog
  // opener that React re-creates on each render.
  // Snapshot {from,to} at open time. The dialog has no backdrop, so a stray
  // click on the editor while typing the URL moves CM6's main selection;
  // applyLink / applyImage previously read `state.selection.main` at *commit*
  // time and inserted the link wherever the user happened to click last,
  // not where they invoked the dialog. Restoring the snapshotted range
  // before calling apply* anchors the insertion to the original spot.
  const [linkDialog, setLinkDialog] = useState<{
    kind: 'link' | 'image';
    defaultLabel: string;
    defaultUrl: string;
    rangeFrom: number;
    rangeTo: number;
  } | null>(null);
  const openLinkDialogRef = useRef<(() => void) | null>(null);
  const openImageDialogRef = useRef<(() => void) | null>(null);
  const openLinkDialog = useCallback(() => {
    const view = viewRef.current;
    if (!view) return;
    const range = view.state.selection.main;
    const sel = view.state.sliceDoc(range.from, range.to);
    // If the selection itself already looks like a URL, route it into the
    // URL field instead of the label field. Common workflow: paste a URL,
    // select it, Ctrl+K to wrap with a friendlier label. Without the swap
    // the URL landed in 顯示文字 and the user had to delete it from there
    // and re-paste it into 連結網址 — exactly the modal-prompt friction
    // the inline dialog was supposed to remove. Detection mirrors the
    // commit-time auto-scheme regex (`^scheme://`), conservative enough to
    // avoid false positives on selections that just happen to contain a
    // colon (e.g. `Note: hello`).
    const trimmed = sel.trim();
    const looksLikeUrl = /^[a-z][a-z0-9+.-]*:\/\//i.test(trimmed);
    setLinkDialog({
      kind: 'link',
      defaultLabel: looksLikeUrl ? '' : sel,
      defaultUrl: looksLikeUrl ? trimmed : '',
      rangeFrom: range.from,
      rangeTo: range.to,
    });
  }, []);
  const openImageDialog = useCallback(() => {
    const view = viewRef.current;
    if (!view) return;
    const range = view.state.selection.main;
    // Pre-fill alt text from the current selection so a workflow like
    // "select 'logo' → click 圖片 → paste URL → Enter" produces
    // `![logo](url)` without forcing the user to retype the alt. The
    // earlier behaviour always seeded an empty defaultLabel, so the
    // selection got replaced by `![](url)` and the descriptive alt the
    // user had highlighted was lost. Mirrors openLinkDialog above and
    // matches the convention in VS Code / Typora / Obsidian (selection
    // becomes label/alt for both link and image insertion).
    //
    // URL-shaped selections short-circuit to the URL field (same rationale
    // as openLinkDialog) so a user who pasted an image URL and selected it
    // doesn't end up with `![https://…](https://…)` placeholder text.
    const sel = view.state.sliceDoc(range.from, range.to);
    const trimmed = sel.trim();
    const looksLikeUrl = /^[a-z][a-z0-9+.-]*:\/\//i.test(trimmed);
    setLinkDialog({
      kind: 'image',
      defaultLabel: looksLikeUrl ? '' : sel,
      defaultUrl: looksLikeUrl ? trimmed : '',
      rangeFrom: range.from,
      rangeTo: range.to,
    });
  }, []);
  useEffect(() => {
    openLinkDialogRef.current = openLinkDialog;
    openImageDialogRef.current = openImageDialog;
  }, [openLinkDialog, openImageDialog]);

  // Re-parse on every content change. Heading regex over <2k lines is sub-ms
  // — cheaper than the marked.parse already running on the same doc — so we
  // skip debouncing and let the outline track typing in real time.
  const outline = useMemo(() => parseOutline(tab.content), [tab.content]);

  // Which outline entry contains the caret? = the last heading whose line
  // number is at-or-before the current cursor line. Adobe Acrobat's
  // Bookmarks pane and VS Code's Outline view both highlight the entry the
  // user is currently "in" so they don't lose their place while scrolling
  // a long document. -1 when the caret sits above the first heading
  // (no section is active yet) or there are no headings. `outline` is
  // already sorted by lineNumber via parseOutline, so a single linear
  // scan suffices.
  const activeOutlineIdx = useMemo(() => {
    if (outline.length === 0) return -1;
    let active = -1;
    for (let i = 0; i < outline.length; i += 1) {
      if (outline[i].lineNumber <= stats.line) active = i;
      else break;
    }
    return active;
  }, [outline, stats.line]);

  // Preview render — debounced so we don't re-parse the whole document on
  // every keystroke. 120ms is short enough that the preview stays "live"
  // (faster than typing rhythm so the user sees output update mid-type) but
  // long enough that big docs don't lag the editor.
  useEffect(() => {
    let cancelled = false;
    const t = setTimeout(() => {
      // R280 — wrap both throw paths: marked.parse can sync-throw (custom
      // extensions, deeply nested input hitting marked's internal stack,
      // future versions tightening input validation) AND can return a
      // rejected Promise (async tokenizer plugins, future API surface
      // changes). The original code had neither try/catch nor `.catch`,
      // so a parse failure escaped as uncaught exception / unhandled
      // rejection — surfacing in DevTools but invisible to the user, who
      // just saw the preview freeze on the last successful render with
      // no recovery path. Same swallow-and-degrade policy as
      // XlsxEditor.writeBack's `.catch(() => {})` (serializer errors
      // surface via the next user action). Keeping the previous
      // previewHtml is less jarring than blanking on every transient
      // parse glitch — failures are rare in practice (typing partial
      // markdown is parse-safe), and the next successful keystroke
      // overwrites the stale preview automatically.
      try {
        Promise.resolve(marked.parse(tab.content))
          .then((html) => {
            if (!cancelled) setPreviewHtml(typeof html === 'string' ? html : '');
          })
          .catch(() => undefined);
      } catch {
        /* sync throw — same swallow policy as the async branch */
      }
    }, 120);
    return () => {
      cancelled = true;
      clearTimeout(t);
    };
  }, [tab.content]);

  // Initialize the editor once per mount (keyed off tab.id — see deps below).
  useEffect(() => {
    if (!containerRef.current) return undefined;

    const updateListener = EditorView.updateListener.of((v) => {
      if (v.docChanged) {
        const newContent = v.state.doc.toString();
        // R275 — short-circuit programmatic-sync dispatches. The
        // [tab.content] effect pre-arms `lastPatchedContentRef` with the
        // about-to-be-inserted string before calling view.dispatch; that
        // dispatch fires this listener with newContent === ref, so we
        // skip the echo patchTab that would otherwise re-render the
        // whole workspace and bump lastEditAt for no real user edit.
        // User-keystroke path: ref still holds the previous content, so
        // newContent !== ref, and we proceed (and update the ref) as
        // before.
        if (newContent === lastPatchedContentRef.current) return;
        lastPatchedContentRef.current = newContent;
        patchTab(tab.id, { content: newContent } as Partial<MarkdownTab>);
        // No direct preview render here — `patchTab` triggers the debounced
        // `tab.content` effect above. Parsing in both places caused a 2x
        // marked.parse per keystroke, visibly lagging large documents.
      }
      if (v.docChanged || v.selectionSet) {
        const sel = v.state.selection.main;
        const lineObj = v.state.doc.lineAt(sel.head);
        setStats({
          words: countWords(v.state.doc.toString()),
          chars: v.state.doc.length,
          line: lineObj.number,
          col: sel.head - lineObj.from + 1,
        });
      }
      if (v.selectionSet) {
        const sel = v.state.selection.main;
        if (!sel.empty) {
          const slice = v.state.doc.sliceString(sel.from, sel.to);
          // R382 — code-point-aware preview slice via slicePreview helper.
          // Pre-fix `.slice(0, 60)` could split an emoji's surrogate pair
          // at the 60-char boundary; the badge then rendered an orphan
          // high surrogate as「�」. AIPanel's `text: slice` carries the
          // full selection (no truncation) so AI side is unaffected — only
          // the visual badge string was broken. The helper's `…` suffix
          // is internal; whitespace-collapse happens before slicing so
          // the boundary calc lands on already-normalised text.
          const preview = slicePreview(slice.replace(/\s+/g, ' '), 60);
          setSelection({
            tabId: tab.id,
            preview: `[md selection] ${preview}`,
            // `text` carries the full raw selection — AIPanel sends this
            // as selectionContext so the model sees everything the user
            // highlighted, not just the 60-char badge preview.
            text: slice,
            payload: { tabId: tab.id, selectionText: slice },
          });
        } else {
          setSelection(null);
        }
      }
    });

    /**
     * Keymap precedence: format / smart-list keys are listed BEFORE
     * `defaultKeymap`, so our Enter handler beats the default newline-insert.
     * Each command returns `false` to fall through when not applicable
     * (e.g. Enter on a non-list line — let CM6 insert a normal newline).
     */
    const formatKeys = keymap.of([
      { key: 'Mod-b', run: (v) => wrapSelection(v, '**', '**') },
      { key: 'Mod-i', run: (v) => wrapSelection(v, '_', '_') },
      { key: 'Mod-`', run: (v) => wrapSelection(v, '`', '`') },
      {
        key: 'Mod-k',
        run: () => {
          openLinkDialogRef.current?.();
          return true;
        },
      },
      { key: 'Mod-1', run: (v) => setLinePrefix(v, '# ') },
      { key: 'Mod-2', run: (v) => setLinePrefix(v, '## ') },
      { key: 'Mod-3', run: (v) => setLinePrefix(v, '### ') },
      { key: 'Mod-Shift-l', run: (v) => setLinePrefix(v, '- ') },
      // Ordered list. Mod-Shift-o pairs with Mod-Shift-l (unordered) — both
      // toggle list shape on the current line. Without this the toolbar
      // button was the only path; keyboard-only users had no way to reach it
      // and the tooltip-promised "consistency with bold/italic/heading" was
      // broken specifically for ordered lists.
      { key: 'Mod-Shift-o', run: (v) => setLinePrefix(v, '1. ') },
      // R141 — Blockquote (`> `). Closes the same gap for blockquote that
      // the Mod-Shift-o block above closed for ordered lists. The toolbar's
      // line-prefix transformation family — Mod-1/2/3 for the three heading
      // levels (lines 430-432), Mod-Shift-l for unordered list (line 433),
      // Mod-Shift-o for ordered list (line 439) — all share the same shape
      // (`setLinePrefix(v, 'X ')` over the current line). Blockquote at
      // MarkdownToolbar.tsx:172 calls the identical helper with `'> '` but
      // had no keymap entry, making it the only line-prefix toolbar button
      // a keyboard-only user couldn't reach. Same words from the
      // Mod-Shift-o comment apply verbatim — "tooltip-promised consistency
      // with bold/italic/heading was broken specifically for [blockquote]."
      // Q for Quote follows the Mod-Shift-{first letter of role} convention
      // the l/o pair established (Mod-Shift-l = list, Mod-Shift-o = ordered);
      // q is unbound elsewhere in the codebase (verified — no other
      // `'q'` / `'Q'` / `Mod-q` / `Mod-Shift-q` keymap or DOM handler).
      // R142 — the original R141 doc-comment above named
      // `MarkdownToolbar.tsx:166`, the line where blockquote's
      // `setLinePrefix(v, '> ')` lived BEFORE R141 added the matching
      // tooltip-comment block in the toolbar file. R141's own 6-line
      // comment insertion in MarkdownToolbar.tsx pushed the Quote button
      // down to line 172 and shifted everything below by +6 — making
      // R141's self-citation stale within the same round (the exact
      // R136/R137 stale-line-ref paradigm replaying in real time inside
      // a new round). R142 also closed the +6 shift in two pre-existing
      // sibling cites (EditorSurface.tsx:341 cited :227→:233 for the
      // `'尋找與取代 (Ctrl+F)'` title, XlsxEditor.tsx:2072 cited the
      // R96 image-button cluster span as :171-204→:184-209).
      { key: 'Mod-Shift-q', run: (v) => setLinePrefix(v, '> ') },
      // Ctrl+G — open CM6's native "Go to line" panel. The other three
      // editors all bind Ctrl+G as their jump-to gesture (DocxEditor →
      // paragraph N, PptxEditor → slide N, XlsxEditor → focus Name Box),
      // so Markdown was the odd one out — `searchKeymap` only ships the
      // Ctrl+Alt+G variant. Power users muscle-trained on "Ctrl+G = jump"
      // across the workspace previously got browser-default "Find Next"
      // (or nothing) when they tried it in a .md tab. formatKeys is
      // ordered above searchKeymap in the extensions list so this binding
      // wins for plain Ctrl+G; Mod-Alt-g still works for users who learned
      // CM6's default.
      { key: 'Mod-g', run: gotoLine },
      { key: 'Enter', run: smartListContinue },
    ]);

    // Restore selection (caret / range) from the previous mount of this
    // tab id by seeding it into the initial EditorState. Doing this here
    // rather than dispatching after `new EditorView` matters because
    // post-create dispatches trigger CM6's "scroll caret into view"
    // behaviour, which fights the scroll restore below. Clamp against the
    // current doc length so an AI changeset that shrunk the doc between
    // unmount and remount can't anchor out of bounds.
    const remembered = viewMemory.get(tab.id);
    const initialSelection = remembered
      ? (() => {
          const docLen = tab.content.length;
          const anchor = Math.min(Math.max(remembered.selFrom, 0), docLen);
          const head = Math.min(Math.max(remembered.selTo, 0), docLen);
          return { anchor, head };
        })()
      : undefined;

    const state = EditorState.create({
      doc: tab.content,
      selection: initialSelection,
      extensions: [
        lineNumbers(),
        highlightActiveLine(),
        history(),
        closeBrackets(),
        // CM6's built-in search panel — handles Ctrl+F (open), F3 (next),
        // Shift+F3 (prev), Ctrl+H (replace) plus regex / case / whole-word
        // toggles for free. Keymap goes ABOVE defaultKeymap so its bindings
        // win when the panel is open.
        search({ top: true }),
        formatKeys,
        // `indentWithTab` binds Tab → indentMore / Shift+Tab → indentLess.
        // CM6 deliberately leaves this out of `defaultKeymap` because it
        // breaks the browser's "Tab moves focus" expectation; we add it back
        // anyway because list nesting (`- foo` → `  - foo`) is the single
        // thing markdown users reach for Tab to do, and silently swallowing
        // the keystroke (focus escapes the editor) is the more confusing
        // failure mode. Keyboard-only users escape via the existing
        // app-level shortcuts (Ctrl+L → AI panel, Ctrl+B → explorer) or
        // can press Esc once before Tab to fall through to the browser
        // default.
        keymap.of([
          ...closeBracketsKeymap,
          ...searchKeymap,
          indentWithTab,
          ...defaultKeymap,
          ...historyKeymap,
        ]),
        markdown(),
        EditorView.lineWrapping,
        updateListener,
      ],
    });
    const view = new EditorView({ state, parent: containerRef.current });
    viewRef.current = view;
    // Seed the stats line on mount — read from the (possibly restored)
    // selection rather than hard-coding 1:1, so the status bar matches the
    // caret the user sees blinking after a tab switch back in.
    const seedSel = state.selection.main;
    const seedLine = state.doc.lineAt(seedSel.head);
    setStats({
      words: countWords(tab.content),
      chars: tab.content.length,
      line: seedLine.number,
      col: seedSel.head - seedLine.from + 1,
    });
    // Restore scroll position from the previous mount of this tab id
    // (see viewMemory comment above). Deferred via queueMicrotask so CM6's
    // initial measurement / layout pass completes first; setting scrollTop
    // synchronously on `.cm-scroller` is overwritten by CM6's own
    // scroll-to-cursor behaviour on the very next layout flush.
    if (remembered && remembered.scroll > 0) {
      queueMicrotask(() => {
        // Bail if a faster tab switch already destroyed this view.
        if (viewRef.current !== view) return;
        view.scrollDOM.scrollTop = remembered.scroll;
      });
    }
    return () => {
      // Capture scroll + selection before destruction. `view.scrollDOM` is
      // the `.cm-scroller` element — same one the split-sync handler reads
      // — so mode toggles (split/source/preview) that don't re-init the
      // editor also benefit because cleanup never runs for them.
      const sel = view.state.selection.main;
      viewMemory.set(tab.id, {
        scroll: view.scrollDOM.scrollTop,
        selFrom: sel.from,
        selTo: sel.to,
      });
      view.destroy();
      viewRef.current = null;
    };
    // We intentionally re-init the editor when switching tabs (tab.id change).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  // Clipboard image paste + drag-drop. The other editors all support "paste
  // a screenshot" (DocxEditor:788, PptxEditor:714, XlsxEditor:1168) — markdown
  // was the odd one out. Without this, a user who copied a screenshot had to
  // open Paint → save → click 圖片 → 從本機選擇檔案 — four steps for what
  // should be Ctrl+V.
  //
  // Paste: when the clipboard carries an image item, embed as a data URL via
  // applyImage at the current selection. preventDefault so CM6 doesn't also
  // paste the textual fallback (Chromium often puts both an image and a
  // HTML/text representation on the clipboard for screenshots).
  //
  // Drop: file drag from File Explorer / Finder. dragover preventDefault is
  // what makes the editor a valid drop target — without it the OS bounces
  // the file back. We listen on view.dom so drop coordinates can be mapped
  // to a CM6 doc position via posAtCoords (image lands where the user
  // actually dropped, not at the previous caret).
  //
  // Keyed off tab.id so we re-attach when the init effect rebuilds the view
  // on a tab swap — without this, the listeners outlive their dom element
  // and a fresh markdown tab silently loses paste support.
  useEffect(() => {
    const view = viewRef.current;
    if (!view) return undefined;
    const dom = view.dom;

    const insertImageAt = async (file: File, pos: number | null) => {
      try {
        const dataUrl = await fileToDataUrl(file);
        if (!dataUrl) {
          notify('讀取貼上的圖片失敗', 'error');
          return;
        }
        const v = viewRef.current;
        if (!v) return;
        if (pos != null) {
          const clamped = Math.max(0, Math.min(pos, v.state.doc.length));
          v.dispatch({ selection: { anchor: clamped, head: clamped } });
        }
        const alt = file.name ? file.name.replace(/\.[^.]+$/, '') : '圖片';
        applyImage(v, dataUrl, alt);
        v.focus();
      } catch (err) {
        notify(`貼上失敗：${err instanceof Error ? err.message : String(err)}`, 'error');
      }
    };

    const onPaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      // R91 — paste-side parallel to the R85 drop-side fix at line ~628.
      // Track whether the clipboard carried any *file* item we declined
      // (i.e. non-image). Plain-text pastes don't surface as file items,
      // so this only flags real "I copied a file from File Explorer"
      // attempts. Without this, pasting a PDF / .docx / etc. produced
      // zero feedback — same silent-swallow class R85 closed for drops.
      let sawNonImageFile = false;
      for (let i = 0; i < items.length; i += 1) {
        const it = items[i];
        if (it.kind === 'file' && it.type.startsWith('image/')) {
          const file = it.getAsFile();
          if (!file) continue;
          e.preventDefault();
          void insertImageAt(file, null);
          return;
        }
        if (it.kind === 'file') sawNonImageFile = true;
      }
      if (sawNonImageFile) {
        e.preventDefault();
        notify('只能貼上圖片檔案', 'warning');
      }
    };

    const isFileDrag = (dt: DataTransfer | null): boolean => {
      if (!dt) return false;
      const types = dt.types;
      for (let i = 0; i < types.length; i += 1) {
        if (types[i] === 'Files') return true;
      }
      return false;
    };
    const onDragOver = (e: DragEvent) => {
      if (!isFileDrag(e.dataTransfer)) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    };
    const onDrop = (e: DragEvent) => {
      if (!isFileDrag(e.dataTransfer)) return;
      const files = Array.from(e.dataTransfer?.files ?? []);
      const imageFile = files.find((f) => f.type.startsWith('image/'));
      if (!imageFile) {
        // R85 silent-swallow fix — see DocxEditor.tsx:1335 for the full
        // rationale. CodeMirror's own dragover sets dropEffect=copy, so the
        // OS shows "+ copy" mouse cue and the user reasonably expects the
        // drop to *do something*. Previously a non-image drop just silently
        // dismissed with no toast. preventDefault here too so CodeMirror's
        // default-handler doesn't paste the file path as text into the buffer.
        e.preventDefault();
        notify('只能拖入圖片檔案', 'warning');
        return;
      }
      e.preventDefault();
      const v = viewRef.current;
      const pos = v?.posAtCoords({ x: e.clientX, y: e.clientY }) ?? null;
      void insertImageAt(imageFile, pos);
    };

    dom.addEventListener('paste', onPaste);
    dom.addEventListener('dragover', onDragOver);
    dom.addEventListener('drop', onDrop);
    return () => {
      dom.removeEventListener('paste', onPaste);
      dom.removeEventListener('dragover', onDragOver);
      dom.removeEventListener('drop', onDrop);
    };
  }, [tab.id]);

  // External edits (e.g. AI ChangeSet apply, undo/redo of an AI changeset)
  // come in via store updates; sync them into the editor when content
  // diverges. Preserve the caret's (line, col) across the swap — CM6's
  // default position-mapping collapses every selection to the boundary
  // of a whole-document replace, so without this the cursor jumps to
  // position 0 every time the AI applies a change. The user clicks Apply
  // expecting to continue editing where they were, then has to scroll
  // back and re-click. Snapshotting line/col and restoring with a clamp
  // keeps them roughly in place even when the change resized the doc.
  useEffect(() => {
    const view = viewRef.current;
    if (!view) return;
    const current = view.state.doc.toString();
    if (current === tab.content) return;
    const sel = view.state.selection.main;
    const lineObj = view.state.doc.lineAt(sel.head);
    const oldLine = lineObj.number;
    const oldCol = sel.head - lineObj.from;
    // R275 — pre-arm the skip flag BEFORE dispatch. The dispatch fires
    // updateListener synchronously; without this, the listener would
    // echo the freshly-applied content back to the store as a patchTab,
    // re-rendering the whole workspace and bumping lastEditAt for what
    // is fundamentally a store→view sync, not a user edit. See ref
    // doc-block for the full loop trace.
    lastPatchedContentRef.current = tab.content;
    view.dispatch({
      changes: { from: 0, to: current.length, insert: tab.content },
    });
    const newDoc = view.state.doc;
    const targetLine = Math.min(oldLine, newDoc.lines);
    const lineInfo = newDoc.line(targetLine);
    const anchor = Math.min(lineInfo.from + oldCol, lineInfo.to);
    view.dispatch({ selection: { anchor } });
  }, [tab.content]);

  const getView = useCallback(() => viewRef.current, []);

  /**
   * Hand the most recently rendered preview HTML over to main, which will
   * spawn a hidden BrowserWindow, drive `printToPDF`, and reveal the result
   * in the OS file manager. Returning `null` from main means the user
   * cancelled the save dialog — silently no-op (no toast clutter for an
   * intentional back-out). Real failures get a single error toast.
   *
   * Default filename: tab.name with the trailing `.md`/`.markdown` stripped
   * — the save dialog will tack `.pdf` back on. The user almost always wants
   * `notes.pdf`, not `notes.md.pdf`.
   */
  const handleExportPdf = useCallback(async () => {
    // R196 — sync ref gate (component-local, drives same-instance double
    // -click block).
    // R226 — module-level gate (`lib/export-pdf-busy.ts`) on TOP, so a
    // tab-switch-out-and-back during an in-flight export doesn't let the
    // remounted MarkdownEditor instance fire a second IPC. The component-
    // local ref alone died with the unmount; the module-level boolean
    // survives. Order matters: check module gate first (more authoritative,
    // covers cross-instance races), then component ref (same-instance
    // micro-batch), then the React state. The module gate is released in
    // `finally` regardless of which path the body took.
    if (!tryEnterExportPdf()) return;
    if (exportPdfBusyRef.current) {
      exitExportPdf();
      return;
    }
    if (exportPdfBusy) {
      exitExportPdf();
      return;
    }
    if (!previewHtml) {
      notify('預覽尚未準備好，請稍候再試', 'warning');
      exitExportPdf();
      return;
    }
    const baseName = tab.name.replace(/\.(md|markdown)$/i, '');
    // R172 — capture workspaceId before the export IPC. Same shape as R171
    // for handleExportTab in App.tsx: the OS save dialog can be open for
    // many seconds while the user picks a destination, and a Ctrl+O during
    // that window would otherwise leave us painting「已匯出 foo.pdf」 on
    // the new workspace's StatusBar. The disk write is correct (it
    // captured the OLD workspace's previewHtml), only the renderer flash
    // needs to be workspace-scoped.
    const exportWorkspaceId = useWorkspace.getState().workspaceId;
    exportPdfBusyRef.current = true;
    setExportPdfBusy(true);
    try {
      const res = await window.gendoc.markdown.exportPdf({
        suggestedName: baseName,
        title: baseName,
        bodyHtml: previewHtml,
      });
      if (res) {
        // Match the native-format export UX (App.tsx::handleExportTab,
        // TabBar::exportSingleTab): a 5-second green flash in the StatusBar
        // showing just the basename. Previously we toasted the full Windows
        // path (`C:\Users\…\Documents\notes.pdf`) which (a) overflowed the
        // toast row, (b) was redundant with the OS save dialog the user just
        // clicked through, and (c) put PDF export in a different visual
        // channel from every other export — users training on "exports flash
        // green at the bottom" had to learn a second pattern just for PDF.
        // R172 — workspace guard. Disk file already wrote OLD workspace's
        // bytes; only the in-renderer StatusBar flash is gated.
        if (useWorkspace.getState().workspaceId !== exportWorkspaceId) return;
        const fileName = res.filePath.split(/[/\\]/).pop() ?? `${baseName}.pdf`;
        useWorkspace.getState().flashExport(fileName, res.filePath);
      }
    } catch (e) {
      notify(`輸出 PDF 失敗：${(e as Error).message}`, 'error');
    } finally {
      exportPdfBusyRef.current = false;
      setExportPdfBusy(false);
      exitExportPdf();
    }
  }, [exportPdfBusy, previewHtml, tab.name]);

  /**
   * Jump to a heading from the outline sidebar. Drives BOTH panes so the user
   * sees the right thing regardless of view mode:
   *   • Source / split: dispatch CM6 selection at the heading's line and call
   *     `EditorView.scrollIntoView` so the line lands near the top.
   *   • Split / preview: locate the Nth `<h1..h6>` in the rendered HTML by
   *     occurrence index (matches the order parseOutline emitted) and
   *     scrollIntoView. Index-by-occurrence avoids needing slugger /id wiring
   *     into `marked`.
   *
   * `programmaticScrollRef` is raised across both pane scrolls so the split-
   * sync handler doesn't ricochet our own scroll back at us.
   */
  const jumpToHeading = useCallback((entry: OutlineEntry) => {
    const view = viewRef.current;
    if (view) {
      const lineCount = view.state.doc.lines;
      const target = Math.min(entry.lineNumber, lineCount);
      const lineInfo = view.state.doc.line(target);
      view.dispatch({
        selection: { anchor: lineInfo.from },
        effects: EditorView.scrollIntoView(lineInfo.from, { y: 'start', yMargin: 8 }),
      });
      view.focus();
    }
    const pre = previewRef.current;
    if (pre) {
      const headings = pre.querySelectorAll('h1, h2, h3, h4, h5, h6');
      const tgt = headings[entry.occurrence] as HTMLElement | undefined;
      if (tgt) {
        programmaticScrollRef.current = true;
        tgt.scrollIntoView({ behavior: 'auto', block: 'start' });
        requestAnimationFrame(() => {
          programmaticScrollRef.current = false;
        });
      }
    }
  }, []);

  // Split-view scroll sync — both panes follow each other proportionally so
  // a long doc keeps the rendered view in step with the source caret. The
  // sync is *fractional* (scrollTop / (scrollHeight - clientHeight)) which
  // is good enough without a heading-anchor map; not exact for variable
  // line heights, but visibly tracks within a paragraph or two.
  useEffect(() => {
    if (viewMode !== 'split') return;
    // CodeMirror's scroller is `.cm-scroller` inside our container — the
    // wrapping div doesn't actually scroll. Wait for the editor to mount.
    const src = containerRef.current?.querySelector('.cm-scroller') as HTMLElement | null;
    const pre = previewRef.current;
    if (!src || !pre) return;

    const driveFrom = (driver: HTMLElement, target: HTMLElement) => {
      if (programmaticScrollRef.current) return;
      const denom = driver.scrollHeight - driver.clientHeight;
      if (denom <= 0) return;
      const frac = driver.scrollTop / denom;
      const tgtDenom = target.scrollHeight - target.clientHeight;
      if (tgtDenom <= 0) return;
      programmaticScrollRef.current = true;
      target.scrollTop = frac * tgtDenom;
      requestAnimationFrame(() => {
        programmaticScrollRef.current = false;
      });
    };
    const onSrc = () => driveFrom(src, pre);
    const onPre = () => driveFrom(pre, src);
    src.addEventListener('scroll', onSrc, { passive: true });
    pre.addEventListener('scroll', onPre, { passive: true });
    // Initial drive: pull the preview into step with source's *current*
    // scrollTop the moment a usable previewHtml exists. Without this, a
    // tab-switch back into split mode looks like this: viewMemory restores
    // source.scrollTop=1200 via queueMicrotask, which fires the `scroll`
    // event — but at that instant `previewHtml` is still the empty seed
    // so `pre.scrollHeight - pre.clientHeight <= 0` and driveFrom early-
    // returns. The marked.parse debounce resolves ~120ms later,
    // previewHtml lands, this effect re-runs with the listeners freshly
    // attached, and... no `scroll` event fires (source's scrollTop didn't
    // change), so preview is stranded at 0 while the source caret sits on
    // line 500. Driving once at attach time covers that gap. The
    // `programmaticScrollRef` window keeps the resulting preview-scroll
    // event from immediately ricocheting the source.
    driveFrom(src, pre);
    return () => {
      src.removeEventListener('scroll', onSrc);
      pre.removeEventListener('scroll', onPre);
    };
    // Re-attach when view mode changes — switching out of split unmounts
    // the preview pane, switching back in remounts it with a fresh ref.
    // `previewHtml` is also a dep so we re-bind once the rendered HTML has
    // produced its real scrollHeight (otherwise first sync uses 0).
  }, [viewMode, previewHtml]);

  const showSource = viewMode !== 'preview';
  const showPreview = viewMode !== 'source';

  return (
    <div className="h-full w-full flex flex-col">
      <MarkdownToolbar
        getView={getView}
        viewMode={viewMode}
        onViewModeChange={setViewMode}
        outlineOpen={outlineOpen}
        onToggleOutline={() => setOutlineOpen((b) => !b)}
        onInsertLink={openLinkDialog}
        onInsertImage={openImageDialog}
        onExportPdf={handleExportPdf}
        exportPdfBusy={exportPdfBusy}
        canExportPdf={previewHtml.length > 0}
      />
      <div className="flex-1 min-h-0 flex">
        {outlineOpen && (
          <OutlinePanel
            tabId={tab.id}
            entries={outline}
            activeIdx={activeOutlineIdx}
            onJump={jumpToHeading}
          />
        )}
        <div
          ref={containerRef}
          className={cn(
            'min-h-0 overflow-auto',
            showSource && showPreview ? 'flex-1 border-r' : showSource ? 'flex-1' : 'hidden',
          )}
          style={
            showSource && showPreview
              ? { flexBasis: 0, minWidth: 0 }
              : undefined
          }
        />
        {showPreview ? (
          <div
            ref={previewRef}
            className={cn(
              'min-h-0 overflow-auto bg-background',
              showSource ? 'flex-1' : 'flex-1',
            )}
            style={
              showSource ? { flexBasis: 0, minWidth: 0 } : undefined
            }
          >
            <MarkdownPreview html={previewHtml} />
          </div>
        ) : null}
      </div>
      <StatusBar stats={stats} />
      {linkDialog && (
        <LinkInsertDialog
          kind={linkDialog.kind}
          defaultLabel={linkDialog.defaultLabel}
          defaultUrl={linkDialog.defaultUrl}
          onClose={() => {
            setLinkDialog(null);
            viewRef.current?.focus();
          }}
          onCommit={(label, url) => {
            const view = viewRef.current;
            if (!view) return;
            // Restore the selection snapshotted at open time so a stray
            // click on the editor while the dialog was up doesn't relocate
            // the insertion point. Clamp against the current doc length —
            // an external edit (AI ChangeSet apply landing mid-dialog) can
            // shrink the doc; CM6 throws on out-of-range positions.
            const docLen = view.state.doc.length;
            const from = Math.min(linkDialog.rangeFrom, docLen);
            const to = Math.min(linkDialog.rangeTo, docLen);
            view.dispatch({ selection: { anchor: from, head: to } });
            if (linkDialog.kind === 'link') applyLink(view, url, label);
            else applyImage(view, url, label);
            setLinkDialog(null);
            view.focus();
          }}
        />
      )}
    </div>
  );
}

/**
 * Inline floating dialog for inserting markdown links / images. Replaces the
 * old `window.prompt` flow (modal, themeless, single-field, no way to also
 * edit the label). Pinned to the upper-right of the editor like the other
 * floating dialogs (Find/Replace, Go to). Auto-prefixes bare URLs with
 * `https://` so users can paste `example.com` without scheme — same
 * convention as the DOCX link toggle.
 */
function LinkInsertDialog({
  kind,
  defaultLabel,
  defaultUrl,
  onClose,
  onCommit,
}: {
  kind: 'link' | 'image';
  defaultLabel: string;
  defaultUrl: string;
  onClose: () => void;
  onCommit: (label: string, url: string) => void;
}) {
  const [url, setUrl] = useState(defaultUrl);
  const [label, setLabel] = useState(defaultLabel);
  // True when the user just attempted submit (Enter / 插入 button) with the
  // URL field empty. Surfaces a red border + inline hint instead of the
  // previous silent-close behaviour, which threw away whatever the user had
  // typed in the label field. Cleared the moment they start typing in URL,
  // matching the calm-on-edit pattern from FindReplaceDialog's regex error.
  const [urlError, setUrlError] = useState(false);
  // Latched while a file → data-URL read is in flight (image mode only). The
  // FileReader is fast for typical screenshots but multi-MB photos can take
  // a beat; without the latch a second click pops a second OS picker and
  // races the readers. Shown to the user as a disabled button + message.
  const [picking, setPicking] = useState(false);
  const [pickError, setPickError] = useState<string | null>(null);
  const urlRef = useRef<HTMLInputElement>(null);
  const labelRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  // Focus restore on close — same contract as GoToDialog / FindReplaceDialog.
  // Without this, an Esc-cancel from the link dialog left focus on <body>:
  // CodeMirror collapses its visible selection on blur, so the user lost
  // their caret position and had to click back into the editor before
  // typing resumed. The two sibling dialogs (FindReplace, GoTo) already do
  // this; LinkInsert was the lone outlier. Only restore when nothing else
  // has claimed focus post-close — onCommit may have legitimately moved
  // focus into the editor and dispatched a transaction.
  const restoreFocusToRef = useRef<HTMLElement | null>(null);
  useEffect(() => {
    restoreFocusToRef.current = document.activeElement as HTMLElement | null;
    return () => {
      const target = restoreFocusToRef.current;
      restoreFocusToRef.current = null;
      if (!target) return;
      if (!document.body.contains(target)) return;
      const ae = document.activeElement;
      if (ae !== null && ae !== document.body) return;
      target.focus();
    };
  }, []);
  useEffect(() => {
    queueMicrotask(() => {
      // When the URL was already detected from the selection (the new
      // openLinkDialog / openImageDialog URL-shaped-selection swap), the
      // user's next action is typing a label / alt — focus that field
      // instead of the pre-filled URL so they don't have to Tab over.
      // Otherwise the URL field gets the previous focus + select-all
      // behaviour for the "I'm about to paste a URL" workflow.
      if (defaultUrl) {
        labelRef.current?.focus();
        labelRef.current?.select();
      } else {
        urlRef.current?.focus();
        urlRef.current?.select();
      }
    });
  }, [defaultUrl]);
  // Document-level Esc — closes even when focus has wandered out of the
  // dialog (e.g. user clicked into the editor to verify what they're
  // about to link). Mirrors the GoToDialog / FindReplaceDialog pattern.
  // Focus inside the dialog itself is handled by the inline onKeyDown
  // below (which also stopPropagation's the synthetic event); focus in
  // a different editable surface defers to that surface so we don't
  // steal an Esc the user meant for somewhere else.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      const ae = document.activeElement as HTMLElement | null;
      if (ae && dialogRef.current?.contains(ae)) return;
      if (ae) {
        const tag = ae.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA' || ae.isContentEditable) return;
      }
      onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);
  const submit = () => {
    const trimmed = url.trim();
    if (!trimmed) {
      // Empty URL — keep the dialog open and flag the URL field. The previous
      // behaviour was to silently `onClose()`, which both swallowed any label
      // the user had typed and gave no signal as to why "插入" did nothing.
      // Esc still cancels for users who genuinely want to abort.
      setUrlError(true);
      urlRef.current?.focus();
      return;
    }
    // Auto-scheme bare URLs. Match the DOCX behaviour exactly so users get
    // consistent results across formats.
    const normalized = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`;
    onCommit(label.trim(), normalized);
  };
  const isImage = kind === 'image';
  // Image-only: open the OS picker, read as data URL, fill the URL field.
  // Auto-fill alt with the file's basename (sans extension) only when the
  // alt field is empty — a user who selected text first ("logo") and then
  // chose a file ("logo-final.png") meant their selection to be the alt,
  // and we'd be overwriting their intent.
  const onPickFile = async () => {
    if (picking) return;
    setPicking(true);
    setPickError(null);
    try {
      const file = await pickImageFile();
      if (!file) return;
      const dataUrl = await fileToDataUrl(file);
      if (!dataUrl) {
        setPickError('讀取檔案失敗');
        return;
      }
      setUrl(dataUrl);
      setUrlError(false);
      if (!label.trim()) {
        const base = file.name.replace(/\.[^.]+$/, '');
        if (base) setLabel(base);
      }
    } catch (err) {
      setPickError(err instanceof Error ? err.message : '讀取檔案失敗');
    } finally {
      setPicking(false);
    }
  };
  return (
    <div
      ref={dialogRef}
      role="dialog"
      aria-label={isImage ? '插入圖片' : '插入連結'}
      className="absolute top-3 right-3 z-30 w-[320px] rounded-md border bg-background shadow-lg p-3 text-sm space-y-2"
      onKeyDown={(e) => {
        // R234 — IME composition guard. Same shape as R231 / R232 / R233.
        // The dialog has alt-text <input> for images and link-text input
        // for links — both are CJK typing surfaces. CJK candidate
        // confirmation (bare Enter) without this guard bubbles to dialog
        // and fires submit() with the raw IME buffer, inserting bopomofo
        // / pinyin characters into the markdown source instead of the
        // chosen Chinese / Japanese / Korean text.
        if (e.nativeEvent.isComposing) return;
        if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        } else if (e.key === 'Enter') {
          e.stopPropagation();
          e.preventDefault();
          submit();
        }
      }}
    >
      <div className="text-xs font-medium">{isImage ? '插入圖片' : '插入連結'}</div>
      <label className="block">
        <span className="block text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
          {isImage ? '圖片網址' : '連結網址'}
        </span>
        <input
          ref={urlRef}
          type="text"
          value={url}
          onChange={(e) => {
            setUrl(e.target.value);
            // Calm the red border the moment the user starts fixing the
            // empty-URL violation — same pattern as FindReplaceDialog's
            // regex error clearing on edit.
            if (urlError) setUrlError(false);
            if (pickError) setPickError(null);
          }}
          aria-invalid={urlError ? true : undefined}
          placeholder={isImage ? 'https://example.com 或選擇本機檔案' : 'https://example.com'}
          className={cn(
            'w-full px-2 py-1 text-xs border rounded bg-background outline-none focus:ring-1',
            urlError
              ? 'border-destructive/60 focus:ring-destructive'
              : 'focus:ring-primary',
          )}
        />
        {urlError && (
          // R154 — `role="alert"` so SR users hear the empty-URL violation
          // when they hit 插入 with the field blank, mirroring the same
          // round's FindReplaceDialog 正規表示式錯誤 (line 853-856) and
          // GoToDialog 請輸入要跳轉的項次 (line ~241). The visual red
          // border + aria-invalid at line 1125 was sighted-only feedback;
          // the alert closes the SR gap.
          <div role="alert" className="mt-1 text-[10px] text-destructive">
            請先填入網址再插入
          </div>
        )}
        {/* Image-only: file picker. Embeds as a base64 data URL so the .md is
            self-contained — see fileToDataUrl rationale. URL value is hidden
            once it's a data URL because a 2 MB string in a single-line
            <input> is unreadable, and the user really only cares that "a
            file was selected". */}
        {isImage && (
          <div className="mt-1.5 flex items-center gap-2">
            <button
              type="button"
              onClick={onPickFile}
              disabled={picking}
              onMouseDown={(e) => e.preventDefault()}
              className={cn(
                'px-2 py-0.5 text-[11px] rounded border bg-secondary/50 hover:bg-secondary transition-colors',
                picking && 'opacity-50 cursor-not-allowed',
              )}
            >
              {picking ? '讀取中…' : '從本機選擇檔案…'}
            </button>
            {url.startsWith('data:') && (
              <span
                className="text-[10px] text-muted-foreground truncate"
                title="已嵌入本機圖片（base64 data URL）"
              >
                已嵌入本機圖片
              </span>
            )}
          </div>
        )}
        {pickError && (
          // R154 — sibling pattern to the urlError alert above. Both errors
          // appear in the same dialog, both are user-recoverable, both were
          // sighted-only.
          <div role="alert" className="mt-1 text-[10px] text-destructive">{pickError}</div>
        )}
      </label>
      <label className="block">
        <span className="block text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
          {isImage ? '替代文字 (alt)' : '顯示文字'}
        </span>
        <input
          ref={labelRef}
          type="text"
          value={label}
          onChange={(e) => setLabel(e.target.value)}
          placeholder={isImage ? '圖片描述' : '連結文字'}
          className="w-full px-2 py-1 text-xs border rounded bg-background outline-none focus:ring-1 focus:ring-primary"
        />
      </label>
      <div className="flex items-center justify-between pt-1">
        <span className="text-[10px] text-muted-foreground">Esc 取消 · Enter 插入</span>
        {/* Sibling-shortcut-in-tooltip parity — same R58 fix applied to
            DocxEditor's LinkEditDialog primary button at DocxEditor.tsx:3217
            this round. See fuller rationale there; in short, every other
            primary-action button in the app (AIPanel send, SettingsDialog
            save, DiffPreview apply, GoToDialog jump, FindReplaceDialog
            replace) advertises its keystroke on the button itself, and these
            two link dialogs were the lone outliers. Title varies by mode
            (連結 vs 圖片) to mirror the dialog's `kind` prop and the aria-
            label set on the dialog wrapper at line 1045. */}
        <button
          type="button"
          onClick={submit}
          title={isImage ? '插入圖片 (Enter)' : '插入連結 (Enter)'}
          className="px-2 py-1 text-xs rounded bg-primary text-primary-foreground hover:bg-primary/90"
        >
          插入
        </button>
      </div>
    </div>
  );
}

/**
 * Prose-styled markdown preview. Uses Tailwind utility classes scoped to
 * descendants of the wrapping div — we avoid `@tailwindcss/typography` here
 * to keep the dependency surface small for MVP.
 */
function MarkdownPreview({ html }: { html: string }): JSX.Element {
  return (
    <div
      className="px-6 py-4 max-w-3xl mx-auto text-sm leading-7 markdown-preview"
      // marked's output is HTML; the source comes from the user's own editor
      // buffer (no remote content) so XSS surface here is the same as them
      // pasting raw HTML into their editor — acceptable for MVP single-user.
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
}

/**
 * Outline / TOC sidebar — Adobe Acrobat's "Bookmarks" pane equivalent.
 *
 * Shows every ATX heading in the source, indented by depth. Clicking a row
 * jumps both editor panes to that heading. Persists open/closed state to
 * localStorage; defaults closed so existing layouts are unchanged.
 *
 * The panel is intentionally a flat list rather than a collapsible tree —
 * collapsing nodes is the single feature most users complain about as
 * gratuitous in equivalent IDE outliners (they'd rather scan the whole
 * thing). The level prefix (the small "H2" badge) plus indentation already
 * communicates structure.
 */
function OutlinePanel({
  tabId,
  entries,
  activeIdx,
  onJump,
}: {
  tabId: string;
  entries: OutlineEntry[];
  /** Index of the entry that contains the caret; -1 = none. Highlighted
   *  with bg + auto-scrolled into view so the user always knows where
   *  they are in the document. */
  activeIdx: number;
  onJump: (e: OutlineEntry) => void;
}): JSX.Element {
  // The outer `<div>` is the actual scroll container (overflow-auto);
  // the `<ul>` inside has no overflow style. We need a separate ref on
  // that div to read/write scrollTop for memory — `listRef` would give
  // us the inner list, whose scrollTop is always 0.
  const containerRef = useRef<HTMLDivElement>(null);
  // Auto-scroll the active entry into view whenever it changes — long
  // documents push later sections below the panel's visible area, and
  // without this the highlight would just be invisible. `block: 'nearest'`
  // means we don't scroll when it's already visible (avoids a jumpy panel
  // every time the caret crosses a heading boundary). Imperative DOM look-up
  // by data-attr keeps OutlinePanel a pure list — no per-row refs needed.
  const listRef = useRef<HTMLUListElement>(null);
  useEffect(() => {
    if (activeIdx < 0) return;
    const el = listRef.current?.querySelector<HTMLElement>(
      `[data-outline-idx="${activeIdx}"]`,
    );
    el?.scrollIntoView({ block: 'nearest' });
  }, [activeIdx]);
  // Outline-pane scroll memory across tab swaps and `outlineOpen`
  // toggles. Restore via `queueMicrotask` so we run AFTER the activeIdx
  // effect above (which declares first and synchronously scrollIntoViews
  // the active row) — microtask flush comes after the current render's
  // effect chain, so the remembered offset wins. Without this ordering
  // trick, mounting on a tab whose active heading is near the top would
  // snap the pane to the top, even though the user had scrolled to read
  // H30-H40 before tab-swapping away. Same pattern as DocxNavPanel
  // (Round 35) and PptxNavPanel (Round 36).
  useEffect(() => {
    const remembered = outlineScrollMemory.get(tabId);
    if (remembered != null) {
      queueMicrotask(() => {
        const el = containerRef.current;
        if (el) el.scrollTop = remembered;
      });
    }
    return () => {
      const el = containerRef.current;
      if (el) outlineScrollMemory.set(tabId, el.scrollTop);
    };
  }, [tabId]);

  // Arrow-key navigation. Tab still cycles through the buttons (browser
  // default), but in long documents Tab is verbose — VSCode's outline lets
  // ↑/↓ scan headings while staying on this panel. Enter is already handled
  // natively by the focused <button>'s onClick. Home/End jump to the
  // boundaries. We listen on the container so keys work even before the
  // user explicitly Tabs into a row.
  const focusEntry = (i: number) => {
    const el = listRef.current?.querySelector<HTMLButtonElement>(
      `[data-outline-idx="${i}"]`,
    );
    el?.focus();
  };
  const onKeyNav = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (entries.length === 0) return;
    // Read which entry currently owns focus (if any) so ↓ from anywhere in
    // the panel still progresses sensibly. Falls back to activeIdx (the
    // caret-tracked entry) so a fresh focus-into-panel lands near "where
    // I'm at" instead of always bouncing to row 0.
    const focused = document.activeElement as HTMLElement | null;
    const ownIdx = focused?.dataset?.outlineIdx
      ? Number(focused.dataset.outlineIdx)
      : Math.max(0, activeIdx);
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusEntry(Math.min(entries.length - 1, ownIdx + 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusEntry(Math.max(0, ownIdx - 1));
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusEntry(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusEntry(entries.length - 1);
    }
  };
  return (
    <div
      ref={containerRef}
      onKeyDown={onKeyNav}
      className="w-56 shrink-0 border-r bg-secondary/20 overflow-auto text-xs"
    >
      {/* Keyboard-nav hint mirrored across all three outline panels this
          round (DocxEditor DocxNavPanel header, PptxEditor PptxNavPanel
          header). onKeyNav above handles ↑/↓/Home/End but nothing in the
          UI told users the arrow keys work — see fuller rationale at
          PptxEditor.tsx near "投影片大綱". */}
      <div
        className="px-3 py-2 text-[10px] uppercase tracking-wider text-muted-foreground font-medium border-b"
        title="↑/↓ 切換 · Home/End 跳到首/末"
      >
        大綱
      </div>
      {entries.length === 0 ? (
        <div className="px-3 py-2 text-[11px] text-muted-foreground italic">
          （沒有標題。輸入 # / ## / ### 開始建立大綱）
        </div>
      ) : (
        <ul ref={listRef} className="py-1">
          {entries.map((e, i) => {
            const isActive = i === activeIdx;
            return (
              <li key={`${e.lineNumber}-${i}`}>
                <button
                  type="button"
                  data-outline-idx={i}
                  onClick={() => onJump(e)}
                  className={cn(
                    'w-full text-left px-2 py-1 truncate flex items-center gap-1.5 transition-colors',
                    isActive
                      ? 'bg-primary/15 text-foreground font-medium'
                      : 'hover:bg-secondary/80',
                  )}
                  style={{ paddingLeft: 8 + (e.level - 1) * 10 }}
                  title={e.text}
                >
                  <span
                    className={cn(
                      'text-[9px] font-mono shrink-0',
                      isActive ? 'text-primary' : 'text-muted-foreground/70',
                    )}
                  >
                    H{e.level}
                  </span>
                  <span className="truncate">{e.text}</span>
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}

function StatusBar({ stats }: { stats: DocStats }): JSX.Element {
  return (
    <div className="flex items-center gap-4 px-3 py-1 text-[11px] text-muted-foreground border-t bg-secondary/20">
      <span>字數 {stats.words.toLocaleString()}</span>
      <span>字元 {stats.chars.toLocaleString()}</span>
      <span className="ml-auto">
        行 {stats.line}, 欄 {stats.col}
      </span>
    </div>
  );
}
