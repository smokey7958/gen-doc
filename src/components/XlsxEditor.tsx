/**
 * Excel editor — multi-sheet table with per-cell editing + formatting toolbar.
 *
 * UX (path B MVP):
 *   - Sheet tabs along the top.
 *   - Formatting toolbar acts on the currently-selected cell (single-cell
 *     selection model — cell becomes "selected" on focus, persists after
 *     blur so the toolbar buttons can stamp formatting onto it).
 *   - Grid of <input> cells. Click to edit, blur to commit text.
 *   - Style writes go through xlsx-js-style on serialize.
 */

import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Image as ImageIcon,
  Italic,
  PaintBucket,
  Plus,
  Rows,
  Columns,
  Merge,
  Search,
  Split,
  Trash2,
  Type,
  Underline,
  X,
} from 'lucide-react';
import type { XlsxTab } from '../types/tab';
import { useWorkspace } from '../store/workspace';
import {
  type XlsxCell,
  type XlsxCellStyle,
  type XlsxImage,
  type XlsxModel,
  type XlsxSheet,
  addSheet,
  colIndexToLetter,
  parseA1,
  deleteColAt,
  deleteRowAt,
  deleteSheet,
  duplicateSheet,
  injectXlsxImages,
  insertColAt,
  insertRowAt,
  isMergeCovered,
  mergeAtAnchor,
  mergeRange,
  moveSheet,
  parseXlsx,
  renameSheet,
  serializeXlsx,
  unmergeAt,
} from '../lib/xlsx-adapter';
import {
  ERROR_CODES,
  isFormulaSource,
  recomputeAllFormulas,
  type FormulaErrorCode,
} from '../lib/xlsx-formula';
import {
  GENDOC_XLSX_MIME,
  applyPaste,
  clearRange,
  clearRangeExcept,
  extractRange,
  isMultiCellTsv,
  serializeRangeToJson,
  serializeRangeToTsv,
  tsvToPayload,
  type RichClipboardPayload,
} from '../lib/xlsx-clipboard';
import { cn, clampToViewport } from '../lib/utils';
import { FONT_FAMILIES, withEmojiFallback } from '../lib/font-families';
import { useFormatShortcuts } from '../lib/use-format-shortcuts';
import { registerEditorFlush } from '../lib/editor-flush';
import { useUndoableState, useUndoShortcuts } from '../lib/use-undoable-state';
import { FindReplaceDialog, type SearchSegment } from './FindReplaceDialog';
import { notify } from '../store/toast';

interface Props {
  tab: XlsxTab;
}

/**
 * Selection state. `(r, c)` is the *anchor* (the cell that has keyboard
 * focus and whose style is shown in the toolbar); `(r2, c2)` is the
 * opposite corner of the selected rectangle. For a single-cell selection
 * `r === r2 && c === c2`. The rectangle is always derived via
 * `rangeOf(selection)` which sorts the corners.
 */
interface Selection {
  r: number;
  c: number;
  r2: number;
  c2: number;
}

interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function rangeOf(sel: Selection): Range {
  return {
    r1: Math.min(sel.r, sel.r2),
    c1: Math.min(sel.c, sel.c2),
    r2: Math.max(sel.r, sel.r2),
    c2: Math.max(sel.c, sel.c2),
  };
}

function isInRange(sel: Selection | null, r: number, c: number): boolean {
  if (!sel) return false;
  const { r1, c1, r2, c2 } = rangeOf(sel);
  return r >= r1 && r <= r2 && c >= c1 && c <= c2;
}

function rangeAddr(sel: Selection): string {
  const { r1, c1, r2, c2 } = rangeOf(sel);
  if (r1 === r2 && c1 === c2) {
    return `${colIndexToLetter(c1)}${r1 + 1}`;
  }
  return `${colIndexToLetter(c1)}${r1 + 1}:${colIndexToLetter(c2)}${r2 + 1}`;
}

/**
 * Stats over a rectangle of cells. Sum / average count *numeric* cells
 * (anything that parses as a finite Number, after a basic strip). Count is
 * all non-empty cells (matches Excel's status-bar Count, vs Numerical Count).
 */
function rangeStats(sheet: XlsxSheet, range: Range): { sum: number; numericCount: number; count: number; cellCount: number } {
  let sum = 0;
  let numericCount = 0;
  let count = 0;
  let cellCount = 0;
  for (let r = range.r1; r <= range.r2; r += 1) {
    const row = sheet.cells[r];
    if (!row) continue;
    for (let c = range.c1; c <= range.c2; c += 1) {
      cellCount += 1;
      const cell = row[c];
      const text = cell?.text ?? '';
      if (text === '') continue;
      count += 1;
      // Strip thousands separators and a trailing % so 1,234 / 50% still
      // parse — Excel's status bar treats those as numeric.
      const cleaned = text.replace(/,/g, '').replace(/%$/, '');
      const num = Number(cleaned);
      if (Number.isFinite(num)) {
        sum += num;
        numericCount += 1;
      }
    }
  }
  return { sum, numericCount, count, cellCount };
}

// 1 px @ 96 DPI = 9525 EMU. Used to convert pixel sizes (what HTMLImageElement
// reports as `naturalWidth/Height`) into the EMU values that OOXML's
// `<xdr:ext cx cy>` expects.
const EMU_PER_PX = 9525;
/**
 * Default cap on inserted-image width (in pixels). Sheet rows are short and
 * tall images dominate the grid quickly, so we shrink anything wider than
 * ~480 px while preserving aspect ratio. Users can resize after insertion in
 * a later round; for now the panel UI surfaces width/height read-only.
 */
const DEFAULT_MAX_WIDTH_PX = 480;

/**
 * Pop the OS file picker for an image. Returns null if the user cancels.
 * Uses a focus-event timeout fallback because Chromium fires no event on
 * cancel — without this we'd hang forever.
 */
function pickImageFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png,image/jpeg,image/gif,image/bmp';
    let settled = false;
    const finish = (f: File | null) => {
      if (settled) return;
      settled = true;
      resolve(f);
    };
    input.onchange = () => finish(input.files?.[0] ?? null);
    // Cancel detection: when the user dismisses the picker, focus returns to
    // the document. Wait one frame so the change event wins if it's coming.
    const onFocus = () => {
      window.removeEventListener('focus', onFocus);
      setTimeout(() => finish(null), 200);
    };
    setTimeout(() => window.addEventListener('focus', onFocus), 0);
    input.click();
  });
}

/** Read natural dimensions from an image data URL. Falls back to a 384-px
 *  square if the browser can't decode the image — the user can still see
 *  the cell anchor, just at a placeholder size. */
function readNaturalSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 384, height: img.naturalHeight || 384 });
    img.onerror = () => resolve({ width: 384, height: 384 });
    img.src = dataUrl;
  });
}

/** Chunked btoa for big binary blobs — vanilla `btoa(String.fromCharCode(...buf))`
 *  blows the call stack on multi-MB images. 32 K chars is well under V8's arg
 *  cap and an order of magnitude faster than concatenating per-byte. */
function uint8ToBase64(buf: Uint8Array): string {
  const CHUNK = 0x8000;
  let out = '';
  for (let i = 0; i < buf.length; i += CHUNK) {
    out += String.fromCharCode(...buf.subarray(i, Math.min(i + CHUNK, buf.length)));
  }
  return btoa(out);
}

function genImageId(): string {
  return `img_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * Per-tab grid-state memory across remounts within the renderer-process
 * session. Mirrors `slideMemory` in PptxEditor and `viewMemory` in
 * MarkdownEditor — `EditorSurface` keys its `<ErrorBoundary key={active.id}>`
 * on the active tab id, so every tab switch unmounts the editor and the
 * tab-id useEffect below resets the grid back to sheet 0 / scrollTop 0.
 * Multi-sheet workbooks where the user is editing "2024 Q3" or some sheet
 * far down the tab strip get bounced back to sheet 0 on every excursion to
 * another tab.
 *
 * What's preserved:
 *   - `sheetIdx`: which sheet was active. Clamped against the freshly-
 *     parsed sheet count on restore so a sheet deleted during the round-
 *     trip can't strand us out of bounds.
 *   - `scrollTop` / `scrollLeft`: the grid scroller's offset. A row-500
 *     edit point is meaningless if the grid lands at A1 every time, even
 *     after Round 13 already brought us back to the right sheet. Captured
 *     from the wrapping `overflow-auto` div at restore + unmount; the
 *     Grid renders all rows in a `<table>` (no virtualization), so the
 *     pixel offset is a stable identifier without needing to remember a
 *     row/col coordinate.
 *
 * What's NOT preserved:
 *   - Cell selection. Re-clicking a cell is trivial and a stale rectangle
 *     on a sheet whose row/col count changed can point at the wrong
 *     content (the same logic that motivated the original Round 13
 *     omission).
 *
 * Map persists for renderer-process lifetime; cleared on app reload.
 */
const tabMemory = new Map<
  string,
  {
    sheetIdx: number;
    selection: Selection | null;
    scrollTop: number;
    scrollLeft: number;
  }
>();

export function XlsxEditor({ tab }: Props): JSX.Element {
  const patchTab = useWorkspace((s) => s.patchTab);

  const initialBytesRef = useRef<Uint8Array>(tab.data);
  // Re-evaluate formulas on parse so display values reflect *our* engine
  // rather than whatever SheetJS cached. Files written by other Excel-likes
  // sometimes ship with stale `cell.v` after edits; recomputing makes the
  // grid authoritative the moment a tab opens.
  const [model, setModel, undoApi] = useUndoableState<XlsxModel>(() => {
    const parsed = parseXlsx(tab.data);
    return { ...parsed, sheets: recomputeAllFormulas(parsed.sheets) };
  });
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [selection, setSelection] = useState<Selection | null>(null);
  // Mirror of activeSheetIdx readable from the tab-swap effect's cleanup —
  // the cleanup closure captures `activeSheetIdx` at effect-run time
  // (always 0, the useState initial), so reading it directly would clobber
  // any actual navigation the user did. Same wire-up as PptxEditor.
  const activeSheetIdxRef = useRef(0);
  useEffect(() => {
    activeSheetIdxRef.current = activeSheetIdx;
  }, [activeSheetIdx]);
  // Same mirror trick for the cell selection — the tab-swap cleanup needs
  // to snapshot the rectangle before unmount, but the closure captures the
  // initial null. Sync via ref so cleanup reads the live value.
  const selectionRef = useRef<Selection | null>(null);
  useEffect(() => {
    selectionRef.current = selection;
  }, [selection]);
  // Per-sheet selection memory within this tab. Without this, switching
  // sheets via SheetTabs carried the previous sheet's rectangle visually
  // onto the new sheet — Excel never does that; each sheet keeps its own
  // active cell. We snapshot on switch-out and restore on switch-in.
  // Cleared on tab-swap (parse effect) and on structural sheet ops where
  // index re-keying isn't worth the complexity.
  const sheetSelectionsRef = useRef<Map<number, Selection>>(new Map());
  // Per-sheet scroll memory within this tab. The Grid's `overflow-auto`
  // container is a single DOM element across sheet swaps, so without
  // this map the previous sheet's scrollTop bleeds into the new sheet —
  // browser auto-clamps to the new max, which on a shorter sheet looks
  // like "you switched and landed at the bottom." Excel keeps per-sheet
  // scroll; mirror that. Cleared alongside sheetSelectionsRef on parse.
  const sheetScrollsRef = useRef<Map<number, { scrollTop: number; scrollLeft: number }>>(
    new Map(),
  );
  const handleSheetSwitch = (idx: number) => {
    if (idx === activeSheetIdx) return;
    if (selection) sheetSelectionsRef.current.set(activeSheetIdx, selection);
    else sheetSelectionsRef.current.delete(activeSheetIdx);
    // Snapshot the leaving sheet's scroll position synchronously — the
    // swap hasn't happened yet, so the DOM still reflects sheet A.
    const leaving = gridScrollRef.current;
    if (leaving) {
      sheetScrollsRef.current.set(activeSheetIdx, {
        scrollTop: leaving.scrollTop,
        scrollLeft: leaving.scrollLeft,
      });
    }
    setActiveSheetIdx(idx);
    setSelection(sheetSelectionsRef.current.get(idx) ?? null);
    // Apply the new sheet's remembered offsets AFTER React commits the
    // Grid render for it; a synchronous write would be overwritten by
    // the re-render. First-time visit (no entry) resets to 0 so a tall
    // previous sheet doesn't leave the new one parked at its floor.
    const target = sheetScrollsRef.current.get(idx) ?? { scrollTop: 0, scrollLeft: 0 };
    requestAnimationFrame(() => {
      const next = gridScrollRef.current;
      if (!next) return;
      next.scrollTop = target.scrollTop;
      next.scrollLeft = target.scrollLeft;
    });
  };

  // Ref to the grid's scroll container (the `overflow-auto` div wrapping
  // the Grid table). Read in this effect's cleanup to capture scroll
  // position before unmount, and on restore to apply the remembered
  // offset after Grid finishes rendering rows for the active sheet.
  const gridScrollRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    initialBytesRef.current = tab.data;
    // Tab swap → fresh history. Undo shouldn't roll back into a different
    // file's state.
    const parsed = parseXlsx(tab.data);
    const recomputed = { ...parsed, sheets: recomputeAllFormulas(parsed.sheets) };
    undoApi.resetHistory(recomputed);
    // Restore the sheet the user was on before unmount. Clamp against the
    // freshly-parsed sheet count so a sheet deleted during the round-trip
    // can't leave us out of bounds; fall through to 0 when nothing's
    // remembered (first open / Ctrl+Shift+T into a tab id we never tracked).
    const remembered = tabMemory.get(tab.id);
    const initialIdx =
      remembered && recomputed.sheets.length > 0
        ? Math.max(0, Math.min(remembered.sheetIdx, recomputed.sheets.length - 1))
        : 0;
    setActiveSheetIdx(initialIdx);
    // Per-sheet selection memory belongs to one mount; a fresh tab swap is
    // a fresh slate (we still rehydrate the active sheet's selection from
    // tabMemory just below). Per-sheet scroll memory shares this lifetime.
    sheetSelectionsRef.current.clear();
    sheetScrollsRef.current.clear();
    // Restore the cell selection too — losing the active rectangle on every
    // tab swap is jarring when bouncing between a spreadsheet and its
    // supporting docs. Clamp to the freshly-parsed sheet's dimensions so a
    // resize-shrink during round-trip can't leave us pointing at a vanished
    // row/col. Selection is sheet-scoped (anchored to whichever sheet was
    // active at unmount), so we only honour it when the restored sheet
    // matches; otherwise fall back to null.
    let restoredSel: Selection | null = null;
    const restoredSheet = recomputed.sheets[initialIdx];
    if (
      remembered?.selection &&
      remembered.sheetIdx === initialIdx &&
      restoredSheet &&
      restoredSheet.rowCount > 0 &&
      restoredSheet.colCount > 0
    ) {
      const rMax = restoredSheet.rowCount - 1;
      const cMax = restoredSheet.colCount - 1;
      const s = remembered.selection;
      restoredSel = {
        r: Math.max(0, Math.min(s.r, rMax)),
        c: Math.max(0, Math.min(s.c, cMax)),
        r2: Math.max(0, Math.min(s.r2, rMax)),
        c2: Math.max(0, Math.min(s.c2, cMax)),
      };
    }
    setSelection(restoredSel);
    // Restore grid scroll AFTER React commits the Grid render for the
    // restored sheet — setting scrollTop synchronously here would
    // immediately get overwritten by Grid mounting empty / rebuilding its
    // table. rAF runs after the next paint, by which point the table has
    // its full pixel height and the scroller can accept the offset.
    if (remembered && (remembered.scrollTop > 0 || remembered.scrollLeft > 0)) {
      requestAnimationFrame(() => {
        const el = gridScrollRef.current;
        if (!el) return;
        el.scrollTop = remembered.scrollTop;
        el.scrollLeft = remembered.scrollLeft;
      });
    }
    return () => {
      // Capture before tear-down. activeSheetIdx via the ref because the
      // cleanup closure's `activeSheetIdx` is always the stale initial 0;
      // scroll positions read directly from the DOM since the element is
      // still attached at cleanup time.
      const el = gridScrollRef.current;
      tabMemory.set(tab.id, {
        sheetIdx: activeSheetIdxRef.current,
        selection: selectionRef.current,
        scrollTop: el?.scrollTop ?? 0,
        scrollLeft: el?.scrollLeft ?? 0,
      });
    };
  }, [tab.id]); // eslint-disable-line react-hooks/exhaustive-deps

  useUndoShortcuts({
    rootSelector: '[data-xlsx-editor-root]',
    undo: () => undoApi.undo(),
    redo: () => undoApi.redo(),
  });

  // Push every model change back to the workspace store as serialized
  // bytes. We coalesce on a 0-ms timer so a burst of commits (rapid
  // Tab/Enter typing through a row) only triggers one serialize at the
  // end of the burst rather than per-commit — `serializeXlsx` rebuilds
  // the entire workbook binary and was visibly stalling rapid typing on
  // workbooks with formulas / many sheets.
  //
  // Save flow: handleSave reads `tab.data` from the workspace store, so
  // a pending writeBack must flush *before* save runs. The keyboard
  // save-shortcut and menu-Save go through `serializeForSave()` on a
  // user-event tick that comes after our timer fires (timers run before
  // the next click/key event), so a normal Save Just Works. The unmount
  // flush below covers tab-switch-then-save corner cases.
  const lastSerialized = useRef<XlsxModel | null>(null);
  const writebackTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const pendingModelRef = useRef<XlsxModel | null>(null);
  useEffect(() => {
    if (lastSerialized.current === model) return;
    pendingModelRef.current = model;
    if (writebackTimerRef.current !== null) return; // already scheduled
    writebackTimerRef.current = setTimeout(() => {
      writebackTimerRef.current = null;
      const m = pendingModelRef.current;
      pendingModelRef.current = null;
      if (!m || lastSerialized.current === m) return;
      lastSerialized.current = m;
      // writeBack is async (image injection requires JSZip). Fire-and-forget
      // the promise — the next save will pick up the bytes once patchTab
      // runs. Errors are swallowed with the same policy as before.
      writeBack(m).catch(() => {
        /* swallow — serializer errors surface via the next user action. */
      });
    }, 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [model]);

  // Flush any pending writeBack on unmount so save / tab-swap doesn't
  // see stale tab.data.
  useEffect(() => {
    return () => {
      if (writebackTimerRef.current !== null) {
        clearTimeout(writebackTimerRef.current);
        writebackTimerRef.current = null;
      }
      const m = pendingModelRef.current;
      pendingModelRef.current = null;
      if (m && lastSerialized.current !== m) {
        lastSerialized.current = m;
        // Fire-and-forget on unmount: we can't await inside a React cleanup,
        // and the bytes still land in the workspace store via patchTab once
        // the promise resolves. Same behaviour as the timer body.
        writeBack(m).catch(() => {
          /* swallow */
        });
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Register with the global editor-flush registry so a save kicked off
  // before the 0-ms writeback timer fires still snapshots the latest
  // bytes. The 0-ms timer normally drains via the event loop before any
  // save IPC, but explicit registration removes the implicit ordering
  // assumption — pptx/docx already do this via their own flush hooks.
  useEffect(() => {
    return registerEditorFlush(async () => {
      if (writebackTimerRef.current !== null) {
        clearTimeout(writebackTimerRef.current);
        writebackTimerRef.current = null;
      }
      const m = pendingModelRef.current;
      if (!m || lastSerialized.current === m) return;
      lastSerialized.current = m;
      pendingModelRef.current = null;
      try {
        await writeBack(m);
      } catch {
        /* swallow — same policy as the timer body. */
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  const sheet = model.sheets[activeSheetIdx] ?? model.sheets[0];

  /**
   * Same race guard as DocxEditor.flushGenRef / PptxEditor.flushGenRef
   * (R158). Without this, the image-bearing async path lost user data:
   *   t=0    writeBack A starts, sheet has images → await injectXlsxImages
   *   t=…    user edits another cell → re-render → useEffect → writeBack B
   *          starts. If B's sheet has no images, B is fully sync:
   *          patchTab(B bytes).
   *   t=…    A's await completes → patchTab(A bytes).  ← STALE overwrite
   * Only the async branch (line 530) actually races, but bumping
   * unconditionally keeps the sequencing invariant clean: latest call
   * always wins, regardless of which branch each took.
   */
  const flushGenRef = useRef(0);

  // R253 — track the bytes the editor last wrote out so the external-change
  // re-parse effect below can tell self-induced tab.data changes (writeBack)
  // apart from external ones (AI Apply, undo, redo). See effect doc-block.
  // Initialize from the FIRST `tab.data` so the re-parse effect's identity
  // check skips on initial mount (the useUndoableState lazy initializer +
  // useEffect[tab.id] already handle the first parse with full setup —
  // running R253's effect on mount too would mean a redundant third parse
  // and an extra resetHistory). useRef's initializer is only consulted on
  // the first render, so subsequent tab.data changes don't reset this ref
  // back to the initial bytes. Editor remounts on tab swap (EditorSurface's
  // ErrorBoundary key={active.id}), so each tab id gets a fresh ref seeded
  // with that tab's bytes.
  const lastWrittenBytesRef = useRef<Uint8Array>(tab.data);

  const writeBack = async (next: XlsxModel): Promise<void> => {
    const myGen = ++flushGenRef.current;
    let bytes = serializeXlsx(next, initialBytesRef.current);
    // Floating images aren't modeled by SheetJS / xlsx-js-style — inject
    // them into the freshly-written bytes ourselves. Only async path; pure
    // text / style edits skip the JSZip round-trip entirely.
    if (next.sheets.some((s) => (s.images?.length ?? 0) > 0)) {
      bytes = await injectXlsxImages(bytes, next);
      // Drop stale result if a newer writeBack has started while we awaited
      // the JSZip round-trip. Sync-path callers (no images) don't reach the
      // await so they always commit on the same task — the guard only
      // matters here.
      if (myGen !== flushGenRef.current) return;
    }
    // R253 — record the bytes we're about to patch so the re-parse effect
    // recognises this tab.data update as self-induced and skips re-parsing.
    // Set BEFORE patchTab because patchTab triggers a synchronous Zustand
    // setState which fires the React subscription immediately; the re-parse
    // effect runs in a microtask after the commit and reads this ref then.
    lastWrittenBytesRef.current = bytes;
    patchTab(tab.id, { data: bytes });
  };

  // R253 — re-parse tab.data when it changes externally (AI Apply on a
  // binary_replace op, undo / redo via App.tsx handleUndo / handleRedo).
  // Without this, the editor's local `model` state stays at the OLD parse
  // forever (the initial useEffect[tab.id] only fires on tab swap), so:
  //   1. Display drift: AI sets cell A1=X, Apply commits tab.data → NEW
  //      bytes. Editor renders cells from `model` (OLD) so A1 still shows
  //      its prior value. The user's "did the AI apply?" check returns
  //      a false negative.
  //   2. Silent revert (the load-bearing data-loss): user clicks B5 to
  //      type "hello" → setModel → writeBack → serializeXlsx(model_OLD +
  //      B5, initialBytesRef=OLD_BYTES) → bytes have OLD A1 + new B5 →
  //      patchTab(tab.id, {data: bytes}) → tab.data is now OLD-A1 + B5.
  //      The AI's A1=X mod is GONE. The user types one cell and the
  //      AI's edit silently reverts.
  // Same shape applies to handleUndo/Redo's tab.data mutations: they
  // call applyChangeset/undoChangeset for binary_replace ops on the
  // active tab, replace tab.data, and the editor never re-parses.
  // The fallback was "tab swap → swap back" which forces a remount via
  // the EditorBoundary key — but no user expects to do that.
  //
  // Skip when tab.data === lastWrittenBytesRef.current — that's our own
  // writeBack landing back in the prop. Re-parsing on every keystroke
  // round-trip would be prohibitively expensive AND wipe the user's
  // in-progress local undo history (resetHistory below) on every keystroke.
  // The reference-equality check is sufficient because writeBack passes
  // the SAME Uint8Array object to patchTab; Zustand stores it as-is, and
  // the next render sees the same object identity. External writers
  // (applyChangeset / undoChangeset) construct fresh Uint8Array objects,
  // so identity diverges. (Even if they happened to pass the same
  // identity by accident, the bytes content would match — re-parsing
  // would be a redundant but correct no-op.)
  //
  // Same data-loss shape exists in DocxEditor.tsx:255-306 and
  // PptxEditor.tsx:210-257 — they parse on tab.id change only with the
  // same writeBack-mutates-tab.data feedback loop. R253's fix here is
  // the canonical shape; DocxEditor / PptxEditor are next-round
  // candidates with the same lastWrittenBytesRef + post-write
  // re-parse-effect pattern.
  useEffect(() => {
    if (tab.data === lastWrittenBytesRef.current) return;
    // R277 — parse FIRST, mutate initialBytesRef only on success. The original
    // order updated initialBytesRef before parseXlsx, so a parse throw (rare
    // but real: cross-version SheetJS round-trip on an old AI binary_replace's
    // `before` bytes, malformed cell metadata from a future tool, corrupt
    // undo entry) left initialBytesRef pointing at unparseable bytes. The
    // ref's contract is "valid xlsx bytes serializeXlsx can use as template"
    // — once it's poisoned, every subsequent writeBack throws inside
    // `XLSXStyle.read(template)`, and writeBack's `.catch(() => {})` (line
    // ~466) swallows it. Net effect: the user types in cells, sees their
    // edits on screen, but patchTab never fires → tab.data stays at the
    // unparseable bytes, save lands those on disk, and the user's edits
    // are silently lost with no notify / setError / console clue.
    //
    // Parse-then-commit reorder + try/catch around the parse keeps the ref
    // pointing at a known-valid byte stream (the previous successful parse's
    // bytes) so writeBack can continue producing usable output. A notify
    // surfaces the parse failure so the user knows the most recent external
    // change (AI Apply, undo) couldn't be loaded — they can undo it,
    // investigate, or just keep working (their edits will overwrite tab.data
    // with OLD+typing, effectively rejecting the malformed external change).
    try {
      const parsed = parseXlsx(tab.data);
      const recomputed = { ...parsed, sheets: recomputeAllFormulas(parsed.sheets) };
      initialBytesRef.current = tab.data;
      undoApi.resetHistory(recomputed);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`解析 xlsx 失敗：${msg}`, 'error');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.data]);

  /**
   * Stamp `text` into a cell, switching it between literal / formula based on
   * whether it begins with `=`. Returns the *patched-but-not-recomputed* sheet
   * — we always recompute the entire workbook in `commitCell` after the patch
   * lands, so dependents update too.
   */
  const writeCell = (cell: XlsxCell, text: string): XlsxCell => {
    if (isFormulaSource(text)) {
      // For a formula the visible `text` is the computed value, but until
      // recompute runs we just echo the source so the cell isn't briefly
      // blank. The recompute pass overwrites this with the actual result.
      return { ...cell, text, formula: text };
    }
    // Plain literal — explicitly drop any prior formula so the cell stops
    // being treated as one.
    const { formula: _drop, ...rest } = cell;
    void _drop;
    return { ...rest, text };
  };

  const commitCell = (r: number, c: number, text: string) => {
    setModel((prev) => {
      const patched: XlsxModel = {
        ...prev,
        sheets: prev.sheets.map((s, i) => {
          if (i !== activeSheetIdx) return s;
          const cells = s.cells.map((row, rr) =>
            rr === r ? row.map((cell, cc) => (cc === c ? writeCell(cell, text) : cell)) : row,
          );
          return { ...s, cells };
        }),
      };
      // Recompute every formula so dependents pick up the change. Costs an
      // O(formulas × passes) sweep per keystroke-commit; sheets are small in
      // MVP and recompute returns early once nothing changed (max 8 passes).
      // writeBack happens via the model-change effect so undo / redo also flush.
      return { ...patched, sheets: recomputeAllFormulas(patched.sheets) };
    });
  };

  const updateStyle = (patch: (current: XlsxCellStyle | undefined) => XlsxCellStyle | undefined) => {
    if (!selection) return;
    // Apply across the entire selected rectangle so multi-cell range select
    // bulk-formats like Excel. Single-cell selections collapse to one cell.
    const { r1, c1, r2, c2 } = rangeOf(selection);
    setModel((prev) => {
      const next: XlsxModel = {
        ...prev,
        sheets: prev.sheets.map((s, i) => {
          if (i !== activeSheetIdx) return s;
          const cells = s.cells.map((row, rr) =>
            rr >= r1 && rr <= r2
              ? row.map((cell, cc) => {
                  if (cc < c1 || cc > c2) return cell;
                  const newStyle = patch(cell.style);
                  return { ...cell, style: newStyle };
                })
              : row,
          );
          return { ...s, cells };
        }),
      };
      // writeBack happens via the model-change effect so undo / redo also flush.
      return next;
    });
  };

  /** Apply a structural sheet op (insert/delete row/col), keeping selection sane. */
  const mutateSheet = (op: (s: XlsxSheet) => XlsxSheet, nextSel?: (s: Selection | null) => Selection | null) => {
    setModel((prev) => {
      const next: XlsxModel = {
        ...prev,
        sheets: prev.sheets.map((s, i) => (i === activeSheetIdx ? op(s) : s)),
      };
      // writeBack happens via the model-change effect so undo / redo also flush.
      return next;
    });
    if (nextSel) setSelection((cur) => nextSel(cur));
  };

  // Workbook-level sheet operations: add / rename / delete / duplicate. They
  // route through setModel so undo captures them, and reset selection because
  // the new active sheet's coordinate system is unrelated to the old one.
  // Snapshot the *current* sheet's selection + scroll into the per-sheet
  // memory maps before any structural op shifts indices around. Mirrors the
  // contract handleSheetSwitch / locateFindResult enforce, so structural ops
  // don't silently drop the leaving sheet's state.
  const snapshotCurrentSheetMemory = () => {
    const cur = selectionRef.current;
    if (cur) sheetSelectionsRef.current.set(activeSheetIdx, cur);
    else sheetSelectionsRef.current.delete(activeSheetIdx);
    const grid = gridScrollRef.current;
    if (grid) {
      sheetScrollsRef.current.set(activeSheetIdx, {
        scrollTop: grid.scrollTop,
        scrollLeft: grid.scrollLeft,
      });
    }
  };
  // Re-key the per-sheet maps when sheets shift positions. `transform`
  // returns the new index for an entry, or null to drop it. Reassigning
  // `current` is fine — these refs are private to this component and read
  // synchronously from event handlers; nothing else holds the prior Map.
  const remapSheetMemory = (transform: (idx: number) => number | null) => {
    const reindex = <V,>(m: Map<number, V>): Map<number, V> => {
      const out = new Map<number, V>();
      for (const [k, v] of m) {
        const nk = transform(k);
        if (nk !== null) out.set(nk, v);
      }
      return out;
    };
    sheetSelectionsRef.current = reindex(sheetSelectionsRef.current);
    sheetScrollsRef.current = reindex(sheetScrollsRef.current);
  };

  const handleAddSheet = () => {
    // Append at end → existing sheet indices are unchanged. Just snapshot
    // the leaving sheet so a later round-trip back to it via SheetTabs
    // restores selection + scroll instead of resetting to A1 / top.
    snapshotCurrentSheetMemory();
    setModel((prev) => addSheet(prev).model);
    setActiveSheetIdx(model.sheets.length); // appended at end
    setSelection(null);
  };
  // Commit a rename produced by SheetTabs' inline editor. The editor handles
  // empty / unchanged / duplicate validation locally so it can stay open and
  // show a hint instead of toggling out via alert + lost edit. We just trust
  // the value here.
  const handleRenameSheetCommit = (idx: number, newName: string) => {
    setModel((prev) => renameSheet(prev, idx, newName));
  };
  const handleDeleteSheet = (idx: number) => {
    // Last-sheet guard: kept as a defensive no-op because the X button and
    // the right-click "刪除工作表" item are both disabled when only one
    // sheet remains, so this branch isn't reachable from the UI any more
    // (Round 26). Earlier this branch fired window.alert — jarring native
    // modal that didn't match the rest of the themed UI; UI-level disable
    // is more discoverable and removes the alert entirely.
    if (model.sheets.length <= 1) return;
    const name = model.sheets[idx]?.name ?? '';
    // Async path through main-process native dialog. Renderer
    // `window.confirm()` on Windows leaves OS focus broken so the cell
    // input we land on after the deletion can't accept typing — same
    // bug class as the new-workspace flow (see App.tsx::handleNew).
    void (async () => {
      // R293 — wrap the confirm-reject path. Same R288/R289/R290/R292
      // idiom for sibling app.confirm callsites. Without this, a main
      // IPC anomaly during the confirm dialog leaves the user with no
      // toast and no indication of whether the delete actually fired.
      try {
        if (!(await window.gendoc.app.confirm(`確定要刪除工作表「${name}」？`))) return;
        snapshotCurrentSheetMemory();
        setModel((prev) => deleteSheet(prev, idx));
      // Drop the deleted sheet's memory; sheets after it shift down by 1.
      remapSheetMemory((k) => (k === idx ? null : k > idx ? k - 1 : k));
      // Clamp the active index so it doesn't dangle past the new last sheet.
      const newIdx =
        activeSheetIdx < idx
          ? activeSheetIdx
          : activeSheetIdx === idx
            ? Math.max(0, idx - 1)
            : activeSheetIdx - 1;
      setActiveSheetIdx(newIdx);
      // Restore the destination sheet's remembered selection + scroll. Same
      // contract as handleSheetSwitch — landing on an existing sheet should
      // resume where the user left off, not reset to A1 / top.
      setSelection(sheetSelectionsRef.current.get(newIdx) ?? null);
      const target = sheetScrollsRef.current.get(newIdx) ?? { scrollTop: 0, scrollLeft: 0 };
      requestAnimationFrame(() => {
        const next = gridScrollRef.current;
        if (!next) return;
        next.scrollTop = target.scrollTop;
        next.scrollLeft = target.scrollLeft;
      });
      } catch (err) {
        notify(`刪除工作表失敗：${(err as Error).message}`, 'error');
      }
    })();
  };
  // Reorder sheets via SheetTabs drag-and-drop. Excel users heavily rely on
  // tab-drag for re-ordering (e.g. moving a "Summary" sheet to the front of
  // a 12-sheet workbook); chevron-only reorder doesn't exist here, so the
  // gap was: rename → delete-and-recreate at the new position. The model
  // operation is the same shape as the row/col `splice` helpers — see
  // moveSheet in xlsx-adapter. The per-sheet memory Maps (selection, scroll)
  // need re-keying because their keys are positional indices, not stable
  // ids; the same remapSheetMemory pipeline that handles delete/duplicate
  // covers this.
  const handleMoveSheet = (from: number, to: number) => {
    if (from === to) return;
    if (from < 0 || from >= model.sheets.length) return;
    if (to < 0 || to >= model.sheets.length) return;
    snapshotCurrentSheetMemory();
    setModel((prev) => moveSheet(prev, from, to));
    // Build the index transform that mirrors the array splice: the moved
    // sheet's old index `from` becomes `to`; everything in [from+1..to]
    // shifts down by 1 (forward move) or [to..from-1] shifts up by 1
    // (backward move). Anything outside the moved span keeps its index.
    remapSheetMemory((k) => {
      if (k === from) return to;
      if (from < to) {
        // forward move: indices in (from, to] shift down by 1
        if (k > from && k <= to) return k - 1;
      } else {
        // backward move: indices in [to, from) shift up by 1
        if (k >= to && k < from) return k + 1;
      }
      return k;
    });
    // Active sheet should follow the user's drag — they grabbed a tab and
    // moved it; landing on a different sheet would feel like the gesture
    // selected the wrong tab. Use the same transform so any sheet (not just
    // the moved one) stays anchored to its new position.
    setActiveSheetIdx((cur) => {
      if (cur === from) return to;
      if (from < to) {
        if (cur > from && cur <= to) return cur - 1;
      } else {
        if (cur >= to && cur < from) return cur + 1;
      }
      return cur;
    });
  };
  const handleDuplicateSheet = (idx: number) => {
    // Duplicate inserts the copy at idx+1; sheets at idx+1+ shift up by 1.
    // Snapshot first, remap, then land on the brand-new copy (no memory).
    snapshotCurrentSheetMemory();
    setModel((prev) => duplicateSheet(prev, idx).model);
    remapSheetMemory((k) => (k > idx ? k + 1 : k));
    setActiveSheetIdx(idx + 1);
    setSelection(null);
  };

  // Insert / delete row & col operations operate on the selection's anchor
  // (single-row / single-col semantics) so multi-row range select doesn't
  // accidentally trigger bulk insertions. They reset r2/c2 to the new anchor.
  const insertRowAbove = () => {
    if (!selection) return;
    mutateSheet((s) => insertRowAt(s, selection.r));
  };
  const insertRowBelow = () => {
    if (!selection) return;
    mutateSheet(
      (s) => insertRowAt(s, selection.r + 1),
      (cur) => (cur ? { r: cur.r + 1, c: cur.c, r2: cur.r + 1, c2: cur.c } : cur),
    );
  };
  const removeSelectedRow = () => {
    if (!selection) return;
    mutateSheet(
      (s) => deleteRowAt(s, selection.r),
      (cur) => {
        if (!cur) return cur;
        const newR = Math.min(cur.r, sheet.rowCount - 2);
        return newR >= 0 ? { r: newR, c: cur.c, r2: newR, c2: cur.c } : null;
      },
    );
  };
  const insertColLeft = () => {
    if (!selection) return;
    mutateSheet((s) => insertColAt(s, selection.c));
  };
  const insertColRight = () => {
    if (!selection) return;
    mutateSheet(
      (s) => insertColAt(s, selection.c + 1),
      (cur) => (cur ? { r: cur.r, c: cur.c + 1, r2: cur.r, c2: cur.c + 1 } : cur),
    );
  };
  const removeSelectedCol = () => {
    if (!selection) return;
    mutateSheet(
      (s) => deleteColAt(s, selection.c),
      (cur) => {
        if (!cur) return cur;
        const newC = Math.min(cur.c, sheet.colCount - 2);
        return newC >= 0 ? { r: cur.r, c: newC, r2: cur.r, c2: newC } : null;
      },
    );
  };

  /**
   * Merge the current selection rectangle into a single cell. Excel's
   * default behaviour is "keep top-left text, blank the rest" — we do the
   * same so subsequent unmerge doesn't resurrect duplicate values. Refuses
   * single-cell selections (nothing to merge).
   */
  const mergeSelected = () => {
    if (!selection) return;
    const { r1, c1, r2, c2 } = rangeOf(selection);
    if (r1 === r2 && c1 === c2) return;
    mutateSheet((s) => {
      // Blank every covered (non-anchor) cell so the merge result is clean.
      const cleared: XlsxSheet = {
        ...s,
        cells: s.cells.map((row, rr) =>
          rr >= r1 && rr <= r2
            ? row.map((cell, cc) =>
                cc >= c1 && cc <= c2 && !(rr === r1 && cc === c1) ? { ...cell, text: '' } : cell,
              )
            : row,
        ),
      };
      return mergeRange(cleared, r1, c1, r2, c2);
    });
  };

  /** Drop any merge whose anchor matches the selection's anchor. */
  const unmergeSelected = () => {
    if (!selection) return;
    mutateSheet((s) => unmergeAt(s, selection.r, selection.c));
  };

  /**
   * Pick an image file and anchor it at the current selection (or A1 if
   * nothing is selected). The image is appended to the active sheet's
   * `images` array; on save we re-emit the OOXML drawing parts via
   * `injectXlsxImages`. Width is clamped to DEFAULT_MAX_WIDTH_PX so a 4K
   * screenshot doesn't immediately swamp the grid.
   */
  /**
   * Stamp a File into the active sheet's images list at the given anchor.
   * Shared by the toolbar button (anchor = selection.r/c) and the grid's
   * native HTML5 drop handler (anchor = cell under cursor at drop time).
   * Width is clamped to DEFAULT_MAX_WIDTH_PX so a 4K screenshot doesn't
   * immediately swamp the grid; aspect ratio preserved.
   */
  const insertImageAt = async (file: File, anchorRow: number, anchorCol: number) => {
    let ext: XlsxImage['ext'];
    if (file.type === 'image/jpeg') ext = 'jpg';
    else if (file.type === 'image/gif') ext = 'gif';
    else if (file.type === 'image/bmp') ext = 'bmp';
    else ext = 'png';
    const mime = file.type || `image/${ext === 'jpg' ? 'jpeg' : ext}`;
    // `file.arrayBuffer()` can reject when the OS revokes the File handle
    // between selection and read (clipboard image whose source page navigated
    // away, OS picker file deleted/locked, oversized blob hitting OOM). Three
    // call sites — handleAddImage (toolbar), the paste handler at ~line 1255,
    // and the drop handler at ~line 1551 — all enter via `void insertImageAt`,
    // so an unhandled rejection bubbles to the unhandledrejection handler with
    // no UI feedback ("button does nothing"). Mirrors PptxEditor.insertPictureFile
    // (line 437) and DocxEditor.insertImageBlockFromFile (Round 15) — the last
    // editor missing this guard. `readNaturalSize` is NOT a failure path here
    // (its img.onerror falls back to default dimensions by design), so we only
    // need to wrap the byte-read step.
    let buf: Uint8Array;
    try {
      buf = new Uint8Array(await file.arrayBuffer());
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`插入圖片失敗：${msg}`, 'error');
      return;
    }
    const dataUrl = `data:${mime};base64,${uint8ToBase64(buf)}`;
    const natural = await readNaturalSize(dataUrl);
    let widthPx = natural.width;
    let heightPx = natural.height;
    if (widthPx > DEFAULT_MAX_WIDTH_PX) {
      const scale = DEFAULT_MAX_WIDTH_PX / widthPx;
      widthPx = DEFAULT_MAX_WIDTH_PX;
      heightPx = Math.max(1, Math.round(heightPx * scale));
    }
    const newImage: XlsxImage = {
      id: genImageId(),
      data: buf,
      ext,
      mime,
      anchorRow,
      anchorCol,
      widthEmu: widthPx * EMU_PER_PX,
      heightEmu: heightPx * EMU_PER_PX,
      dataUrl,
    };
    setModel((prev) => ({
      ...prev,
      sheets: prev.sheets.map((s, i) =>
        i === activeSheetIdx ? { ...s, images: [...(s.images ?? []), newImage] } : s,
      ),
    }));
    // Auto-select the freshly inserted image so resize / move handles show
    // immediately — matches native Office behaviour where a fresh paste is
    // pre-selected for follow-up positioning.
    setSelectedImageId(newImage.id);
  };

  const handleAddImage = async () => {
    const file = await pickImageFile();
    if (!file) return;
    await insertImageAt(file, selection?.r ?? 0, selection?.c ?? 0);
  };

  const handleRemoveImage = (id: string) => {
    setModel((prev) => ({
      ...prev,
      sheets: prev.sheets.map((s, i) => {
        if (i !== activeSheetIdx) return s;
        const next = (s.images ?? []).filter((img) => img.id !== id);
        return { ...s, images: next.length > 0 ? next : undefined };
      }),
    }));
    setSelectedImageId((cur) => (cur === id ? null : cur));
  };

  /**
   * Patch a single image's mutable fields. Used by the in-grid overlay's
   * drag (anchorRow / anchorCol) and resize (widthEmu / heightEmu). Each
   * call captures into the undo stack — drag / resize commit *once* at
   * gesture end so a single Ctrl+Z reverts the whole move.
   */
  const patchImage = (id: string, patch: Partial<XlsxImage>) => {
    setModel((prev) => ({
      ...prev,
      sheets: prev.sheets.map((s, i) =>
        i === activeSheetIdx
          ? {
              ...s,
              images: (s.images ?? []).map((img) => (img.id === id ? { ...img, ...patch } : img)),
            }
          : s,
      ),
    }));
  };

  // Selected overlay image — highlighted in the grid, deletable via the
  // Delete / Backspace key. Lives in editor state (not the model) so undo
  // doesn't restore stale "selected" markers.
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null);

  // Active sheet swap or image disappearing → drop selection. Avoids a
  // dangling selectedImageId when the model loses the referenced image.
  useEffect(() => {
    if (!selectedImageId) return;
    const exists = sheet.images?.some((img) => img.id === selectedImageId);
    if (!exists) setSelectedImageId(null);
  }, [activeSheetIdx, sheet.images, selectedImageId]);

  // Delete / Backspace on a selected overlay image → remove. Scoped to the
  // editor root so it doesn't fire while typing into a cell.
  useEffect(() => {
    if (!selectedImageId) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Delete' && e.key !== 'Backspace') return;
      const ae = document.activeElement as HTMLElement | null;
      // Don't hijack Delete while a cell input is focused — that would
      // delete the whole image when the user really meant to clear text.
      if (ae?.tagName === 'INPUT' || ae?.isContentEditable) return;
      if (!ae?.closest?.('[data-xlsx-editor-root]')) return;
      e.preventDefault();
      handleRemoveImage(selectedImageId);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedImageId]);

  /**
   * Selection rectangle straddles at least one merge — used to enable the
   * "Merge" button only when a real merge is possible (>1 cell selected).
   * The "Unmerge" button is enabled when the anchor sits on a merge anchor.
   */
  const canMerge = selection && (selection.r !== selection.r2 || selection.c !== selection.c2);
  const canUnmerge = !!(selection && mergeAtAnchor(sheet, selection.r, selection.c));

  const selectedCell: XlsxCell | null =
    selection && sheet.cells[selection.r]?.[selection.c] ? sheet.cells[selection.r][selection.c] : null;
  // Anchor address (for the formula bar — always a single cell, even when
  // a range is selected, since the formula bar edits one cell).
  const anchorAddr = selection ? `${colIndexToLetter(selection.c)}${selection.r + 1}` : null;
  // Range address for the toolbar / status bar — "A1" for single-cell or
  // "A1:C3" for a range.
  const selectedAddr = selection ? rangeAddr(selection) : null;

  // Ctrl/Cmd+B / +I / +U — flips the matching style key on the active cell.
  // Mirrors the toolbar `toggle()` logic but lifted to editor scope so the
  // hook can dispatch it.
  useFormatShortcuts({
    rootSelector: '[data-xlsx-editor-root]',
    isActive: () => !!selection,
    toggle: (key) =>
      updateStyle((cur) => normalizeStyle({ ...(cur ?? {}), [key]: !cur?.[key] })),
  });

  // Ctrl/Cmd+F: open Find & Replace and (re)focus the query input. We
  // don't toggle-close on Ctrl+F — Esc closes, matching VS Code / Chrome
  // convention. `findFocusNonce` lets the dialog re-focus its input on
  // every Ctrl+F so users who clicked into a cell to navigate a match
  // can hit Ctrl+F to come back to the query field.
  const [findOpen, setFindOpen] = useState(false);
  const [findFocusNonce, setFindFocusNonce] = useState(0);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'f') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-xlsx-editor-root]')) return;
      e.preventDefault();
      setFindOpen(true);
      setFindFocusNonce((n) => n + 1);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // Ctrl/Cmd+G: focus the Name Box. PptxEditor and DocxEditor both bind
  // Ctrl+G to a GoToDialog for slide / paragraph navigation; the muscle
  // memory broke when switching to Excel because nothing was wired up.
  // The Name Box already accepts cell addresses ("B5", "AA12") and ranges,
  // so rather than introduce a second dialog we just send focus there —
  // selecting the existing text means the user can immediately overtype.
  // Same scope guard as Ctrl+F: only fires when focus is somewhere inside
  // the Excel editor surface.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'g') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-xlsx-editor-root]')) return;
      const nameBox = document.querySelector<HTMLInputElement>('[data-xlsx-namebox]');
      if (!nameBox) return;
      e.preventDefault();
      nameBox.focus();
      nameBox.select();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // ── Clipboard: copy / cut / paste / delete-range ─────────────────────
  // Native ClipboardEvents bubble from the focused input to the editor
  // root. We intercept *only* when the action is range-shaped — single
  // cells with no internal text-selection overlap fall through to the
  // input's native text editing so users can still copy / paste *inside*
  // a cell (e.g. partial text within a long string).
  //
  // The "cut origin" ref records the source range when Ctrl+X fired;
  // the next paste blanks those cells. Paste from a different app
  // ignores the ref (different content on the system clipboard).
  const cutOriginRef = useRef<
    | { sheetIdx: number; r1: number; c1: number; r2: number; c2: number; signature: string }
    | null
  >(null);

  // Canonical "is the user actually typing into a cell?" signal. CellInput
  // toggles this via onEditingChange; clipboard handlers consult it instead
  // of guessing from input.value vs cell.text. The ref form survives across
  // setState/render boundaries so paste handlers see the current state.
  const editingCellRef = useRef(false);

  /**
   * If the destination selection is bigger than a 1×1 paste, replicate the
   * source cell to fill the selected rectangle. Same as Excel's "select
   * range, paste single value" behavior. Multi-cell payloads pass through
   * untouched — partial-tiling logic (paste 1×3 over 2×3) is not implemented.
   */
  const maybeTilePayload = (
    payload: RichClipboardPayload,
    sel: Selection | null,
  ): RichClipboardPayload => {
    if (!sel) return payload;
    const rows = payload.cells.length;
    const cols = payload.cells[0]?.length ?? 0;
    if (rows !== 1 || cols !== 1) return payload;
    const { r1, c1, r2, c2 } = rangeOf(sel);
    const tr = r2 - r1 + 1;
    const tc = c2 - c1 + 1;
    if (tr === 1 && tc === 1) return payload;
    const src = payload.cells[0][0];
    const cells: XlsxCell[][] = [];
    for (let r = 0; r < tr; r += 1) {
      const row: XlsxCell[] = [];
      for (let c = 0; c < tc; c += 1) {
        row.push({ ...src, style: src.style ? { ...src.style } : undefined });
      }
      cells.push(row);
    }
    return { origin: payload.origin, cells };
  };

  // Visual marker for the most recent copy / cut range — shown as a dashed
  // border on the source cells so users can tell the clipboard is "armed".
  // Cleared on Esc, on next paste, or on next mutation.
  const [copyMarker, setCopyMarker] = useState<{
    sheetIdx: number;
    r1: number;
    c1: number;
    r2: number;
    c2: number;
    mode: 'copy' | 'cut';
  } | null>(null);

  /** Stamp a payload into the active sheet at the anchor + recompute. */
  const pasteAtAnchor = (anchorR: number, anchorC: number, payload: RichClipboardPayload): void => {
    // The focused cell input keeps a local copy of its value while
    // `editing` is true; without a blur it would render the pre-paste
    // text even after the model updates. Blur first so the next render
    // re-reads from the cell.
    (document.activeElement as HTMLElement | null)?.blur?.();

    // If the destination range is bigger than a 1×1 paste, tile the
    // single cell to fill the entire selected rectangle — matches Excel
    // when you select 5×3 and paste a single cell over it.
    payload = maybeTilePayload(payload, selection);

    setModel((prev) => {
      const cutOrigin = cutOriginRef.current;
      const next: XlsxModel = {
        ...prev,
        sheets: prev.sheets.map((s, i) => {
          if (i === activeSheetIdx) return applyPaste(s, anchorR, anchorC, payload);
          // If a cut was outstanding on a different sheet, blank that source too.
          if (cutOrigin && i === cutOrigin.sheetIdx) {
            return clearRange(s, cutOrigin.r1, cutOrigin.c1, cutOrigin.r2, cutOrigin.c2);
          }
          return s;
        }),
      };
      // Same-sheet cut: clear the old origin AFTER paste so overlapping
      // ranges (paste over part of the cut origin) keep the pasted data.
      if (cutOrigin && cutOrigin.sheetIdx === activeSheetIdx) {
        next.sheets = next.sheets.map((s, i) => {
          if (i !== activeSheetIdx) return s;
          // Mask out any cells we just wrote into.
          const { r1, c1, r2, c2 } = cutOrigin;
          const writtenRows = payload.cells.length;
          const writtenCols = payload.cells[0]?.length ?? 0;
          return clearRangeExcept(
            s,
            r1,
            c1,
            r2,
            c2,
            anchorR,
            anchorC,
            anchorR + writtenRows - 1,
            anchorC + writtenCols - 1,
          );
        });
      }
      cutOriginRef.current = null;
      // Recompute so any relative-shifted formulas land on real values.
      return { ...next, sheets: recomputeAllFormulas(next.sheets) };
    });
    // Move selection to cover the pasted rectangle so the user can
    // immediately follow up with formatting / arrow nav.
    const writtenRows = payload.cells.length;
    const writtenCols = payload.cells[0]?.length ?? 0;
    setSelection({
      r: anchorR,
      c: anchorC,
      r2: anchorR + writtenRows - 1,
      c2: anchorC + writtenCols - 1,
    });
    // Clear the visual copy marker — paste is the natural end of the
    // "this was copied" cycle (cut already cleared the source above).
    setCopyMarker(null);
  };

  useEffect(() => {
    const root = document.querySelector('[data-xlsx-editor-root]') as HTMLElement | null;
    if (!root) return;

    /**
     * Decide whether to defer a clipboard event to the focused <input>.
     * We step aside when:
     *
     *   1. The cell is in *edit mode* (the user pressed F2, double-clicked,
     *      or typed a printable char to begin editing) — Ctrl+C/X/V should
     *      then act on the text inside the input, not the cell range.
     *
     *   2. The focus is on an input outside our editor (formula bar,
     *      sheet-rename prompt, etc.).
     *
     *   3. The user has a *partial* highlight inside the active input —
     *      they're mid-text-edit and want a substring copy/cut.
     *
     * Read-only "selected" cells are NOT deferred: in select-mode the
     * input is `readOnly` and the whole text is auto-selected, so range
     * copy is the natural intent.
     */
    const shouldDeferToInput = (): boolean => {
      if (editingCellRef.current) return true;
      const ae = document.activeElement as HTMLInputElement | null;
      if (!ae || ae.tagName !== 'INPUT') return false;
      if (!ae.closest('[data-xlsx-editor-root]')) return true; // formula bar etc.
      // Partial text highlight inside the input?
      const start = ae.selectionStart ?? 0;
      const end = ae.selectionEnd ?? 0;
      const len = ae.value.length;
      const noneSelected = start === end;
      const allSelected = start === 0 && end === len;
      return !noneSelected && !allSelected;
    };

    const onCopy = (e: ClipboardEvent) => {
      if (shouldDeferToInput()) return;
      if (!selection || !e.clipboardData) return;
      const { r1, c1, r2, c2 } = rangeOf(selection);
      const cells = extractRange(sheet, r1, c1, r2, c2);
      const payload: RichClipboardPayload = { origin: { r: r1, c: c1 }, cells };
      const tsv = serializeRangeToTsv(cells);
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      e.clipboardData.setData(GENDOC_XLSX_MIME, serializeRangeToJson(payload));
      cutOriginRef.current = null; // copy clears any pending cut
      setCopyMarker({ sheetIdx: activeSheetIdx, r1, c1, r2, c2, mode: 'copy' });
    };

    const onCut = (e: ClipboardEvent) => {
      if (shouldDeferToInput()) return;
      if (!selection || !e.clipboardData) return;
      const { r1, c1, r2, c2 } = rangeOf(selection);
      const cells = extractRange(sheet, r1, c1, r2, c2);
      const payload: RichClipboardPayload = { origin: { r: r1, c: c1 }, cells };
      const tsv = serializeRangeToTsv(cells);
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      e.clipboardData.setData(GENDOC_XLSX_MIME, serializeRangeToJson(payload));
      // Mark the source for later blanking; we don't blank now because
      // the user may paste back in place (cut+paste-in-place = no-op).
      cutOriginRef.current = {
        sheetIdx: activeSheetIdx,
        r1,
        c1,
        r2,
        c2,
        signature: tsv,
      };
      setCopyMarker({ sheetIdx: activeSheetIdx, r1, c1, r2, c2, mode: 'cut' });
    };

    const onPaste = (e: ClipboardEvent) => {
      // Image-on-clipboard short-circuits the cell-paste path. We check
      // *before* shouldDeferToInput so a screenshot pasted while a cell
      // input is focused still lands as an overlay — image bytes have no
      // sensible interpretation as cell text. Anchored to the current
      // selection (or A1 if nothing is selected). Falls through to the
      // existing TSV / JSON range-paste below when no image is present.
      const items = e.clipboardData?.items;
      if (items) {
        for (let i = 0; i < items.length; i += 1) {
          const it = items[i];
          if (it.kind === 'file' && it.type.startsWith('image/')) {
            const file = it.getAsFile();
            if (!file) continue;
            e.preventDefault();
            const r = selection?.r ?? 0;
            const c = selection?.c ?? 0;
            void insertImageAt(file, r, c);
            return;
          }
        }
      }
      if (shouldDeferToInput()) return;
      if (!selection || !e.clipboardData) return;
      const json = e.clipboardData.getData(GENDOC_XLSX_MIME);
      const tsv = e.clipboardData.getData('text/plain');

      let payload: RichClipboardPayload | null = null;
      if (json) {
        try {
          const parsed = JSON.parse(json) as RichClipboardPayload;
          if (parsed && parsed.cells && parsed.origin) payload = parsed;
        } catch {
          /* fall through to TSV */
        }
      }
      if (!payload) {
        if (!tsv) return;
        // Single-line / single-cell text → defer to native so the cell
        // input gets a normal text overwrite. Multi-cell TSV always goes
        // through range-paste regardless of focus state.
        if (!isMultiCellTsv(tsv)) return;
        payload = tsvToPayload(tsv, selection.r, selection.c);
      }
      e.preventDefault();
      pasteAtAnchor(selection.r, selection.c, payload);
    };

    /** Clear-on-Delete for multi-cell selections + Esc to drop the
     *  copy marker. Single-cell delete is handled by the input natively
     *  (we focus + select-all on focus, so Delete blanks the input). */
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && copyMarker) {
        // Don't preventDefault — let the cell input also see Esc to
        // cancel its in-progress edit (existing behavior).
        setCopyMarker(null);
        return;
      }
      if (e.key !== 'Delete') return;
      if (!selection) return;
      const isMultiCell = selection.r !== selection.r2 || selection.c !== selection.c2;
      if (!isMultiCell) return; // let single-cell Backspace/Delete trigger the input's edit-mode-blank flow
      // If the user is editing inside a cell, Delete should act on the
      // input contents, not nuke the entire range.
      if (editingCellRef.current) return;
      e.preventDefault();
      const { r1, c1, r2, c2 } = rangeOf(selection);
      setModel((prev) => {
        const next: XlsxModel = {
          ...prev,
          sheets: prev.sheets.map((s, i) => (i === activeSheetIdx ? clearRange(s, r1, c1, r2, c2) : s)),
        };
        return { ...next, sheets: recomputeAllFormulas(next.sheets) };
      });
    };

    root.addEventListener('copy', onCopy);
    root.addEventListener('cut', onCut);
    root.addEventListener('paste', onPaste);
    root.addEventListener('keydown', onKeyDown);
    return () => {
      root.removeEventListener('copy', onCopy);
      root.removeEventListener('cut', onCut);
      root.removeEventListener('paste', onPaste);
      root.removeEventListener('keydown', onKeyDown);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, selection, activeSheetIdx, copyMarker]);

  // Flatten ALL sheets' cells into searchable segments. Composite id
  // "sheetIdx:r:c" lets onUpdateSegment / onLocate route back to the
  // right sheet. The label embeds the sheet name so the user can tell
  // "A1" on Sheet1 from "A1" on Sheet2 in the result counter.
  const findSegments = useMemo<SearchSegment[]>(() => {
    const out: SearchSegment[] = [];
    model.sheets.forEach((sh, si) => {
      sh.cells.forEach((row, r) => {
        row.forEach((cell, c) => {
          out.push({
            id: `${si}:${r}:${c}`,
            text: cell.text ?? '',
            label: `${sh.name} · ${colIndexToLetter(c)}${r + 1}`,
          });
        });
      });
    });
    return out;
  }, [model.sheets]);

  /** Cross-sheet F&R replacement. Switches to the target sheet if needed,
   * then writes the cell. Reuses the same `commitCell` path as direct
   * editing so styles, undo, and serialization are unchanged. */
  const applyFindReplace = (id: string, text: string): void => {
    const [ss, rs, cs] = id.split(':');
    const si = Number(ss);
    const r = Number(rs);
    const c = Number(cs);
    if (!Number.isFinite(si) || !Number.isFinite(r) || !Number.isFinite(c)) return;
    if (si !== activeSheetIdx) {
      // Switch sheet first; the next commit lands on the now-active sheet.
      setActiveSheetIdx(si);
      // commitCell uses activeSheetIdx via closure; we need to bypass that
      // for cross-sheet writes. Inline a sheet-targeted mutation instead.
      setModel((prev) => {
        const patched: XlsxModel = {
          ...prev,
          sheets: prev.sheets.map((s, i) =>
            i === si
              ? {
                  ...s,
                  cells: s.cells.map((row, rr) =>
                    rr === r ? row.map((cell, cc) => (cc === c ? writeCell(cell, text) : cell)) : row,
                  ),
                }
              : s,
          ),
        };
        const next: XlsxModel = { ...patched, sheets: recomputeAllFormulas(patched.sheets) };
        // Floating image injection is async; this is a write-through path
        // (find/replace cross-sheet stamp) so we just kick the promise off.
        writeBack(next).catch(() => {
          /* swallow */
        });
        return next;
      });
    } else {
      commitCell(r, c, text);
    }
  };

  /** Cross-sheet locate: switch sheets and select the target cell so the
   * underlying input has focus when the dialog navigates. */
  const locateFindResult = (id: string): void => {
    const [ss, rs, cs] = id.split(':');
    const si = Number(ss);
    const r = Number(rs);
    const c = Number(cs);
    if (!Number.isFinite(si) || !Number.isFinite(r) || !Number.isFinite(c)) return;
    if (si !== activeSheetIdx) {
      // Mirror handleSheetSwitch's contract: snapshot the leaving sheet's
      // selection into sheetSelectionsRef before navigating away. Without
      // this, jumping cross-sheet via Find drops the previous sheet's
      // selection — a later SheetTabs click back to it would resume with
      // null instead of the cell the user had highlighted.
      if (selection) sheetSelectionsRef.current.set(activeSheetIdx, selection);
      else sheetSelectionsRef.current.delete(activeSheetIdx);
      // Same idea for scroll: preserve the leaving sheet's scrollTop so
      // bouncing back via SheetTabs lands where the user was. Drop any
      // stored entry for the target sheet — the scrollIntoView below is
      // the source of truth for where Find should land, and a stale
      // stored offset would race against it on the rAF that follows.
      const leaving = gridScrollRef.current;
      if (leaving) {
        sheetScrollsRef.current.set(activeSheetIdx, {
          scrollTop: leaving.scrollTop,
          scrollLeft: leaving.scrollLeft,
        });
      }
      sheetScrollsRef.current.delete(si);
      setActiveSheetIdx(si);
    }
    const target: Selection = { r, c, r2: r, c2: c };
    setSelection(target);
    // Also persist the new selection into the per-sheet map so that a
    // subsequent SheetTabs round-trip away from and back to `si` restores
    // the located cell rather than whatever was there before the locate.
    sheetSelectionsRef.current.set(si, target);
    // Scroll the matched cell into view. Without this, clicking a result
    // beyond the current viewport (long sheet, user scrolled away) just
    // moved the selection rectangle behind the scrollport and the user
    // saw no visible feedback — Find Next felt like a no-op even though
    // the model state had advanced. rAF defers until React has committed
    // the sheet swap / new selection so the target <td> exists in the
    // DOM. `block: 'nearest'` / `inline: 'nearest'` keep already-visible
    // hits from jiggling. We don't focus the input — the dialog still
    // owns keyboard focus so the user can press Enter for Find Next.
    requestAnimationFrame(() => {
      const el = gridScrollRef.current?.querySelector<HTMLElement>(
        `[data-cell-r="${r}"][data-cell-c="${c}"]`,
      );
      el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    });
  };

  return (
    <div data-xlsx-editor-root className="h-full w-full flex flex-col bg-background">
      <Banner />
      <FormatToolbar
        cell={selectedCell}
        selectionAddr={selectedAddr}
        onUpdate={updateStyle}
        onInsertRowAbove={insertRowAbove}
        onInsertRowBelow={insertRowBelow}
        onDeleteRow={removeSelectedRow}
        onInsertColLeft={insertColLeft}
        onInsertColRight={insertColRight}
        onDeleteCol={removeSelectedCol}
        // Mirror the >1 floors enforced by deleteRowAt / deleteColAt
        // (xlsx-adapter.ts:687 / :718) — without these, the buttons stay
        // active on the last row/col and clicking silently no-ops.
        canDeleteRow={sheet.rowCount > 1}
        canDeleteCol={sheet.colCount > 1}
        canMerge={!!canMerge}
        canUnmerge={canUnmerge}
        onMerge={mergeSelected}
        onUnmerge={unmergeSelected}
        onInsertImage={() => {
          void handleAddImage();
        }}
        onOpenFind={() => {
          setFindOpen(true);
          setFindFocusNonce((n) => n + 1);
        }}
      />
      <FormulaBar
        addr={anchorAddr}
        cell={selectedCell}
        rowCount={sheet.rowCount}
        colCount={sheet.colCount}
        onCommit={(text) => {
          if (!selection) return;
          // Compare against the formula source (when present) since that's
          // what the bar is showing — otherwise typing the same formula in
          // would never re-fire the commit.
          const cur = selectedCell?.formula ?? selectedCell?.text ?? '';
          if (text !== cur) commitCell(selection.r, selection.c, text);
        }}
        onJumpToAddr={(r, c) => {
          // Collapse to a single-cell selection at the typed address.
          // Excel's name box also accepts ranges ("A1:B5"); we keep the
          // MVP scope to single-cell because that's the dominant
          // workflow and avoids surprising auto-resizes of the existing
          // anchor.
          setSelection({ r, c, r2: r, c2: c });
          // Scroll the target into view — without this, typing "AA500"
          // moves the selection rectangle behind the scrollport and the
          // jump feels like a no-op. rAF defers until React commits the
          // new selection so the target <td> exists. Mirrors the Find
          // navigation handler above.
          requestAnimationFrame(() => {
            const el = gridScrollRef.current?.querySelector<HTMLElement>(
              `[data-cell-r="${r}"][data-cell-c="${c}"]`,
            );
            el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
          });
        }}
      />
      <SheetTabs
        sheets={model.sheets}
        activeIdx={activeSheetIdx}
        onSelect={handleSheetSwitch}
        onAdd={handleAddSheet}
        onRenameCommit={handleRenameSheetCommit}
        onDelete={handleDeleteSheet}
        onDuplicate={handleDuplicateSheet}
        onReorder={handleMoveSheet}
      />
      <div ref={gridScrollRef} className="relative flex-1 min-h-0 overflow-auto">
        <FindReplaceDialog
          open={findOpen}
          focusNonce={findFocusNonce}
          onClose={() => setFindOpen(false)}
          segments={findSegments}
          onUpdateSegment={applyFindReplace}
          onLocateSegment={locateFindResult}
        />
        <Grid
          sheet={sheet}
          selection={selection}
          copyMarker={copyMarker?.sheetIdx === activeSheetIdx ? copyMarker : null}
          // Click: collapse selection to that cell. If the click landed on
          // a merge anchor, expand the selection rectangle to cover the
          // whole merge so the toolbar's "Unmerge" enables and stats over
          // the merge work as expected.
          onSelect={(r, c) => {
            const m = mergeAtAnchor(sheet, r, c);
            if (m) {
              setSelection({ r: m.r1, c: m.c1, r2: m.r2, c2: m.c2 });
            } else {
              setSelection({ r, c, r2: r, c2: c });
            }
          }}
          // Tab-cycle within the existing rect: move the anchor without
          // collapsing the selection. Falls back to a fresh single-cell
          // selection if nothing is selected yet (defensive).
          onMoveAnchor={(r, c) =>
            setSelection((cur) => (cur ? { ...cur, r, c } : { r, c, r2: r, c2: c }))
          }
          onExtend={(r2, c2) =>
            setSelection((cur) => (cur ? { ...cur, r2, c2 } : { r: r2, c: c2, r2, c2 }))
          }
          onCommit={commitCell}
          onEditingChange={(v) => { editingCellRef.current = v; }}
          images={sheet.images}
          selectedImageId={selectedImageId}
          onImageSelect={setSelectedImageId}
          onImagePatch={patchImage}
          onDropImageFile={(file, r, c) => {
            void insertImageAt(file, r, c);
          }}
        />
      </div>
      {sheet.images && sheet.images.length > 0 ? (
        <ImagePanel
          images={sheet.images}
          onRemove={handleRemoveImage}
          onJump={(r, c) => setSelection({ r, c, r2: r, c2: c })}
        />
      ) : null}
      <StatusBar selection={selection} sheet={sheet} addr={selectedAddr} />
    </div>
  );
}

/**
 * Bottom bar showing range address + Sum / Average / Count for the current
 * selection. Mirrors Excel / Numbers / Sheets — only the *selection-aware*
 * stats; cursor position / zoom belongs elsewhere. When nothing is
 * selected the bar still renders (empty) so the layout doesn't jump.
 */
function StatusBar({
  selection,
  sheet,
  addr,
}: {
  selection: Selection | null;
  sheet: XlsxSheet;
  addr: string | null;
}): JSX.Element {
  if (!selection) {
    return (
      <div className="flex items-center justify-end gap-4 px-3 py-1 border-t bg-secondary/30 text-xs text-muted-foreground">
        <span>未選取</span>
      </div>
    );
  }
  const range = rangeOf(selection);
  const stats = rangeStats(sheet, range);
  const formatNum = (n: number): string => {
    // Up to 6 significant decimals, trimming trailing zeros so 3.5 stays 3.5
    // not 3.500000. Locale separators match the surrounding UI (zh-TW).
    if (!Number.isFinite(n)) return '—';
    const s = n.toLocaleString(undefined, { maximumFractionDigits: 6 });
    return s;
  };
  return (
    <div className="flex items-center justify-end gap-4 px-3 py-1 border-t bg-secondary/30 text-xs text-muted-foreground">
      <span className="font-mono text-foreground">{addr}</span>
      {stats.numericCount > 0 ? (
        <>
          <span>
            加總: <span className="font-mono text-foreground">{formatNum(stats.sum)}</span>
          </span>
          <span>
            平均:{' '}
            <span className="font-mono text-foreground">
              {formatNum(stats.sum / stats.numericCount)}
            </span>
          </span>
        </>
      ) : null}
      <span>
        計數: <span className="font-mono text-foreground">{stats.count}</span>
      </span>
      {stats.cellCount > 1 ? (
        <span>
          儲存格: <span className="font-mono text-foreground">{stats.cellCount}</span>
        </span>
      ) : null}
    </div>
  );
}

function Banner(): JSX.Element {
  return (
    <div className="flex items-center gap-2 px-3 py-1.5 text-xs bg-amber-100 border-b border-amber-300 text-amber-800">
      <AlertTriangle className="h-3.5 w-3.5 shrink-0" />
      <span>
        Excel MVP 編輯：cell 文字／數值 round-trip 保留樣式。粗體 / 對齊 / 顏色寫入 styles.xml；重新開啟後 toolbar 顯示可能不全。
      </span>
    </div>
  );
}

/**
 * Strip showing every image inserted into the active sheet. We don't render
 * images in-grid this round (the SheetJS-backed grid has no overlay layer);
 * instead the panel surfaces a thumbnail + anchor + size + delete control so
 * the user can confirm the inserted state before saving. Clicking the
 * anchor jumps the selection back to the anchor cell.
 */
function ImagePanel({
  images,
  onRemove,
  onJump,
}: {
  images: XlsxImage[];
  onRemove: (id: string) => void;
  onJump: (r: number, c: number) => void;
}): JSX.Element {
  return (
    <div className="flex items-center gap-2 px-3 py-1.5 border-t bg-secondary/30 overflow-x-auto">
      <span className="text-[11px] text-muted-foreground shrink-0">圖片 ({images.length}):</span>
      {images.map((img) => {
        const widthPx = Math.round(img.widthEmu / EMU_PER_PX);
        const heightPx = Math.round(img.heightEmu / EMU_PER_PX);
        const addr = `${colIndexToLetter(img.anchorCol)}${img.anchorRow + 1}`;
        return (
          <div
            key={img.id}
            className="flex items-center gap-1.5 rounded border border-border bg-background px-1.5 py-1 shrink-0"
          >
            {img.dataUrl ? (
              // eslint-disable-next-line @next/next/no-img-element
              <img
                src={img.dataUrl}
                alt=""
                className="h-8 w-8 object-cover rounded-sm border border-border"
              />
            ) : (
              <div className="h-8 w-8 rounded-sm border border-border bg-muted" />
            )}
            <button
              type="button"
              onClick={() => onJump(img.anchorRow, img.anchorCol)}
              className="text-[11px] font-mono text-foreground hover:underline"
              title="跳至錨點儲存格"
            >
              {addr}
            </button>
            <span className="text-[11px] text-muted-foreground tabular-nums">
              {widthPx}×{heightPx}
            </span>
            {/* R120 — cross-editor verb parity. Previously this read
                `移除圖片`, the lone outlier among the three image-delete
                buttons in the codebase:
                  DocxEditor.tsx:3275  Trash2 + `刪除圖片`
                  PptxEditor.tsx:2191  X      + `刪除此文字框`
                  XlsxEditor.tsx:1713  X      + `移除圖片`        ← was here
                handleRemoveImage at line 917-927 filters the image out of
                `sheet.images` entirely — the binary is gone after save, no
                undo pool, no "soft remove" semantic that would justify the
                softer verb. Same content-destructive action as the docx
                Trash2 button (which calls onRemove → strips the floating
                image from `paragraph.runs`). The project reserves 移除 for
                non-content attribute/pointer removal where the underlying
                content survives:
                  DocxEditor.tsx:3420   `移除目前連結` (link annotation,
                                         text stays)
                  SettingsDialog.tsx:278 `移除已儲存的 API key` (stored
                                         credential, not document content)
                The PptxEditor frame button proves the icon-verb axis is
                independent: X icon + 刪除 verb is fine. So the verb here
                aligns with the operation's destructiveness (content gone),
                not the icon shape. */}
            <button
              type="button"
              onClick={() => onRemove(img.id)}
              className="h-5 w-5 inline-flex items-center justify-center rounded text-muted-foreground hover:text-destructive hover:bg-destructive/10"
              title="刪除圖片"
            >
              <X className="h-3 w-3" />
            </button>
          </div>
        );
      })}
    </div>
  );
}

/** Common Excel number-format strings. Empty value = "General" (auto). */
const NUMBER_FORMATS: Array<{ label: string; value: string }> = [
  { label: '一般', value: '' },
  { label: '整數 (1234)', value: '0' },
  { label: '小數 (1234.56)', value: '0.00' },
  { label: '千分位 (1,234)', value: '#,##0' },
  { label: '千分位.小數 (1,234.56)', value: '#,##0.00' },
  { label: '貨幣 NT$', value: '"NT$"#,##0.00' },
  { label: '貨幣 $', value: '"$"#,##0.00' },
  { label: '百分比 (12%)', value: '0%' },
  { label: '百分比.小數 (12.34%)', value: '0.00%' },
  { label: '日期 yyyy-mm-dd', value: 'yyyy-mm-dd' },
  { label: '日期時間 yyyy-mm-dd hh:mm', value: 'yyyy-mm-dd hh:mm' },
];

function FormatToolbar({
  cell,
  selectionAddr,
  onUpdate,
  onInsertRowAbove,
  onInsertRowBelow,
  onDeleteRow,
  onInsertColLeft,
  onInsertColRight,
  onDeleteCol,
  canDeleteRow,
  canDeleteCol,
  canMerge,
  canUnmerge,
  onMerge,
  onUnmerge,
  onInsertImage,
  onOpenFind,
}: {
  cell: XlsxCell | null;
  selectionAddr: string | null;
  onUpdate: (fn: (cur: XlsxCellStyle | undefined) => XlsxCellStyle | undefined) => void;
  onInsertRowAbove: () => void;
  onInsertRowBelow: () => void;
  onDeleteRow: () => void;
  onInsertColLeft: () => void;
  onInsertColRight: () => void;
  onDeleteCol: () => void;
  /** False when the sheet has only one row — adapter would silently no-op. */
  canDeleteRow: boolean;
  /** False when the sheet has only one column — adapter would silently no-op. */
  canDeleteCol: boolean;
  canMerge: boolean;
  canUnmerge: boolean;
  onMerge: () => void;
  onUnmerge: () => void;
  /** Insert image at the current selection's anchor (or A1 when no selection). */
  onInsertImage: () => void;
  onOpenFind: () => void;
}): JSX.Element {
  const style = cell?.style ?? {};
  const disabled = !selectionAddr;

  const toggle = (key: 'bold' | 'italic' | 'underline') =>
    onUpdate((cur) => normalizeStyle({ ...cur, [key]: !cur?.[key] }));

  const setAlign = (align: 'left' | 'center' | 'right') =>
    onUpdate((cur) => normalizeStyle({ ...cur, align: cur?.align === align ? undefined : align }));

  const setColor = (key: 'fontColor' | 'bgColor', hex: string) => {
    const cleaned = hex.replace(/^#/, '').toUpperCase();
    onUpdate((cur) => normalizeStyle({ ...cur, [key]: cleaned || undefined }));
  };

  const clearColor = (key: 'fontColor' | 'bgColor') =>
    onUpdate((cur) => normalizeStyle({ ...cur, [key]: undefined }));

  const setNumberFormat = (value: string) =>
    onUpdate((cur) => normalizeStyle({ ...cur, numberFormat: value || undefined }));

  const setFontFamily = (value: string) =>
    onUpdate((cur) => normalizeStyle({ ...cur, fontFamily: value || undefined }));

  // Wrapper no longer dims the entire bar when there's no selection. The
  // image-insert button (and the Find button below) are always-enabled —
  // applying opacity-50 to the wrapper made them visually inseparable from
  // the disabled style/format buttons. Per-control disabled styling lives
  // on `ToolbarBtn` and the select / color picker individually.
  return (
    <div className="flex items-center gap-0.5 px-2 py-1 border-b bg-secondary/30">
      {/* Bold/italic/underline tooltips spell out keyboard shortcuts that
          the editor binds via useFormatShortcuts. Mirrors DocxEditor /
          PptxEditor / MarkdownToolbar for cross-editor parity.

          R92 — when `disabled` (no cell selected, line 1780), flip the title
          to explain *why* the button is dimmed instead of repeating the
          enabled-state shortcut. Mirrors R87 on DocxEditor:2027-2034 where
          the same Bold/Italic/Underline + align cluster already does this.
          Wording '請先選取一個儲存格' matches the cell-selection vocabulary
          this editor uses elsewhere. */}
      <ToolbarBtn active={!!style.bold} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '粗體 (Ctrl+B)'} onClick={() => toggle('bold')}>
        <Bold className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.italic} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '斜體 (Ctrl+I)'} onClick={() => toggle('italic')}>
        <Italic className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.underline} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '底線 (Ctrl+U)'} onClick={() => toggle('underline')}>
        <Underline className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn active={style.align === 'left'} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '靠左對齊'} onClick={() => setAlign('left')}>
        <AlignLeft className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={style.align === 'center'} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '置中對齊'} onClick={() => setAlign('center')}>
        <AlignCenter className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={style.align === 'right'} disabled={disabled} title={disabled ? '請先選取一個儲存格' : '靠右對齊'} onClick={() => setAlign('right')}>
        <AlignRight className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      {/* R97 — disabled-state tooltip flip extended to the colour pickers,
          font select and number-format select. R92 already wired the same
          flip onto the sibling Bold/Italic/Underline/Align buttons (lines
          1812-1837); this batch closes the rest of the toolbar so every
          control sharing `disabled = !selectionAddr` explains itself on
          hover instead of leaving the user staring at a greyed action
          label. Wording 「請先選取一個儲存格」 is reused verbatim from R92
          to keep the toolbar speaking one voice. */}
      <ColorPickerBtn
        disabled={disabled}
        title="文字顏色"
        disabledTitle="請先選取一個儲存格"
        icon={<Type className="h-3.5 w-3.5" />}
        value={style.fontColor}
        onChange={(hex) => setColor('fontColor', hex)}
        onClear={() => clearColor('fontColor')}
      />
      <ColorPickerBtn
        disabled={disabled}
        title="背景顏色"
        disabledTitle="請先選取一個儲存格"
        icon={<PaintBucket className="h-3.5 w-3.5" />}
        value={style.bgColor}
        onChange={(hex) => setColor('bgColor', hex)}
        onClear={() => clearColor('bgColor')}
      />
      <Divider />
      <select
        disabled={disabled}
        value={style.fontFamily ?? ''}
        onChange={(e) => setFontFamily(e.target.value)}
        className={cn(
          'h-7 text-xs rounded border border-border bg-background px-1.5 max-w-[170px]',
          disabled && 'cursor-not-allowed',
        )}
        title={disabled ? '請先選取一個儲存格' : '字型'}
        style={style.fontFamily ? { fontFamily: style.fontFamily } : undefined}
      >
        <option value="">預設字型</option>
        {FONT_FAMILIES.map((f) => (
          <option key={f.value} value={f.value} style={{ fontFamily: f.value }}>
            {f.label}
          </option>
        ))}
      </select>
      <select
        disabled={disabled}
        value={style.numberFormat ?? ''}
        onChange={(e) => setNumberFormat(e.target.value)}
        className={cn(
          'h-7 text-xs rounded border border-border bg-background px-1.5 max-w-[160px]',
          disabled && 'cursor-not-allowed',
        )}
        title={disabled ? '請先選取一個儲存格' : '數值格式'}
      >
        {NUMBER_FORMATS.map((f) => (
          <option key={f.value} value={f.value}>
            {f.label}
          </option>
        ))}
      </select>
      <Divider />
      {/* Insert row/col tooltips mirror the delete-row/col 2-way pattern at
          line 1968-1983/2004-2019 — same toolbar, same disable cause
          (`disabled` = `!selectionAddr`), so showing a static "在上方插入列"
          while the button is greyed leaves the user staring at a dead button
          with no explanation. The delete buttons already explain "請先選取一
          個儲存格" in the same disabled state; the insert buttons sat right
          next to them advertising only the action label, breaking the
          parity that the merge/unmerge comment at line 2021-2034 explicitly
          calls out as the toolbar's standard. (Insert is 2-way not 3-way
          because there is no row/col-count floor for insertion — adding a
          row never violates an "at-least-1" invariant.) */}
      <ToolbarBtn
        disabled={disabled}
        title={disabled ? '請先選取一個儲存格' : '在上方插入列'}
        onClick={onInsertRowAbove}
      >
        <span className="relative inline-flex">
          <Rows className="h-3.5 w-3.5" />
          <span className="absolute -top-1 -right-1 text-[8px] font-bold leading-none">↑</span>
        </span>
      </ToolbarBtn>
      <ToolbarBtn
        disabled={disabled}
        title={disabled ? '請先選取一個儲存格' : '在下方插入列'}
        onClick={onInsertRowBelow}
      >
        <span className="relative inline-flex">
          <Rows className="h-3.5 w-3.5" />
          <span className="absolute -bottom-1 -right-1 text-[8px] font-bold leading-none">↓</span>
        </span>
      </ToolbarBtn>
      {/* Row/column delete: mirror the sibling sheet-delete button at line
          ~2696 (`sheets.length <= 1 ? '至少要保留一個工作表' : '刪除工作表'`)
          — both adapters refuse to drop below 1 (xlsx-adapter.ts:687 for
          rows, :718 for cols, both early-return the unchanged sheet on
          boundary), and the sheet-delete button already syncs that
          constraint to the UI. The row/col delete buttons did not, so a
          user with a single-row or single-col sheet would click "刪除此列"
          and watch nothing happen — destructive intent + zero feedback is
          the worst possible failure mode. The 3-way state-aware tooltip
          mirrors the merge/unmerge pair just below for parity:
            disabled = !selectionAddr   → no cell selected
            !canDeleteRow/Col           → at the 1-row / 1-col floor
          The destructive-red Trash2 sub-icon stays — it's a visual cue,
          not the affordance the user reads to predict behavior. */}
      <ToolbarBtn
        disabled={disabled || !canDeleteRow}
        title={
          disabled
            ? '請先選取一個儲存格'
            : !canDeleteRow
              ? '至少要保留一列'
              : '刪除此列'
        }
        onClick={onDeleteRow}
      >
        <span className="relative inline-flex">
          <Rows className="h-3.5 w-3.5" />
          <Trash2 className="h-2.5 w-2.5 absolute -bottom-0.5 -right-0.5 text-destructive" />
        </span>
      </ToolbarBtn>
      <ToolbarBtn
        disabled={disabled}
        title={disabled ? '請先選取一個儲存格' : '在左側插入欄'}
        onClick={onInsertColLeft}
      >
        <span className="relative inline-flex">
          <Columns className="h-3.5 w-3.5" />
          <span className="absolute -top-1 -left-1 text-[8px] font-bold leading-none">←</span>
        </span>
      </ToolbarBtn>
      <ToolbarBtn
        disabled={disabled}
        title={disabled ? '請先選取一個儲存格' : '在右側插入欄'}
        onClick={onInsertColRight}
      >
        <span className="relative inline-flex">
          <Columns className="h-3.5 w-3.5" />
          <span className="absolute -top-1 -right-1 text-[8px] font-bold leading-none">→</span>
        </span>
      </ToolbarBtn>
      <ToolbarBtn
        disabled={disabled || !canDeleteCol}
        title={
          disabled
            ? '請先選取一個儲存格'
            : !canDeleteCol
              ? '至少要保留一欄'
              : '刪除此欄'
        }
        onClick={onDeleteCol}
      >
        <span className="relative inline-flex">
          <Columns className="h-3.5 w-3.5" />
          <Trash2 className="h-2.5 w-2.5 absolute -bottom-0.5 -right-0.5 text-destructive" />
        </span>
      </ToolbarBtn>
      <Divider />
      {/* State-aware tooltips for the merge / unmerge pair: when the button
          is disabled, the tooltip explains *why* and *what the user has to
          do to enable it*, instead of just repeating the button label. The
          previous static strings were the only dead-end on this toolbar —
          every other disabled button in the same file (sheet delete at line
          2696) and across the app (App.tsx Undo/Redo/Export at 1229/1238/
          1219) already follows this pattern. Merge/unmerge is the failure
          mode users hit most: Excel's mental model is "select cells then
          click", so a greyed button with no explanation reads as a bug.
          Three states per button (enabled / no-selection / wrong-shape)
          match the actual disabled-cause check at line 2037-2043:
            disabled = !selectionAddr   → no cell selected at all
            !canMerge                   → only one cell selected (line 986)
            !canUnmerge                 → selection has no merge (line 987) */}
      <ToolbarBtn
        disabled={disabled || !canMerge}
        title={
          disabled
            ? '請先選取儲存格範圍才能合併'
            : !canMerge
              ? '請先選取兩個以上的儲存格才能合併'
              : '合併儲存格（保留左上角內容）'
        }
        onClick={onMerge}
      >
        <Merge className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        disabled={disabled || !canUnmerge}
        title={
          disabled
            ? '請先選取已合併的儲存格'
            : !canUnmerge
              ? '目前選取的範圍沒有合併儲存格'
              : '取消合併'
        }
        onClick={onUnmerge}
      >
        <Split className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      {/* Image insertion is allowed even with no selection — falls back to A1.
          We keep the button always-enabled so users can stamp an image onto
          a fresh sheet without first having to click into the grid.
          R99 — tooltip is now state-aware: previously the static string
          「錨定於選取的儲存格」 contradicted the very behaviour the comment
          two lines up defends — a user with nothing selected reads it and
          assumes they must click into the grid first, the exact friction
          the always-enabled button was added to prevent. Same-file
          smoking-gun mismatch (author's defensive comment vs.
          user-facing tooltip), in the spirit of R96's tooltip-honesty
          fix at MarkdownToolbar.tsx:184-209 where the picker accepted
          local files but the entry-point tooltip claimed URL-only.
          When `selectionAddr` is null we now disclose the A1 fallback
          inline so the tooltip stops lying about the click outcome. */}
      <ToolbarBtn
        title={selectionAddr
          ? `插入圖片（錨定於 ${selectionAddr}）`
          : '插入圖片（未選取儲存格時錨定於 A1）'}
        onClick={onInsertImage}
      >
        <ImageIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      {/* Find & Replace — Ctrl+F is bound (line ~711) but the keyboard
          shortcut is invisible to mouse-first users. Stays enabled even
          without a selection: searching across all sheets is meaningful
          regardless of which (if any) cell currently has focus. */}
      <ToolbarBtn disabled={false} title="尋找與取代 (Ctrl+F)" onClick={onOpenFind}>
        <Search className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Selection address pushed to the right edge so toolbar action
          buttons stay anchored regardless of address string length;
          whitespace-nowrap guarantees we don't wrap "Sheet1!A1:C3". */}
      <span className="ml-auto text-[11px] text-muted-foreground font-mono whitespace-nowrap">
        {selectionAddr ?? '選取一個儲存格以開始編輯'}
      </span>
    </div>
  );
}

/**
 * Color-picker button: native HTML5 `<input type="color">` hidden underneath
 * a styled button so we get the OS picker without a popover library. A small
 * "×" appears beside the button when a color is set, to clear it.
 */
function ColorPickerBtn({
  title,
  disabledTitle,
  icon,
  value,
  disabled,
  onChange,
  onClear,
}: {
  title: string;
  /** R97 — when present and `disabled` is true, both the picker label and
   *  the clear-colour × button show this string instead of the default
   *  「文字顏色」 / 「清除文字顏色」 vocabulary. Mirrors the inline
   *  `disabled ? ... : ...` flip the R92 batch wired across the rest of
   *  this toolbar (bold/italic/underline/align — see XlsxEditor.tsx:1812-
   *  1837). Without this prop, hovering a greyed-out colour button gave
   *  the user the action label on a dead control — the exact gap R92
   *  closed for the sibling buttons two rows up. Optional so callsites
   *  outside the cell-style toolbar (none today) stay opt-in. */
  disabledTitle?: string;
  icon: React.ReactNode;
  value: string | undefined;
  disabled?: boolean;
  onChange: (hex: string) => void;
  onClear: () => void;
}): JSX.Element {
  const effectiveTitle = disabled && disabledTitle ? disabledTitle : title;
  return (
    <span className="relative inline-flex items-center">
      <label
        title={effectiveTitle}
        // Mousedown preventDefault keeps the cell focused — same reason
        // as ToolbarBtn. The native color-picker dialog still opens
        // because click fires regardless of focus.
        onMouseDown={(e) => e.preventDefault()}
        className={cn(
          'h-7 w-7 inline-flex items-center justify-center rounded transition-colors cursor-pointer',
          'text-muted-foreground hover:text-foreground hover:bg-secondary',
          disabled && 'cursor-not-allowed opacity-50 pointer-events-none',
        )}
      >
        <span className="relative">
          {icon}
          <span
            className="absolute -bottom-0.5 left-0 right-0 h-0.5 rounded"
            style={{ background: value ? `#${value}` : 'currentColor' }}
          />
        </span>
        <input
          type="color"
          value={value ? `#${value}` : '#000000'}
          onChange={(e) => onChange(e.target.value)}
          className="absolute inset-0 opacity-0 cursor-pointer"
          disabled={disabled}
        />
      </label>
      {value ? (
        <button
          type="button"
          title={disabled && disabledTitle ? disabledTitle : `清除 ${title}`}
          onClick={onClear}
          onMouseDown={(e) => e.preventDefault()}
          disabled={disabled}
          className="ml-px text-muted-foreground hover:text-destructive text-[10px] leading-none px-0.5"
        >
          ×
        </button>
      ) : null}
    </span>
  );
}

function ToolbarBtn({
  children,
  title,
  active,
  disabled,
  onClick,
}: {
  children: React.ReactNode;
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onClick={onClick}
      // R153 — toggle-state SR exposure. Same treatment as the sibling
      // ToolbarBtn definitions in MarkdownToolbar / DocxEditor / PptxEditor
      // (this round). Format buttons in this editor's toolbar (粗體 / 斜體 /
      // 底線 line 1841-1848、對齊三聯 line 1851-1858) all pass `active=
      // {!!style.X}` / `active={style.align === 'X'}`. The visual 「bg-
      // primary/20 text-primary」 active-highlight was a sighted-only signal
      // of「the currently focused cell is bold / italic / left-aligned」.
      // `aria-pressed={undefined}` renders no attribute, so action-only
      // ToolbarBtn callsites in this file keep clean action-button semantics
      // — only the format toggles pick up the proper toggle role.
      aria-pressed={active}
      // preventDefault on mousedown stops the button from stealing focus
      // from the active cell. Without this, clicking Bold/Italic/etc.
      // blurs the cell input → in edit mode that means a forced commit
      // on every format toggle, and the user has to re-click the cell to
      // continue typing. Excel keeps focus on the cell; we mirror that.
      onMouseDown={(e) => e.preventDefault()}
      className={cn(
        'h-7 w-7 inline-flex items-center justify-center rounded transition-colors',
        active
          ? 'bg-primary/20 text-primary'
          : 'text-muted-foreground hover:text-foreground hover:bg-secondary',
        disabled && 'cursor-not-allowed',
      )}
    >
      {children}
    </button>
  );
}

function Divider() {
  return <div className="w-px h-4 bg-border mx-1" />;
}

/** Strip falsy fields so we don't carry empty objects around. */
function normalizeStyle(s: XlsxCellStyle): XlsxCellStyle | undefined {
  const out: XlsxCellStyle = {};
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  if (s.align) out.align = s.align;
  if (s.fontColor) out.fontColor = s.fontColor;
  if (s.bgColor) out.bgColor = s.bgColor;
  if (s.fontSize) out.fontSize = s.fontSize;
  if (s.fontFamily) out.fontFamily = s.fontFamily;
  if (s.numberFormat) out.numberFormat = s.numberFormat;
  return Object.keys(out).length === 0 ? undefined : out;
}

/**
 * Formula bar (fx bar) — shows the address + raw text of the selected cell and
 * lets the user edit from a single, always-visible input. Mirrors Excel's
 * "Name Box | fx | content" row above the grid.
 */
function FormulaBar({
  addr,
  cell,
  rowCount,
  colCount,
  onCommit,
  onJumpToAddr,
}: {
  addr: string | null;
  cell: XlsxCell | null;
  rowCount: number;
  colCount: number;
  onCommit: (text: string) => void;
  /** Excel-style "Name Box" jump — typed cell address (e.g. "B5") is
   *  parsed and the selection moved there. Out-of-bounds addresses are
   *  rejected via toast so the caller doesn't have to coordinate. */
  onJumpToAddr: (r: number, c: number) => void;
}): JSX.Element {
  // Prefer the formula source over the computed text — the formula bar is
  // where users *edit* the cell, so showing "=A1+B1" beats showing "42".
  const editable = cell?.formula ?? cell?.text ?? '';
  const [value, setValue] = useState('');
  useEffect(() => {
    setValue(editable);
  }, [editable, addr]);
  // Name-box draft. Editable input that reflects the current address while
  // unfocused, and accepts a typed address ("B5", "AA12") on Enter to jump.
  // Cancel-via-Esc resets to the current address — same dance as the fx
  // input below. Without this, users had to use mouse + scroll to navigate
  // big sheets; Excel's name box is a top-3 keyboard shortcut for power
  // users (F5 → name box on real Excel). Storing as draft (vs. controlled)
  // keeps the input usable while the user is typing — we only commit on
  // Enter / blur.
  const [addrDraft, setAddrDraft] = useState(addr ?? '');
  useEffect(() => {
    setAddrDraft(addr ?? '');
  }, [addr]);
  const addrCancelRef = useRef(false);
  const commitAddr = () => {
    const raw = addrDraft.trim().toUpperCase();
    if (!raw || raw === (addr ?? '')) {
      // No-op — restore display in case casing differed (user typed "b5").
      setAddrDraft(addr ?? '');
      return;
    }
    const parsed = parseA1(raw);
    if (!parsed) {
      notify(`無法解析儲存格位址：${raw}`, 'error');
      setAddrDraft(addr ?? '');
      return;
    }
    if (parsed.r < 0 || parsed.r >= rowCount || parsed.c < 0 || parsed.c >= colCount) {
      notify(`位址超出工作表範圍：${raw}`, 'error');
      setAddrDraft(addr ?? '');
      return;
    }
    onJumpToAddr(parsed.r, parsed.c);
  };
  // Surface formula errors directly in the bar — without this the user sees
  // "#DIV/0!" only as a computed cell value (also red, in the grid) and the
  // formula bar shows the source unchanged. When they click into the bar to
  // fix the formula they get no signal that the *current* state is broken.
  // Showing a small error pill + red border on the input gives the same
  // red-marker feedback the grid does, in the place users actually edit.
  // Detection: only flag when the cell actually has a formula AND its
  // computed text is a known error code. A user who typed the literal
  // string "#VALUE!" as a plain text label shouldn't see the bar light up.
  const errorCode: FormulaErrorCode | null =
    cell?.formula && cell.text && ERROR_CODES.has(cell.text as FormulaErrorCode)
      ? (cell.text as FormulaErrorCode)
      : null;
  // Tooltip text — short, user-facing. Kept terse so it fits a native title=
  // without wrapping awkwardly.
  const errorHints: Record<FormulaErrorCode, string> = {
    '#DIV/0!': '除以零或空儲存格',
    '#VALUE!': '參數型別錯誤（例如數字運算遇到文字）',
    '#REF!': '參照已不存在的儲存格',
    '#NAME?': '無法識別的函數或名稱',
    '#N/A': '查無資料',
    '#NUM!': '數值無效或超出範圍',
    '#CYCLE!': '公式互相循環參照',
    '#ERROR!': '公式語法錯誤',
  };
  // Escape-to-cancel needs to defeat the input's onBlur — when Escape calls
  // .blur(), onBlur fires synchronously with the closure's still-typed
  // `value` (setValue's update hasn't landed yet) and onCommit then writes
  // exactly the text the user wanted to revert. Same hazard TabBar's
  // `cancelRenameRef` solves for tab rename: a ref flag short-circuits the
  // commit branch on the way out.
  const cancelRef = useRef(false);

  return (
    <div className="flex items-stretch border-b border-border bg-background text-xs">
      {/* Name box — type a cell address ("B5") and press Enter to jump.
          Disabled when there's no selection (shouldn't happen in practice,
          but keeps the affordance honest). Esc reverts the draft so a
          mid-edit user can bail without committing a typo. */}
      <input
        type="text"
        value={addrDraft}
        disabled={!addr}
        // R156 — explicit accessible name. Placeholder is just「—」 (line
        // 2378), useless as an SR fallback; the title at line 2377 carries
        // the full usage instruction, but `title` isn't reliable as the
        // primary accessible name on inputs (SR engines either skip it or
        // append it as description). The name「Name Box / 儲存格位址」 is
        // Excel-canonical — using the formal Chinese term keeps SR users
        // who know Excel's terminology oriented immediately.
        aria-label="儲存格位址"
        // Marked so XlsxEditor's Ctrl+G keymap can focus this input via
        // querySelector — Excel's "Go To" (Ctrl+G / F5) historically opens
        // a dedicated dialog, but since the Name Box already accepts the
        // same address syntax, focusing it gives keyboard parity with
        // PptxEditor / DocxEditor's Ctrl+G without a redundant second UI.
        data-xlsx-namebox
        onChange={(e) => setAddrDraft(e.target.value)}
        onFocus={(e) => e.currentTarget.select()}
        onBlur={() => {
          if (addrCancelRef.current) {
            addrCancelRef.current = false;
            setAddrDraft(addr ?? '');
            return;
          }
          commitAddr();
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            (e.target as HTMLInputElement).blur();
          } else if (e.key === 'Escape') {
            addrCancelRef.current = true;
            (e.target as HTMLInputElement).blur();
          }
        }}
        title="輸入儲存格位址後按 Enter 跳至（例：B5）。Ctrl+G 可從鍵盤聚焦至此"
        placeholder="—"
        className="w-24 shrink-0 px-2 py-1 border-r border-border bg-secondary/50 font-mono text-muted-foreground outline-none focus:bg-background focus:text-foreground focus:ring-1 focus:ring-primary"
      />
      <div className="w-8 shrink-0 px-2 py-1 border-r border-border bg-secondary/50 italic font-serif flex items-center text-muted-foreground">
        fx
      </div>
      <input
        type="text"
        value={value}
        disabled={!addr}
        // R156 — explicit accessible name. Same rationale as the Name Box
        // sibling above. The visible「fx」 glyph at line 2381 is a sighted-
        // only signal that this is the formula bar; SR users see neither
        // the Name Box's role nor the fx separator and would tab into a
        // bare edit field. Using the canonical Excel term「資料編輯列」 /
        // formula bar keeps consistency with the addr sibling's naming.
        aria-label="資料編輯列"
        onChange={(e) => setValue(e.target.value)}
        onBlur={() => {
          if (cancelRef.current) {
            cancelRef.current = false;
            setValue(editable);
            return;
          }
          onCommit(value);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            (e.target as HTMLInputElement).blur();
          } else if (e.key === 'Escape') {
            cancelRef.current = true;
            (e.target as HTMLInputElement).blur();
          }
        }}
        // Mirror the name box's tooltip (line 2201) so the fx bar advertises
        // the same Enter-commit / Esc-revert contract its sibling does. Both
        // inputs share the cancelRef pattern explained at lines 2159-2165:
        // Esc reverts the draft, Enter writes it. Without this hint the
        // revert-on-Esc behaviour is discoverable only by experiment, while
        // the name box right next to it loudly announces the same shortcut —
        // a cross-input inconsistency that makes the bar feel half-finished.
        title="輸入內容或 = 公式後按 Enter 寫入；按 Esc 取消還原"
        placeholder={addr ? '輸入內容或 =公式…' : '選一個儲存格'}
        className={cn(
          'flex-1 px-2 py-1 bg-transparent outline-none font-mono',
          errorCode
            ? 'text-destructive focus:bg-destructive/5'
            : 'focus:bg-primary/5',
        )}
      />
      {errorCode && (
        <div
          // Pinned to the right edge of the bar — gives the user the same
          // red marker the grid cell already shows, plus the error code so
          // they can map "what's broken" to "what to fix" without crossing
          // the screen.
          title={`${errorCode}：${errorHints[errorCode]}`}
          className="shrink-0 self-center mr-2 px-1.5 py-0.5 rounded text-[10px] font-mono bg-destructive/15 text-destructive border border-destructive/40"
        >
          {errorCode}
        </div>
      )}
    </div>
  );
}

/**
 * Workbook tab strip. Scrolls horizontally when the workbook has many sheets
 * — without `overflow-x-auto` + `shrink-0`, flex layout was squeezing tabs
 * together and clipping later ones, which looked like "only one sheet
 * opened" on imports of multi-sheet files.
 *
 * Each tab supports double-click to rename and a small trailing × button to
 * delete (after confirm). The "+" button at the strip's end appends a fresh
 * blank sheet.
 */
function SheetTabs({
  sheets,
  activeIdx,
  onSelect,
  onAdd,
  onRenameCommit,
  onDelete,
  onDuplicate,
  onReorder,
}: {
  sheets: XlsxSheet[];
  activeIdx: number;
  onSelect: (i: number) => void;
  onAdd: () => void;
  onRenameCommit: (i: number, newName: string) => void;
  onDelete: (i: number) => void;
  onDuplicate: (i: number) => void;
  /** Drag-and-drop reorder. Excel's tab strip lets users grab a tab and drop
   *  it at a new position; we mirror the same gesture here using the TabBar
   *  / SlideRail HTML5 drag-drop pattern. Right-click menu is unaffected
   *  (still anchors to the clicked tab via stopPropagation on the rename
   *  input). */
  onReorder: (from: number, to: number) => void;
}): JSX.Element {
  // Drag state — local because the parent only needs to know the final
  // (from, to) pair, not every dragover tick. Keeps the spreadsheet from
  // re-rendering on hover.
  const [dragIdx, setDragIdx] = useState<number | null>(null);
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null);
  // Inline rename — replaces the previous `window.prompt('重新命名工作表', cur)`,
  // which forced a modal browser dialog that broke focus, blocked the rest of
  // the UI, and didn't match the in-place rename the rest of the app uses
  // (TabBar, file explorer). Now: enter rename mode, edit on top of the tab,
  // Enter / blur commits, Esc / dup-name reverts.
  const [renamingIdx, setRenamingIdx] = useState<number | null>(null);
  // Esc commits via blur, but we have to short-circuit the commit when the
  // user explicitly cancelled. State setting in the keydown would race the
  // blur fired in the same tick, so use a ref. Mirrors TabBar.tsx.
  const cancelRenameRef = useRef(false);
  const startRename = (i: number) => {
    cancelRenameRef.current = false;
    setRenamingIdx(i);
  };
  // Floating right-click menu state. The previous implementation popped a
  // browser `window.prompt` asking the user to type 'r' / 'd' / 'c' — which
  // is jarring (no other right-click menu in the app does that, Excel /
  // Sheets / Numbers all show real menus) and forces users to remember
  // letter codes. We mirror the TabBar pattern: anchor a small floating menu
  // at the click point, dismiss on outside-click or Esc.
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; idx: number } | null>(null);
  useEffect(() => {
    if (!ctxMenu) return;
    const off = () => setCtxMenu(null);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setCtxMenu(null);
    };
    window.addEventListener('mousedown', off);
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('mousedown', off);
      window.removeEventListener('keydown', onKey);
    };
  }, [ctxMenu]);

  // Auto-scroll the active sheet tab into view, mirroring the SlideRail fix
  // in PptxEditor and PptxNavPanel's existing implementation. The strip is
  // horizontally-scrollable (`overflow-x-auto`) and a workbook with 12+
  // sheets will park later tabs out of frame; without this, Round 13's
  // tab-switch sheet-index restore brings the spreadsheet back to sheet 8
  // but the strip still shows sheets 1-5, leaving the active highlight
  // invisible. Same triggers also bite duplicate-sheet (which appends to
  // the end) and Ctrl+PageDown navigation.
  // `inline: 'nearest'` handles the horizontal axis ('block: nearest' on
  // the vertical avoids any incidental page scroll if the strip itself is
  // partially clipped). Re-runs on sheets.length so an add/duplicate that
  // shifted later indices still settles to the right offset.
  const stripScrollRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = stripScrollRef.current?.querySelector<HTMLElement>(
      `[data-xlsx-sheet-idx="${activeIdx}"]`,
    );
    el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [activeIdx, sheets.length]);

  return (
    <div
      ref={stripScrollRef}
      className="flex items-stretch border-b border-border bg-secondary/40 overflow-x-auto"
      style={{ scrollbarWidth: 'thin' }}
    >
      <div className="flex items-stretch gap-px px-1 py-0.5 min-w-max">
        {sheets.map((s, i) => {
          const active = i === activeIdx;
          return (
            <div
              key={`${s.name}-${i}`}
              data-xlsx-sheet-idx={i}
              // Drag-reorder. Disabled while renaming so the inline rename
              // input stays interactive (otherwise the mousedown that goes
              // to focus the input also starts a drag and the field never
              // gets text-selection focus). Also disabled when a context
              // menu is open over this tab.
              draggable={renamingIdx !== i}
              onDragStart={(e) => {
                if (renamingIdx === i) {
                  e.preventDefault();
                  return;
                }
                setDragIdx(i);
                e.dataTransfer.effectAllowed = 'move';
                // Some browsers refuse to start a drag without setData; the
                // payload itself is unused since we read dragIdx from state.
                try { e.dataTransfer.setData('text/plain', String(i)); } catch { /* noop */ }
              }}
              onDragOver={(e) => {
                if (dragIdx === null || dragIdx === i) return;
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                if (dragOverIdx !== i) setDragOverIdx(i);
              }}
              onDragLeave={() => {
                if (dragOverIdx === i) setDragOverIdx(null);
              }}
              onDragEnd={() => {
                setDragIdx(null);
                setDragOverIdx(null);
              }}
              onDrop={(e) => {
                e.preventDefault();
                const from = dragIdx;
                setDragIdx(null);
                setDragOverIdx(null);
                if (from === null || from === i) return;
                onReorder(from, i);
              }}
              className={cn(
                'group flex items-center shrink-0 whitespace-nowrap rounded-t-md border-b-2 transition-colors',
                active
                  ? 'border-primary bg-background text-foreground shadow-sm'
                  : 'border-transparent bg-secondary/60 text-muted-foreground hover:text-foreground hover:bg-secondary',
                dragIdx === i && 'opacity-40',
                dragOverIdx === i && dragIdx !== i && 'ring-2 ring-primary',
              )}
            >
              {renamingIdx === i ? (
                <input
                  autoFocus
                  defaultValue={s.name}
                  // Select on focus so user can immediately overtype.
                  onFocus={(e) => e.currentTarget.select()}
                  onBlur={(e) => {
                    if (cancelRenameRef.current) {
                      cancelRenameRef.current = false;
                      setRenamingIdx(null);
                      return;
                    }
                    const trimmed = e.currentTarget.value.trim();
                    setRenamingIdx(null);
                    // Same-name no-op stays quiet — user just clicked away
                    // without changing anything. But empty submit was
                    // silently dropping the rename; symmetry with the
                    // duplicate-name notify below means an empty entry
                    // should also surface a reason. Otherwise users see
                    // their tab snap back to the old name and assume the
                    // app ate their keystrokes.
                    if (trimmed === s.name) return;
                    if (!trimmed) {
                      notify('工作表名稱不能為空', 'warning');
                      return;
                    }
                    if (sheets.some((other, j) => j !== i && other.name === trimmed)) {
                      // Duplicate name. Previously this returned silently —
                      // the tab snapped back to the old name and users were
                      // left wondering whether their keystrokes were lost
                      // or the rename actually committed. A transient
                      // warning toast names the conflicting sheet
                      // explicitly so they can re-aim immediately. The
                      // toast store coalesces identical messages, so rapid
                      // retries don't stack up.
                      notify(`已有同名工作表「${trimmed}」，請改用其他名稱`, 'warning');
                      return;
                    }
                    onRenameCommit(i, trimmed);
                  }}
                  onKeyDown={(e) => {
                    // R232 — same IME composition guard as R231 (TabBar
                    // rename) and the cell-edit branch above. Sheet
                    // rename is the secondary CJK rename surface in
                    // XlsxEditor (cell input being primary); both share
                    // the same Enter-commits-on-blur shape.
                    if (e.nativeEvent.isComposing) return;
                    if (e.key === 'Enter' || e.key === 'Tab') {
                      e.preventDefault();
                      (e.target as HTMLInputElement).blur();
                    } else if (e.key === 'Escape') {
                      cancelRenameRef.current = true;
                      (e.target as HTMLInputElement).blur();
                    }
                  }}
                  className="bg-transparent outline-none border-b border-primary px-3 py-1 text-xs font-medium w-32"
                />
              ) : (
                <button
                  type="button"
                  onClick={() => onSelect(i)}
                  onDoubleClick={() => {
                    onSelect(i);
                    startRename(i);
                  }}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    // Make the right-clicked sheet active first — the menu's
                    // delete / duplicate actions otherwise feel disconnected
                    // from where the user pointed.
                    onSelect(i);
                    setCtxMenu({ x: e.clientX, y: e.clientY, idx: i });
                  }}
                  // Right-click menu (lines 2638-2673) is the ONLY entry point
                  // for 複製工作表 — there is no toolbar button, the `+` at
                  // line 2611 adds a BLANK sheet (handleAddSheet → addSheet),
                  // and no keyboard shortcut covers duplicate. A user who
                  // doesn't reflexively right-click a sheet tab simply cannot
                  // find the feature. Delete has a visible × button (2586)
                  // and rename has a documented double-click (this tooltip),
                  // but duplicate was tribal knowledge — same R66 (TabBar
                  // 中鍵關閉) / R67 (drag-drop file open) / R68 (Ctrl+F)
                  // pattern: feature 100% implemented, 0% advertised at the
                  // discovery surface. Middle-dot 「·」 separator and bracketed
                  // gesture-list match TabBar.tsx:350's R66 convention so
                  // tab-like surfaces across the app speak the same tooltip
                  // dialect (TabBar = 雙擊重新命名 · 中鍵關閉, sheet tab =
                  // 雙擊重新命名 · 右鍵選單). 「右鍵選單」 stays abstract so
                  // future menu items don't churn the tooltip — the sibling
                  // PptxEditor SlideRail at PptxEditor.tsx:2752 has the same
                  // right-click → menu wiring but its menu duplicates already-
                  // visible rail-icon affordances (上移/下移/複製/刪除 all
                  // exposed when active), so the sheet tab is the surface
                  // where the gap actually loses functionality.
                  // 拖曳排序 added in R75 alongside the matching TabBar.tsx:
                  // 372 fix. Drag is the only reorder path — handleMoveSheet
                  // doc-comment at line 705-707 explicitly says "chevron-only
                  // reorder doesn't exist here, so the gap was: rename →
                  // delete-and-recreate at the new position." A user who
                  // doesn't realise the tabs are draggable has *no* way to
                  // rearrange a workbook short of recreating sheets, which is
                  // exactly the "tribal-knowledge gesture" failure mode the
                  // R69 right-click fix above already targeted. Insertion
                  // slot mirrors TabBar (drag right before 右鍵選單) so the
                  // three tab-like surfaces (TabBar / sheet tab / SlideRail)
                  // speak the same tooltip dialect.
                  title={`${s.name}（雙擊重新命名 · 拖曳排序 · 右鍵選單）`}
                  className="px-3 py-1 text-xs font-medium"
                >
                  {s.name}
                </button>
              )}
              <button
                type="button"
                onClick={() => onDelete(i)}
                // Disable when this is the last remaining sheet — XLSX requires
                // at least one sheet, and the parent's `handleDeleteSheet`
                // refuses anyway. Disabling the affordance is more discoverable
                // than firing a native window.alert after the click; matches
                // the same guard already in place on the right-click menu's
                // 刪除 item below.
                disabled={sheets.length <= 1}
                title={sheets.length <= 1 ? '至少要保留一個工作表' : '刪除工作表'}
                className={cn(
                  'mr-1 grid h-4 w-4 place-items-center rounded text-muted-foreground hover:bg-destructive/10 hover:text-destructive disabled:pointer-events-none',
                  sheets.length <= 1
                    ? 'opacity-0 group-hover:opacity-20'
                    : active
                      ? 'opacity-60 hover:opacity-100'
                      : 'opacity-0 group-hover:opacity-60 hover:opacity-100',
                )}
              >
                <X size={11} />
              </button>
            </div>
          );
        })}
        {/* R148 — flat-index drift correction inside this same file
            (R136/R137/R140/R145/R146/R147 paradigm). Five nested
            doc-comments collectively cited 14 wrong line numbers about
            this file's own structure plus cross-file App.tsx and
            PptxEditor.tsx targets. Smoking gun: the sheet-delete X
            button at L2696 was cited at THREE different wrong line
            numbers — `~2498` in two comments (L1955/L2025), `2629`
            in four comments (L2798/L2835/L2846/L2852) — six cites of
            one target with two stale values, classic R145/R146 self-
            contradiction. delete-row/col toolbar (cited 1910-1924/
            1938-1953) actually lives at L1968-1983/2004-2019; merge/
            unmerge comment (cited 1955-1968) actually lives at L2021-
            2034; merge title evaluation (cited 1911-1918) actually
            lives at L2037-2043; rename-menu sibling labels (cited
            2792, 2826) actually live at L2831, L2865; disabled-state
            hint (cited 2819) actually lives at L2858. Cross-file:
            App.tsx Undo/Redo/Export at L2026-2027 (1206/1216/1225 →
            1229/1238/1219); PptxEditor R80 menu scope at L2840 (2904-
            2941 → 3302-3341); PptxEditor wording-precedent cite at
            L2845 (2823-2824 → 3193). xlsx-adapter.ts:687 / :718 cites
            at L1956-1957 verified still correct — adapter file
            unchanged. All 14 cites updated in-place; no JSX or behavior
            change. Anchor sits at L2711+, ABOVE three cited targets
            below it (L2845/L2872/L2879 — visible labels and disabled-
            state hint of the right-click menu). When this anchor
            block expanded during R148's own write, those targets
            shifted thrice: L2831/L2858/L2865 → L2837/L2864/L2871 →
            L2842/L2869/L2876 → L2845/L2872/L2879 (anchor +6, then
            narrative +5, then this meta-narrative +2). Cites at
            L2799-2805 self-corrected thrice within the same round
            (R142 paradigm — same-round-self-drift loop closed in-
            place via strict 4-digit swaps so no further line count
            changes propagate). All other cited
            targets (L1968-1983, L2004-2019, L2021-2034, L2037-2043,
            L2696) sit ABOVE the anchor and are unshifted. */}
        {/* R110 — disclose position + content + naming, mirroring the
            cross-editor sibling. The exact-analog button — a `+` at the
            bottom edge of a list/rail of items — at PptxEditor.tsx:3221
            carries `title="在目前投影片之後新增一張（複製目前投影片的
            內容）"`, naming WHERE the new item lands and WHAT'S IN IT.
            Prior shape here was just "新增工作表": tooltip = action verb,
            zero info on hover beyond the visible `<Plus>` icon, leaving
            three hidden behavioral facts:

            1. Position. handleAddSheet at line 649-656 appends to the
               end (the comment there literally says "Append at end" and
               sets `setActiveSheetIdx(model.sheets.length)`). PptxEditor's
               sibling explicitly says「在目前投影片之後」— same disclosure
               principle applies, and end-of-list vs after-active is a
               real difference users should not have to discover by trial.

            2. Content. addSheet at xlsx-adapter.ts:801-818 creates an
               empty grid (`emptyGrid(MIN_ROWS, MIN_COLS)`). The very
               comment at PptxEditor.tsx:3074-3076 already calls out that
               this editor's `+` "really does add a blank sheet" —
               codebase knows the contrast vs PptxEditor's duplicate-
               from-active, but the user never sees it on hover.

            3. Naming. addSheet auto-names via `uniqueSheetName` with
               collision suffix `(n)` (xlsx-adapter.ts:802-807). User
               might not realize the action commits without prompting
               for a name; double-click-to-rename is a separate
               affordance that this tooltip leaves undiscoverable.

            Wording mirrors PptxEditor.tsx:3221 structurally
            (「在 + 位置 + 新增一張 + 內容描述」) but with the parenthesis
            carrying the naming hint instead of content (since "空白" is
            already baked into the noun phrase). */}
        <button
          type="button"
          onClick={onAdd}
          title="在末端新增一張空白工作表（自動命名）"
          className="shrink-0 ml-1 grid h-7 w-7 place-items-center rounded text-muted-foreground hover:bg-secondary hover:text-foreground"
        >
          <Plus size={14} />
        </button>
      </div>
      <div className="ml-auto flex items-center gap-2 px-2 text-[10px] text-muted-foreground shrink-0">
        共 {sheets.length} 張
      </div>
      {ctxMenu && (() => {
        // Estimate bounds the menu's rendered size: 4 items + 1 separator
        // at ~28px each ≈ 140px tall; min-w 160px + padding ≈ 180px wide.
        // Clamp at render time avoids the flicker of measure-and-adjust.
        const pos = clampToViewport(ctxMenu.x, ctxMenu.y, 180, 180);
        return (
        <div
          role="menu"
          // Stop mousedown from reaching the window-level dismiss handler so
          // clicking an item actually fires its onClick — without this the
          // outside-click listener closes the menu before the click lands.
          onMouseDown={(e) => e.stopPropagation()}
          style={{ left: pos.left, top: pos.top }}
          className="fixed z-50 min-w-[160px] rounded-md border bg-popover text-popover-foreground shadow-md py-1 text-xs"
        >
          {/* Same-menu visible-label parity: the other two items in this
              3-row context menu read `複製工作表` / `刪除工作表` (line
              2845, 2879) — `動詞 + 工作表`. The rename row used to read
              just `重新命名`, the lone bare verb in a menu where the
              other two rows already disclosed scope, and the lone bare
              verb among the four sheet-targeted strings users encounter
              (the X button at line 2696 reads `刪除工作表`, the disabled-
              state hint at line 2872 reads `至少要保留一個工作表`). The
              earlier `window.prompt` form documented at line 2465 even
              used `重新命名工作表` verbatim — the bare label was a
              shortening regression introduced when the prompt was
              replaced by inline rename. Aligned to match the menu's own
              two siblings and the historical precedent. (Note: this
              menu uses `動詞 + 工作表` without `此`, unlike the PPTX
              slide menu's `動詞 + 此 + 投影片` from R133 — sheets have
              names and the menu anchors to the named sheet whose tab
              was right-clicked, so `此` is redundant here. The choice
              is preserved verbatim from the existing two siblings.) */}
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              const idx = ctxMenu.idx;
              setCtxMenu(null);
              startRename(idx);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-accent"
          >
            重新命名工作表
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onDuplicate(ctxMenu.idx);
              setCtxMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 hover:bg-accent"
          >
            複製工作表
          </button>
          <div className="my-1 h-px bg-border" />
          {/* Boundary-aware tooltip mirrors the same-file X button at line
              2696 above (`title={sheets.length <= 1 ? '至少要保留一個工作表'
              : '刪除工作表'}`) — both surfaces drive onDelete with the
              identical `sheets.length <= 1` floor, but only the X button
              had been explaining why when disabled. Same in-file
              inconsistency R80 just closed for PptxEditor's SlideRail
              right-click menu (PptxEditor.tsx:3302-3341): rail-icon
              sibling already had boundary-aware titles, the context menu
              hadn't gotten the same treatment. The R80 fix even cited
              this exact XlsxEditor X-button tooltip as the wording
              precedent for its `'至少要保留一張投影片'` boundary message
              (PptxEditor.tsx:3193). Boundary string is reused
              verbatim from line 2696 so the same disabled state reads
              the same hint across both surfaces. `: undefined` (rather
              than `: '刪除工作表'`) for the same reason as R80: the
              menu item's visible text already labels the action, an
              active-state tooltip would just stack a redundant browser
              tooltip on top of the button text. The X button at line
              2696 keeps `: '刪除工作表'` because it's icon-only and
              needs the active tooltip as a label. */}
          <button
            type="button"
            role="menuitem"
            disabled={sheets.length <= 1}
            title={sheets.length <= 1 ? '至少要保留一個工作表' : undefined}
            onClick={() => {
              onDelete(ctxMenu.idx);
              setCtxMenu(null);
            }}
            className="w-full text-left px-3 py-1.5 text-destructive hover:bg-destructive/10 disabled:opacity-40 disabled:pointer-events-none"
          >
            刪除工作表
          </button>
        </div>
        );
      })()}
    </div>
  );
}

function Grid({
  sheet,
  selection,
  copyMarker,
  onSelect,
  onMoveAnchor,
  onExtend,
  onCommit,
  onEditingChange,
  images,
  selectedImageId,
  onImageSelect,
  onImagePatch,
  onDropImageFile,
}: {
  sheet: XlsxSheet;
  selection: Selection | null;
  /** Most-recent copy / cut rectangle on this sheet, if any. Renders as
   *  a dashed border around the source cells. */
  copyMarker: { r1: number; c1: number; r2: number; c2: number; mode: 'copy' | 'cut' } | null;
  onSelect: (r: number, c: number) => void;
  /** Move the anchor inside the existing selection rect without collapsing
   *  it — used by Tab/Enter cycling within a multi-cell selection. */
  onMoveAnchor: (r: number, c: number) => void;
  /** Extend the current selection's far corner — fires while dragging. */
  onExtend: (r: number, c: number) => void;
  onCommit: (r: number, c: number, text: string) => void;
  /** Bubbles cell edit-mode transitions up to the editor. */
  onEditingChange: (editing: boolean) => void;
  /** Floating images on this sheet — rendered as overlays anchored to cells. */
  images: XlsxImage[] | undefined;
  /** Currently-selected overlay (for highlight + Delete-key removal). */
  selectedImageId: string | null;
  onImageSelect: (id: string | null) => void;
  /** Commit a drag (anchor change) or resize (size change) to the model. */
  onImagePatch: (id: string, patch: Partial<XlsxImage>) => void;
  /** Native HTML5 drop of an image file from the OS — anchor cell is the
   *  one under the cursor at drop time. */
  onDropImageFile: (file: File, r: number, c: number) => void;
}): JSX.Element {
  // While true, mouseenter on a cell extends the selection. Set on
  // mousedown and cleared on the next document-level mouseup so the
  // gesture survives the cursor briefly leaving the table.
  const draggingRef = useRef(false);
  const tableRef = useRef<HTMLTableElement>(null);
  // Set by focusCell({keepRect:true}) so the next focus event keeps the
  // surrounding selection rectangle (Tab/Enter cycling inside a range)
  // instead of collapsing to a 1×1 selection on the new anchor.
  const movingInRangeRef = useRef(false);

  const beginDrag = (r: number, c: number) => {
    draggingRef.current = true;
    onSelect(r, c);
    const stop = () => {
      draggingRef.current = false;
      window.removeEventListener('mouseup', stop);
    };
    window.addEventListener('mouseup', stop);
  };

  /**
   * Spreadsheet-style navigation: clamp into the sheet, find the input via
   * its data-cell-r / -c attributes, and focus it. Select-mode's onFocus
   * already select-alls so typing replaces. When `keepRect` is true the
   * surrounding selection rectangle is preserved (Tab/Enter inside a
   * multi-cell selection).
   */
  const focusCell = (r: number, c: number, opts?: { keepRect?: boolean }) => {
    const nr = Math.max(0, Math.min(sheet.rowCount - 1, r));
    const nc = Math.max(0, Math.min(sheet.colCount - 1, c));
    const input = tableRef.current?.querySelector<HTMLInputElement>(
      `[data-cell-r="${nr}"][data-cell-c="${nc}"]`,
    );
    if (input) {
      if (opts?.keepRect) movingInRangeRef.current = true;
      input.focus();
    }
  };

  /**
   * Per-key navigation policy:
   *   - Enter / Shift+Enter:    down / up
   *   - Tab / Shift+Tab:        right / left (preventDefault stops native focus traversal)
   *   - ArrowUp / ArrowDown:    always navigate (cells are short — text scrolling is rare)
   *   - ArrowLeft / ArrowRight: only navigate when caret is at the value's edge,
   *                             so users can still move the caret inside text
   *   - Shift+Arrow*:            extend the selection's far corner without
   *                              moving focus — Excel-style range grow.
   * Returns true when the key was handled so CellInput can preventDefault.
   */
  const handleNav = (
    r: number,
    c: number,
    e: React.KeyboardEvent<HTMLInputElement>,
  ): boolean => {
    const target = e.currentTarget;
    const atStart = target.selectionStart === 0 && target.selectionEnd === 0;
    const atEnd =
      target.selectionStart === target.value.length &&
      target.selectionEnd === target.value.length;

    // Shift+Arrow → extend the selection's far corner. Anchor stays
    // focused; we only move (r2, c2). Allowed regardless of caret
    // position because the user is clearly in "range mode" not "type mode".
    if (e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown' || e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      if (!selection) return false;
      let nr = selection.r2;
      let nc = selection.c2;
      if (e.key === 'ArrowUp') nr = Math.max(0, selection.r2 - 1);
      else if (e.key === 'ArrowDown') nr = Math.min(sheet.rowCount - 1, selection.r2 + 1);
      else if (e.key === 'ArrowLeft') nc = Math.max(0, selection.c2 - 1);
      else if (e.key === 'ArrowRight') nc = Math.min(sheet.colCount - 1, selection.c2 + 1);
      onExtend(nr, nc);
      return true;
    }

    // Tab / Enter inside a multi-cell selection: cycle the anchor within
    // the rectangle (Excel keyboard idiom). Tab walks row-major, Enter
    // walks column-major; Shift reverses. The selection rectangle is
    // preserved via the keepRect flag so the user can keep entering data
    // into the prepared range without losing it.
    if ((e.key === 'Tab' || e.key === 'Enter') && selection) {
      const range = rangeOf(selection);
      const isMulti = range.r1 !== range.r2 || range.c1 !== range.c2;
      if (isMulti) {
        const rows = range.r2 - range.r1 + 1;
        const cols = range.c2 - range.c1 + 1;
        const total = rows * cols;
        const dir = e.shiftKey ? -1 : 1;
        // Linear-index cycle, skipping cells covered by a merge (only
        // anchor cells render an input — landing on a covered cell
        // would silently fail). Bail after `total` steps so a range
        // entirely inside a merge doesn't loop forever.
        const indexOf = (rr: number, cc: number): number =>
          e.key === 'Tab'
            ? (rr - range.r1) * cols + (cc - range.c1)
            : (cc - range.c1) * rows + (rr - range.r1);
        const fromIdx = (idx: number): { r: number; c: number } => {
          if (e.key === 'Tab') {
            return { r: range.r1 + Math.floor(idx / cols), c: range.c1 + (idx % cols) };
          }
          return { r: range.r1 + (idx % rows), c: range.c1 + Math.floor(idx / rows) };
        };
        let idx = indexOf(r, c);
        for (let step = 0; step < total; step += 1) {
          idx = ((idx + dir) % total + total) % total;
          const { r: nr, c: nc } = fromIdx(idx);
          if (!isMergeCovered(sheet, nr, nc)) {
            focusCell(nr, nc, { keepRect: true });
            return true;
          }
        }
        return true; // no nav target; consume key anyway
      }
    }

    switch (e.key) {
      case 'Enter':
        focusCell(e.shiftKey ? r - 1 : r + 1, c);
        return true;
      case 'Tab':
        focusCell(r, e.shiftKey ? c - 1 : c + 1);
        return true;
      case 'ArrowUp':
        focusCell(r - 1, c);
        return true;
      case 'ArrowDown':
        focusCell(r + 1, c);
        return true;
      case 'ArrowLeft':
        if (!atStart) return false;
        focusCell(r, c - 1);
        return true;
      case 'ArrowRight':
        if (!atEnd) return false;
        focusCell(r, c + 1);
        return true;
      default:
        return false;
    }
  };

  const colLetters = useMemo(
    () => Array.from({ length: sheet.colCount }, (_, c) => colIndexToLetter(c)),
    [sheet.colCount],
  );

  // Native HTML5 drag-and-drop of an image file from the OS into the grid.
  // We accept the first image File on drop, locate the cell under the
  // cursor, and hand off to the editor's `onDropImageFile`. `dragOverDepth`
  // counts dragenter / dragleave to keep the indicator visible while the
  // cursor crosses internal element boundaries (a single-flag toggle would
  // flicker each time the user moves over a cell border).
  const [draggingFile, setDraggingFile] = useState(false);
  const dragDepthRef = useRef(0);

  const isFileDrag = (e: React.DragEvent): boolean => {
    // `e.dataTransfer.types` is always populated during drag (unlike `.files`,
    // which is empty until drop for security). "Files" is the canonical type
    // when the OS file manager is the source.
    const types = e.dataTransfer.types;
    return Array.from(types).includes('Files');
  };

  const cellAtPoint = (clientX: number, clientY: number): { r: number; c: number } | null => {
    if (!tableRef.current) return null;
    const tds = tableRef.current.querySelectorAll<HTMLElement>('td');
    for (const td of Array.from(tds)) {
      const inp = td.querySelector<HTMLElement>('[data-cell-r][data-cell-c]');
      if (!inp) continue;
      const rect = td.getBoundingClientRect();
      if (clientX < rect.left || clientX > rect.right) continue;
      if (clientY < rect.top || clientY > rect.bottom) continue;
      const rr = Number(inp.dataset.cellR);
      const cc = Number(inp.dataset.cellC);
      if (Number.isFinite(rr) && Number.isFinite(cc)) return { r: rr, c: cc };
    }
    return null;
  };

  return (
    <div
      className="relative inline-block"
      // Click on the wrapper outside any image overlay → drop image selection.
      // The table catches its own clicks via cell mousedown handlers, so this
      // only fires for the small empty regions around the grid.
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onImageSelect(null);
      }}
      onDragEnter={(e) => {
        if (!isFileDrag(e)) return;
        e.preventDefault();
        dragDepthRef.current += 1;
        setDraggingFile(true);
      }}
      onDragOver={(e) => {
        if (!isFileDrag(e)) return;
        // preventDefault is what makes the browser treat us as a valid drop
        // target — without this, `drop` never fires and the OS bounces the
        // file back to its source. dropEffect="copy" gives the right cursor.
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
      }}
      onDragLeave={(e) => {
        if (!isFileDrag(e)) return;
        dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
        if (dragDepthRef.current === 0) setDraggingFile(false);
      }}
      onDrop={(e) => {
        if (!isFileDrag(e)) return;
        e.preventDefault();
        dragDepthRef.current = 0;
        setDraggingFile(false);
        const files = Array.from(e.dataTransfer.files ?? []);
        const imageFile = files.find((f) => f.type.startsWith('image/'));
        if (!imageFile) return;
        // Drop coordinate → cell. Falls back to A1 when the cursor was
        // somewhere we can't resolve (header band, off-grid padding).
        const cell = cellAtPoint(e.clientX, e.clientY) ?? { r: 0, c: 0 };
        onDropImageFile(imageFile, cell.r, cell.c);
      }}
    >
    <table ref={tableRef} className="border-collapse text-xs font-mono select-text">
      <thead className="sticky top-0 z-10">
        <tr>
          <th className="w-10 sticky left-0 z-20 bg-secondary border border-border" />
          {colLetters.map((letter) => (
            <th
              key={letter}
              className="min-w-[80px] px-2 py-1 bg-secondary border border-border text-muted-foreground font-normal"
            >
              {letter}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {sheet.cells.map((row, r) => (
          <tr key={r}>
            <th className="sticky left-0 z-10 bg-secondary border border-border text-muted-foreground font-normal px-2">
              {r + 1}
            </th>
            {row.map((cell, c) => {
              // Cells covered by a merge (but not the anchor) are skipped
              // entirely — the anchor's <td> spans across them via rowSpan
              // / colSpan. Without this guard the rowSpan'd cell would
              // collide with the next row's <td> and the table would
              // visibly drift right.
              if (isMergeCovered(sheet, r, c)) return null;
              const merge = mergeAtAnchor(sheet, r, c);
              const rowSpan = merge ? merge.r2 - merge.r1 + 1 : 1;
              const colSpan = merge ? merge.c2 - merge.c1 + 1 : 1;
              const isAnchor = selection?.r === r && selection?.c === c;
              const inRange = isInRange(selection, r, c);
              // Selection rectangle perimeter — only the outer cells of
              // the selection draw the heavy primary border, so a multi-
              // cell selection looks like one continuous rectangle rather
              // than every cell having its own outline.
              const sr = selection ? rangeOf(selection) : null;
              const onSelTop = !!sr && r === sr.r1 && c >= sr.c1 && c <= sr.c2;
              const onSelBottom = !!sr && r === sr.r2 && c >= sr.c1 && c <= sr.c2;
              const onSelLeft = !!sr && c === sr.c1 && r >= sr.r1 && r <= sr.r2;
              const onSelRight = !!sr && c === sr.c2 && r >= sr.r1 && r <= sr.r2;
              // Copy / cut marker — same idea, dashed accent border.
              const onCopyTop = !!copyMarker && r === copyMarker.r1 && c >= copyMarker.c1 && c <= copyMarker.c2;
              const onCopyBottom = !!copyMarker && r === copyMarker.r2 && c >= copyMarker.c1 && c <= copyMarker.c2;
              const onCopyLeft = !!copyMarker && c === copyMarker.c1 && r >= copyMarker.r1 && r <= copyMarker.r2;
              const onCopyRight = !!copyMarker && c === copyMarker.c2 && r >= copyMarker.r1 && r <= copyMarker.r2;
              const copyColor = copyMarker?.mode === 'cut' ? 'border-amber-500' : 'border-primary';
              return (
                <td
                  key={c}
                  rowSpan={rowSpan}
                  colSpan={colSpan}
                  className={cn(
                    'relative border border-border p-0',
                    isAnchor && 'outline outline-2 outline-primary outline-offset-[-2px] z-[1]',
                  )}
                  style={{
                    background: cell.style?.bgColor ? `#${cell.style.bgColor}` : undefined,
                  }}
                  onMouseDown={(e) => {
                    if (e.shiftKey) {
                      // Extend the existing selection's far corner — same
                      // anchor, new (r2, c2). If there's no selection yet,
                      // treat as a normal click.
                      if (selection) {
                        e.preventDefault();
                        onExtend(r, c);
                        return;
                      }
                    }
                    beginDrag(r, c);
                  }}
                  onMouseEnter={() => {
                    if (draggingRef.current) onExtend(r, c);
                  }}
                >
                  {/* Range tint overlay — drawn above the cell bg color but
                      below the input so styling and text stay legible.
                      pointer-events:none keeps the input clickable. */}
                  {inRange && !isAnchor ? (
                    <div className="absolute inset-0 bg-primary/15 pointer-events-none" />
                  ) : null}
                  {/* Selection rectangle perimeter — drawn as 4 absolute
                      edge bars so the multi-cell selection has one clean
                      outer border instead of N×M outlines stacking. */}
                  {sr ? (
                    <>
                      {onSelTop ? <div className="absolute left-0 right-0 top-0 h-[2px] bg-primary pointer-events-none z-[2]" /> : null}
                      {onSelBottom ? <div className="absolute left-0 right-0 bottom-0 h-[2px] bg-primary pointer-events-none z-[2]" /> : null}
                      {onSelLeft ? <div className="absolute top-0 bottom-0 left-0 w-[2px] bg-primary pointer-events-none z-[2]" /> : null}
                      {onSelRight ? <div className="absolute top-0 bottom-0 right-0 w-[2px] bg-primary pointer-events-none z-[2]" /> : null}
                    </>
                  ) : null}
                  {/* Copy / cut marker — dashed border on the source rect. */}
                  {copyMarker ? (
                    <>
                      {onCopyTop ? <div className={cn('absolute left-0 right-0 top-0 border-t-2 border-dashed pointer-events-none z-[2]', copyColor)} /> : null}
                      {onCopyBottom ? <div className={cn('absolute left-0 right-0 bottom-0 border-b-2 border-dashed pointer-events-none z-[2]', copyColor)} /> : null}
                      {onCopyLeft ? <div className={cn('absolute top-0 bottom-0 left-0 border-l-2 border-dashed pointer-events-none z-[2]', copyColor)} /> : null}
                      {onCopyRight ? <div className={cn('absolute top-0 bottom-0 right-0 border-r-2 border-dashed pointer-events-none z-[2]', copyColor)} /> : null}
                    </>
                  ) : null}
                  <CellInput
                    r={r}
                    c={c}
                    cell={cell}
                    onFocus={() => {
                      // Tab/Enter cycling sets movingInRangeRef so we
                      // keep the surrounding rectangle; any other focus
                      // (click, arrow nav, keyboard tab from outside)
                      // collapses the selection to this cell.
                      if (movingInRangeRef.current) {
                        movingInRangeRef.current = false;
                        onMoveAnchor(r, c);
                      } else {
                        onSelect(r, c);
                      }
                    }}
                    onCommit={(text) => onCommit(r, c, text)}
                    onNav={handleNav}
                    onEditingChange={onEditingChange}
                  />
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
    <ImageOverlayLayer
      tableRef={tableRef}
      sheet={sheet}
      images={images}
      selectedImageId={selectedImageId}
      onImageSelect={onImageSelect}
      onImagePatch={onImagePatch}
    />
    {/* File-drop visual: dashed-border tint covering the whole grid while
        the user is dragging an OS file overhead. Pointer-events:none so it
        never intercepts the actual drop event (which lives on the wrapper). */}
    {draggingFile ? (
      <div
        className="absolute inset-0 bg-primary/10 border-2 border-dashed border-primary pointer-events-none flex items-start justify-center"
        style={{ zIndex: 20 }}
      >
        <div className="mt-3 px-3 py-1.5 bg-primary text-primary-foreground text-xs rounded shadow">
          放開以將圖片插入到游標所在儲存格
        </div>
      </div>
    ) : null}
    </div>
  );
}

/**
 * Absolute-positioned overlay layer for floating images.
 *
 * The grid uses a real `<table>` with auto-sized cells, so we can't compute
 * cell positions from a fixed row-height / col-width — we have to measure.
 * On every layout-affecting change (image set, sheet structure, table
 * resize) we walk the images, find each anchor `<td>` via the cell input's
 * data-cell-r / -c attrs, and read its `getBoundingClientRect()` relative
 * to the table.
 *
 * Drag: mousedown captures every cell's bounding box once; mousemove
 * updates a transient `dragDelta`; mouseup finds the cell containing the
 * cursor and commits a new anchor. Snap-to-cell — sheet anchors are cell
 * coordinates, not free pixel offsets.
 *
 * Resize: bottom-right handle. Aspect-locked by default; hold Shift while
 * dragging to free the ratio. Live preview via a `resizeDelta` ref.
 */
function ImageOverlayLayer({
  tableRef,
  sheet,
  images,
  selectedImageId,
  onImageSelect,
  onImagePatch,
}: {
  tableRef: React.RefObject<HTMLTableElement>;
  sheet: XlsxSheet;
  images: XlsxImage[] | undefined;
  selectedImageId: string | null;
  onImageSelect: (id: string | null) => void;
  onImagePatch: (id: string, patch: Partial<XlsxImage>) => void;
}): JSX.Element | null {
  // Per-image measured rect, in coords relative to the table.
  const [positions, setPositions] = useState<
    Array<{ id: string; left: number; top: number; widthPx: number; heightPx: number }>
  >([]);

  // Transient gesture state — set on mousedown, cleared on mouseup. Stored
  // in refs because the window-level mousemove handler reads from a stale
  // closure if we used setState directly.
  const dragRef = useRef<{
    id: string;
    startX: number;
    startY: number;
    cells: Array<{ r: number; c: number; left: number; top: number; right: number; bottom: number }>;
  } | null>(null);
  const resizeRef = useRef<{
    id: string;
    startX: number;
    startY: number;
    baseW: number;
    baseH: number;
    aspect: number;
    curW: number;
    curH: number;
  } | null>(null);
  // Force-rerender ticks for live preview during drag / resize.
  const [dragTick, setDragTick] = useState(0);
  const [resizeTick, setResizeTick] = useState(0);
  const [dragDelta, setDragDelta] = useState<{ id: string; dx: number; dy: number } | null>(null);

  // Measure positions whenever the image set or sheet structure changes,
  // and on table resize. We use useLayoutEffect so the overlay paints in
  // the same frame as the table's relayout — no flicker on sheet swap.
  useLayoutEffect(() => {
    const recompute = () => {
      if (!images || images.length === 0 || !tableRef.current) {
        setPositions([]);
        return;
      }
      const table = tableRef.current;
      const tableRect = table.getBoundingClientRect();
      const next: typeof positions = [];
      for (const img of images) {
        const inp = table.querySelector<HTMLElement>(
          `[data-cell-r="${img.anchorRow}"][data-cell-c="${img.anchorCol}"]`,
        );
        const td = inp?.closest('td');
        if (!td) continue;
        const r = td.getBoundingClientRect();
        next.push({
          id: img.id,
          left: r.left - tableRect.left,
          top: r.top - tableRect.top,
          widthPx: img.widthEmu / EMU_PER_PX,
          heightPx: img.heightEmu / EMU_PER_PX,
        });
      }
      setPositions(next);
    };
    recompute();
    if (!tableRef.current) return;
    const ro = new ResizeObserver(() => recompute());
    ro.observe(tableRef.current);
    return () => ro.disconnect();
  }, [images, sheet.cells, sheet.rowCount, sheet.colCount, sheet.merges, tableRef]);

  if (!images || images.length === 0) return null;

  const startDrag = (e: React.MouseEvent, id: string) => {
    e.preventDefault();
    e.stopPropagation();
    onImageSelect(id);
    if (!tableRef.current) return;
    // Snapshot cell rectangles up-front so mouseup's "what cell am I over"
    // lookup doesn't pay layout costs per move.
    const cells: Array<{ r: number; c: number; left: number; top: number; right: number; bottom: number }> = [];
    const tds = tableRef.current.querySelectorAll<HTMLElement>('td');
    tds.forEach((td) => {
      const inp = td.querySelector<HTMLElement>('[data-cell-r][data-cell-c]');
      if (!inp) return;
      const rr = Number(inp.dataset.cellR);
      const cc = Number(inp.dataset.cellC);
      if (!Number.isFinite(rr) || !Number.isFinite(cc)) return;
      const r = td.getBoundingClientRect();
      cells.push({ r: rr, c: cc, left: r.left, top: r.top, right: r.right, bottom: r.bottom });
    });
    dragRef.current = { id, startX: e.clientX, startY: e.clientY, cells };

    const onMove = (ev: MouseEvent) => {
      if (!dragRef.current) return;
      setDragDelta({
        id: dragRef.current.id,
        dx: ev.clientX - dragRef.current.startX,
        dy: ev.clientY - dragRef.current.startY,
      });
      setDragTick((n) => n + 1);
    };
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      const drag = dragRef.current;
      dragRef.current = null;
      setDragDelta(null);
      if (!drag) return;
      const movedX = Math.abs(ev.clientX - drag.startX);
      const movedY = Math.abs(ev.clientY - drag.startY);
      // <3 px ≡ click; selection was already set on mousedown.
      if (movedX < 3 && movedY < 3) return;
      const hit = drag.cells.find(
        (cell) =>
          ev.clientX >= cell.left &&
          ev.clientX <= cell.right &&
          ev.clientY >= cell.top &&
          ev.clientY <= cell.bottom,
      );
      if (!hit) return;
      onImagePatch(drag.id, { anchorRow: hit.r, anchorCol: hit.c });
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const startResize = (e: React.MouseEvent, id: string, baseW: number, baseH: number) => {
    e.preventDefault();
    e.stopPropagation();
    onImageSelect(id);
    resizeRef.current = {
      id,
      startX: e.clientX,
      startY: e.clientY,
      baseW,
      baseH,
      aspect: baseW / Math.max(1, baseH),
      curW: baseW,
      curH: baseH,
    };
    const onMove = (ev: MouseEvent) => {
      if (!resizeRef.current) return;
      const dx = ev.clientX - resizeRef.current.startX;
      const dy = ev.clientY - resizeRef.current.startY;
      let w = Math.max(8, resizeRef.current.baseW + dx);
      let h = Math.max(8, resizeRef.current.baseH + dy);
      // Default aspect-lock; Shift releases it. Drive the lock from whichever
      // axis the user moved more so the image doesn't "snap" sideways.
      if (!ev.shiftKey) {
        if (Math.abs(dx) >= Math.abs(dy)) {
          h = Math.max(8, w / resizeRef.current.aspect);
        } else {
          w = Math.max(8, h * resizeRef.current.aspect);
        }
      }
      resizeRef.current.curW = w;
      resizeRef.current.curH = h;
      setResizeTick((n) => n + 1);
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      const r = resizeRef.current;
      resizeRef.current = null;
      setResizeTick((n) => n + 1);
      if (!r) return;
      // Skip commit if the gesture didn't actually change anything (e.g. the
      // user clicked the handle without moving — gracious no-op).
      if (Math.round(r.curW) === Math.round(r.baseW) && Math.round(r.curH) === Math.round(r.baseH)) return;
      onImagePatch(r.id, {
        widthEmu: Math.round(r.curW * EMU_PER_PX),
        heightEmu: Math.round(r.curH * EMU_PER_PX),
      });
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  // Reference dragTick / resizeTick so React knows these renders are tied
  // to the gesture state. The values themselves are unused in JSX.
  void dragTick;
  void resizeTick;

  return (
    <div
      className="absolute inset-0 pointer-events-none"
      style={{ zIndex: 5 }}
    >
      {positions.map((p) => {
        const img = images.find((i) => i.id === p.id);
        if (!img) return null;
        const isSelected = p.id === selectedImageId;
        let left = p.left;
        let top = p.top;
        let widthPx = p.widthPx;
        let heightPx = p.heightPx;
        if (dragDelta && dragDelta.id === p.id) {
          left += dragDelta.dx;
          top += dragDelta.dy;
        }
        if (resizeRef.current?.id === p.id) {
          widthPx = resizeRef.current.curW;
          heightPx = resizeRef.current.curH;
        }
        return (
          <div
            key={p.id}
            // The whole image acts as a drag target (cursor: grab → grabbing
            // mid-drag), but until now the gesture had no title hint — the
            // grab cursor changes only on hover-into, so a user just looking
            // at an idle selected image had no surface telling them they
            // could pick it up. Every other cursor-grab drag surface in
            // the project is titled:
            //   DocxEditor.tsx:2486  title="拖曳到任意位置"  (block grip — BlockRow)
            //   DocxEditor.tsx:2827  title="拖曳到任意位置"  (table grip — TableBlockRow with KIND_LABEL.table)
            //   DocxEditor.tsx:3164  title="拖曳到任意位置"  (floating image grip — ImageBlockRow with KIND_LABEL.image)
            //   PptxEditor.tsx:2169  title="拖曳移動 · Shift 鎖軸 · Alt 拖曳複製"
            // — same R78 cross-editor consistency rationale: the docx
            // floating-image grip and this xlsx image wrapper drive the
            // exact same model (an image with a cell / paragraph anchor
            // that can be moved freely on the page), so they should read
            // the same way on hover. Reusing the docx wording verbatim
            // (rather than coining yet a third phrasing) keeps the
            // floating-image hover copy identical across docx and xlsx,
            // mirroring what R78 just did for the resize handle next door
            // (XlsxEditor.tsx:3591 = DocxEditor.tsx:3200 verbatim).
            // Note: title sits on the wrapper, not the resize square — the
            // resize handle has its own title at line 3591 below, and a
            // child title overrides the parent's during hover, so the two
            // hints don't compete.
            // R140 — refreshed line numbers (2401/2663/2999 → 2486/2827/3164,
            // 2043 → 2169, 3396/3035 → 3591/3200, 3413 → 3591) and corrected
            // two swapped semantic labels: the previous block had the smaller
            // line number marked (table grip) and the middle line marked
            // (block grip), but the current source order is BlockRow first
            // (line 2486 — generic block / heading / list / paragraph), then
            // a table-rendering component (line 2827, identifiable via
            // KIND_LABEL.table at line 2835), then an image-rendering
            // component (line 3164, identifiable via KIND_LABEL.image at
            // line 3172). Same self-quoting paradigm flagged in R133 / R135:
            // a stale cross-reference comment that named non-existent line
            // numbers AND mis-tagged two of the three sites it was citing
            // for consistency. Fixing both in one pass so the next reader
            // jumping to a referenced line lands on the actual title= they
            // expected, with the right component label next to it.
            title="拖曳到任意位置"
            className={cn(
              'absolute pointer-events-auto select-none',
              isSelected ? 'ring-2 ring-primary' : 'ring-1 ring-transparent hover:ring-primary/60',
            )}
            style={{
              left,
              top,
              width: widthPx,
              height: heightPx,
              cursor: dragRef.current?.id === p.id ? 'grabbing' : 'grab',
            }}
            onMouseDown={(e) => startDrag(e, p.id)}
          >
            {img.dataUrl ? (
              <img
                src={img.dataUrl}
                alt=""
                draggable={false}
                className="block w-full h-full object-fill"
              />
            ) : (
              // No preview data URL (image came from a parsed file) — show a
              // neutral placeholder so the user can still position / resize.
              <div className="w-full h-full bg-muted flex items-center justify-center text-[10px] text-muted-foreground">
                IMG
              </div>
            )}
            {isSelected ? (
              // Match the canonical image-resize tooltip wording / format used
              // by the sibling editors so the same gesture reads the same way
              // across all three:
              //   DocxEditor.tsx:3035   title="拖曳調整大小 · Shift 解除等比例"
              //   PptxEditor.tsx:2043   title="拖曳移動 · Shift 鎖軸 · Alt 拖曳複製"
              //   PptxEditor.tsx:2208+  title="拖曳調整大小 · Shift 鎖定等比例"
              //                          (corner handles, R118)
              // The pptx corner-resize gesture DOES toggle aspect with Shift —
              // verified at PptxEditor.tsx:1972-1989's `if (shift && isCorner)`
              // ratio-locked branch — but with INVERTED semantic: pptx shapes
              // default-free + Shift locks (PowerPoint / Adobe convention),
              // while docx / xlsx images default-locked + Shift releases (Word
              // convention). Both are intentional, matching their host
              // platform's muscle memory; the verb after `Shift` flips
              // accordingly (`鎖定` vs `解除`). R118 added tooltips to all 8
              // pptx resize handles closing the prior 2-vs-1 silent-outlier
              // gap; this comment was updated then to reflect reality.
              // Previously this lone xlsx handle used `拖曳以調整大小（按住
              // Shift 取消等比）` — same gesture, same Shift semantics, but
              // parenthesized rather than middle-dotted, and with `取消等比`
              // instead of `解除等比例`. The middle-dot `·` separator is the
              // documented project convention for compound tooltip hint
              // strings (App.tsx:1160-1162 explicitly cites the R39
              // ContextItem layout, and R76's `saveAsHint` reuses it: `儲存
              // (Ctrl+S) · 另存新檔 (Ctrl+Shift+S)`). The image-resize
              // gesture is identical across docx / xlsx (aspect-locked by
              // default, Shift releases the lock) — there's no editor-
              // specific reason for the wording to differ. Aligning with
              // the docx wording verbatim gives a fresh user opening a
              // workspace with both an embedded docx tab and an xlsx tab
              // identical hover copy on the same handle, instead of having
              // to re-parse two different phrasings of the same shortcut.
              <div
                role="button"
                tabIndex={-1}
                title="拖曳調整大小 · Shift 解除等比例"
                className="absolute -bottom-1.5 -right-1.5 h-3 w-3 bg-primary border border-background rounded-sm"
                style={{ cursor: 'nwse-resize' }}
                onMouseDown={(e) => startResize(e, p.id, widthPx, heightPx)}
              />
            ) : null}
          </div>
        );
      })}
    </div>
  );
}

function CellInput({
  r,
  c,
  cell,
  onCommit,
  onFocus,
  onNav,
  onEditingChange,
}: {
  r: number;
  c: number;
  cell: XlsxCell;
  onCommit: (text: string) => void;
  onFocus: () => void;
  /** Spreadsheet-nav callback. Returns true when the key was consumed. */
  onNav: (r: number, c: number, e: React.KeyboardEvent<HTMLInputElement>) => boolean;
  /** Notifies the editor when this cell enters / leaves edit mode so
   *  clipboard handlers can defer to the input while typing. */
  onEditingChange: (editing: boolean) => void;
}): JSX.Element {
  // Two distinct states for a focused cell:
  //   - select mode (default):  input shows the computed value with the
  //     caret hidden + cursor:cell; the whole text is auto-selected so
  //     typing replaces. Ctrl+C/V act on the cell range.
  //   - edit mode:              input shows the formula source so the
  //     user edits "=A1+B1" not "42"; Enter / Tab commit + navigate.
  //     Entered via F2, double-click, Backspace/Delete, IME composition,
  //     or typing any printable key.
  //
  // We deliberately do NOT use `readOnly` for select mode because
  // readOnly inputs reject IME composition, which would force Chinese /
  // Japanese / Korean users to F2 before every cell — that was the
  // single largest source of "卡頓" in the prior design. Instead, the
  // input is always editable and the first onChange / compositionstart
  // auto-transitions to edit mode.
  const [editMode, setEditMode] = useState(false);
  // Ref mirror — needed because blur() can fire after setEditMode(false)
  // is queued but before the closure sees the new state, leading to a
  // double commit. Reading from the ref keeps transitions idempotent.
  const editModeRef = useRef(false);
  const editable = cell.formula ?? cell.text;
  const display = cell.text;
  const [value, setValue] = useState(display);
  // Pull external changes back into the input only while in select mode;
  // mid-edit the user is the source of truth.
  useEffect(() => {
    if (!editModeRef.current) setValue(display);
  }, [display]);

  const enterEdit = (initial?: string) => {
    if (editModeRef.current) return;
    editModeRef.current = true;
    setEditMode(true);
    if (initial !== undefined) setValue(initial);
    else setValue(editable);
    onEditingChange(true);
  };

  const exitEdit = (commit: boolean) => {
    if (!editModeRef.current) return;
    editModeRef.current = false;
    setEditMode(false);
    onEditingChange(false);
    if (commit && value !== editable) {
      onCommit(value);
      // The model update on the next render refreshes display and the
      // useEffect resyncs value. Don't snap to display here — there's
      // a brief gap where display is still the old value.
    } else {
      setValue(display);
    }
  };

  const styleAttr = cell.style;
  return (
    <input
      type="text"
      data-cell-r={r}
      data-cell-c={c}
      value={value}
      // Native onChange auto-transitions to edit mode on the first edit.
      // Covers all input paths: regular keys, paste, IME composition end
      // (which triggers an input event in all major browsers).
      onChange={(e) => {
        const v = e.target.value;
        setValue(v);
        if (!editModeRef.current) {
          editModeRef.current = true;
          setEditMode(true);
          // Switch to formula source so further edits modify the source,
          // not whatever literal text the user just typed onto the
          // computed display. But only if they typed onto a non-empty
          // *display* — for an empty display the typed char already is
          // the source, no point swapping.
          // Caveat: in the "type to overwrite" idiom we WANT the typed
          // char to stand alone, not get prefixed by editable. Since
          // the input started in select mode with all-text-selected,
          // typing replaces the selected display, leaving exactly the
          // typed chars in `v` — so we keep `v` as-is. ✓
          onEditingChange(true);
        }
      }}
      onCompositionStart={() => {
        // IME composition (Chinese / Japanese / Korean). Treat the same
        // as onChange — flip into edit mode so subsequent recompute
        // sweeps don't snap value back to display mid-composition.
        if (!editModeRef.current) {
          editModeRef.current = true;
          setEditMode(true);
          onEditingChange(true);
        }
      }}
      onFocus={(e) => {
        // Auto select-all so a follow-up printable key / IME composition
        // replaces content (matches Excel's "selected cell" behavior).
        e.currentTarget.select();
        onFocus();
      }}
      onClick={(e) => {
        // Single-click on a focused cell in select mode: re-select all
        // so the caret doesn't drift into the middle of the displayed
        // text. Excel does the same — clicking a selected cell is a
        // no-op visually. Doesn't fire on double-click first half
        // because dblclick is a separate event.
        if (!editModeRef.current) {
          e.currentTarget.select();
        }
      }}
      onDoubleClick={(e) => {
        // Excel parity: double-click puts the cell into edit mode with
        // the caret at the end of the current text.
        if (!editModeRef.current) {
          enterEdit();
          const el = e.currentTarget;
          requestAnimationFrame(() => {
            const len = el.value.length;
            el.setSelectionRange(len, len);
          });
        }
      }}
      onBlur={() => {
        // If focus left while editing, commit. exitEdit is idempotent
        // so calling it after Enter/Tab already committed is a no-op.
        if (editModeRef.current) exitEdit(true);
      }}
      onKeyDown={(e) => {
        // Esc cancels an in-progress edit (no commit). In select mode
        // it does nothing here so the editor-level handler can clear
        // the copy marker.
        if (e.key === 'Escape') {
          if (editModeRef.current) {
            exitEdit(false);
            e.preventDefault();
          }
          return;
        }

        if (!editModeRef.current) {
          // F2 → enter edit mode preserving content, caret at end.
          if (e.key === 'F2') {
            const el = e.currentTarget;
            enterEdit();
            requestAnimationFrame(() => {
              const len = el.value.length;
              el.setSelectionRange(len, len);
            });
            e.preventDefault();
            return;
          }
          // Backspace / Delete on a select-mode cell blanks it and
          // immediately enters edit mode so the next keystroke is a
          // fresh edit on an empty cell.
          if (e.key === 'Backspace' || e.key === 'Delete') {
            enterEdit('');
            e.preventDefault();
            return;
          }
          // Other keys: let onNav decide (Enter/Tab/Arrow* navigate).
          // Printable keys + IME aren't intercepted here — they fall
          // through to native input handling, which (because text is
          // all-selected) replaces display with the typed char and
          // fires onChange → auto-transitions to edit mode.
          if (onNav(r, c, e)) e.preventDefault();
          return;
        }

        // Edit mode: commit + navigate keys.
        // Enter / Shift+Enter:    down / up
        // Tab / Shift+Tab:        right / left (or cycle inside range)
        // ArrowUp / ArrowDown:    Excel-style commit + move
        // R232 — skip the commit-and-navigate path during IME composition.
        // Excel-on-Windows cell editing is a top-3 CJK typing context (使
        // 用 wrap 注音 / 拼音 / かな / 한글 IME). Enter confirms an IME
        // candidate; without isComposing guard, exitEdit(true) commits
        // the raw pre-confirmation buffer (bopomofo / pinyin / etc.)
        // instead of the chosen CJK glyph, then onNav moves to the next
        // cell. User loses the typed Chinese / Japanese / Korean
        // entirely — among the most damaging IME bugs in any editor.
        // Tab / ArrowUp / ArrowDown are also conceptually commit-and-
        // move keystrokes; some IMEs use Tab for candidate cycling
        // (Korean), and arrow keys navigate the IME candidate list, so
        // we skip ALL four during composition. Mirrors R231's TabBar
        // rename guard, which adopted the same `if (e.nativeEvent
        // .isComposing) return;` short-circuit.
        if (e.nativeEvent.isComposing) return;
        if (
          e.key === 'Enter' ||
          e.key === 'Tab' ||
          e.key === 'ArrowUp' ||
          e.key === 'ArrowDown'
        ) {
          exitEdit(true);
          if (onNav(r, c, e)) e.preventDefault();
          return;
        }
        // ArrowLeft / ArrowRight stay in the input for caret movement.
        // Other printable keys flow to native input → onChange.
      }}
      className={cn(
        'w-full min-w-[80px] px-2 py-1 bg-transparent outline-none',
        editMode ? 'cursor-text bg-primary/5' : 'cursor-cell',
      )}
      style={{
        // Hide the caret in select mode so the cell looks "selected"
        // not "being typed into" — visual signal that maps to Excel's
        // distinct selected-vs-edit cursor.
        caretColor: editMode ? undefined : 'transparent',
        fontWeight: styleAttr?.bold ? 'bold' : undefined,
        fontStyle: styleAttr?.italic ? 'italic' : undefined,
        textDecoration: styleAttr?.underline ? 'underline' : undefined,
        color: styleAttr?.fontColor ? `#${styleAttr.fontColor}` : undefined,
        textAlign: styleAttr?.align,
        fontSize: styleAttr?.fontSize ? `${styleAttr.fontSize}px` : undefined,
        fontFamily: withEmojiFallback(styleAttr?.fontFamily),
      }}
    />
  );
}
