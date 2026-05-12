/**
 * PowerPoint editor — slide-by-slide text run editor.
 *
 * Path B MVP+ UX:
 *   - Left rail: slide list (numbered).
 *   - Right pane: each `<a:r>` text run from the active slide as a textarea.
 *   - Formatting toolbar (粗體 / 斜體 / 字色 / 字級) acts on the focused run.
 *   - Empty (zero-byte) pptx tabs show a placeholder.
 *
 * Edits flow through a debounced re-serialize that re-loads the original zip
 * and only patches `<a:r>` content for modified slides; everything else is
 * byte-preserved.
 */

import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  AlignCenter,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignLeft,
  AlignRight,
  AlignStartHorizontal,
  ArrowRight,
  Bold,
  ChevronDown,
  ChevronUp,
  Circle,
  Copy,
  GripVertical,
  Image as ImageIcon,
  Italic,
  LayoutTemplate,
  List,
  Palette,
  Play,
  Plus,
  Presentation,
  Search,
  Shapes,
  Square,
  Trash2,
  Triangle,
  Type,
  Underline,
  X,
} from 'lucide-react';
import type { PptxTab } from '../types/tab';
import { useWorkspace } from '../store/workspace';
import {
  PPTX_LAYOUTS,
  type PptxFrame,
  type PptxLayoutId,
  type PptxModel,
  type PptxRunStyle,
  type PptxShapeKind,
  type PptxTextRun,
  addPictureToSlide,
  addShapeToSlide,
  addTextBoxToSlide,
  applyLayoutToSlide,
  createBlankPptx,
  deleteSlide,
  deleteTextBoxFromSlide,
  duplicateShapeOnSlide,
  duplicateSlide,
  moveShapeOnSlide,
  parsePptx,
  reorderSlides,
  serializePptx,
} from '../lib/pptx-adapter';
import { clampToViewport, cn } from '../lib/utils';
import { notify } from '../store/toast';
import { FONT_FAMILIES, withEmojiFallback } from '../lib/font-families';
import { useFormatShortcuts } from '../lib/use-format-shortcuts';
import { useUndoableState, useUndoShortcuts } from '../lib/use-undoable-state';
import { registerEditorFlush } from '../lib/editor-flush';
import { FindReplaceDialog, type SearchSegment } from './FindReplaceDialog';
import { GoToDialog } from './GoToDialog';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from './ui/dropdown';

interface Props {
  tab: PptxTab;
}

/**
 * Per-tab active-slide memory across remounts within the renderer-process
 * session. Mirrors `scrollMemory` in MarkdownEditor and is here for the
 * same reason: `EditorSurface` keys its `<ErrorBoundary key={active.id}>`
 * on the active tab id, so every tab switch unmounts the editor and the
 * fresh mount calls `setActiveIdx(0)` after the async parse completes —
 * meaning a user editing slide 12 of a 30-slide deck, glancing at another
 * tab, and coming back lands on slide 1 with no record of where they were.
 * Stashing on unmount and clamping on restore (against the freshly-parsed
 * slide count) handles the case where AI/undo shrunk the deck during the
 * round-trip. Map persists for the renderer-process lifetime; cleared on
 * app reload — same volatility as MarkdownEditor's `scrollMemory`. Bounded
 * by the # of pptx tabs touched in a session, so GC isn't worth the wire-up.
 */
const slideMemory = new Map<string, number>();
/**
 * Per-tab outline-pane scroll memory — same shape and lifetime as
 * `slideMemory` above, but for the Nav panel on the left. Without it,
 * the user could scroll the slide outline to read titles for slides
 * 30-40 in a long deck, switch tabs, switch back, and the pane jumps
 * back to the top while the deck itself remembered where they were
 * (slideMemory). Mirrors the same fix in DocxEditor.tsx (Round 35,
 * `navScrollMemory`). Renderer-process lifetime; cleared on reload.
 */
const navScrollMemory = new Map<string, number>();
/**
 * Per-tab slide-rail scroll memory — the rail (left thumbnail column)
 * is a different scroll container from the nav panel that Round 36
 * fixed. Round 30 already resets desk scroll on slide change, and the
 * rail auto-scrolls the active row into view via scrollIntoView; what
 * was still missing is preserving the user's *manual* scroll across
 * tab swaps. In a 50-slide deck, browsing rail slides 30-45 to pick
 * which one to duplicate, switching tabs and switching back, snapped
 * the rail back to wherever the active row was — losing the user's
 * "I was looking at this section" context. Same shape and lifetime as
 * `navScrollMemory`.
 */
const railScrollMemory = new Map<string, number>();

export function PptxEditor({ tab }: Props): JSX.Element {
  const patchTab = useWorkspace((s) => s.patchTab);
  const markTabDirty = useWorkspace((s) => s.markTabDirty);
  const initialBytesRef = useRef<Uint8Array>(tab.data);
  const [model, setModel, undoApi] = useUndoableState<PptxModel | null>(null);
  const [activeIdx, setActiveIdx] = useState(0);
  // Mirror of activeIdx readable from cleanup. The parse effect's cleanup
  // captures `activeIdx` at the time the effect ran (always 0, the initial
  // useState value), so we can't read it directly there. A ref updated via
  // its own effect gives us the latest value at unmount time.
  const activeIdxRef = useRef(0);
  useEffect(() => {
    activeIdxRef.current = activeIdx;
  }, [activeIdx]);
  // Desk scrollport — switching slides should drop the user back to the top
  // of the new slide so the canvas is visible. Without this, navigating
  // (rail click / Ctrl+PgDn / Ctrl+G / NavPanel jump) inherits the previous
  // slide's scrollTop: a user who scrolled to the notes editor of slide 5
  // and then jumps to slide 12 sees slide 12's notes pane while slide 12's
  // canvas sits invisibly above the fold — and Ctrl+PgDn looks like it did
  // nothing. PowerPoint's slide view re-centers the new slide on every
  // navigation; mirroring that here. Find→Locate's scrollIntoView still
  // composes correctly: this effect resets to 0, then locateFindResult's
  // rAF scrolls only if the matched run is below the canvas-fold.
  const deskScrollRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const el = deskScrollRef.current;
    if (el) el.scrollTop = 0;
  }, [activeIdx]);
  // Slide title outline pane (Round 76) — mirrors Word's Navigation Pane.
  // Default closed: a heading-less Word doc shows "(沒有標題)" which is
  // friction; PowerPoint decks usually have at least one titled slide so
  // we could default-open, but keep parity with Word's choice and let
  // localStorage carry the user's preference.
  const [navOpen, setNavOpen] = useState<boolean>(() => {
    try {
      return localStorage.getItem('gendoc.pptxNavOpen') === '1';
    } catch {
      return false;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem('gendoc.pptxNavOpen', navOpen ? '1' : '0');
    } catch {
      /* no-op: private mode */
    }
  }, [navOpen]);
  // OS-level drag-and-drop of an image file from File Explorer (Round 73).
  // `dragDepthRef` counters dragenter/dragleave so the indicator stays
  // visible while the cursor crosses internal boundaries (toolbar, slide
  // strip, shape borders).
  const [draggingFile, setDraggingFile] = useState(false);
  const dragDepthRef = useRef(0);
  const [activeRunId, setActiveRunId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [presenting, setPresenting] = useState(false);
  const serializeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  // Latest model awaiting flush. The flush callback (registered with the
  // editor-flush registry) reads this so a save issued mid-debounce picks
  // up the freshest edits, not whatever the most recent React closure
  // happened to capture.
  const pendingModelRef = useRef<PptxModel | null>(null);
  // R95 — reentrancy guard for structural slide ops (duplicate / delete /
  // reorder / add text-box / add shape / add picture). All of these flow
  // through runStructuralOp, which awaits three async steps in series
  // (serializePptx → op → parsePptx) while sharing one mutable
  // initialBytesRef.current. A second click before the first completes
  // races on that ref: op #1 reads bytes B0, computes B1; op #2 reads B0
  // (still — #1 hasn't written yet), computes B1' from the SAME starting
  // bytes; whichever finishes second wins, the other op's slide vanishes.
  // The smoking gun is the per-row rail "複製此投影片" button at line 2867
  // which has no boundary disable (上移/下移 disable at i===0 / slideCount-1,
  // 刪除 disables at slideCount<=1, 複製 always-clickable) — rapid double-
  // clicks on duplicate are the path most likely to fire it. Same shape as
  // App.tsx::performSave's `savingRef` guard at App.tsx:316-328 — early
  // return + try/finally so a thrown error still releases the flag.
  const structuralOpInFlightRef = useRef(false);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    // Brand-new (zero-byte) tabs get bootstrapped with a one-slide blank deck
    // so the user lands on an editable text frame instead of the placeholder.
    const ensureBytes = async (): Promise<Uint8Array> => {
      if (tab.data.byteLength > 0) return tab.data;
      const blank = await createBlankPptx();
      // R254 — mark self-induced so the external re-parse effect skips
      // the bootstrap-bytes change.
      lastWrittenBytesRef.current = blank;
      patchTab(tab.id, { data: blank });
      return blank;
    };
    // R257 — capture this parse's generation so a later R254-driven re-parse
    // for an externally-changed tab.data (AI Apply / undo / redo during our
    // initial parse window) can win the resetHistory race. See parseGenRef
    // doc-block.
    const myGen = ++parseGenRef.current;
    ensureBytes()
      .then(async (bytes) => {
        if (cancelled) return;
        initialBytesRef.current = bytes;
        const m = await parsePptx(bytes);
        if (cancelled) return;
        // R260 — clear loading BEFORE the R257 gen check; see DocxEditor
        // sibling for the full doc-block. If R254's parse won the gen
        // race, our gen-drop here would leave loading=true forever ("正在
        // 解析 pptx…" stuck) because R254's parse-then doesn't touch
        // loading. Loading is "any parse finished" — clear it here
        // regardless of which parse owns the model commit.
        setLoading(false);
        // R257 — newer parse won; drop ours so we don't override R254's
        // resetHistory with our OLD-bytes model.
        if (myGen !== parseGenRef.current) return;
        // Initial parse: clear any prior history (e.g., switching tabs) so
        // undo can't roll back to the previous file.
        // R267 — arm the skip flag before resetHistory; see doc-block on
        // skipNextScheduleRef.
        skipNextScheduleRef.current = true;
        undoApi.resetHistory(m);
        // Restore the slide the user was on before this tab was unmounted.
        // Clamp against the freshly-parsed slide count so AI / undo edits
        // that shrank the deck during the round-trip can't drop us out of
        // bounds; fall through to 0 when there's nothing remembered.
        const remembered = slideMemory.get(tab.id);
        const initialIdx =
          typeof remembered === 'number' && m.slides.length > 0
            ? Math.max(0, Math.min(remembered, m.slides.length - 1))
            : 0;
        setActiveIdx(initialIdx);
        setActiveRunId(null);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        // R272 — match the then arm's setLoading + gen-check structure;
        // see DocxEditor sibling doc-block for the full race trace.
        // Without this gate, a stale [tab.id] parse's throw shows an
        // error banner over R254's already-resolved clean state.
        setLoading(false);
        if (myGen !== parseGenRef.current) return;
        setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
      // Capture the active slide before this mount tears down. Reads from
      // activeIdxRef (kept in sync above) because the cleanup closure
      // captured `activeIdx` at effect-run time — always 0 from the
      // initial useState — and would clobber any actual navigation.
      slideMemory.set(tab.id, activeIdxRef.current);
    };
  }, [tab.id]); // eslint-disable-line react-hooks/exhaustive-deps

  // F5 to enter presentation mode, Escape to leave it. Use capture phase so
  // we beat any other handler, and preventDefault so Chromium doesn't try to
  // reload the page on F5. Must live above any early-return so hook order
  // stays stable between renders (React error #310).
  //
  // Ctrl+PageUp / Ctrl+PageDown navigates between slides during edit — Adobe
  // Acrobat / browser-tab convention for "previous/next page". The Ctrl
  // modifier disambiguates from a focused textarea / contentEditable that
  // otherwise consumes raw PageUp/PageDown to scroll its own content.
  const totalSlides = model?.slides.length ?? 0;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'F5') {
        e.preventDefault();
        e.stopPropagation();
        if (!presenting) setPresenting(true);
      } else if (e.key === 'Escape' && presenting) {
        e.preventDefault();
        e.stopPropagation();
        setPresenting(false);
      } else if (
        !presenting &&
        (e.ctrlKey || e.metaKey) &&
        (e.key === 'PageDown' || e.key === 'PageUp')
      ) {
        if (totalSlides <= 1) return;
        e.preventDefault();
        setActiveIdx((cur) => {
          const next = e.key === 'PageDown' ? cur + 1 : cur - 1;
          return Math.max(0, Math.min(totalSlides - 1, next));
        });
        setActiveRunId(null);
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
    // Re-bind when slide count changes so the bounds stay accurate after
    // add/delete/reorder. `presenting` flips the F5/Escape branch behavior.
  }, [presenting, totalSlides]);

  const scheduleSerialize = (m: PptxModel) => {
    pendingModelRef.current = m;
    // Mark the tab dirty *now* even though the bytes haven't been written
    // back yet. Without this, closing the tab within 400ms of a keystroke
    // skips the unsaved-changes prompt because tab.dirty is still false.
    markTabDirty(tab.id);
    if (serializeTimer.current) clearTimeout(serializeTimer.current);
    serializeTimer.current = setTimeout(() => {
      void flushPendingSerialize();
    }, 400);
  };

  /**
   * Monotonic generation counter — same race guard as DocxEditor's
   * flushGenRef (DocxEditor.tsx:~330 doc-block has the full scenario
   * trace). `serializePptx` calls `JSZip.generateAsync` which is
   * genuinely async (uses worker-style streaming for compression), so
   * a slide-heavy deck can have multi-second serializes; a user typing
   * across two debounce windows hits the same A-finishes-after-B race
   * where stale bytes overwrite latest. Identical fix shape to keep
   * the three editors symmetric.
   */
  const flushGenRef = useRef(0);

  // R254 — track the bytes the editor last wrote so the external-change
  // re-parse effect below recognises self-induced tab.data updates and
  // skips re-parsing. Same shape as R253 (XlsxEditor) and R254
  // (DocxEditor); see effect doc-block at the bottom of this section.
  // Initialize from the FIRST `tab.data` so the re-parse effect skips
  // on initial mount (the existing useEffect[tab.id] above already
  // handles the first parse with full setup — slide-idx restore,
  // ensureBytes blank-bootstrap, loading flag — that R254's minimal
  // re-parse effect doesn't replicate). Editor remounts on tab swap
  // via EditorBoundary's key, so each tab gets a fresh ref.
  // R254 — initialize lazily via the lazy initializer pattern; tab.data
  // for a fresh blank pptx is the empty Uint8Array(0) — that's still a
  // valid Uint8Array and works as the initial sentinel. The
  // ensureBytes() flow inside the parse effect will patchTab a real
  // blank deck and the writeBack-aware lastWrittenBytesRef update at
  // patchTab below keeps the ref in sync with that bootstrap write.
  const lastWrittenBytesRef = useRef<Uint8Array>(tab.data);

  // R257 — monotonic parse-generation counter so that when AI Apply (or
  // undo/redo) changes tab.data during the [tab.id] effect's async parse
  // window, the OLD parse drops itself instead of racing R254's NEW
  // parse for the resetHistory call. parsePptx is async (JSZip-based,
  // 50-500ms for typical decks), so the race is real: user has a
  // PendingChange queued, switches to the pptx tab (mount triggers
  // parse), clicks Apply within the parse window. Both [tab.id]'s parse
  // (with OLD bytes captured at effect-fire) and R254's parse (with
  // NEW bytes from the post-Apply tab.data) run in parallel; whichever
  // .then completes LAST wins. If [tab.id] wins, model = OLD, tab.data
  // = NEW → same silent-revert R253/R254 was supposed to close, just
  // re-opened during the initial-load window. Each parse captures its
  // own gen at start; before resetHistory it checks the gen is still
  // current. Newer parse wins.
  const parseGenRef = useRef(0);

  // R267 — gate the [model] auto-serialize effect (line ~468) against
  // resetHistory-induced model swaps. The auto-serialize was designed
  // to push undo/redo's model change back to tab.data via
  // scheduleSerialize, with `if (loading) return;` intended to skip the
  // initial parse. But [tab.id]'s `.then` commits setLoading(false) AND
  // undoApi.resetHistory(m) in the SAME React-18-batched commit, so
  // [model] sees loading=false already and the gate is open. Net effect:
  // opening a pptx triggers an auto-serialize 400ms after parse →
  // markTabDirty flips dirty=true with no user edit, and AI Apply's
  // AI_BYTES get re-serialized back over themselves. Same shape as the
  // DocxEditor sibling fix; see that doc-block for the full bug trace.
  // Set true immediately before each resetHistory (initial [tab.id]
  // parse + R254 [tab.data] external re-parse + any other bytes→model
  // sync); the [model] effect consumes it once.
  const skipNextScheduleRef = useRef(false);

  /**
   * Run any pending serialize *now* and clear the debounce timer. Used
   * both as the timer body and as the flush callback that the save flow
   * invokes via `flushEditors()` — without this hook a Ctrl+S issued
   * within the 400ms window would snapshot stale `tab.data`.
   */
  const flushPendingSerialize = async (): Promise<void> => {
    if (serializeTimer.current) {
      clearTimeout(serializeTimer.current);
      serializeTimer.current = null;
    }
    const m = pendingModelRef.current;
    if (!m) return;
    pendingModelRef.current = null;
    const myGen = ++flushGenRef.current;
    try {
      const bytes = await serializePptx(m, initialBytesRef.current);
      if (myGen !== flushGenRef.current) return;
      // R254 — record bytes before patchTab so the re-parse effect
      // recognises self-induced.
      lastWrittenBytesRef.current = bytes;
      patchTab(tab.id, { data: bytes });
    } catch (err) {
      // R270 — same gen-check on the catch arm as DocxEditor sibling.
      // A stale flush's serializer throw (e.g., JSZip stream blip on an
      // OLD-model serialize after the user typed more) would otherwise
      // overwrite a newer flush's success with a false error banner.
      // See DocxEditor.flushPendingSerialize R270 doc-block for the full
      // failure trace.
      if (myGen !== flushGenRef.current) return;
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  // R254 — re-parse tab.data on external changes (AI Apply binary_replace,
  // undo / redo). Without this, the local `model` state stays stale after
  // any external mutation: display shows old slides + the next user edit's
  // writeBack serializes the stale model and silently reverts the AI's
  // change. See R253 (XlsxEditor) doc-block for the full data-loss trace —
  // pptx has the same shape. structuralOpInFlightRef-driven structural
  // ops (duplicate / delete / reorder slide at line ~388-415) ALSO patchTab
  // and re-parse manually; we update lastWrittenBytesRef there too so this
  // effect doesn't double-parse on top of the structural op's own setModel.
  useEffect(() => {
    if (tab.data === lastWrittenBytesRef.current) return;
    if (tab.data.byteLength === 0) return; // blank-bootstrap path is handled by useEffect[tab.id]
    let cancelled = false;
    // R257 — bump generation so any in-flight [tab.id] parse (or older
    // R254 parse) drops itself rather than racing this one for resetHistory.
    const myGen = ++parseGenRef.current;
    void parsePptx(tab.data)
      .then((m) => {
        if (cancelled) return;
        if (myGen !== parseGenRef.current) return; // newer parse won
        // R278 — sync the differential-serialize template to the freshly
        // parsed bytes. `initialBytesRef` is what writeBack's serializePptx
        // uses as the JSZip template; the contract is "its slide-zipPath
        // set must be a superset of model.slides' zipPaths". After an
        // external mutation lands new bytes (AI binary_replace, undo/redo
        // of a structural change), the new model's slides may not exist
        // in the OLD template — writeBack's `zip.file(slide.zipPath)` then
        // returns null and the slide silently falls off the output, losing
        // AI's structural changes on the next user edit. runStructuralOp
        // already maintains this invariant (line ~539); [tab.id] does too
        // (line ~233); R254 was the lone path that updated the model
        // without updating its template. XlsxEditor's R277 sibling handles
        // the same shape for SheetJS's `originalBytes` template parameter.
        initialBytesRef.current = tab.data;
        // R267 — same skip-arm as [tab.id]: AI Apply's bytes are
        // authoritative, don't let the [model] auto-serialize re-write them.
        skipNextScheduleRef.current = true;
        undoApi.resetHistory(m);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        // R272 — gen-check; mirrors DocxEditor R272 [tab.data] sibling.
        // cancelled covers same-effect re-runs but not cross-effect
        // [tab.id] vs [tab.data] races.
        if (myGen !== parseGenRef.current) return;
        setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.data]);

  // Register with the global editor-flush registry. The save flow awaits
  // every registered flush before snapshotting tab bytes, so debounced
  // edits never get silently dropped on save.
  useEffect(() => {
    return registerEditorFlush(flushPendingSerialize);
    // flushPendingSerialize is stable enough — it reads model/tab via refs
    // and closures over patchTab (which is itself stable from zustand).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  // On unmount (tab close, workspace swap), force-flush so an in-flight
  // edit isn't lost when the editor goes away. Fire-and-forget — by the
  // time React tears the component down, the user has either confirmed
  // the unsaved-changes prompt (so the data loss is intentional) or this
  // unmount comes from a tab swap where we want to preserve the bytes.
  useEffect(() => {
    return () => {
      if (serializeTimer.current) {
        void flushPendingSerialize();
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Ctrl/Cmd+Z / +Y for editor-local undo/redo. Scoped to PPT root so it
  // doesn't fight the markdown editor's CM6 history or other tabs.
  useUndoShortcuts({
    rootSelector: '[data-pptx-editor-root]',
    undo: () => undoApi.undo(),
    redo: () => undoApi.redo(),
  });

  // After undo/redo (or any model swap) push the model back to bytes via
  // the existing debounced serializer. Skip during initial load — that
  // setModel was driven by a fresh parse, not a user edit.
  // R267 — `loading` alone is insufficient: React-18-batched
  // setLoading(false) + resetHistory's setState reach this effect in the
  // same commit with loading=false already, opening the gate.
  // skipNextScheduleRef is armed before each resetHistory and consumed
  // here so the bytes→model sync paths don't trigger a redundant
  // model→bytes re-serialize. See ref doc-block for the bug shape.
  useEffect(() => {
    if (!model || loading) return;
    if (skipNextScheduleRef.current) {
      skipNextScheduleRef.current = false;
      return;
    }
    scheduleSerialize(model);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [model]);

  /**
   * Run a structural slide op (duplicate / delete / reorder). We must flush
   * any pending text edits first so they make it into the new bytes.
   */
  const runStructuralOp = async (op: (bytes: Uint8Array) => Promise<Uint8Array>, nextActiveIdx: (n: number) => number) => {
    if (!model) return;
    // R95 — drop reentrant calls. See structuralOpInFlightRef declaration
    // above for the race detail.
    if (structuralOpInFlightRef.current) return;
    structuralOpInFlightRef.current = true;
    if (serializeTimer.current) {
      clearTimeout(serializeTimer.current);
      serializeTimer.current = null;
    }
    try {
      const flushed = await serializePptx(model, initialBytesRef.current);
      const next = await op(flushed);
      // R282 — validate `next` via parsePptx BEFORE committing refs /
      // patchTab. The original order committed initialBytesRef +
      // lastWrittenBytesRef + tab.data, then validated; a parsePptx throw
      // (op produced bytes that aren't round-trip-safe — edge slide content,
      // future structural-op regression, version-mismatched undo bytes)
      // left tab.data and the refs at the malformed `next` while `model`
      // stayed at OLD. Subsequent user edits then walked serializeXlsx /
      // serializePptx with a poisoned template, every keystroke surfaced
      // a writeBack error banner, and the only recovery was tab close +
      // reload. Mirrors R277 (XlsxEditor R253 parse-then-commit) and R278
      // (PptxEditor R254 ditto): commit state only after the bytes are
      // proven parseable. Bonus: collapses the "tab.data updated but
      // model still OLD" window — both writes now land in the same React
      // commit tick instead of straddling the parsePptx await.
      const newModel = await parsePptx(next);
      initialBytesRef.current = next;
      // R254 — structural op path also writes tab.data; mark
      // self-induced so the external re-parse effect doesn't fire on
      // top of our own parsePptx + setModel below.
      lastWrittenBytesRef.current = next;
      patchTab(tab.id, { data: next });
      // R309 — arm the [model] auto-serialize skip-flag. setModel(newModel)
      // below is a bytes→model sync: newModel was just parsed FROM `next`
      // and tab.data is already `next`, so the existing R267 invariant
      // ("any other bytes→model sync" per skipNextScheduleRef doc-block
      // at line 393) applies — without this arm the [model] effect fires
      // scheduleSerialize 400ms later and redundantly re-serializes
      // newModel with template=next, producing bytes that are
      // byte-equivalent to `next` but with a fresh Uint8Array ref, then
      // patchTab's it again. Cost: an extra serializePptx (50-500ms CPU
      // for large decks) + a 400ms autoSave-timer push-back via
      // markTabDirty's lastEditAt bump in scheduleSerialize. The skip-arm
      // MUST sit between patchTab and setModel — patchTab affects
      // tab.data (not the [model] effect's dep), setModel triggers the
      // effect. Arming BEFORE the await chain would be unsafe: if
      // serializePptx / op / parsePptx throws into the catch arm below,
      // setModel never fires, the flag stays armed, and the NEXT real
      // user-edit-driven [model] effect would consume the stale arm and
      // skip its own scheduleSerialize — that edit then lives in `model`
      // but never lands in tab.data, save snapshots stale bytes, edits
      // lost. Arming HERE (post-await, pre-setModel) is the only safe
      // window. Same one-shot semantics as R267's [tab.id] and [tab.data]
      // sibling sites.
      skipNextScheduleRef.current = true;
      setModel(newModel);
      setActiveIdx((cur) => Math.max(0, Math.min(nextActiveIdx(cur), newModel.slides.length - 1)));
      setActiveRunId(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      structuralOpInFlightRef.current = false;
    }
  };

  const handleDuplicate = (idx: number) =>
    runStructuralOp(
      (b) => duplicateSlide(b, idx),
      () => idx + 1,
    );

  const handleDelete = (idx: number) => {
    // R151 — confirm carries same Ctrl+Z warning as handleDeleteTextBox
    // (line 510): serializePptx (pptx-adapter.ts:286-301) only edits
    // existing runs, never re-adds dropped slides; sheet delete in
    // XlsxEditor:679 is correctly silent — serializeXlsx round-trips.
    void (async () => {
      // R292 — wrap the confirm-reject path. Same R288/R289/R290 idiom for
      // sibling app.confirm callsites; this destructive operation
      // ("不可 Ctrl+Z 還原") is especially user-unfriendly to fail silently.
      try {
        if (!(await window.gendoc.app.confirm(`刪除第 ${idx + 1} 張投影片？此動作無法用 Ctrl+Z 還原。`))) return;
        await runStructuralOp(
          (b) => deleteSlide(b, idx),
          (cur) => (cur > idx ? cur - 1 : cur === idx ? Math.max(0, idx - 1) : cur),
        );
      } catch (err) {
        notify(`刪除投影片失敗：${(err as Error).message}`, 'error');
      }
    })();
  };

  const handleReorder = (from: number, to: number) =>
    runStructuralOp(
      (b) => reorderSlides(b, from, to),
      () => to,
    );

  const handleAddTextBox = () => {
    if (!model) return;
    void runStructuralOp(
      (b) => addTextBoxToSlide(b, activeIdx),
      () => activeIdx,
    );
  };

  const handleAddShape = (kind: PptxShapeKind) => {
    if (!model) return;
    void runStructuralOp(
      (b) => addShapeToSlide(b, activeIdx, kind),
      () => activeIdx,
    );
  };

  /**
   * Insert a picture from a local file. We open a hidden file picker for
   * jpg/png/gif/svg/webp/bmp, read the bytes, decode the dimensions via
   * an off-screen `Image()` so `addPictureToSlide` can size the shape to
   * the source aspect ratio, and dispatch through `runStructuralOp` so the
   * insertion shares the same flush-then-reparse pipeline as text-boxes
   * and shapes. Errors (read failure, oversized image) surface as a toast
   * — NOT via setError(), which is reserved for fatal parse failures and
   * blanks the entire editor (see the `if (error)` branch in the render).
   * A bad image file should not lose the user access to the rest of the
   * deck.
   */
  const insertPictureFile = async (file: File, anchor?: { x: number; y: number }) => {
    try {
      const arr = new Uint8Array(await file.arrayBuffer());
      const ext = inferImageExt(file.name, file.type);
      const mime = file.type || extToMime(ext);
      const size = await readImageSize(file);
      await runStructuralOp(
        (b) => addPictureToSlide(b, activeIdx, arr, ext, mime, size, anchor),
        () => activeIdx,
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`插入圖片失敗：${msg}`, 'error');
    }
  };

  const handleAddPicture = () => {
    if (!model) return;
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/jpeg,image/png,image/gif,image/svg+xml,image/webp,image/bmp';
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return;
      await insertPictureFile(file);
    };
    input.click();
  };

  const handleApplyLayout = (layoutId: PptxLayoutId) => {
    if (!model) return;
    const label = PPTX_LAYOUTS.find((l) => l.id === layoutId)?.label ?? layoutId;
    // Async path — see handleDelete comment.
    void (async () => {
      // R292 — wrap the confirm-reject path; see handleDelete sibling.
      try {
        if (
          !(await window.gendoc.app.confirm(
            `套用「${label}」會清除目前投影片的文字框內容，確定套用？此動作無法用 Ctrl+Z 還原。`,
          ))
        )
          return;
        await runStructuralOp(
          (b) => applyLayoutToSlide(b, activeIdx, layoutId),
          () => activeIdx,
        );
      } catch (err) {
        notify(`套用版面失敗：${(err as Error).message}`, 'error');
      }
    })();
  };

  const handleDeleteTextBox = (runIdx: number) => {
    // Async path — see handleDelete comment.
    void (async () => {
      // R292 — wrap the confirm-reject path; see handleDelete sibling.
      try {
        if (
          !(await window.gendoc.app.confirm(
            `刪除這個文字框？此動作無法用 Ctrl+Z 還原。`,
          ))
        )
          return;
        await runStructuralOp(
          (b) => deleteTextBoxFromSlide(b, activeIdx, runIdx),
          () => activeIdx,
        );
      } catch (err) {
        notify(`刪除文字框失敗：${(err as Error).message}`, 'error');
      }
    })();
  };

  const handleMoveShape = (shapeIdx: number, x: number, y: number) => {
    if (shapeIdx < 0) return; // table runs / loose runs aren't repositionable
    void runStructuralOp(
      (b) => moveShapeOnSlide(b, activeIdx, shapeIdx, { x, y }),
      () => activeIdx,
    );
  };

  // Adobe PS/AI/ID/Figma convention: Alt-drag duplicates the source shape and
  // lands the copy at the drop position; the original stays put. We can't
  // duplicate runs that don't sit inside a <p:sp> (loose runs / table cells)
  // because there's no shape XML to clone.
  const handleDuplicateShape = (shapeIdx: number, x: number, y: number) => {
    if (shapeIdx < 0) return;
    void runStructuralOp(
      (b) => duplicateShapeOnSlide(b, activeIdx, shapeIdx, { x, y }),
      () => activeIdx,
    );
  };

  /**
   * Align the active shape to the slide along one of six axes — Adobe
   * InDesign / Illustrator / PowerPoint standard alignment palette. Single-
   * shape align uses the slide rectangle as the alignment container, which
   * is the most common case (centering a hero image, snapping a footer to
   * the bottom edge). Multi-shape alignment (align to each other / to
   * selection) requires a multi-select state we don't yet have — queued.
   *
   * Computes the new x/y from the *current* frame so successive clicks are
   * idempotent. Reuses moveShapeOnSlide so the byte-level patch path is the
   * same as drag-move; the user's undo stack stays uniform.
   */
  const handleAlignShape = (
    axis: 'left' | 'hcenter' | 'right' | 'top' | 'vmiddle' | 'bottom',
  ) => {
    if (!model || !activeRun || activeRun.shapeIndex < 0) return;
    const { frame } = activeRun;
    const { cx: slideCx, cy: slideCy } = model.slideSize;
    let nx = frame.x;
    let ny = frame.y;
    switch (axis) {
      case 'left':
        nx = 0;
        break;
      case 'hcenter':
        // Round to integer EMU so successive alignments don't accrete a
        // half-EMU drift from non-integer slide widths.
        nx = Math.round((slideCx - frame.cx) / 2);
        break;
      case 'right':
        nx = slideCx - frame.cx;
        break;
      case 'top':
        ny = 0;
        break;
      case 'vmiddle':
        ny = Math.round((slideCy - frame.cy) / 2);
        break;
      case 'bottom':
        ny = slideCy - frame.cy;
        break;
    }
    void runStructuralOp(
      (b) => moveShapeOnSlide(b, activeIdx, activeRun.shapeIndex, { x: nx, y: ny }),
      () => activeIdx,
    );
  };

  // Resize: same code path as move (the adapter accepts a partial
  // {x, y, cx, cy} patch — we just send all four).
  const handleResizeShape = (shapeIdx: number, x: number, y: number, cx: number, cy: number) => {
    if (shapeIdx < 0) return;
    void runStructuralOp(
      (b) => moveShapeOnSlide(b, activeIdx, shapeIdx, { x, y, cx, cy }),
      () => activeIdx,
    );
  };

  const updateRun = (slideIdx: number, runId: string, patch: Partial<PptxTextRun>) => {
    setModel((prev) => {
      if (!prev) return prev;
      const slides = prev.slides.map((s) =>
        s.index !== slideIdx
          ? s
          : { ...s, runs: s.runs.map((r) => (r.id === runId ? { ...r, ...patch } : r)) },
      );
      const next = { ...prev, slides };
      scheduleSerialize(next);
      return next;
    });
  };

  const updateNotes = (slideIdx: number, notesText: string) => {
    setModel((prev) => {
      if (!prev) return prev;
      const slides = prev.slides.map((s) => (s.index !== slideIdx ? s : { ...s, notesText }));
      const next = { ...prev, slides };
      scheduleSerialize(next);
      return next;
    });
  };

  // ── HOOKS ABOVE EARLY RETURNS ─────────────────────────────────────────
  // React error #310: hook count must stay constant between renders.
  // First render has model=null/loading=true; subsequent renders flip the
  // gate. All hooks below — useFormatShortcuts, useState(findOpen),
  // useEffect for Ctrl+F, useMemo(findSegments) — run on every render and
  // guard their bodies with model nullability checks.

  const active = model?.slides[activeIdx] ?? model?.slides[0] ?? null;
  const activeRun = active?.runs.find((r) => r.id === activeRunId) ?? null;

  // Slide-title outline (Round 76). Heuristic: pick the topmost shape
  // (smallest frame.y) that has any text — title placeholders are almost
  // always positioned at the top of the layout. We don't currently parse
  // <p:ph type="title"> in the adapter; this heuristic is robust enough
  // for the common case of "real" decks (placeholder, blank-with-title,
  // section-header layouts) while degrading gracefully for hand-rolled
  // slides. Empty / image-only slides label as "(無標題)".
  const outline = useMemo(() => {
    if (!model) return [] as Array<{ idx: number; title: string }>;
    return model.slides.map((slide, idx) => {
      const byShape = new Map<number, { y: number; texts: string[] }>();
      for (const r of slide.runs) {
        if (r.shapeIndex < 0) continue;
        const t = r.text.trim();
        if (!t) continue;
        const cur = byShape.get(r.shapeIndex);
        if (cur) cur.texts.push(t);
        else byShape.set(r.shapeIndex, { y: r.frame.y, texts: [t] });
      }
      let bestY = Number.POSITIVE_INFINITY;
      let bestText = '';
      for (const { y, texts } of byShape.values()) {
        if (y < bestY) {
          bestY = y;
          bestText = texts.join(' ').trim();
        }
      }
      return { idx, title: bestText || '(無標題)' };
    });
  }, [model]);

  const toggleStyle = (key: 'bold' | 'italic' | 'underline') => {
    if (!activeRun || !active) return;
    const cur = activeRun.style ?? {};
    const next: PptxRunStyle = { ...cur, [key]: !cur[key] };
    updateRun(active.index, activeRun.id, { style: normalizeStyle(next) });
  };

  // Ctrl/Cmd+B / +I / +U shortcuts. PptxRunStyle now models underline as
  // a boolean (parser collapses any non-`none` <a:rPr u="..."> value;
  // serializer writes back u="sng" when true).
  useFormatShortcuts({
    rootSelector: '[data-pptx-editor-root]',
    isActive: () => !!activeRun,
    toggle: (k) => toggleStyle(k),
  });

  // Ctrl/Cmd+F: open Find & Replace across ALL slides and (re)focus the
  // query input. Composite ids ("slideIdx:runId") let the dialog jump
  // between slides on Locate / Replace without losing the per-run
  // identity. We don't toggle-close on Ctrl+F — Esc closes, matching
  // VS Code / Chrome. `findFocusNonce` re-fires the focus on every
  // Ctrl+F so users who clicked into a slide text box can hit Ctrl+F
  // to come back to the query field.
  const [findOpen, setFindOpen] = useState(false);
  const [findFocusNonce, setFindFocusNonce] = useState(0);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'f') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-pptx-editor-root]')) return;
      e.preventDefault();
      // Mutex with GoTo (Round 79): both dialogs use `absolute top-3 right-3`
      // and would visually overlap if open simultaneously. Ctrl+F coming in
      // while GoTo is open is a clear "switch modes" signal; close the other.
      setGotoOpen(false);
      setFindOpen(true);
      setFindFocusNonce((n) => n + 1);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // Ctrl+G — "Go to slide" (Round 77). Mirrors VS Code's "Go to Line" and
  // PowerPoint's own slide-number shortcut. Same scoping as Ctrl+F so it
  // doesn't fire from other tabs / the markdown preview. focusNonce lets
  // re-pressing Ctrl+G refocus the input even if the dialog is already
  // open. Browsers map Ctrl+G to "Find Next" by default — preventDefault
  // swallows that since we have Enter-in-FindReplace for the same job.
  const [gotoOpen, setGotoOpen] = useState(false);
  const [gotoFocusNonce, setGotoFocusNonce] = useState(0);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'g') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-pptx-editor-root]')) return;
      e.preventDefault();
      // See mutex note in Ctrl+F handler above — close Find when GoTo opens.
      setFindOpen(false);
      setGotoOpen(true);
      setGotoFocusNonce((n) => n + 1);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // Ctrl+V image paste: clipboard image files (e.g. screenshot copies) get
  // dropped onto the active slide as a picture shape. Mirrors the Excel/Word
  // paste flow (Round 71/72). Goes through the same `insertPictureFile`
  // helper as the toolbar button so it respects the flush-then-reparse
  // pipeline. Scoped to the editor root so we don't intercept pastes from
  // other tabs.
  useEffect(() => {
    const root = document.querySelector('[data-pptx-editor-root]');
    if (!root) return;
    const onPaste = (ev: Event) => {
      const e = ev as ClipboardEvent;
      const items = e.clipboardData?.items;
      if (!items) return;
      // R91 — paste-side parallel to the R85 drop-side fix at the onDrop
      // handler near line 628. See MarkdownEditor.tsx (~line 600) for the
      // canonical doc-comment. Pasting a non-image file used to no-op
      // silently, while dragging the same file already warned — this aligns
      // both paths.
      let sawNonImageFile = false;
      for (let i = 0; i < items.length; i += 1) {
        const it = items[i];
        if (it.kind === 'file' && it.type.startsWith('image/')) {
          const file = it.getAsFile();
          if (!file) continue;
          e.preventDefault();
          void insertPictureFile(file);
          return;
        }
        if (it.kind === 'file') sawNonImageFile = true;
      }
      if (sawNonImageFile) {
        e.preventDefault();
        notify('只能貼上圖片檔案', 'warning');
      }
    };
    root.addEventListener('paste', onPaste);
    return () => root.removeEventListener('paste', onPaste);
    // insertPictureFile closes over `activeIdx`/`runStructuralOp`; rebind
    // when activeIdx changes so paste lands on the right slide.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeIdx]);

  const findSegments = useMemo<SearchSegment[]>(() => {
    const out: SearchSegment[] = [];
    if (!model) return out;
    for (const slide of model.slides) {
      slide.runs.forEach((r, i) => {
        out.push({
          id: `${slide.index}:${r.id}`,
          text: r.text,
          label: `Slide ${slide.index + 1} · Box ${i + 1}`,
        });
      });
    }
    return out;
  }, [model]);

  // ── EARLY RETURNS ─────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="h-full w-full flex items-center justify-center text-muted-foreground text-sm">
        正在解析 pptx…
      </div>
    );
  }
  if (error) {
    return (
      <div className="h-full w-full flex items-center justify-center text-destructive text-sm p-8 text-center">
        無法解析 pptx：{error}
      </div>
    );
  }
  if (!model || model.empty || !active) {
    return (
      <div className="h-full w-full flex flex-col items-center justify-center text-muted-foreground gap-3 p-8 text-center">
        <Presentation className="h-12 w-12" />
        <div className="text-sm">這個 pptx 頁籤還是空的。</div>
        <div className="text-xs max-w-md">
          MVP 的 pptx 編輯需要從既有 .pptx 檔開始（會保留 layout / 圖片 / 樣式）。
          請從「檔案 → 開啟」載入已有的 pptx，或先用 Markdown 頁籤編寫內容、之後再匯出。
        </div>
      </div>
    );
  }

  const applyFindReplace = (id: string, text: string): void => {
    const [si, runId] = splitFindId(id);
    if (si === null || !runId) return;
    updateRun(si, runId, { text });
  };

  const locateFindResult = (id: string): void => {
    const [si, runId] = splitFindId(id);
    if (si === null || !runId) return;
    if (si !== activeIdx) setActiveIdx(si);
    setActiveRunId(runId);
    // Scroll the matched run's textarea into view. Without this, a Find
    // result on a long slide (canvas taller than scrollport — common when
    // the deck wrapper carries notes + alignment row + a 16:9 canvas at
    // maxWidth 960) just flips the active-ring class on a textbox that
    // remains below the desk's scroll fold, and Find Next looks like a
    // no-op even though state advanced. rAF defers until React has
    // committed the (possibly new) slide so [data-run-id] exists in the
    // DOM. scrollIntoView walks scrollable ancestors itself, so we don't
    // need to thread a scrollport ref. block: 'nearest' / inline: 'nearest'
    // skip same-slide hits that are already on screen.
    requestAnimationFrame(() => {
      const el = document.querySelector<HTMLElement>(
        `[data-pptx-slide-canvas] [data-run-id="${CSS.escape(runId)}"]`,
      );
      el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    });
  };

  const setColor = (hex: string) => {
    if (!activeRun) return;
    const cleaned = hex.replace(/^#/, '').toUpperCase();
    const cur = activeRun.style ?? {};
    updateRun(active.index, activeRun.id, {
      style: normalizeStyle({ ...cur, color: cleaned || undefined }),
    });
  };

  const clearColor = () => {
    if (!activeRun) return;
    const cur = activeRun.style ?? {};
    updateRun(active.index, activeRun.id, {
      style: normalizeStyle({ ...cur, color: undefined }),
    });
  };

  const setSize = (sizePt: number | undefined) => {
    if (!activeRun) return;
    const cur = activeRun.style ?? {};
    updateRun(active.index, activeRun.id, { style: normalizeStyle({ ...cur, size: sizePt }) });
  };

  const setFontFamily = (name: string | undefined) => {
    if (!activeRun) return;
    const cur = activeRun.style ?? {};
    updateRun(active.index, activeRun.id, {
      style: normalizeStyle({ ...cur, fontFamily: name }),
    });
  };

  // OS file-drag predicate. `dataTransfer.types` is populated during the drag
  // (unlike `.files`, which is empty until drop). "Files" is the canonical
  // type when File Explorer / Finder is the source.
  const isFileDrag = (e: React.DragEvent): boolean => {
    return Array.from(e.dataTransfer.types).includes('Files');
  };

  return (
    <div
      data-pptx-editor-root
      className="h-full w-full flex flex-col bg-background relative"
      onDragEnter={(e) => {
        if (!isFileDrag(e)) return;
        e.preventDefault();
        dragDepthRef.current += 1;
        setDraggingFile(true);
      }}
      onDragOver={(e) => {
        if (!isFileDrag(e)) return;
        // preventDefault makes the editor a valid drop target; without it,
        // `drop` never fires and the OS bounces the file back.
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
        if (!imageFile) {
          // R85 silent-swallow fix — see DocxEditor.tsx:1335 for the full
          // rationale. Previously: dragOver advertised the editor as a valid
          // drop target → overlay flashed → drop with non-image silently
          // dismissed the overlay with no toast. MarkdownEditor.tsx:628
          // carries the same one-liner.
          notify('只能拖入圖片檔案', 'warning');
          return;
        }
        // Drop precision (Round 75): convert the drop point's pixel coords
        // (relative to the slide canvas) to EMU and pass as anchor so the
        // picture lands at the cursor instead of centered. Falls back to
        // centered placement when the canvas isn't queryable (e.g. drop
        // hit the slide rail or the toolbar — only the canvas itself maps
        // to slide-space). The model's `slideSize.cx/cy` gives EMU/canvas
        // ratio: emuPerPx = slideSize.cx / canvasRect.width.
        const canvas = document.querySelector<HTMLElement>('[data-pptx-slide-canvas]');
        let anchor: { x: number; y: number } | undefined;
        if (canvas && model) {
          const rect = canvas.getBoundingClientRect();
          if (
            rect.width > 0 &&
            e.clientX >= rect.left &&
            e.clientX <= rect.right &&
            e.clientY >= rect.top &&
            e.clientY <= rect.bottom
          ) {
            const emuPerPx = model.slideSize.cx / rect.width;
            anchor = {
              x: (e.clientX - rect.left) * emuPerPx,
              y: (e.clientY - rect.top) * emuPerPx,
            };
          }
        }
        void insertPictureFile(imageFile, anchor);
      }}
    >
      <Banner />
      <FormatToolbar
        run={activeRun}
        navOpen={navOpen}
        onToggleNav={() => setNavOpen((v) => !v)}
        onToggleStyle={toggleStyle}
        onSetColor={setColor}
        onClearColor={clearColor}
        onSetSize={setSize}
        onSetFontFamily={setFontFamily}
        onPresent={() => setPresenting(true)}
        onOpenFind={() => {
          setFindOpen(true);
          setFindFocusNonce((n) => n + 1);
        }}
      />
      <div className="flex-1 min-h-0 flex">
        <SlideRail
          tabId={tab.id}
          slideCount={model.slides.length}
          activeIdx={activeIdx}
          titles={outline}
          onSelect={(i) => {
            setActiveIdx(i);
            setActiveRunId(null);
          }}
          onDuplicate={handleDuplicate}
          onDelete={handleDelete}
          onMoveUp={(i) => i > 0 && handleReorder(i, i - 1)}
          onMoveDown={(i) => i < model.slides.length - 1 && handleReorder(i, i + 1)}
          onReorder={(from, to) => {
            if (from !== to) void handleReorder(from, to);
          }}
        />
        {navOpen && (
          <PptxNavPanel
            tabId={tab.id}
            outline={outline}
            activeIdx={activeIdx}
            onJump={(i) => {
              setActiveIdx(i);
              setActiveRunId(null);
            }}
          />
        )}
        {/* Slide canvas backdrop — same rationale as DocxEditor: the
            work-canvas wrapper handles palette CSS vars, but hardcoded zinc
            tokens are var-free, so we keep the desk light in any theme. */}
        <div
          ref={deskScrollRef}
          className="relative flex-1 min-w-0 overflow-auto bg-zinc-200"
        >
          <FindReplaceDialog
            open={findOpen}
            focusNonce={findFocusNonce}
            onClose={() => setFindOpen(false)}
            segments={findSegments}
            onUpdateSegment={applyFindReplace}
            onLocateSegment={locateFindResult}
            title="尋找與取代 · 全部投影片"
          />
          <GoToDialog
            open={gotoOpen}
            focusNonce={gotoFocusNonce}
            onClose={() => setGotoOpen(false)}
            max={model.slides.length}
            label="跳到第幾張投影片？"
            onJump={(oneBased) => {
              setActiveIdx(oneBased - 1);
              setActiveRunId(null);
            }}
          />
          <div className="mx-auto py-6 px-4" style={{ maxWidth: 960 }}>
            <div className="flex items-center justify-between mb-3 gap-2">
              <button
                type="button"
                onClick={() => {
                  // Mirror the Ctrl+G handler — close any conflicting dialog
                  // and bump the focus nonce so re-clicking re-focuses the
                  // input even if the dialog is already open. Same dance as
                  // onOpenFind at line 920-923.
                  setFindOpen(false);
                  setGotoOpen(true);
                  setGotoFocusNonce((n) => n + 1);
                }}
                // "第 N / M 張投影片" — Adobe-style "Page N of M" so the user
                // knows their position in the deck without scanning the
                // SlideRail. Especially helpful in long decks where the rail
                // has scrolled the active thumbnail off-screen.
                //
                // Promoted from <div> to <button> so clicking it opens the
                // GoTo dialog — that's exactly what the Adobe-style indicator
                // does in Acrobat / Reader, which is the UX inspiration this
                // comment cites. Doubles as a Ctrl+G discoverability hook:
                // the indicator already shows current/total, so a tooltip
                // surfacing the shortcut + click affordance turns this passive
                // readout into the keyboard-mouse parity surface that
                // XlsxEditor's name box (line 2201) already provides for
                // Ctrl+G. Without this the Ctrl+G binding (line 686-708)
                // is invisible to mouse users — there's no toolbar button or
                // any other UI hint pointing at it. Hover-only background tint
                // keeps the indicator looking like a readout in steady state
                // so it doesn't visually compete with the action buttons next
                // to it (LayoutPicker / ShapePicker / 插入圖片 / 新增文字框).
                //
                // Ctrl+PgUp / Ctrl+PgDn (bound at line 264-276) sits naturally
                // alongside Ctrl+G as the second slide-navigation shortcut —
                // same family (move-to-slide), same UX lineage cited in that
                // binding's comment ("Adobe Acrobat / browser-tab convention").
                // It was previously the only slide-nav binding with zero in-
                // app surfacing — the readout button is already the established
                // home for this category, so extending its tooltip closes the
                // discoverability gap without spawning a new UI element. Same
                // pattern as R47 (GoToDialog Enter on its primary button) and
                // R54 (outline pane ↑/↓/Home/End on its header).
                title="跳至投影片… (Ctrl+G) · 切換上/下一張 (Ctrl+PgUp / Ctrl+PgDn)"
                className="text-xs text-muted-foreground hover:text-foreground hover:bg-secondary px-1.5 py-0.5 -mx-1.5 -my-0.5 rounded transition-colors"
              >
                第 {active.index + 1} / {model.slides.length} 張投影片 · {active.runs.length} 個文字框
              </button>
              <div className="flex items-center gap-1.5">
                {/* Alignment palette — Adobe AI/ID/PowerPoint standard six-
                    axis align-to-slide. Only renders when the user has a
                    real shape selected (shapeIndex >= 0); table runs / loose
                    runs can't be repositioned so the buttons would no-op. */}
                {activeRun && activeRun.shapeIndex >= 0 && (
                  <AlignmentPalette onAlign={handleAlignShape} />
                )}
                <LayoutPicker onApply={handleApplyLayout} />
                <ShapePicker onAdd={handleAddShape} />
                <button
                  type="button"
                  onClick={handleAddPicture}
                  className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded border border-border hover:bg-secondary text-muted-foreground hover:text-foreground transition-colors"
                  // R100 — tooltip mirrors the picker's `accept` list at
                  // PptxEditor.tsx:478 (BMP added in that round so the
                  // format list matches the file dialog exactly).
                  // R107 — also disclose WHERE the image lands. Two
                  // sibling editors already do this at their own image-
                  // insert toolbar buttons:
                  //   • XlsxEditor.tsx:2054-2058 (R99) flips between
                  //     「插入圖片（錨定於 ${selectionAddr}）」 and
                  //     「插入圖片（未選取儲存格時錨定於 A1）」
                  //   • DocxEditor.tsx:2152-2199 (R102) flips between
                  //     「在此段後插入圖片(...)」 and
                  //     「於文件結尾插入圖片(...)」
                  // Pptx was the lone holdout still naming only the
                  // formats — leaving a click-and-find-out gap for a
                  // user reading the toolbar to confirm "where will
                  // this image land?" before triggering a file picker.
                  // Same-file precedent makes the omission especially
                  // pointed: the drop-target overlay 16 lines below at
                  // PptxEditor.tsx:1147 already discloses「放開即插入
                  // 圖片到目前投影片」 — drop and click are the two
                  // entry points for the same `insertPictureFile`
                  // helper (PptxEditor.tsx:458-472), so they should
                  // describe the destination in matching terms. The
                  // overlay uses「到目前投影片」 verbatim; reused
                  // verbatim here so both surfaces speak with one
                  // voice. Pptx has no analogue to Word's per-paragraph
                  // anchor or Excel's selection cell — `handleAddPicture`
                  // (PptxEditor.tsx:474) always inserts into `activeIdx`
                  // — so a single static string is correct (no need for
                  // the 2-way state-aware shape that Doc/Excel needed).
                  // Format list stays in trailing parens to preserve
                  // R100's "additions append in one place" property.
                  title="插入圖片到目前投影片(PNG / JPG / GIF / SVG / WebP / BMP)"
                >
                  <ImageIcon className="h-3 w-3" />
                  插入圖片
                </button>
                {/* R112 — lone same-row holdout after R107/R108/R109
                    polished the other three buttons. Two converging gaps
                    that R107-R109 already established as fix-worthy on
                    siblings sitting just to the left:

                    1. Tooltip = visible label (zero info on hover). The
                       button shows `<Plus> 新增文字框` and the tooltip
                       repeated「新增文字框」 verbatim — exact same shape
                       as ShapePicker pre-R108 ('新增圖形' visible label
                       + 'title="新增圖形"' tooltip). Hover gave the user
                       nothing the eyes hadn't already read.

                    2. Destination undisclosed. Same row carries three
                       buttons that explicitly disclose where the action
                       lands:
                         • LayoutPicker (R109, line 2486) —
                           「套用版面配置到目前投影片」
                         • ShapePicker (R108, line 2417) —
                           「新增圖形到目前投影片」
                         • 插入圖片 (R107, line 1115) —
                           「插入圖片到目前投影片(...)」
                       handleAddTextBox at line 430-436 likewise calls
                       `addTextBoxToSlide(b, activeIdx)` — destination is
                       `activeIdx` exactly like the three siblings — so
                       leaving this one silent splits a same-row 3-vs-1.

                    addTextBoxToSlide at pptx-adapter.ts:307-340 also
                    inserts a placeholder run with default text「新文字框」
                    (parameter default `text = '新文字框'` at line 310,
                    written into the `<a:t>` at line 335). Disclosing the
                    placeholder lets the user predict what'll appear on
                    the canvas before they click — same parens-as-
                    auxiliary-info shape R107 used to surface the picker's
                    format list. Wording mirrors R107's compact non-
                    dropdown structure: 「{verb}{dest}（{aux info}）」. */}
                <button
                  type="button"
                  onClick={handleAddTextBox}
                  className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded border border-border hover:bg-secondary text-muted-foreground hover:text-foreground transition-colors"
                  title="新增文字框到目前投影片（預設文字：新文字框）"
                >
                  <Plus className="h-3 w-3" />
                  新增文字框
                </button>
              </div>
            </div>
            <SlideCanvas
              slideSize={model.slideSize}
              runs={active.runs}
              pictures={active.pictures}
              activeRunId={activeRunId}
              onRunFocus={(id) => setActiveRunId(id)}
              onRunChange={(id, text) => updateRun(active.index, id, { text })}
              onRunDelete={(runIdx) => handleDeleteTextBox(runIdx)}
              onShapeMove={handleMoveShape}
              onShapeResize={handleResizeShape}
              onShapeDuplicate={handleDuplicateShape}
            />
            <NotesEditor
              key={`notes-${active.index}`}
              value={active.notesText ?? ''}
              onChange={(t) => updateNotes(active.index, t)}
            />
          </div>
        </div>
      </div>
      {presenting ? (
        <PresentationMode
          slides={model.slides}
          startIdx={activeIdx}
          onClose={() => setPresenting(false)}
        />
      ) : null}
      {draggingFile && (
        // Drop indicator. `pointer-events-none` is critical — without it, the
        // overlay would intercept dragenter/leave from descendants and the
        // depth counter would never settle.
        <div className="pointer-events-none absolute inset-0 bg-primary/5 border-2 border-dashed border-primary/40 z-50 flex items-center justify-center">
          <div className="px-3 py-1.5 rounded bg-background/90 text-xs text-primary border border-primary/40 shadow">
            放開即插入圖片到目前投影片
          </div>
        </div>
      )}
    </div>
  );
}

function Banner(): JSX.Element {
  return (
    <div className="flex items-center gap-2 px-3 py-1.5 text-xs bg-amber-100 border-b border-amber-300 text-amber-800">
      <AlertTriangle className="h-3.5 w-3.5 shrink-0" />
      <span>
        PowerPoint MVP 編輯：可改文字框 + 粗體 / 斜體 / 字色 / 字級；左側可新增 / 複製 / 刪除 / 上下移動投影片。新增 pptx 會自動帶一張預設投影片。
      </span>
    </div>
  );
}

/** Common pptx font sizes in points. */
const FONT_SIZES_PT = [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 66, 80];

function FormatToolbar({
  run,
  navOpen,
  onToggleNav,
  onToggleStyle,
  onSetColor,
  onClearColor,
  onSetSize,
  onSetFontFamily,
  onPresent,
  onOpenFind,
}: {
  run: PptxTextRun | null;
  navOpen: boolean;
  onToggleNav: () => void;
  onToggleStyle: (k: 'bold' | 'italic' | 'underline') => void;
  onSetColor: (hex: string) => void;
  onClearColor: () => void;
  onSetSize: (pt: number | undefined) => void;
  onSetFontFamily: (name: string | undefined) => void;
  onPresent: () => void;
  onOpenFind: () => void;
}): JSX.Element {
  const disabled = !run;
  const style = run?.style ?? {};
  return (
    <div className={cn('flex items-center gap-0.5 px-2 py-1 border-b bg-secondary/30')}>
      {/* Outline toggle — first item, mirrors DocxEditor. NOT disabled when
          no run is selected: navigating between slides shouldn't require a
          text-box selection. The wrapping <div>'s opacity-50 was eating the
          button's affordance, so we drop the global opacity gate and let
          each control manage its own disabled visual. */}
      <ToolbarBtn
        active={navOpen}
        title={navOpen ? '隱藏投影片大綱' : '顯示投影片大綱'}
        onClick={onToggleNav}
      >
        <List className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <div className={cn('flex items-center gap-0.5', disabled && 'opacity-50')}>
      {/* Bold/italic/underline tooltips spell out the keyboard shortcuts the
          editor binds via useFormatShortcuts (see line ~549). Without this
          the bindings exist but are invisible to mouse-first users. Mirrors
          DocxEditor / MarkdownToolbar tooltip style for cross-editor parity.

          R92 — when `disabled` (no run selected, line 1156), flip the title
          to explain *why* the button is dimmed. Wording '請先點選一個文字框'
          mirrors the existing status hint at line ~1270
          ('點選一個文字框以開始編輯'), keeping the disabled-state vocabulary
          consistent within this toolbar. Pairs with R87 (DocxEditor) and the
          XlsxEditor R92 sibling. */}
      <ToolbarBtn active={!!style.bold} disabled={disabled} title={disabled ? '請先點選一個文字框' : '粗體 (Ctrl+B)'} onClick={() => onToggleStyle('bold')}>
        <Bold className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.italic} disabled={disabled} title={disabled ? '請先點選一個文字框' : '斜體 (Ctrl+I)'} onClick={() => onToggleStyle('italic')}>
        <Italic className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.underline} disabled={disabled} title={disabled ? '請先點選一個文字框' : '底線 (Ctrl+U)'} onClick={() => onToggleStyle('underline')}>
        <Underline className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      {/* R97 — extend the R92 disabled-state-tooltip flip to the rest of the
          toolbar. R92 only covered Bold/Italic/Underline (lines 1206-1213);
          the colour picker, clear-colour ×, font family select and font
          size select right next to them all share the same `disabled` gate
          but kept their tooltips static. Result: hovering any of these
          while no text-box is selected returned the action label («文字
          顏色» / «字型» / etc.) on a greyed-out control, breaking the «所
          有 disabled 控制都解釋為什麼被 disable» rule that R92 set for this
          same toolbar. Wording reuses «請先點選一個文字框» verbatim from
          R92 (lines 1206-1213) and from the toolbar's own status hint at
          line 1298 («點選一個文字框以開始編輯»), keeping the disabled-
          state vocabulary identical across the whole toolbar. */}
      <span className="relative inline-flex items-center">
        <label
          title={disabled ? '請先點選一個文字框' : '文字顏色'}
          // Preserve current run selection — the label was eating mousedown
          // and clearing the active run before the picker change fired.
          onMouseDown={(e) => e.preventDefault()}
          className={cn(
            'h-7 w-7 inline-flex items-center justify-center rounded transition-colors cursor-pointer',
            'text-muted-foreground hover:text-foreground hover:bg-secondary',
            disabled && 'cursor-not-allowed opacity-50 pointer-events-none',
          )}
        >
          <span className="relative">
            <Palette className="h-3.5 w-3.5" />
            <span
              className="absolute -bottom-0.5 left-0 right-0 h-0.5 rounded"
              style={{ background: style.color ? `#${style.color}` : 'currentColor' }}
            />
          </span>
          <input
            type="color"
            value={style.color ? `#${style.color}` : '#000000'}
            onChange={(e) => onSetColor(e.target.value)}
            disabled={disabled}
            className="absolute inset-0 opacity-0 cursor-pointer"
          />
        </label>
        {style.color ? (
          <button
            type="button"
            title={disabled ? '請先點選一個文字框' : '清除文字顏色'}
            onMouseDown={(e) => e.preventDefault()}
            onClick={onClearColor}
            disabled={disabled}
            className="ml-px text-muted-foreground hover:text-destructive text-[10px] leading-none px-0.5"
          >
            ×
          </button>
        ) : null}
      </span>
      <Divider />
      <select
        disabled={disabled}
        value={style.fontFamily ?? ''}
        onChange={(e) => {
          const v = e.target.value;
          onSetFontFamily(v === '' ? undefined : v);
        }}
        className={cn(
          'h-7 text-xs rounded border border-border bg-background px-1.5 ml-1 max-w-[170px]',
          disabled && 'cursor-not-allowed',
        )}
        title={disabled ? '請先點選一個文字框' : '字型'}
        style={style.fontFamily ? { fontFamily: style.fontFamily } : undefined}
      >
        <option value="">預設字型</option>
        {FONT_FAMILIES.map((f) => (
          <option key={f.value} value={f.value} style={{ fontFamily: f.value }}>
            {f.label}
          </option>
        ))}
      </select>
      <Type className="h-3.5 w-3.5 text-muted-foreground ml-1" />
      <select
        disabled={disabled}
        value={style.size ?? ''}
        onChange={(e) => {
          const v = e.target.value;
          onSetSize(v === '' ? undefined : Number(v));
        }}
        className={cn(
          'h-7 text-xs rounded border border-border bg-background px-1.5',
          disabled && 'cursor-not-allowed',
        )}
        title={disabled ? '請先點選一個文字框' : '字級 (pt)'}
      >
        <option value="">預設</option>
        {FONT_SIZES_PT.map((s) => (
          <option key={s} value={s}>{s} pt</option>
        ))}
      </select>
      <span className="ml-2 text-[11px] text-muted-foreground whitespace-nowrap">
        {run ? `已選：第 ${run.runIndex + 1} 個文字框` : '點選一個文字框以開始編輯'}
      </span>
      </div>
      <div className="ml-auto" />
      {/* Find & Replace — Ctrl+F is bound (line ~570) but the keyboard
          shortcut is invisible to mouse users. Surface it in the toolbar
          for parity with DocxEditor / MarkdownEditor. Stays enabled even
          without a focused run: searching across all slides is meaningful
          regardless of which text box (if any) currently has the caret. */}
      <ToolbarBtn disabled={false} title="尋找與取代 (Ctrl+F)" onClick={onOpenFind}>
        <Search className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <button
        type="button"
        onClick={onPresent}
        title="開始放映 (F5)"
        className="h-7 inline-flex items-center gap-1 px-2 rounded text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
      >
        <Play className="h-3.5 w-3.5" />
        放映 (F5)
      </button>
    </div>
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
      // R153 — toggle-state SR exposure. Same treatment as the sibling
      // ToolbarBtn definitions in MarkdownToolbar / DocxEditor / XlsxEditor
      // (this round). Format buttons in this editor's toolbar (粗體 / 斜體 /
      // 底線 line 1273-1280) all pass `active={!!style.X}`; the visual
      // 「bg-primary/20 text-primary」 active-highlight was a sighted-only
      // signal of「currently bold / italic / underlined」 in the focused text
      // frame. React renders `aria-pressed={undefined}` as no attribute, so
      // action-only ToolbarBtn callsites in this file keep clean action-
      // button semantics — only the format toggles (and any future toggles)
      // pick up the proper toggle role.
      aria-pressed={active}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={cn(
        'h-7 w-7 inline-flex items-center justify-center rounded transition-colors',
        active ? 'bg-primary/20 text-primary' : 'text-muted-foreground hover:text-foreground hover:bg-secondary',
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

/** Parse a F&R composite id of form "slideIndex:runId" — runIds may
 * themselves contain hyphens or other punctuation, so we split on the
 * FIRST colon only. Returns [null, ''] for malformed inputs. */
function splitFindId(id: string): [number | null, string] {
  const idx = id.indexOf(':');
  if (idx <= 0) return [null, ''];
  const si = Number(id.slice(0, idx));
  if (!Number.isFinite(si)) return [null, ''];
  return [si, id.slice(idx + 1)];
}

/**
 * Pick the OOXML media file extension to use for a picked image. Prefer the
 * file's name extension (canonical for SVG, where MIME may be application/*
 * or empty on some platforms); fall back to MIME, then default to png so
 * the upload never fails on an unrecognized type.
 */
function inferImageExt(filename: string, mime: string): string {
  const fromName = filename.match(/\.([a-z0-9]+)$/i)?.[1]?.toLowerCase();
  if (fromName && /^(png|jpg|jpeg|gif|svg|webp|bmp)$/.test(fromName)) return fromName;
  if (mime === 'image/jpeg') return 'jpg';
  if (mime === 'image/png') return 'png';
  if (mime === 'image/gif') return 'gif';
  if (mime === 'image/svg+xml') return 'svg';
  if (mime === 'image/webp') return 'webp';
  if (mime === 'image/bmp') return 'bmp';
  return 'png';
}

function extToMime(ext: string): string {
  switch (ext) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'svg':
      return 'image/svg+xml';
    case 'webp':
      return 'image/webp';
    case 'bmp':
      return 'image/bmp';
    default:
      return 'image/png';
  }
}

/**
 * Read an image's natural dimensions via an off-screen Image() loader. We
 * use these to set the shape's aspect ratio in the slide. Falls back to
 * a square 1×1 on any decode failure (rather than throwing) so the upload
 * still succeeds — the user can resize manually after.
 */
function readImageSize(file: File): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth || 1;
      const h = img.naturalHeight || 1;
      URL.revokeObjectURL(url);
      resolve({ width: w, height: h });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({ width: 1, height: 1 });
    };
    img.src = url;
  });
}

function normalizeStyle(s: PptxRunStyle): PptxRunStyle | undefined {
  const out: PptxRunStyle = {};
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  if (s.color) out.color = s.color;
  if (s.size) out.size = s.size;
  if (s.fontFamily) out.fontFamily = s.fontFamily;
  return Object.keys(out).length === 0 ? undefined : out;
}

/**
 * 16:9 canvas — renders the slide as a scaled-down rectangle and absolutely
 * positions each `<p:sp>` frame at its actual coordinates so users see the
 * spatial layout instead of a flat list. Runs are grouped by `shapeIndex`
 * (multiple `<a:r>` inside the same `<p:sp>` stack vertically inside one
 * frame). Font sizes scale linearly with canvas width via a measured
 * `pxPerEmu` ratio so 24-pt text on the slide looks like ~24-pt-equivalent
 * on the canvas regardless of the canvas's rendered size.
 */
/**
 * Snap-guide marker emitted by ShapeFrame during a live drag/resize and
 * rendered by SlideCanvas as a dotted alignment line. `pos` is in EMU on
 * the relevant axis: x-axis guides are vertical lines at column `pos`,
 * y-axis guides are horizontal lines at row `pos`.
 */
interface SnapGuide {
  axis: 'x' | 'y';
  pos: number;
}

function SlideCanvas({
  slideSize,
  runs,
  pictures,
  activeRunId,
  onRunFocus,
  onRunChange,
  onRunDelete,
  onShapeMove,
  onShapeResize,
  onShapeDuplicate,
}: {
  slideSize: { cx: number; cy: number };
  runs: PptxTextRun[];
  pictures?: Record<number, string>;
  activeRunId: string | null;
  onRunFocus: (id: string) => void;
  onRunChange: (id: string, text: string) => void;
  onRunDelete: (runIndex: number) => void;
  onShapeMove: (shapeIndex: number, x: number, y: number) => void;
  onShapeResize: (shapeIndex: number, x: number, y: number, cx: number, cy: number) => void;
  onShapeDuplicate: (shapeIndex: number, x: number, y: number) => void;
}): JSX.Element {
  const canvasRef = useRef<HTMLDivElement>(null);
  const [canvasWidth, setCanvasWidth] = useState(0);
  useEffect(() => {
    const el = canvasRef.current;
    if (!el) return;
    setCanvasWidth(el.getBoundingClientRect().width);
    const ro = new ResizeObserver((entries) => {
      const w = entries[0]?.contentRect.width ?? 0;
      setCanvasWidth(w);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  const pxPerEmu = canvasWidth > 0 && slideSize.cx > 0 ? canvasWidth / slideSize.cx : 0;

  // Active snap guides — one entry per matched axis target. ShapeFrame
  // pushes/clears these via onGuidesChange while the user drags/resizes.
  const [guides, setGuides] = useState<SnapGuide[]>([]);

  // Group runs by shapeIndex while preserving slide-XML order. Runs that
  // didn't sit inside a <p:sp> (shapeIndex === -1) each get their own
  // singleton group so they still surface in the canvas.
  const groups: { key: string; frame: PptxFrame; runs: PptxTextRun[] }[] = [];
  for (const run of runs) {
    if (run.shapeIndex >= 0) {
      const last = groups[groups.length - 1];
      if (last && last.key === `s${run.shapeIndex}`) {
        last.runs.push(run);
        continue;
      }
      groups.push({ key: `s${run.shapeIndex}`, frame: run.frame, runs: [run] });
    } else {
      groups.push({ key: `o${run.id}`, frame: run.frame, runs: [run] });
    }
  }

  return (
    <div
      ref={canvasRef}
      data-pptx-slide-canvas
      className="relative w-full bg-white border border-zinc-300 shadow-lg rounded-md overflow-hidden"
      style={{ aspectRatio: `${slideSize.cx} / ${slideSize.cy}` }}
    >
      {groups.length === 0 ? (
        <div className="absolute inset-0 flex items-center justify-center text-sm text-muted-foreground italic px-4 text-center">
          這張投影片沒有可編輯的文字。可使用上方「新增文字框」加入一個。
        </div>
      ) : (
        groups.map((group, groupIdx) => {
          const shapeIdx = group.runs[0]?.shapeIndex ?? -1;
          // Pass every other group's frame as a snap candidate. We exclude
          // the active group itself so the shape can't snap to its own
          // edges (which would freeze it in place).
          const siblingFrames = groups
            .filter((_, i) => i !== groupIdx)
            .map((g) => g.frame);
          const pictureUrl = shapeIdx >= 0 ? pictures?.[shapeIdx] : undefined;
          return (
            <ShapeFrame
              key={group.key}
              frame={group.frame}
              slideSize={slideSize}
              pxPerEmu={pxPerEmu}
              runs={group.runs}
              pictureUrl={pictureUrl}
              activeRunId={activeRunId}
              onRunFocus={onRunFocus}
              onRunChange={onRunChange}
              onRunDelete={onRunDelete}
              onMove={(x, y) => onShapeMove(shapeIdx, x, y)}
              onResize={(x, y, cx, cy) => onShapeResize(shapeIdx, x, y, cx, cy)}
              onDuplicate={(x, y) => onShapeDuplicate(shapeIdx, x, y)}
              canMove={shapeIdx >= 0}
              siblingFrames={siblingFrames}
              onGuidesChange={setGuides}
            />
          );
        })
      )}
      {/* Render snap guides as dotted lines spanning the canvas. Each guide
          is positioned by % so it tracks zoom/aspect changes. pointer-events
          off so it never intercepts clicks meant for the shape underneath. */}
      {guides.map((g, i) => (
        <div
          key={i}
          className="absolute bg-fuchsia-500/80 pointer-events-none"
          style={
            g.axis === 'x'
              ? {
                  left: `${(g.pos / slideSize.cx) * 100}%`,
                  top: 0,
                  bottom: 0,
                  width: 1,
                }
              : {
                  top: `${(g.pos / slideSize.cy) * 100}%`,
                  left: 0,
                  right: 0,
                  height: 1,
                }
          }
        />
      ))}
    </div>
  );
}

function ShapeFrame({
  frame,
  slideSize,
  pxPerEmu,
  runs,
  pictureUrl,
  activeRunId,
  onRunFocus,
  onRunChange,
  onRunDelete,
  onMove,
  onResize,
  onDuplicate,
  canMove,
  siblingFrames,
  onGuidesChange,
}: {
  frame: PptxFrame;
  slideSize: { cx: number; cy: number };
  pxPerEmu: number;
  runs: PptxTextRun[];
  pictureUrl?: string;
  activeRunId: string | null;
  onRunFocus: (id: string) => void;
  onRunChange: (id: string, text: string) => void;
  onRunDelete: (runIndex: number) => void;
  onMove: (x: number, y: number) => void;
  onResize: (x: number, y: number, cx: number, cy: number) => void;
  onDuplicate: (x: number, y: number) => void;
  canMove: boolean;
  siblingFrames: PptxFrame[];
  onGuidesChange: (guides: SnapGuide[]) => void;
}): JSX.Element {
  // Pre-compute the union of snap targets per axis: every sibling shape's
  // left/center/right (X) or top/middle/bottom (Y), plus the slide's own
  // edges and center. This list is stable across a drag (we recompute when
  // siblingFrames or slideSize change) so the inner pointermove handler
  // doesn't have to rebuild it every frame.
  const xTargets = useMemo(() => {
    const t = new Set<number>([0, slideSize.cx / 2, slideSize.cx]);
    for (const f of siblingFrames) {
      t.add(f.x);
      t.add(f.x + f.cx / 2);
      t.add(f.x + f.cx);
    }
    return Array.from(t);
  }, [siblingFrames, slideSize.cx]);
  const yTargets = useMemo(() => {
    const t = new Set<number>([0, slideSize.cy / 2, slideSize.cy]);
    for (const f of siblingFrames) {
      t.add(f.y);
      t.add(f.y + f.cy / 2);
      t.add(f.y + f.cy);
    }
    return Array.from(t);
  }, [siblingFrames, slideSize.cy]);
  // Live drag state: a non-null `dragOffset` means we're mid-drag and the
  // textarea drag-overlay should render the frame at frame.x+dx / frame.y+dy
  // *visually* without committing to the model. Commit happens on pointerup.
  const [dragOffset, setDragOffset] = useState<{ dx: number; dy: number } | null>(null);
  // Tracked separately so the live readout can show a "+ 複製" badge while
  // the user holds Alt mid-drag (Adobe-style duplicate-mode feedback).
  const [altDuplicating, setAltDuplicating] = useState(false);
  // Resize state: signed deltas applied to the frame's x / y / cx / cy
  // while the user is dragging a corner / edge handle. Cleared on commit.
  const [resizeOffset, setResizeOffset] = useState<{
    dx: number;
    dy: number;
    dcx: number;
    dcy: number;
  } | null>(null);

  // EMU coords → percentage of the canvas. During a drag we apply the offset
  // (also in EMU) so the frame visibly tracks the cursor.
  const liveX = frame.x + (dragOffset?.dx ?? 0) + (resizeOffset?.dx ?? 0);
  const liveY = frame.y + (dragOffset?.dy ?? 0) + (resizeOffset?.dy ?? 0);
  const liveCx = frame.cx + (resizeOffset?.dcx ?? 0);
  const liveCy = frame.cy + (resizeOffset?.dcy ?? 0);
  const left = (liveX / slideSize.cx) * 100;
  const top = (liveY / slideSize.cy) * 100;
  const width = (liveCx / slideSize.cx) * 100;
  const height = (liveCy / slideSize.cy) * 100;
  const anyActive = runs.some((r) => r.id === activeRunId);

  /**
   * Pointer-driven drag from the move-handle. We attach window-level
   * listeners on pointerdown so the drag continues even if the cursor
   * leaves the frame, and convert pixel deltas back to EMU using
   * `pxPerEmu` (1 EMU == 1/pxPerEmu px). Commits on pointerup via
   * `onMove(absX, absY)` then clears the offset; the parent re-parses the
   * pptx and the new frame.x/y come back through props.
   */
  /**
   * Drag-snap helper. Given the proposed delta (in EMU), look at the active
   * shape's three candidate edges on each axis (left/center/right and
   * top/middle/bottom) and find the closest snap target within
   * `thresholdEmu`. If one matches, return the adjusted delta plus a guide
   * marker for visual feedback. The first matching edge wins per axis so
   * one drag never produces two competing guides on the same axis.
   */
  const computeDragSnap = (
    dxEmu: number,
    dyEmu: number,
    thresholdEmu: number,
  ): { dx: number; dy: number; guides: SnapGuide[] } => {
    const guides: SnapGuide[] = [];
    let dx = dxEmu;
    let dy = dyEmu;
    const liveLeft = frame.x + dxEmu;
    const liveRight = liveLeft + frame.cx;
    const liveCenter = liveLeft + frame.cx / 2;
    for (const [edge, name] of [
      [liveLeft, 'left'],
      [liveCenter, 'center'],
      [liveRight, 'right'],
    ] as const) {
      let best: { target: number; diff: number } | null = null;
      for (const t of xTargets) {
        const diff = t - edge;
        if (Math.abs(diff) <= thresholdEmu && (!best || Math.abs(diff) < Math.abs(best.diff))) {
          best = { target: t, diff };
        }
      }
      if (best) {
        dx = dxEmu + best.diff;
        guides.push({ axis: 'x', pos: best.target });
        void name;
        break;
      }
    }
    const liveTop = frame.y + dyEmu;
    const liveBottom = liveTop + frame.cy;
    const liveMiddle = liveTop + frame.cy / 2;
    for (const edge of [liveTop, liveMiddle, liveBottom]) {
      let best: { target: number; diff: number } | null = null;
      for (const t of yTargets) {
        const diff = t - edge;
        if (Math.abs(diff) <= thresholdEmu && (!best || Math.abs(diff) < Math.abs(best.diff))) {
          best = { target: t, diff };
        }
      }
      if (best) {
        dy = dyEmu + best.diff;
        guides.push({ axis: 'y', pos: best.target });
        break;
      }
    }
    return { dx, dy, guides };
  };

  const startDrag = (e: React.PointerEvent) => {
    if (!canMove || pxPerEmu <= 0) return;
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startY = e.clientY;
    // 6 px screen-space → EMU. Roughly one pixel of "stickiness" on either
    // side at 1× zoom; tightens at higher zoom because pxPerEmu grows.
    const SNAP_PX = 6;
    const thresholdEmu = SNAP_PX / pxPerEmu;
    /**
     * Adobe PS/AI/ID/Figma/Keynote/PowerPoint shared convention: holding
     * Shift while dragging locks motion to the dominant axis (whichever
     * displacement is larger in absolute terms). Used for "move this
     * exactly horizontally" / "move this exactly vertically" intents that
     * would otherwise need post-edit cleanup. We bypass snap when the
     * user is constraining — explicit lock-to-axis intent should beat
     * convenience snapping that might pull off-axis. Re-evaluate Shift
     * on every pointermove so the user can press/release mid-drag and
     * see the constraint engage/disengage live.
     */
    const constrainOrSnap = (rawDx: number, rawDy: number, shift: boolean) => {
      if (shift) {
        return Math.abs(rawDx) >= Math.abs(rawDy)
          ? { dx: rawDx, dy: 0, guides: [] as SnapGuide[] }
          : { dx: 0, dy: rawDy, guides: [] as SnapGuide[] };
      }
      return computeDragSnap(rawDx, rawDy, thresholdEmu);
    };
    const onMoveDoc = (ev: PointerEvent) => {
      const dxPx = ev.clientX - startX;
      const dyPx = ev.clientY - startY;
      const rawDx = Math.round(dxPx / pxPerEmu);
      const rawDy = Math.round(dyPx / pxPerEmu);
      const result = constrainOrSnap(rawDx, rawDy, ev.shiftKey);
      setDragOffset({ dx: result.dx, dy: result.dy });
      setAltDuplicating(ev.altKey);
      onGuidesChange(result.guides);
    };
    const onUpDoc = (ev: PointerEvent) => {
      const dxPx = ev.clientX - startX;
      const dyPx = ev.clientY - startY;
      const rawDx = Math.round(dxPx / pxPerEmu);
      const rawDy = Math.round(dyPx / pxPerEmu);
      window.removeEventListener('pointermove', onMoveDoc);
      window.removeEventListener('pointerup', onUpDoc);
      setDragOffset(null);
      setAltDuplicating(false);
      onGuidesChange([]);
      // Skip commit if the user barely moved (treat as a misclick).
      if (Math.abs(dxPx) < 3 && Math.abs(dyPx) < 3) return;
      // Read modifier keys fresh from the up event — user may have pressed
      // or released between the last move and the release.
      const result = constrainOrSnap(rawDx, rawDy, ev.shiftKey);
      // Clamp into the slide so a wild drag can't push the box off-canvas.
      const nx = Math.max(0, Math.min(slideSize.cx - frame.cx, frame.x + result.dx));
      const ny = Math.max(0, Math.min(slideSize.cy - frame.cy, frame.y + result.dy));
      // Adobe PS/AI/ID/Figma convention: Alt held on release means duplicate
      // — clone the source shape at the drop position and leave the original
      // in place. We dispatch via the duplicate path; otherwise it's a plain
      // move. Shift can compose with Alt (axis-constrained duplicate), which
      // is also Adobe-standard, because the snap/constrain step runs first.
      if (ev.altKey) onDuplicate(nx, ny);
      else onMove(nx, ny);
    };
    window.addEventListener('pointermove', onMoveDoc);
    window.addEventListener('pointerup', onUpDoc);
  };

  /**
   * Resize from one of 8 handles. `handle` encodes which edges are anchored
   * vs. moving — e.g. 'nw' moves both x and y while shrinking cx/cy by the
   * same amount; 'e' only changes cx. Live deltas live in `resizeOffset`
   * and commit through `onResize` on pointerup.
   *
   * Minimum size guard: keep cx/cy >= MIN_EMU so the shape doesn't collapse
   * to a click target the user can't grab again. Also clamp into the slide
   * so a wild drag can't push the frame past the canvas.
   */
  const startResize = (
    e: React.PointerEvent,
    handle: 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w',
  ) => {
    if (!canMove || pxPerEmu <= 0) return;
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startY = e.clientY;
    // Min ~0.2 inch (914400 EMU = 1 inch).
    const MIN_EMU = 200000;
    const SNAP_PX = 6;
    const thresholdEmu = SNAP_PX / pxPerEmu;
    // Which edges are active for snapping based on which handle was grabbed.
    // Left edge handles snap liveX; right edge handles snap (liveX + liveCx).
    const xEdge: 'left' | 'right' | null =
      handle === 'nw' || handle === 'w' || handle === 'sw'
        ? 'left'
        : handle === 'ne' || handle === 'e' || handle === 'se'
          ? 'right'
          : null;
    const yEdge: 'top' | 'bottom' | null =
      handle === 'nw' || handle === 'n' || handle === 'ne'
        ? 'top'
        : handle === 'sw' || handle === 's' || handle === 'se'
          ? 'bottom'
          : null;

    const compute = (
      dxPx: number,
      dyPx: number,
      shift: boolean,
    ): { dx: number; dy: number; dcx: number; dcy: number; guides: SnapGuide[] } => {
      const dxEmu = Math.round(dxPx / pxPerEmu);
      const dyEmu = Math.round(dyPx / pxPerEmu);
      // Per-handle decomposition: which side moves (anchor offset), which
      // side stretches (size delta).
      let dx = 0;
      let dy = 0;
      let dcx = 0;
      let dcy = 0;
      // Horizontal axis
      if (handle === 'nw' || handle === 'w' || handle === 'sw') {
        // Left edge moves with cursor; width shrinks by the same amount.
        dx = dxEmu;
        dcx = -dxEmu;
      } else if (handle === 'ne' || handle === 'e' || handle === 'se') {
        // Right edge moves with cursor; width grows.
        dcx = dxEmu;
      }
      // Vertical axis
      if (handle === 'nw' || handle === 'n' || handle === 'ne') {
        dy = dyEmu;
        dcy = -dyEmu;
      } else if (handle === 'sw' || handle === 's' || handle === 'se') {
        dcy = dyEmu;
      }

      const isCorner =
        handle === 'nw' || handle === 'ne' || handle === 'sw' || handle === 'se';

      // Adobe PS/AI/ID/Figma/Keynote/PowerPoint shared convention: Shift
      // held while dragging a corner handle locks the aspect ratio. Edge
      // handles ignore Shift (single-axis resize is by definition
      // unconstrained — Adobe matches this). The lead axis is whichever
      // pixel movement is dominant; the other dimension is derived from
      // the original aspect ratio so resizing 1600×900 with Shift always
      // produces a 16:9 box. We skip snap entirely when constraining —
      // explicit ratio-lock intent should override convenience snapping
      // that might pull off-ratio (same rationale as Shift-drag move).
      // MIN_EMU is enforced inside this branch too so a tiny lead-axis
      // gesture can't produce a sub-min derived axis (which the regular
      // post-snap guard wouldn't catch in a ratio-aware way).
      if (shift && isCorner) {
        const ratio = frame.cx / frame.cy;
        let newCx: number;
        let newCy: number;
        if (Math.abs(dxPx) >= Math.abs(dyPx)) {
          newCx = Math.max(MIN_EMU, frame.cx + dcx);
          newCy = newCx / ratio;
          if (newCy < MIN_EMU) {
            newCy = MIN_EMU;
            newCx = MIN_EMU * ratio;
          }
        } else {
          newCy = Math.max(MIN_EMU, frame.cy + dcy);
          newCx = newCy * ratio;
          if (newCx < MIN_EMU) {
            newCx = MIN_EMU;
            newCy = MIN_EMU / ratio;
          }
        }
        const dcxNew = newCx - frame.cx;
        const dcyNew = newCy - frame.cy;
        // Translate size deltas back into anchor offsets for handles whose
        // left/top edges move (nw moves both; ne only top; sw only left;
        // se anchors top-left so dx = dy = 0).
        dx = handle === 'nw' || handle === 'sw' ? -dcxNew : 0;
        dy = handle === 'nw' || handle === 'ne' ? -dcyNew : 0;
        dcx = dcxNew;
        dcy = dcyNew;
        return { dx, dy, dcx, dcy, guides: [] };
      }

      // Snap the active edge(s) to nearest target, then propagate the snap
      // adjustment to the matching size delta so the opposite edge stays
      // anchored.
      const guides: SnapGuide[] = [];
      if (xEdge) {
        const edgePos = xEdge === 'left' ? frame.x + dx : frame.x + frame.cx + dcx;
        let best: { target: number; diff: number } | null = null;
        for (const t of xTargets) {
          const diff = t - edgePos;
          if (Math.abs(diff) <= thresholdEmu && (!best || Math.abs(diff) < Math.abs(best.diff))) {
            best = { target: t, diff };
          }
        }
        if (best) {
          if (xEdge === 'left') {
            dx += best.diff;
            dcx -= best.diff;
          } else {
            dcx += best.diff;
          }
          guides.push({ axis: 'x', pos: best.target });
        }
      }
      if (yEdge) {
        const edgePos = yEdge === 'top' ? frame.y + dy : frame.y + frame.cy + dcy;
        let best: { target: number; diff: number } | null = null;
        for (const t of yTargets) {
          const diff = t - edgePos;
          if (Math.abs(diff) <= thresholdEmu && (!best || Math.abs(diff) < Math.abs(best.diff))) {
            best = { target: t, diff };
          }
        }
        if (best) {
          if (yEdge === 'top') {
            dy += best.diff;
            dcy -= best.diff;
          } else {
            dcy += best.diff;
          }
          guides.push({ axis: 'y', pos: best.target });
        }
      }

      // Enforce minimum size by capping how far the active edge can travel.
      // For left/top edges: clamp dx so frame.cx + dcx >= MIN  =>  dx <= cx - MIN.
      if (dcx < 0 && frame.cx + dcx < MIN_EMU) {
        const allowed = frame.cx - MIN_EMU;
        dcx = -allowed;
        if (dx !== 0) dx = allowed;
      }
      if (dcy < 0 && frame.cy + dcy < MIN_EMU) {
        const allowed = frame.cy - MIN_EMU;
        dcy = -allowed;
        if (dy !== 0) dy = allowed;
      }
      return { dx, dy, dcx, dcy, guides };
    };

    const onMoveDoc = (ev: PointerEvent) => {
      const dxPx = ev.clientX - startX;
      const dyPx = ev.clientY - startY;
      const r = compute(dxPx, dyPx, ev.shiftKey);
      setResizeOffset({ dx: r.dx, dy: r.dy, dcx: r.dcx, dcy: r.dcy });
      onGuidesChange(r.guides);
    };
    const onUpDoc = (ev: PointerEvent) => {
      const dxPx = ev.clientX - startX;
      const dyPx = ev.clientY - startY;
      window.removeEventListener('pointermove', onMoveDoc);
      window.removeEventListener('pointerup', onUpDoc);
      setResizeOffset(null);
      onGuidesChange([]);
      if (Math.abs(dxPx) < 3 && Math.abs(dyPx) < 3) return;
      // Re-read shift on the up event so press/release between the last
      // move and release still reflects the user's final intent.
      const { dx, dy, dcx, dcy } = compute(dxPx, dyPx, ev.shiftKey);
      // Final clamp into slide bounds.
      let nx = frame.x + dx;
      let ny = frame.y + dy;
      let ncx = frame.cx + dcx;
      let ncy = frame.cy + dcy;
      if (nx < 0) {
        ncx += nx;
        nx = 0;
      }
      if (ny < 0) {
        ncy += ny;
        ny = 0;
      }
      if (nx + ncx > slideSize.cx) ncx = slideSize.cx - nx;
      if (ny + ncy > slideSize.cy) ncy = slideSize.cy - ny;
      ncx = Math.max(MIN_EMU, ncx);
      ncy = Math.max(MIN_EMU, ncy);
      onResize(nx, ny, ncx, ncy);
    };
    window.addEventListener('pointermove', onMoveDoc);
    window.addEventListener('pointerup', onUpDoc);
  };

  return (
    <div
      className={cn(
        'absolute group flex flex-col gap-0.5 p-1 rounded-sm border border-dashed transition-colors',
        anyActive ? 'border-primary bg-primary/5' : 'border-border/60 hover:border-primary/40',
        (dragOffset || resizeOffset) && 'opacity-80 ring-2 ring-primary',
      )}
      style={{
        left: `${left}%`,
        top: `${top}%`,
        width: `${width}%`,
        height: `${height}%`,
      }}
    >
      {/* Live coord/size readout — Adobe AI/ID/PS, Figma, Keynote all show
          "X 120, Y 80" while moving and "240 × 180" while resizing so the
          user has precision feedback before releasing instead of having to
          guess and verify via inspector. EMU → pt conversion: 1pt = 12700
          EMU; pt is the design-standard unit and matches the font-size
          spinner below. Anchored top-left inside the frame so the badge
          rides with the shape; pointer-events-none so it never eats the
          drag itself. Only renders during an active gesture, so idle
          slides stay clean. */}
      {(dragOffset || resizeOffset) && (
        <div
          className={cn(
            'absolute top-0 left-0 px-1.5 py-0.5 text-[10px] font-mono rounded-br whitespace-nowrap pointer-events-none z-10 shadow-sm',
            // Switch to a green badge when Alt-duplicating so the user has
            // unmistakable feedback that releasing now clones rather than
            // moves — matches Adobe/Figma's green "+" cursor convention.
            altDuplicating ? 'bg-emerald-600/90 text-white' : 'bg-zinc-900/90 text-zinc-50',
          )}
        >
          {resizeOffset
            ? `${Math.round(liveCx / 12700)} × ${Math.round(liveCy / 12700)} pt`
            : `${altDuplicating ? '+ 複製 ' : ''}X ${Math.round(liveX / 12700)} · Y ${Math.round(liveY / 12700)} pt`}
        </div>
      )}
      {/* Picture overlay — when this shape carries a `<a:blipFill>` we render
          the resolved data-URL as a stretched <img> behind the runs. The
          empty synthetic run inserted by `addPictureToSlide` sits on top so
          the user can still type a caption inside the picture frame. We use
          object-cover so the image fills like PowerPoint's stretch fill;
          pointer-events-none keeps clicks routed to the underlying handles. */}
      {pictureUrl && (
        <img
          src={pictureUrl}
          alt=""
          draggable={false}
          className="absolute inset-0 w-full h-full object-cover pointer-events-none rounded-sm"
        />
      )}
      {runs.map((run) => (
        <RunInputCanvas
          key={run.id}
          run={run}
          active={run.id === activeRunId}
          pxPerEmu={pxPerEmu}
          onFocus={() => onRunFocus(run.id)}
          onChange={(t) => onRunChange(run.id, t)}
        />
      ))}
      {/* Move handle — only on shapes that map to a real <p:sp>. */}
      {canMove ? (
        <button
          type="button"
          onPointerDown={startDrag}
          title="拖曳移動 · Shift 鎖軸 · Alt 拖曳複製"
          tabIndex={-1}
          className={cn(
            'absolute -top-2 -left-2 h-5 w-5 inline-flex items-center justify-center rounded-full',
            'bg-background border border-border text-muted-foreground hover:text-primary hover:border-primary',
            'cursor-grab active:cursor-grabbing',
            'opacity-0 pointer-events-none group-hover:opacity-100 group-hover:pointer-events-auto transition-opacity',
          )}
        >
          <GripVertical className="h-3 w-3" />
        </button>
      ) : null}
      {/* Frame-level delete: removes the whole shape (containing all its
          runs). Uses the first run's runIndex as the delete anchor since
          deleteTextBoxFromSlide drops the entire <p:sp>. */}
      {runs[0] ? (
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            onRunDelete(runs[0].runIndex);
          }}
          title="刪除此文字框"
          tabIndex={-1}
          className={cn(
            'absolute -top-2 -right-2 h-5 w-5 inline-flex items-center justify-center rounded-full',
            'bg-background border border-border text-muted-foreground hover:text-destructive hover:border-destructive',
            'opacity-0 pointer-events-none group-hover:opacity-100 group-hover:pointer-events-auto transition-opacity',
          )}
        >
          <X className="h-3 w-3" />
        </button>
      ) : null}
      {/* 8-handle resize. Corners scale both dimensions, edges scale one.
          Visible only on hover or when the shape is active so they don't
          clutter idle slides. Each handle catches its own pointerdown so
          startResize knows which edges are anchored.

          R118 — tooltip honesty / cross-editor parity. Previously these
          eight <span>s carried no `title=` at all, even though:
            • The corner handles silently support Shift-aspect-lock
              (startResize at line 1972-1989: `if (shift && isCorner)`
              switches into the ratio-preserving branch). A user dragging
              a corner had no way to discover the modifier without
              reading the source.
            • The move handle 30px above (line 2169) DOES disclose its
              modifiers (`title="拖曳移動 · Shift 鎖軸 · Alt 拖曳複製"`),
              so the same shape exposes a rich hover hint on one grip
              and zero on the eight grips wrapping it — a same-component
              parity break, not just a missing-tooltip one.
            • Cross-editor: DocxEditor.tsx:3200 and XlsxEditor.tsx:3542
              both expose their image-resize Shift behaviour via
              `title="拖曳調整大小 · Shift 解除等比例"`. PptxEditor's
              eight handles were the silent outliers across the three
              editors that share the resize-handle idiom.
            • The XlsxEditor cross-ref comment at XlsxEditor.tsx:3522-
              3523 explicitly claimed pptx resize handles "don't toggle
              aspect with Shift so they intentionally don't carry that
              suffix". That claim is contradicted by the L1972 branch —
              stale comment, fixed in the same round.

          Note the Shift semantic is INVERTED between editors and that
          inversion is intentional, matching each host application:
            • Word/Pages images: default = aspect-locked, Shift releases
              (DocxEditor / XlsxEditor wording: "Shift 解除等比例").
            • PowerPoint/Adobe shapes: default = free, Shift locks
              (PptxEditor wording: "Shift 鎖定等比例").
          So we can't reuse the docx/xlsx string verbatim — the verb
          flips. Edge handles get a separate string because Shift has
          no effect on them (line 1962 comment: "single-axis resize is
          by definition unconstrained — Adobe matches this"); pretending
          otherwise would import a fresh honesty gap. The shape `動詞
          · Shift 修飾語` mirrors the move handle's tooltip above so
          all nine grips on the same shape now read with one voice. */}
      {canMove ? (
        <>
          {(
            [
              ['nw', '-top-1 -left-1', 'cursor-nwse-resize'],
              ['n', '-top-1 left-1/2 -translate-x-1/2', 'cursor-ns-resize'],
              ['ne', '-top-1 -right-1', 'cursor-nesw-resize'],
              ['e', 'top-1/2 -right-1 -translate-y-1/2', 'cursor-ew-resize'],
              ['se', '-bottom-1 -right-1', 'cursor-nwse-resize'],
              ['s', '-bottom-1 left-1/2 -translate-x-1/2', 'cursor-ns-resize'],
              ['sw', '-bottom-1 -left-1', 'cursor-nesw-resize'],
              ['w', 'top-1/2 -left-1 -translate-y-1/2', 'cursor-ew-resize'],
            ] as const
          ).map(([h, pos, cur]) => {
            const isCorner = h === 'nw' || h === 'ne' || h === 'sw' || h === 'se';
            return (
              <span
                key={h}
                role="presentation"
                onPointerDown={(e) => startResize(e, h)}
                title={isCorner ? '拖曳調整大小 · Shift 鎖定等比例' : '拖曳調整大小（單軸）'}
                className={cn(
                  'absolute h-2 w-2 rounded-sm bg-background border border-primary',
                  pos,
                  cur,
                  'opacity-0 pointer-events-none group-hover:opacity-100 group-hover:pointer-events-auto transition-opacity',
                  anyActive && 'opacity-100 pointer-events-auto',
                )}
              />
            );
          })}
        </>
      ) : null}
    </div>
  );
}

function RunInputCanvas({
  run,
  active,
  pxPerEmu,
  onChange,
  onFocus,
}: {
  run: PptxTextRun;
  active: boolean;
  pxPerEmu: number;
  onChange: (text: string) => void;
  onFocus: () => void;
}): JSX.Element {
  const style = run.style ?? {};
  const taRef = useRef<HTMLTextAreaElement>(null);
  const ensureFocus = () => {
    const el = taRef.current;
    if (el && document.activeElement !== el) el.focus();
  };
  // Scale font: 1 pt = 12700 EMU. fontPx = pt * 12700 * pxPerEmu.
  // Cap to avoid 0-px text before the canvas measures, and never go below
  // 8 px so the textarea is always clickable.
  const ptToPx = pxPerEmu > 0 ? 12700 * pxPerEmu : 0;
  const sizePt = style.size ?? 18;
  const fontSizePx = ptToPx > 0 ? Math.max(8, Math.round(sizePt * ptToPx)) : 14;
  return (
    <div className="relative flex-1 min-h-0" onMouseDown={ensureFocus}>
      <textarea
        ref={taRef}
        // Hook for cross-slide Find→Locate: the parent's locateFindResult
        // queries this attribute to scroll the matched run into view (the
        // desk scrollport scrolls vertically when toolbar + 16:9 canvas +
        // notes exceed the wrapper height, so an off-fold match would
        // otherwise just show an invisible ring above/below the scrollport).
        data-run-id={run.id}
        value={run.text}
        onChange={(e) => onChange(e.target.value)}
        onFocus={onFocus}
        onClick={ensureFocus}
        className={cn(
          'absolute inset-0 w-full h-full bg-transparent rounded-sm px-1 py-0.5 outline-none resize-none cursor-text leading-tight',
          active ? 'ring-1 ring-primary' : 'focus:ring-1 focus:ring-primary',
        )}
        style={{
          fontWeight: style.bold ? 'bold' : undefined,
          fontStyle: style.italic ? 'italic' : undefined,
          textDecoration: style.underline ? 'underline' : undefined,
          color: style.color ? `#${style.color}` : undefined,
          fontSize: `${fontSizePx}px`,
          fontFamily: withEmojiFallback(style.fontFamily),
        }}
      />
    </div>
  );
}

/**
 * Built-in shape menu — order/labels are PowerPoint-ish so users don't have
 * to think about OOXML preset names. Icons borrow lucide's basic set; for
 * shapes lucide has no exact match (roundRect) we reuse Square so the row
 * stays visually balanced.
 */
const SHAPE_OPTIONS: ReadonlyArray<{
  kind: PptxShapeKind;
  label: string;
  Icon: typeof Square;
}> = [
  { kind: 'rect', label: '矩形', Icon: Square },
  { kind: 'roundRect', label: '圓角矩形', Icon: Square },
  { kind: 'ellipse', label: '橢圓', Icon: Circle },
  { kind: 'triangle', label: '三角形', Icon: Triangle },
  { kind: 'rightArrow', label: '右箭頭', Icon: ArrowRight },
];

/**
 * Six-button alignment palette. Adobe Illustrator / InDesign / PowerPoint
 * all expose this as a horizontal strip of identical-sized icon buttons in
 * exactly this order (left/center/right ⊕ top/middle/bottom), with a
 * vertical divider between the two axis groups. Sticking to the convention
 * means the buttons are recognisable on first sight without reading labels.
 */
function AlignmentPalette({
  onAlign,
}: {
  onAlign: (
    axis: 'left' | 'hcenter' | 'right' | 'top' | 'vmiddle' | 'bottom',
  ) => void;
}): JSX.Element {
  // Tooltip wording calls out "對齊投影片" so the user knows the alignment
  // container is the slide rectangle (vs. selection or sibling shapes — the
  // multi-select-relative variants come once we add marquee selection).
  return (
    <div className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded border border-border">
      <AlignBtn title="對齊投影片左緣" onClick={() => onAlign('left')}>
        <AlignLeft className="h-3.5 w-3.5" />
      </AlignBtn>
      <AlignBtn title="水平置中於投影片" onClick={() => onAlign('hcenter')}>
        <AlignCenter className="h-3.5 w-3.5" />
      </AlignBtn>
      <AlignBtn title="對齊投影片右緣" onClick={() => onAlign('right')}>
        <AlignRight className="h-3.5 w-3.5" />
      </AlignBtn>
      <div className="w-px h-4 bg-border mx-0.5" />
      <AlignBtn title="對齊投影片頂端" onClick={() => onAlign('top')}>
        <AlignStartHorizontal className="h-3.5 w-3.5" />
      </AlignBtn>
      <AlignBtn title="垂直置中於投影片" onClick={() => onAlign('vmiddle')}>
        <AlignCenterVertical className="h-3.5 w-3.5" />
      </AlignBtn>
      <AlignBtn title="對齊投影片底端" onClick={() => onAlign('bottom')}>
        <AlignEndHorizontal className="h-3.5 w-3.5" />
      </AlignBtn>
    </div>
  );
}

function AlignBtn({
  title,
  onClick,
  children,
}: {
  title: string;
  onClick: () => void;
  children: React.ReactNode;
}): JSX.Element {
  return (
    <button
      type="button"
      onClick={onClick}
      title={title}
      className="inline-flex items-center justify-center w-6 h-6 rounded text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
    >
      {children}
    </button>
  );
}

function ShapePicker({ onAdd }: { onAdd: (kind: PptxShapeKind) => void }): JSX.Element {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button
          type="button"
          className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded border border-border hover:bg-secondary text-muted-foreground hover:text-foreground transition-colors"
          // R108 — disclose the dropdown nature + the five shape options
          // + the destination, in one tooltip. Two converging gaps left
          // this trigger silent on every dimension that matters:
          //
          // 1. Dropdown discoverability. The button has no chevron icon
          //    (just <Shapes/>), so visually it reads as a single-action
          //    button rather than a menu trigger. Same-codebase precedent
          //    at TabBar.tsx:564 sets the convention for dropdown-trigger
          //    tooltips:「新增頁籤 — 點擊選擇格式 (Markdown / Word /
          //    Excel / PowerPoint)」 — the trailing chevron in TabBar
          //    makes the menu nature visible, but the tooltip still
          //    enumerates the options because the chevron alone doesn't
          //    name them. Here there's no chevron AND no enumeration,
          //    so users can't tell either that a click reveals choices
          //    OR what those choices are without committing to the
          //    click and dismissing if uninterested.
          //
          // 2. Tooltip = label (zero info on hover). The previous
          //    `title="新增圖形"` matched the visible label byte-for-
          //    byte (line 2354), so a hover-to-discover user got
          //    nothing the button text didn't already say. Compare
          //    sibling LayoutPicker (line 2486) whose tooltip
          //    「套用版面配置」at least adds the「套用」verb absent
          //    from its「版面配置」label — this trigger was the lone
          //    redundant-tooltip outlier in this same dropdown family.
          //
          // 3. Destination undisclosed. R107 just landed
          //    「插入圖片到目前投影片」 on the same-row 插入圖片 button
          //    (line 1115); ShapePicker invokes the exact same
          //    runStructuralOp(activeIdx) pipeline (handleAddShape at
          //    line 438-444), so the「目前投影片」disclosure applies
          //    word-for-word. Reusing「目前投影片」 verbatim keeps
          //    the row's destination vocabulary in one voice.
          //
          // Wording mirrors TabBar.tsx:564's structure verbatim:
          // 「{verb}{dest} — 點擊選擇 X (option list)」. Option order
          // matches the SHAPE_OPTIONS array literal at line 2271-2275
          // so future shape additions only need to be appended in one
          // place mentally — same property R100 protected for the
          // image format list.
          title="新增圖形到目前投影片 — 點擊選擇形狀 (矩形 / 圓角矩形 / 橢圓 / 三角形 / 右箭頭)"
        >
          <Shapes className="h-3 w-3" />
          新增圖形
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="min-w-[160px]">
        {SHAPE_OPTIONS.map(({ kind, label, Icon }) => (
          <DropdownMenuItem
            key={kind}
            onSelect={() => onAdd(kind)}
            className="flex items-center gap-2"
          >
            <Icon className="h-3.5 w-3.5 text-muted-foreground" />
            <span className="text-sm">{label}</span>
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function LayoutPicker({ onApply }: { onApply: (id: PptxLayoutId) => void }): JSX.Element {
  // R145 — flat-index drift correction (R136/R137/R140 paradigm).
  // Two upstream doc-comments cited this function at two DIFFERENT
  // wrong line numbers — internal contradiction visible without even
  // checking the actual definition: line 1135 said `line 2466`, line
  // 2445 said `line 2380`, but `function LayoutPicker` exists exactly
  // once. ShapePicker had a sibling +26 drift (line 2391 → 2417) in
  // the same R109/R108/R107 enumeration block. All three citations
  // corrected in-place to the post-shift positions; this anchor sits
  // INSIDE LayoutPicker's body so the function-definition line itself
  // (now 2486) does not move and no upstream cite needs a second
  // revision — the same minimum-blast-radius placement R142 used to
  // avoid same-round-self-drift loops.
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        {/* R109 — disclose dropdown nature + the six layouts + the
            destination + the destructive side-effect, in one tooltip.
            Prior shape: `title="套用版面配置"` told the user the verb
            (套用) but nothing else, leaving four gaps stacked on this
            single trigger:

            1. Dropdown discoverability. Like ShapePicker pre-R108, the
               button has no chevron icon (just <LayoutTemplate/>), so
               visually it reads as a single-action button rather than
               a menu trigger. TabBar.tsx:564 sets the codebase
               convention for dropdown-trigger tooltips:
               「新增頁籤 — 點擊選擇格式 (Markdown / Word / Excel /
               PowerPoint)」 — the「點擊選擇 X (option list)」suffix
               makes the dropdown visible on hover.

            2. Hidden options. PPTX_LAYOUTS at pptx-adapter.ts:803-810
               holds six choices (標題投影片 / 標題與內容 / 兩欄內容 /
               章節標題 / 僅標題 / 空白). Same-row sibling 新增圖形
               just got its five shapes enumerated in R108 by the same
               argument — there is no reason this dropdown's options
               should remain hidden when the next button over now lists
               its own.

            3. Destination undisclosed. Same-row R107 just landed
               「插入圖片到目前投影片」 on 插入圖片, and R108 followed
               with「新增圖形到目前投影片」on 新增圖形. handleApplyLayout
               at line 487-503 likewise applies to `activeIdx`, so the
               same「到目前投影片」disclosure is owed here — three
               consecutive same-row buttons should not split 2-vs-1 on
               whether they name where the action lands.

            4. Destructive-side-effect honesty (unique to this picker).
               handleApplyLayout at line 492-497 prompts a confirm
               dialog「套用「${label}」會清除目前投影片的文字框內容，
               是否繼續？」 — i.e. clicking a layout silently destroys
               existing text-box content. ShapePicker / 插入圖片 have no
               such side-effect, so this is the one place tooltip-vs-
               actual-behavior honesty (R37 / R57 family) demands an
               extra clause beyond the dropdown-trigger template. The
               trailing「· 會清除目前投影片的文字框內容」warns of the
               confirm BEFORE the user commits to clicking, matching the
               same middle-dot separator pattern used elsewhere
               (TabBar.tsx:485 keeps its option list parenthesized; the
               destructive warning rides as a separate clause after「·」
               so it doesn't get parsed as a seventh layout option).

            Wording mirrors TabBar.tsx:485 / R108 verbatim for the
            first half (「{verb}{dest} — 點擊選擇 X (option list)」)
            then appends the unique destructive-honesty clause.

            R144 — destructive clause was previously 「· 會清除既有文字框
            內容」, drifting from the confirm dialog body 「會清除目前投影
            片的文字框內容」 in two ways: (a) 「既有」 vs 「目前投影片的」
            scope vocabulary, and (b) the comment self-quote here had
            also dropped 「內容」 to read 「· 會清除既有文字框」 (R133/
            R135 self-quoting paradigm). Now identical to the confirm
            dialog's destructive substring verbatim — same R138 cross-
            entry-point alignment principle (TabBar disabled tooltip ==
            non-active-tab toast == Ctrl+E toast). The same comment block
            already cites the confirm dialog's exact string at item 4
            opening, so any reader who scrolled top-to-bottom would see
            the two phrasings within five lines of each other and notice
            the scope-vocabulary mismatch — closing it. */}
        <button
          type="button"
          className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded border border-border hover:bg-secondary text-muted-foreground hover:text-foreground transition-colors"
          title="套用版面配置到目前投影片 — 點擊選擇配置 (標題投影片 / 標題與內容 / 兩欄內容 / 章節標題 / 僅標題 / 空白) · 會清除目前投影片的文字框內容"
        >
          <LayoutTemplate className="h-3 w-3" />
          版面配置
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="min-w-[180px]">
        {PPTX_LAYOUTS.map((l) => (
          <DropdownMenuItem key={l.id} onSelect={() => onApply(l.id)} className="flex flex-col items-start gap-0.5">
            <span className="text-sm">{l.label}</span>
            <span className="text-[10px] text-muted-foreground">{l.hint}</span>
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

function NotesEditor({
  value,
  onChange,
}: {
  value: string;
  onChange: (t: string) => void;
}): JSX.Element {
  // Local mirror so typing doesn't fight the debounced re-serialize.
  const [local, setLocal] = useState(value);
  useEffect(() => {
    setLocal(value);
  }, [value]);
  return (
    <div className="pt-3 mt-3 border-t border-border">
      <div className="text-xs text-muted-foreground mb-1">備忘稿（speaker notes，僅在放映模式可見）</div>
      <textarea
        value={local}
        onChange={(e) => {
          setLocal(e.target.value);
          onChange(e.target.value);
        }}
        rows={4}
        placeholder="在這裡寫下這張投影片的備忘稿…"
        // R156 — explicit accessible name. The visible「備忘稿（speaker
        // notes，僅在放映模式可見）」label at line 2611 sits above the
        // textarea but isn't programmatically associated (it's a `<div>`,
        // not a `<label htmlFor=…>`). SR users tabbing here would miss it.
        // The shorter「備忘稿」 stays anchored when placeholder vanishes
        // on input — same rationale as R155's FindReplaceDialog inputs and
        // R156's AIPanel / XlsxEditor inputs.
        aria-label="備忘稿"
        className="w-full bg-secondary/20 border border-border rounded px-3 py-2 text-sm outline-none resize-y focus:ring-1 focus:ring-primary"
      />
    </div>
  );
}

function PresentationMode({
  slides,
  startIdx,
  onClose,
}: {
  slides: PptxModel['slides'];
  startIdx: number;
  onClose: () => void;
}): JSX.Element {
  const [idx, setIdx] = useState(Math.max(0, Math.min(startIdx, slides.length - 1)));
  const [showNotes, setShowNotes] = useState(false);
  /**
   * Two-step end-of-deck exit. Without this, Right / Space / PageDown on the
   * last slide silently no-ops (the old `Math.min(len-1, i+1)` clamp) and
   * mid-presentation the user spams Space wondering if the keyboard died,
   * then has to remember Esc to bail. PowerPoint / Keynote / Google Slides
   * / reveal.js all show a black "End of slide show" screen on the first
   * advance past the end, then exit on the next advance — a deliberate
   * pause so the speaker registers "deck is over" before being dropped
   * back to edit mode (which would otherwise be a jarring transition mid-
   * presentation, especially with an audience watching).
   *
   * The state is reset to false whenever idx moves (Home / End / Left /
   * PageUp re-enters the deck), and Left / PageUp from the end-screen
   * brings the user back to the last slide rather than the second-to-last
   * — they were "past" slide N, not on N-1.
   */
  const [atEnd, setAtEnd] = useState(false);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      } else if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
        e.preventDefault();
        if (atEnd) {
          onClose();
          return;
        }
        setIdx((i) => {
          if (i >= slides.length - 1) {
            setAtEnd(true);
            return i;
          }
          return i + 1;
        });
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        if (atEnd) {
          setAtEnd(false);
          return;
        }
        setIdx((i) => Math.max(0, i - 1));
      } else if (e.key === 'Home') {
        // Jump to the first slide. PowerPoint / Keynote / Google Slides /
        // reveal.js all bind Home here — without it, getting back to the
        // start of a long deck mid-presentation means spamming ← which is
        // visibly distracting to the audience. End is symmetric.
        e.preventDefault();
        setAtEnd(false);
        setIdx(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        setAtEnd(false);
        setIdx(Math.max(0, slides.length - 1));
      } else if (e.key.toLowerCase() === 'n') {
        e.preventDefault();
        setShowNotes((v) => !v);
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [slides.length, onClose, atEnd]);

  const slide = slides[idx];
  return (
    <div
      // Marker for the App-level focus-mode Esc handler: when this overlay
      // is mounted it owns the Esc gesture (closes the presentation), so
      // App.tsx must bail and not also exit focus mode in the same press.
      data-pptx-presenting="true"
      className="fixed inset-0 z-50 bg-black text-white flex flex-col"
      // Click anywhere on the end-screen to dismiss, mirroring PowerPoint.
      // Scoped to atEnd so a stray click during the actual presentation
      // doesn't accidentally exit (the surrounding bar buttons keep their
      // own onClick — clicks on them stop propagating via React's default
      // bubbling, which is fine because their handlers fire first).
      onClick={atEnd ? onClose : undefined}
    >
      <div className="flex items-center justify-between px-4 py-2 text-xs bg-black/70 border-b border-white/10">
        <div>
          {atEnd ? '簡報結束' : `投影片 ${idx + 1} / ${slides.length}`}
        </div>
        <div className="flex items-center gap-3 text-white/70">
          <span>
            {atEnd
              ? '按任意鍵 / 點擊離開 · ← / PageUp 返回最後一張'
              // Surface Space alongside →: it's the universal slideshow-
              // advance key (PowerPoint / Keynote / reveal.js — see the
              // atEnd comment block above) and what every wireless presenter
              // clicker sends. Previously the hint only mentioned arrows, so
              // first-time users hit Space, got no feedback (it works, but
              // the bar didn't say so), and assumed the keyboard was dead —
              // exactly the failure mode the implementation comment at
              // line 2308 already calls out. Splitting forward / back keys
              // because Space only advances; lumping it under "切換"
              // would imply Space goes backward too, which it doesn't.
              : '→ / Space / PageDown 下一張 · ← / PageUp 上一張 · Home / End 首末張 · N 顯示備忘稿 · Esc 結束'}
          </span>
          <button
            type="button"
            onClick={(e) => {
              // Stop propagation so the wrapper's atEnd onClick doesn't
              // double-fire onClose (idempotent today, but the redundant
              // call would also trigger any future onClose side-effects
              // twice — cheap to guard).
              e.stopPropagation();
              onClose();
            }}
            className="inline-flex items-center gap-1 px-2 py-1 rounded hover:bg-white/10"
            title="結束放映 (Esc)"
          >
            <X className="h-3.5 w-3.5" />
            結束
          </button>
        </div>
      </div>
      <div className="flex-1 min-h-0 flex">
        <div className="flex-1 min-w-0 flex items-center justify-center p-10 overflow-auto">
          {atEnd ? (
            // Black end-screen — convention from every major presentation
            // tool. Big calm message, no slide chrome, no white frame; the
            // visual hand-off says "the deck is over" without yanking the
            // audience back into edit-mode UI mid-room.
            <div className="text-center select-none">
              <div className="text-3xl font-medium tracking-wide">簡報結束</div>
              <div className="mt-3 text-sm text-white/60">
                按任意鍵 / 點擊離開 · ← / PageUp 返回最後一張
              </div>
            </div>
          ) : (
            <div className="w-full max-w-4xl aspect-[16/9] bg-white text-black rounded shadow-2xl p-10 overflow-auto flex flex-col gap-4">
              {slide.runs.length === 0 ? (
                <div className="text-gray-400 italic m-auto">（這張投影片沒有文字）</div>
              ) : (
                slide.runs.map((run) => {
                  const st = run.style ?? {};
                  return (
                    <div
                      key={run.id}
                      className="whitespace-pre-wrap leading-snug"
                      style={{
                        fontWeight: st.bold ? 'bold' : undefined,
                        fontStyle: st.italic ? 'italic' : undefined,
                        textDecoration: st.underline ? 'underline' : undefined,
                        color: st.color ? `#${st.color}` : undefined,
                        fontSize: st.size ? `${st.size}px` : '24px',
                        fontFamily: withEmojiFallback(st.fontFamily),
                      }}
                    >
                      {run.text || ' '}
                    </div>
                  );
                })
              )}
            </div>
          )}
        </div>
        {!atEnd && showNotes ? (
          // Render the panel whenever showNotes is true, even on slides with
          // no notes — mirrors PowerPoint / Keynote speaker view, where the
          // notes pane is the user's persistent preference, not a per-slide
          // toggle.
          //
          // The previous gate (`showNotes && slide.notesText`) made the panel
          // silently disappear when advancing onto a notes-less slide, while
          // the bottom-bar button kept saying "隱藏備忘稿" (line 2885 reflects
          // showNotes, not actual visibility) — button label and screen state
          // disagreed, and pressing N appeared to do nothing. Worse, pressing
          // N on a notes-less slide as a first interaction flipped showNotes
          // to true with zero visible feedback, reading as a broken keystroke.
          //
          // With the unconditional render + placeholder, N is now a stable
          // self-describing toggle: opening it on a notes-less slide shows
          // "（這張投影片沒有備忘稿）" so the cause is clear; advancing into
          // notes-less territory keeps the pane reserved so the layout doesn't
          // jump between slides; closing reflects the same intent everywhere.
          <div className="w-80 shrink-0 border-l border-white/10 bg-black/80 p-4 overflow-auto">
            <div className="text-[11px] uppercase tracking-wider text-white/50 mb-2">備忘稿</div>
            {slide.notesText ? (
              <div className="text-sm whitespace-pre-wrap text-white/90">{slide.notesText}</div>
            ) : (
              <div className="text-sm italic text-white/40">（這張投影片沒有備忘稿）</div>
            )}
          </div>
        ) : null}
      </div>
      {/* Sibling-shortcut-in-tooltip parity inside PresentationMode: the close
          button at line ~2453 carries `title="結束放映 (Esc)"`, and the notes
          toggle two slots over inlines `(N)` to disambiguate that the
          parenthetical is a keystroke, not decoration. The bottom-bar prev /
          next buttons surfaced only the `←` / `→` glyphs, which read
          ambiguously as directional symbols rather than literal arrow-key
          hints — and they hid four other wired keystrokes (PageUp/PageDown
          per line 2648/2661, Space per line 2648, Home/End per line 2668/2676
          for the prev jump-to-start case). The header status line at line
          2722 enumerates all of them ("→ / Space 下一張 · ← 上一張 · Home /
          End 首末張 · N 顯示備忘稿 · Esc 結束") but that hint disappears the
          moment the user reaches the end-of-deck black screen, where the
          prev button is the one control still keyboard-navigable to bounce
          back into the deck. Tooltips on the buttons themselves are the only
          hover-stable surface, mirroring the close-button pattern. The
          atEnd branches retain the existing label semantics: prev becomes
          "返回最後一張" with the same `←` key still wired (line 2661-2667),
          next becomes "結束放映" sharing the close button's Esc keystroke
          (the keymap at line 2648-2660 lets →/Space exit from atEnd too,
          but matching the close button's tooltip wording keeps the exit
          gesture consistent across both surfaces). */}
      {/* R101 — boundary-aware tooltips for the disabled states on this nav
          row, closing the gap with the same-file SlideRail rail icons
          (line 3155/3162/3193) and the right-click menu (line 3306/3319/
          3333) which both already explain "已經是第一張" / "已經是最後一張"
          / "至少要保留一張投影片" instead of repeating the action label
          on a greyed-out control. Three surfaces in this file share the
          disabled-when-at-boundary pattern; SlideRail and the context
          menu self-explain, but PresentationMode's nav row sat outside
          the convention — same as the R92/R97/R98 batches that closed
          this anti-pattern across the editor toolbars. The boundary
          string for the prev button is reused verbatim from line 3155
          ("已經是第一張") so identical states read the same string
          across all three surfaces. The notes-toggle's atEnd disable is
          a softer boundary (end-screen, no slide to read notes for),
          so its message points the user at the recovery action — the
          ← key — instead of just naming the state. */}
      <div className="flex items-center justify-between px-4 py-2 bg-black/70 border-t border-white/10">
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            if (atEnd) setAtEnd(false);
            else setIdx((i) => Math.max(0, i - 1));
          }}
          disabled={!atEnd && idx === 0}
          title={
            !atEnd && idx === 0
              ? '已經是第一張'
              : atEnd
                ? '返回最後一張 (← / PageUp)'
                : '上一張 (← / PageUp)'
          }
          className="px-3 py-1 text-xs rounded hover:bg-white/10 disabled:opacity-30"
        >
          {atEnd ? '← 返回最後一張' : '← 上一張'}
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            setShowNotes((v) => !v);
          }}
          disabled={atEnd}
          title={
            atEnd
              ? '簡報已結束 — 按 ← / PageUp 返回最後一張即可切換備忘稿'
              : showNotes
                ? '隱藏備忘稿 (N)'
                : '顯示備忘稿 (N)'
          }
          className="px-3 py-1 text-xs rounded hover:bg-white/10 disabled:opacity-30"
        >
          {showNotes ? '隱藏備忘稿' : '顯示備忘稿'} (N)
        </button>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            if (atEnd) {
              onClose();
              return;
            }
            if (idx >= slides.length - 1) setAtEnd(true);
            else setIdx((i) => i + 1);
          }}
          title={atEnd ? '結束放映 (Esc)' : '下一張 (→ / Space / PageDown)'}
          className="px-3 py-1 text-xs rounded hover:bg-white/10"
        >
          {atEnd ? '結束放映' : '下一張 →'}
        </button>
      </div>
    </div>
  );
}

function SlideRail({
  tabId,
  slideCount,
  activeIdx,
  titles,
  onSelect,
  onDuplicate,
  onDelete,
  onMoveUp,
  onMoveDown,
  onReorder,
}: {
  tabId: string;
  slideCount: number;
  activeIdx: number;
  /** Heuristic title per slide (same source the outline panel uses).
   *  Empty / missing entry falls back to the generic "投影片 N" label. */
  titles: Array<{ idx: number; title: string }>;
  onSelect: (i: number) => void;
  onDuplicate: (i: number) => void;
  onDelete: (i: number) => void;
  onMoveUp: (i: number) => void;
  onMoveDown: (i: number) => void;
  /** HTML5 drag-and-drop reorder. The chevron buttons remain (good for
   *  precise single-position moves and keyboard-only users), but a 30-slide
   *  deck is unworkable with chevrons alone — drag is the fallback for
   *  bulk reorders. Mirrors TabBar's drag pattern for consistency. */
  onReorder: (from: number, to: number) => void;
}): JSX.Element {
  // Drag state lives inside the rail (not lifted to PptxEditor) — outside
  // observers don't need to know which slide is mid-drag, and keeping it
  // local avoids re-rendering the slide canvas every dragover tick.
  const [dragIdx, setDragIdx] = useState<number | null>(null);
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null);
  // Right-click menu on slide rows. Mirrors the XlsxEditor sheet-tab pattern
  // (and TabBar's tab right-click) so users learn one gesture for "act on a
  // row I'm pointing at" across the app. Without this, deleting / duplicating
  // a non-active slide takes two clicks: first to make it active so the
  // hover-action icons appear (those icons only render on the active row,
  // line 3119-3198), then a second click on the icon. Right-click collapses
  // that to one gesture against the slide the user actually pointed at.
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
  // Auto-scroll the active slide row into view, the same way PptxNavPanel
  // already does for the title-outline pane below. Without this, in a long
  // deck (>15 slides) the rail's viewport keeps showing slides 1-15 while
  // the user is editing slide 30 — the highlighted row sits invisibly
  // below the fold. Triggers that move activeIdx off-screen include:
  //   • Round 12's tab-switch restore (slide 30 selected, rail starts at top)
  //   • "+ 新增投影片" duplicating near the bottom of the visible window
  //   • GoToDialog (Ctrl+G) jumping to an arbitrary slide number
  //   • PresentationMode close returning at the last-shown index
  // `block: 'nearest'` mirrors PptxNavPanel — no jiggling when already
  // visible. Re-runs on slideCount as well so a duplicate-then-jump that
  // adds a row above the active one (rare, but possible via reorder)
  // still settles to the right offset.
  const railScrollRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = railScrollRef.current?.querySelector<HTMLElement>(
      `[data-pptx-rail-idx="${activeIdx}"]`,
    );
    el?.scrollIntoView({ block: 'nearest' });
  }, [activeIdx, slideCount]);
  // Rail scroll memory across tab swaps. Restore via `queueMicrotask`
  // so we run AFTER the activeIdx effect above (which declares first
  // and synchronously scrollIntoViews the active row) — microtask flush
  // comes after the current render's effect chain, so the remembered
  // offset wins. Without this ordering trick, mounting on a tab whose
  // active slide is near the top would snap the rail to the top, even
  // though the user had scrolled to study slides 30-45 before tab-
  // swapping away. Same pattern as PptxNavPanel (Round 36).
  useEffect(() => {
    const remembered = railScrollMemory.get(tabId);
    if (remembered != null) {
      queueMicrotask(() => {
        const el = railScrollRef.current;
        if (el) el.scrollTop = remembered;
      });
    }
    return () => {
      const el = railScrollRef.current;
      if (el) railScrollMemory.set(tabId, el.scrollTop);
    };
  }, [tabId]);

  return (
    <div
      ref={railScrollRef}
      className="w-36 shrink-0 border-r border-border overflow-y-auto bg-secondary/20 flex flex-col"
    >
      <div className="p-2 space-y-1 flex-1">
        {Array.from({ length: slideCount }, (_, i) => {
          const active = i === activeIdx;
          // Reuse the outline panel's heuristic title so a long deck doesn't
          // collapse to a column of indistinguishable "投影片 N" labels. The
          // rail is narrow (w-36 = 144px) so the title gets truncated, but a
          // truncated real title still tells the user 100x more than the
          // bare slide number. `(無標題)` from the outline survives when a
          // slide has no detectable title shape — we replace that with the
          // generic label so the rail doesn't show a literal "(無標題)"
          // tooltip pretending to be useful information.
          const titleEntry = titles.find((t) => t.idx === i);
          const rawTitle = titleEntry?.title ?? '';
          const hasTitle = rawTitle && rawTitle !== '(無標題)';
          const display = hasTitle ? rawTitle : `投影片 ${i + 1}`;
          return (
            <div
              key={i}
              data-pptx-rail-idx={i}
              draggable
              onDragStart={(e) => {
                setDragIdx(i);
                // 'move' = in-app rearrange, not file copy — same OS-level
                // cue TabBar uses. Required so the cursor isn't a "no-drop"
                // icon throughout the drag.
                e.dataTransfer.effectAllowed = 'move';
              }}
              onDragOver={(e) => {
                if (dragIdx === null || dragIdx === i) return;
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                if (dragOverIdx !== i) setDragOverIdx(i);
              }}
              onDragLeave={() => {
                // dragleave fires on every nested child crossing (icon, label,
                // chevron buttons) so a naive clear flickers. Same fix as
                // TabBar — only clear if we're still the marked target.
                if (dragOverIdx === i) setDragOverIdx(null);
              }}
              onDragEnd={() => {
                setDragIdx(null);
                setDragOverIdx(null);
              }}
              onDrop={() => {
                if (dragIdx === null || dragIdx === i) {
                  setDragOverIdx(null);
                  return;
                }
                onReorder(dragIdx, i);
                setDragIdx(null);
                setDragOverIdx(null);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                // Make the right-clicked slide active first — the menu's
                // delete / duplicate actions otherwise feel disconnected
                // from where the user pointed. Same pattern as XlsxEditor's
                // sheet-tab right-click (line 2466-2473).
                onSelect(i);
                setCtxMenu({ x: e.clientX, y: e.clientY, idx: i });
              }}
              className={cn(
                'group rounded text-xs transition-colors',
                active ? 'bg-primary text-primary-foreground' : 'hover:bg-secondary text-foreground',
                // Faded source + outlined drop target — same vocabulary as
                // TabBar so users who've reordered tabs read the rail
                // without re-learning. Only show drop ring when the index
                // would actually change (skip self-drop).
                dragIdx === i && 'opacity-40',
                dragOverIdx === i && dragIdx !== i && 'ring-2 ring-primary',
              )}
            >
              <button
                type="button"
                onClick={() => onSelect(i)}
                // Advertise the two hidden gestures the row supports —
                // drag-to-reorder (line 3029 `draggable`) and the right-
                // click context menu (line 3062). Both are 100% wired but
                // 0% discoverable from a static slide-title tooltip; users
                // who don't know about them fall back to clicking the up/
                // down chevrons one step at a time, which is unworkable for
                // a 30-slide deck (the exact scenario the drag handler was
                // added for, per the doc-comment at line 2937-2939). Same
                // dialect as the sibling row-style controls that already
                // got this fix:
                //   • XlsxEditor sheet tab (XlsxEditor.tsx:2580):
                //       title={`${s.name}（雙擊重新命名 · 右鍵選單）`}
                //   • TabBar workspace tab (TabBar.tsx:371):
                //       …（雙擊重新命名 · 中鍵關閉 · 右鍵選單）
                // Slide rows have no double-click-rename gesture (titles
                // are derived from slide content, not user-editable
                // labels) so the hint set is just drag + right-click,
                // matching what's actually available here.
                title={`${display}（拖曳排序 · 右鍵選單）`}
                className="w-full flex items-center gap-2 px-2 py-2"
              >
                <span className={cn('font-mono w-6 text-right shrink-0', active ? 'opacity-90' : 'text-muted-foreground')}>
                  {i + 1}
                </span>
                <Presentation className="h-3.5 w-3.5" />
                <span
                  className={cn(
                    'truncate flex-1 text-left',
                    !hasTitle && (active ? 'opacity-80' : 'text-muted-foreground italic'),
                  )}
                >
                  {display}
                </span>
              </button>
              {active && (
                // State-aware tooltips on the four rail-icon buttons mirror
                // the R49/R50 fixes that already landed for XlsxEditor's
                // row/col delete (XlsxEditor.tsx:1910-1924, 1938-1953) and
                // DocxEditor's per-row trash (DocxEditor.tsx:2657-2675):
                // any button that disables at a boundary should explain why
                // instead of leaving the user staring at a greyed icon. The
                // slide rail was the lone format still showing static "上移"
                // / "下移" / "刪除" labels — same disable mechanic, same
                // failure mode (user clicks repeatedly, wonders if app is
                // broken, no signal that they're at the edge / deck floor).
                // Slide-delete in particular is destructive intent; the
                // adapter throws "pptx_min_one_slide" at the floor (pptx-
                // adapter.ts:247-249), so the button's disable is load-
                // bearing — surfacing "至少要保留一張投影片" matches the
                // sibling sheet-delete tooltip at XlsxEditor.tsx:2696 and
                // its destructive-floor message exactly. Enabled-state
                // wording for the Trash2 button was tightened in a
                // follow-up: see the inline comment attached to it below
                // for the destination-disclosure rationale.
                <div className="flex items-center justify-around px-1 pb-1">
                  {/* Same-row scope-disclosure parity across the 4-button
                      rail. The Trash2 comment ~30 lines below already lays
                      out the `動詞 + 此 + 投影片` shape and explicitly
                      enumerates the desired set ending in `上移此投影片` /
                      `下移此投影片` / `複製此投影片` / `刪除此投影片`, but
                      these two move buttons used to read just `上移` / `下移`
                      — i.e. the same comment that prescribed the shape was
                      contradicted by the two siblings sitting directly above
                      it. The disabled-state messages already structurally
                      name the target (`已經是第一張` / `已經是最後一張`,
                      with `張` implying 投影片), so the enabled state was
                      the lone scope drop in a row that otherwise consistently
                      discloses its target. Aligned to match the rail's
                      established vocabulary verbatim. */}
                  <RailIconBtn
                    title={i === 0 ? '已經是第一張' : '上移此投影片'}
                    disabled={i === 0}
                    onClick={() => onMoveUp(i)}
                  >
                    <ChevronUp className="h-3 w-3" />
                  </RailIconBtn>
                  <RailIconBtn
                    title={i === slideCount - 1 ? '已經是最後一張' : '下移此投影片'}
                    disabled={i === slideCount - 1}
                    onClick={() => onMoveDown(i)}
                  >
                    <ChevronDown className="h-3 w-3" />
                  </RailIconBtn>
                  <RailIconBtn title="複製此投影片" onClick={() => onDuplicate(i)}>
                    <Copy className="h-3 w-3" />
                  </RailIconBtn>
                  {/* Destination disclosure on the enabled-state tooltip:
                      previously this read just `刪除` while every sibling
                      disclosed its target — the same-row neighbour at line
                      3168 reads `複製此投影片`, the in-file frame-delete at
                      PptxEditor.tsx:2191 reads `刪除此文字框`, and the
                      cross-editor twin XlsxEditor.tsx:2696 reads
                      `刪除工作表`. The disabled-state message already
                      structurally names the target (`至少要保留一張投影
                      片`, see comment ~25 lines up), so the enabled state
                      silently dropped the disclosure that every sibling
                      provides — and the destructive button (Trash2) was
                      the one with the LEAST clarity, the inversion of
                      what destructive-action affordances should do. The
                      adopted `刪除此投影片` mirrors the sibling Copy
                      tooltip's `動詞 + 此 + 投影片` shape verbatim, so
                      hovering across the four-button rail now reads as a
                      consistent set of slide-targeted verbs (`上移此投影片`
                      / `下移此投影片` / `複製此投影片` / `刪除此投影片`)
                      instead of two slide-aware strings flanking two bare
                      verbs. (The two move buttons were aligned in a follow-
                      up — see the comment attached to them above.) */}
                  <RailIconBtn
                    title={slideCount <= 1 ? '至少要保留一張投影片' : '刪除此投影片'}
                    disabled={slideCount <= 1}
                    onClick={() => onDelete(i)}
                  >
                    <Trash2 className="h-3 w-3" />
                  </RailIconBtn>
                </div>
              )}
            </div>
          );
        })}
      </div>
      {/* Label honesty: this button calls `onDuplicate(activeIdx)`, which
          routes through duplicateSlide() at pptx-adapter.ts:199-223 and
          performs a literal XML clone of the active slide (content +
          shapes + rels), NOT a blank-slide insert. The previous label
          "+ 新增投影片" implied a fresh slide and surprised users who
          clicked it expecting an empty canvas — the tooltip said
          "（複製）" but tooltips need hover and the label is the first
          read. Rename so the visible label matches behavior, mirroring
          the per-row rail icon's "複製此投影片" string at line 3168.
          XlsxEditor's `+` (XlsxEditor.tsx:2746) really does add a blank
          sheet (handleAddSheet → addSheet), so the asymmetry was a
          genuine inconsistency, not a parallel pattern. */}
      <button
        type="button"
        onClick={() => onDuplicate(activeIdx)}
        className="text-xs px-2 py-2 border-t border-border hover:bg-secondary text-muted-foreground hover:text-foreground transition-colors"
        title="在目前投影片之後新增一張（複製目前投影片的內容）"
      >
        + 新增（複製目前頁）
      </button>
      {/* R146 — flat-index drift cluster correction (R136/R137/R140/R145
          paradigm). Two upstream comments (line 3173 and line 3213) each
          cited `複製此投影片` at DIFFERENT wrong line numbers (3084 and
          2742) — the file directly contradicted itself: `複製此投影片`
          appears exactly once in the JSX (now correctly cited as line
          3168). Same comment cluster also cited XlsxEditor.tsx:2674 for
          `刪除工作表` (actual 2696, +22) and XlsxEditor.tsx:2513 for the
          XlsxEditor `+` button (actual 2746, +233 after major insertions).
          Reverse cite from XlsxEditor.tsx:2740 to PptxEditor.tsx:3081 was
          equally stale (actual 3221, +140). All five cites updated to
          their post-shift positions. This anchor sits BELOW line 3221 so
          neither cited target in this file (3168 and 3221) is shifted by
          the anchor itself — same minimum-blast-radius placement R145
          used to avoid R142's same-round-self-drift loop. */}
      {/* R147 — extends R145/R146 to PresentationMode + SlideRail comment
          cluster. 14 stale line-number citations across 4 nested doc-
          comments (the showNotes panel rationale at L2789-2792, the prev/
          next-button comment at L2814-2832, the R101 boundary-aware-
          tooltip block at L2833-2848, the right-click-menu rationale at
          L2942-2948, and the slide-row title-button hint comment at
          L3085-3092) all drifted +253-310 lines after upstream insertions
          pushed the presentation-mode keystroke handlers and the SlideRail
          JSX downward. Smoking-gun: `Ctrl+G 2366` for the cited PageDown
          handler landed on completely unrelated content — the comment was
          actively misleading, not just imprecise. All 14 cites updated to
          their actual post-shift positions (PageDown handler L2648, Home
          L2668, End L2676, status line L2722, showNotes-button label
          L2885, draggable L3029, onContextMenu L3062, drag-state doc-
          comment L2937-2939, rail icons L3155/3162/3193, ctxMenu items
          L3306/3319/3333, active-row icons render gate L3119-3198). This
          anchor sits BELOW line 3221 (same R145/R146 placement) so none
          of the cited targets in this file are shifted by the anchor —
          all targets sit either above the anchor or in the JSX tree this
          anchor is a sibling of. Pure doc-comment fix; no JSX or behavior
          change. */}
      {ctxMenu && (() => {
        // Same bound-estimate trick as XlsxEditor's sheet-tab menu — clamp at
        // render time so a right-click near the viewport edge doesn't push
        // the menu off-screen, no flicker from measure-and-adjust.
        // 3 items + 1 separator at ~28px ≈ 112px tall; min-w 160px ≈ 180px.
        const pos = clampToViewport(ctxMenu.x, ctxMenu.y, 180, 140);
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
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                onDuplicate(ctxMenu.idx);
                setCtxMenu(null);
              }}
              className="w-full text-left px-3 py-1.5 hover:bg-accent"
            >
              複製此投影片
            </button>
            {/* Boundary-aware tooltips + visible-label parity mirror the
                sibling rail-icon buttons ~140 lines up in this same
                component:
                  line 3155: title={i === 0 ? '已經是第一張' : '上移此投影片'}
                  line 3162: title={i === slideCount - 1 ? '已經是最後一張' : '下移此投影片'}
                  line 3168: title="複製此投影片"
                  line 3193: title={slideCount <= 1 ? '至少要保留一張投影片' : '刪除此投影片'}
                The rail-icon design comment at lines 3120-3138 already
                articulates the principle this menu was violating: "any
                button that disables at a boundary should explain why
                instead of leaving the user staring at a greyed icon."
                Both surfaces drive the SAME `onMoveUp / onMoveDown /
                onDelete / onDuplicate` callbacks with the SAME disable
                conditions, so a user who right-clicks a slide at the
                deck floor and sees a greyed "刪除投影片" with no hint
                has the exact problem the rail-icon comment was written
                to prevent — the menu items just hadn't gotten the same
                treatment. The visible labels were the second half of
                the same parity break: the rail's four entries all read
                `動詞 + 此 + 投影片` (上移此投影片 / 下移此投影片 /
                複製此投影片 / 刪除此投影片) but this menu had only
                `複製此投影片` matching, with `上移` / `下移` bare and
                `刪除投影片` missing the `此` — three of the menu's four
                rows broke the shape that the rail had aligned to in a
                prior round. Aligned verbatim so right-click and rail
                read the same string for the same action.
                Tooltip stays conditional-on-disabled (`: undefined`)
                rather than the rail's `: '上移此投影片'`/`: '刪除此投
                影片'`, because the rail icons are icon-only and need
                the active-state tooltip as a label, but here the menu
                item's visible text IS the label — emitting an active-
                state tooltip with the same wording would just stack a
                redundant browser tooltip on top of the visible button.
                Boundary messages are reused verbatim from the rail
                icons so identical states read the same string across
                both surfaces. */}
            <button
              type="button"
              role="menuitem"
              disabled={ctxMenu.idx === 0}
              title={ctxMenu.idx === 0 ? '已經是第一張' : undefined}
              onClick={() => {
                onMoveUp(ctxMenu.idx);
                setCtxMenu(null);
              }}
              className="w-full text-left px-3 py-1.5 hover:bg-accent disabled:opacity-40 disabled:pointer-events-none"
            >
              上移此投影片
            </button>
            <button
              type="button"
              role="menuitem"
              disabled={ctxMenu.idx === slideCount - 1}
              title={ctxMenu.idx === slideCount - 1 ? '已經是最後一張' : undefined}
              onClick={() => {
                onMoveDown(ctxMenu.idx);
                setCtxMenu(null);
              }}
              className="w-full text-left px-3 py-1.5 hover:bg-accent disabled:opacity-40 disabled:pointer-events-none"
            >
              下移此投影片
            </button>
            <div className="my-1 h-px bg-border" />
            <button
              type="button"
              role="menuitem"
              disabled={slideCount <= 1}
              title={slideCount <= 1 ? '至少要保留一張投影片' : undefined}
              onClick={() => {
                onDelete(ctxMenu.idx);
                setCtxMenu(null);
              }}
              className="w-full text-left px-3 py-1.5 text-destructive hover:bg-destructive/10 disabled:opacity-40 disabled:pointer-events-none"
            >
              刪除此投影片
            </button>
          </div>
        );
      })()}
    </div>
  );
}

/**
 * Slide-title outline pane (Round 76). Mirrors DocxEditor's Navigation Pane:
 * a flat list of slide titles, click-to-jump, current slide auto-highlighted
 * and auto-scrolled into view. The titles are derived heuristically (topmost
 * shape with text) since we don't currently parse `<p:ph type="title">` —
 * the heuristic is robust for templated decks and degrades to "(無標題)"
 * for image-only / hand-rolled slides.
 */
function PptxNavPanel({
  tabId,
  outline,
  activeIdx,
  onJump,
}: {
  tabId: string;
  outline: Array<{ idx: number; title: string }>;
  activeIdx: number;
  onJump: (idx: number) => void;
}): JSX.Element {
  // The `<aside>` is the actual scroll container (overflow-auto); the
  // `<ul>` inside has no overflow style. We need a separate ref on the
  // aside to read/write its scrollTop — `listRef` would give us the
  // inner list, whose scrollTop is always 0.
  const asideRef = useRef<HTMLElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  // Auto-scroll the active slide into view in the panel itself when the
  // selection changes. `block: 'nearest'` avoids jiggling when already visible.
  useEffect(() => {
    const el = listRef.current?.querySelector<HTMLElement>(
      `[data-pptx-nav-idx="${activeIdx}"]`,
    );
    el?.scrollIntoView({ block: 'nearest' });
  }, [activeIdx]);
  // Outline-pane scroll memory across tab swaps. Restore via
  // `queueMicrotask` so we run AFTER the activeIdx effect above (which
  // declares first and synchronously scrollIntoViews the active row) —
  // microtask flush comes after the current render's effect chain, so
  // the remembered offset wins. Without this ordering trick, mounting on
  // a tab whose active slide is near the top would snap the pane to the
  // top, even though the user had scrolled to read slides 30-40 before
  // tab-swapping away. Same pattern as DocxNavPanel (Round 35).
  useEffect(() => {
    const remembered = navScrollMemory.get(tabId);
    if (remembered != null) {
      queueMicrotask(() => {
        const el = asideRef.current;
        if (el) el.scrollTop = remembered;
      });
    }
    return () => {
      const el = asideRef.current;
      if (el) navScrollMemory.set(tabId, el.scrollTop);
    };
  }, [tabId]);

  // Arrow-key navigation across the outline list. Mirrors DocxNavPanel
  // (line 1720-1739) and MarkdownEditor's OutlinePanel — both support
  // ↑ / ↓ to cycle entries and Home / End to jump to boundaries when the
  // outline pane has focus. Without this, users with 50+ slides had to
  // click each row individually; PowerPoint was the lone outlier among
  // the three outline panes. We focus the row's <button> directly so the
  // browser's native focus-ring shows the cursor and the existing
  // onClick/onKeyDown(Enter) commit path stays in charge of "actually
  // jump to that slide" — arrow keys browse, Enter / Space jumps.
  const focusEntry = (i: number) => {
    const el = listRef.current?.querySelector<HTMLButtonElement>(
      `[data-pptx-nav-idx="${i}"]`,
    );
    el?.focus();
  };
  const onKeyNav = (e: React.KeyboardEvent<HTMLElement>) => {
    if (outline.length === 0) return;
    const focused = document.activeElement as HTMLElement | null;
    // Walk back from focused button to its outline.idx via data-attr —
    // outline indices are sparse for image-only slides, so we can't infer
    // position from an array offset.
    const focusedIdx = focused?.dataset?.pptxNavIdx
      ? Number(focused.dataset.pptxNavIdx)
      : activeIdx;
    const order = outline.map((o) => o.idx);
    const pos = order.indexOf(focusedIdx);
    const safePos = pos === -1 ? Math.max(0, order.indexOf(activeIdx)) : pos;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusEntry(order[Math.min(order.length - 1, safePos + 1)]);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusEntry(order[Math.max(0, safePos - 1)]);
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusEntry(order[0]);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusEntry(order[order.length - 1]);
    }
  };

  return (
    <aside
      ref={asideRef}
      onKeyDown={onKeyNav}
      className="w-60 shrink-0 border-r bg-secondary/20 overflow-auto text-xs"
    >
      {/* Surface the keyboard nav (onKeyNav at line 2965-2989: ↑/↓ scan,
          Home/End jump to first/last) — feature was added to "bring it to
          parity with the other two outline panes" per the comment at
          line 2950-2958, but parity at the feature level isn't parity at
          the discoverability level. None of the three nav panels (this
          one, DocxNavPanel, MarkdownEditor's OutlinePanel) ever told users
          the arrow keys work; the panel-toggle tooltips at PptxEditor:1137
          / DocxEditor:1960 / MarkdownToolbar:234 advertise show/hide but
          not how to navigate once shown. Same pattern as the R47 GoToDialog
          Enter advertisement and R52 TabBar double-click-rename hint:
          fully-implemented gesture with zero user-facing surface. Fix
          mirrors across all three panels in this round so cross-format
          consistency is preserved. */}
      <div
        className="px-3 py-2 text-[10px] uppercase tracking-wider text-muted-foreground font-medium border-b"
        title="↑/↓ 切換 · Home/End 跳到首/末"
      >
        投影片大綱
      </div>
      {outline.length === 0 ? (
        <div className="px-3 py-2 text-[11px] text-muted-foreground italic">
          （沒有投影片）
        </div>
      ) : (
        <ul ref={listRef} className="py-1">
          {outline.map((e) => {
            const isActive = e.idx === activeIdx;
            return (
              <li key={e.idx}>
                <button
                  type="button"
                  data-pptx-nav-idx={e.idx}
                  onClick={() => onJump(e.idx)}
                  className={cn(
                    'w-full text-left px-2 py-1 truncate flex items-center gap-1.5 transition-colors',
                    isActive
                      ? 'bg-primary/15 text-foreground font-medium'
                      : 'hover:bg-secondary/80',
                  )}
                  title={e.title}
                >
                  <span
                    className={cn(
                      'text-[9px] font-mono shrink-0 w-6 text-right',
                      isActive ? 'text-primary' : 'text-muted-foreground/70',
                    )}
                  >
                    {e.idx + 1}
                  </span>
                  <span className="truncate">{e.title}</span>
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </aside>
  );
}

function RailIconBtn({
  children,
  title,
  disabled,
  onClick,
}: {
  children: React.ReactNode;
  title: string;
  disabled?: boolean;
  onClick: () => void;
}): JSX.Element {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onClick={(e) => {
        e.stopPropagation();
        if (!disabled) onClick();
      }}
      className={cn(
        'p-1 rounded hover:bg-black/10',
        disabled && 'opacity-30 cursor-not-allowed',
      )}
    >
      {children}
    </button>
  );
}
