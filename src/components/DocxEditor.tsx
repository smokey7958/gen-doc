/**
 * Word editor — paragraph-level structured editor with selection-aware
 * inline formatting.
 *
 * Path B MVP UX:
 *   - Each block is a contentEditable RichBlock tagged with its kind
 *     (paragraph / headingN / bullet). Clicking the kind tag cycles types.
 *   - Bold / italic / underline apply to the *current selection* when the
 *     user has highlighted text (Phase F-2), and fall back to block-level
 *     toggling for caret-only state. Color / font / size still live at
 *     block level — mammoth doesn't surface them per-run.
 *   - Yellow banner spells out the remaining round-trip caveats.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignVerticalSpaceAround,
  Bold,
  Eye,
  Code2,
  FileCode,
  FileText,
  GripVertical,
  Highlighter,
  Image as ImageIcon,
  Italic,
  Link as LinkIcon,
  List,
  Palette,
  Pilcrow,
  Plus,
  Rows,
  Columns,
  Search,
  SeparatorHorizontal,
  Strikethrough,
  Trash2,
  Type,
  Underline,
} from 'lucide-react';
import type { DocxTab } from '../types/tab';
import { useWorkspace } from '../store/workspace';
import {
  type DocxAlign,
  type DocxBlock,
  type DocxBlockKind,
  type DocxBlockStyle,
  type DocxHighlightColor,
  type DocxImage,
  type DocxModel,
  type DocxPageSize,
  type DocxPageMargins,
  type DocxRun,
  parseDocx,
  serializeDocx,
} from '../lib/docx-adapter';
import { cn } from '../lib/utils';
import { notify } from '../store/toast';
import { FONT_FAMILIES, withEmojiFallback } from '../lib/font-families';
import { useFormatShortcuts } from '../lib/use-format-shortcuts';
import { useUndoableState, useUndoShortcuts } from '../lib/use-undoable-state';
import { registerEditorFlush } from '../lib/editor-flush';
import { applyStyleToRange, getCharRange } from '../lib/rich-text';
import { markdownToDocxBlocks } from '../lib/markdown-to-docx';
import { htmlToDocxBlocks } from '../lib/html-to-docx';
import { RichBlock } from './RichBlock';
import { FindReplaceDialog, type SearchSegment } from './FindReplaceDialog';
import { GoToDialog } from './GoToDialog';
import { useT, tImp } from '../lib/i18n';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from './ui/dropdown';

/** Page-size presets in twips (= twentieths of a point). */
const PAGE_PRESETS: { id: string; label: string; w: number; h: number }[] = [
  { id: 'a4', label: 'A4 (210×297mm)', w: 11906, h: 16838 },
  { id: 'letter', label: 'Letter (8.5×11")', w: 12240, h: 15840 },
  { id: 'legal', label: 'Legal (8.5×14")', w: 12240, h: 20160 },
  { id: 'a3', label: 'A3 (297×420mm)', w: 16838, h: 23811 },
  { id: 'a5', label: 'A5 (148×210mm)', w: 8390, h: 11906 },
  { id: 'b5', label: 'B5 (176×250mm)', w: 9978, h: 14173 },
];

/** Margin presets — also in twips. 1440 twips = 1 inch. */
const MARGIN_PRESETS: { id: string; labelZh: string; labelEn: string; v: number }[] = [
  { id: 'narrow', labelZh: '窄 (0.5")', labelEn: 'Narrow (0.5")', v: 720 },
  { id: 'normal', labelZh: '標準 (1")', labelEn: 'Normal (1")', v: 1440 },
  { id: 'wide', labelZh: '寬 (1.5")', labelEn: 'Wide (1.5")', v: 2160 },
];

interface Props {
  tab: DocxTab;
}

const KIND_CYCLE: DocxBlockKind[] = [
  'paragraph',
  'heading1',
  'heading2',
  'heading3',
  'bullet',
  'numbered',
];

const KIND_LABEL_ZH: Record<DocxBlockKind, string> = {
  paragraph: '段落',
  heading1: 'H1',
  heading2: 'H2',
  heading3: 'H3',
  heading4: 'H4',
  heading5: 'H5',
  heading6: 'H6',
  bullet: '• 列表',
  numbered: '1. 列表',
  table: '表格',
  image: '圖片',
};
const KIND_LABEL_EN: Record<DocxBlockKind, string> = {
  paragraph: 'Paragraph',
  heading1: 'H1',
  heading2: 'H2',
  heading3: 'H3',
  heading4: 'H4',
  heading5: 'H5',
  heading6: 'H6',
  bullet: '• List',
  numbered: '1. List',
  table: 'Table',
  image: 'Image',
};
function kindLabel(k: DocxBlockKind): string {
  return tImp(KIND_LABEL_ZH[k], KIND_LABEL_EN[k]);
}
// Backwards-compatible alias for the old name — same proxy-style lookup that
// reads the current locale on each access. Kept so the many existing
// `KIND_LABEL[xxx]` callsites in this file continue to work without per-site
// rewrites.
const KIND_LABEL: Record<DocxBlockKind, string> = new Proxy({} as Record<DocxBlockKind, string>, {
  get(_, p: string) {
    return kindLabel(p as DocxBlockKind);
  },
});

/** Style-dropdown entries. Heading 4-6 stay reachable via parse / existing
 *  mechanisms but are omitted here to keep the menu tight. */
const STYLE_OPTIONS: { kind: DocxBlockKind; zh: string; en: string }[] = [
  { kind: 'paragraph', zh: '內文', en: 'Body' },
  { kind: 'heading1', zh: '標題 1', en: 'Heading 1' },
  { kind: 'heading2', zh: '標題 2', en: 'Heading 2' },
  { kind: 'heading3', zh: '標題 3', en: 'Heading 3' },
  { kind: 'bullet', zh: '項目符號', en: 'Bulleted list' },
  { kind: 'numbered', zh: '編號清單', en: 'Numbered list' },
];

/** The six named docx highlight colors the palette offers + their CSS render. */
const HIGHLIGHT_COLORS: { id: DocxHighlightColor; zh: string; en: string; css: string }[] = [
  { id: 'yellow', zh: '黃色', en: 'Yellow', css: '#FFFF00' },
  { id: 'green', zh: '綠色', en: 'Green', css: '#00FF00' },
  { id: 'cyan', zh: '青色', en: 'Cyan', css: '#00FFFF' },
  { id: 'magenta', zh: '洋紅', en: 'Magenta', css: '#FF00FF' },
  { id: 'red', zh: '紅色', en: 'Red', css: '#FF0000' },
  { id: 'darkYellow', zh: '深黃', en: 'Dark yellow', css: '#808000' },
];
const HIGHLIGHT_CSS: Record<string, string> = Object.fromEntries(
  HIGHLIGHT_COLORS.map((h) => [h.id, h.css]),
);

/** Line-spacing multipliers (docx `spacing.line` = 240 × multiplier). */
const LINE_SPACING_OPTIONS: { v: number; zh: string; en: string }[] = [
  { v: 1, zh: '單行', en: 'Single' },
  { v: 1.15, zh: '1.15', en: '1.15' },
  { v: 1.5, zh: '1.5 倍', en: '1.5×' },
  { v: 2, zh: '2 倍', en: 'Double' },
];

let nextLocalId = 0;
/**
 * Per-tab scroll memory for the desk surface across remounts within the
 * renderer-process session. Mirrors the same-shaped Maps in MarkdownEditor
 * (`scrollMemory`) and PptxEditor (`slideMemory`): `EditorSurface` keys
 * its `<ErrorBoundary key={active.id}>` on the active tab id, so every tab
 * switch unmounts the editor subtree and the next mount renders the paper
 * with `scrollTop=0`. A user reading page 8 of a long Word doc, glancing at
 * another tab, and switching back lands at page 1. Map persists for the
 * renderer-process lifetime; cleared on app reload.
 */
const scrollMemory = new Map<string, number>();
/**
 * Per-tab outline-pane scroll memory, parallel to `scrollMemory` above
 * but for the Navigation Pane on the left. Without this, the user could
 * scroll the outline of a 50-heading doc to read sections H30-H40,
 * switch tabs, switch back — and the pane resets to the top while the
 * desk surface itself remembered exactly where the page was. The
 * existing `scrollIntoView({ block: 'nearest' })` on activeIdx only
 * tracks the *active* heading, so it can't carry over a position the
 * user picked manually for context-reading. Same lifetime as
 * `scrollMemory` (renderer-process; cleared on reload).
 */
const navScrollMemory = new Map<string, number>();

function genId(): string {
  nextLocalId += 1;
  return `local-${nextLocalId}`;
}

export function DocxEditor({ tab }: Props): JSX.Element {
  const t = useT();
  const patchTab = useWorkspace((s) => s.patchTab);
  const markTabDirty = useWorkspace((s) => s.markTabDirty);
  const [model, setModel, undoApi] = useUndoableState<DocxModel | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [activeBlockId, setActiveBlockId] = useState<string | null>(null);
  // OS-level drag-and-drop of an image file from File Explorer (Round 73).
  // `dragDepthRef` counters dragenter/dragleave to keep the indicator visible
  // while the cursor crosses internal element boundaries (toolbar, ruler,
  // block borders) — a single boolean flag would flicker every transition.
  const [draggingFile, setDraggingFile] = useState(false);
  const dragDepthRef = useRef(0);
  /**
   * Preview/print-preview mode — hides every editing affordance (banner,
   * toolbar, navigation pane, per-block grip / kind-tag / +/- chrome,
   * active-block ring) so only the page sheet + text content remain.
   * Lets users see "what would actually print" without leaving the editor.
   * Esc exits; the floating exit pill in the corner is the mouse path.
   * State stays per-tab (component instance) — switching tabs naturally
   * resets it, since each docx tab mounts its own DocxEditor.
   */
  const [previewMode, setPreviewMode] = useState(false);
  useEffect(() => {
    if (!previewMode) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        setPreviewMode(false);
      }
    };
    // Capture phase so we beat any nested editor / dialog Esc handlers —
    // the user pressed Esc *to leave preview*, and nothing else should
    // intercept it. Find/Goto dialogs aren't reachable in preview anyway
    // (their open buttons live on the hidden toolbar).
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [previewMode]);
  /**
   * Document navigation pane (Word's "View → Navigation Pane" / Adobe
   * Acrobat's "Bookmarks" equivalent) — left-side outline of every heading
   * in the document. Click a heading to jump and focus that block. We
   * persist the open/closed flag the same way MarkdownEditor does so the
   * user's preference survives tab switches and restarts. Default closed
   * because most documents start with no headings; opening it on a
   * heading-less doc would just show "(沒有標題)" which is friction.
   */
  const [navOpen, setNavOpen] = useState<boolean>(() => {
    try {
      return localStorage.getItem('gendoc.docxNavOpen') === '1';
    } catch {
      return false;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem('gendoc.docxNavOpen', navOpen ? '1' : '0');
    } catch {
      /* private mode / quota — fine to lose */
    }
  }, [navOpen]);
  /**
   * Per-block free-position state. Coordinates are in twips relative to the
   * page sheet's top-left (including page margins), so they survive scale
   * changes from the responsive ResizeObserver. `wTwip` freezes the block's
   * width at the moment it became free-positioned, otherwise an absolutely-
   * positioned block would collapse to its content width and visually jump.
   * NOTE: positions live only in editor state — they are NOT serialised to
   * docx (option C: pure visual free-positioning).
   */
  const [positions, setPositions] = useState<
    Record<string, { xTwip: number; yTwip: number; wTwip: number }>
  >({});
  /**
   * Marquee multi-selection — populated by drag-rectangle on the page
   * background, consumed by `startBlockMove` to translate every selected
   * block by the same delta. Cleared by Esc, by a single-block drag, by a
   * non-selected block focus, or by a fresh marquee that hits nothing.
   * Pure editor state; never serialised.
   */
  const [selectedBlockIds, setSelectedBlockIds] = useState<Set<string>>(
    () => new Set(),
  );
  /**
   * Live marquee rectangle in page-content-relative pixels. Non-null only
   * while the user is actively dragging the rubber-band. Rendered inside
   * `PageSheet` children as an absolute-positioned overlay so it shares
   * the same coordinate space as the blocks it's selecting.
   */
  const [marquee, setMarquee] = useState<
    { left: number; top: number; width: number; height: number } | null
  >(null);
  const [pxPerTwip, setPxPerTwip] = useState(1 / 15);
  const serializeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  // Latest model awaiting flush. Read by the registered flush callback so
  // a save issued mid-debounce sees the freshest edits regardless of
  // closure capture timing.
  const pendingModelRef = useRef<DocxModel | null>(null);
  // Ref onto the scrollable desk surface (the zinc-200 backdrop around the
  // paper). Used by the parse effect below to capture scrollTop on
  // unmount and restore it on the next mount of the same tab id — see
  // `scrollMemory` at the top of the file.
  const deskScrollRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);
    // R257 — capture parse gen, see parseGenRef doc-block.
    const myGen = ++parseGenRef.current;
    parseDocx(tab.data)
      .then((m) => {
        if (cancelled) return;
        // R260 — clear loading BEFORE the R257 gen check. If R254's parse
        // (started after AI Apply during initial load) completed first
        // and won the resetHistory race via gen bump, our OLD parse
        // arrives here with stale gen and the R257 check returns. But
        // setLoading(false) hasn't run yet; the gen-drop would leave
        // the editor stuck on the「正在解析 docx…」 placeholder forever
        // (R254's own parse-then doesn't touch loading because it's
        // designed for after-initial-load external changes). loading
        // is a UI state about「any parse has finished」 — true regardless
        // of which parse won the resetHistory contest. Clearing it
        // unconditionally on either parse-then path closes the stuck-
        // placeholder bug. The gen check still owns the model commit
        // below.
        setLoading(false);
        // R257 — newer parse won; drop our model commit.
        if (myGen !== parseGenRef.current) return;
        // Initial load: bypass the undo stack so the user can't "undo back to
        // null"; subsequent edits flow through `setModel` which records.
        // R267 — arm the skip flag before resetHistory so the [model] auto-
        // serialize effect doesn't re-serialize freshly-parsed bytes back
        // onto tab.data (which would flip dirty + diverge disk bytes).
        skipNextScheduleRef.current = true;
        undoApi.resetHistory(m);
        // Pre-select the first block of an empty document so the user can
        // type immediately without clicking. RichBlock auto-focuses when it
        // becomes `active`, which sidesteps the Chromium quirk where a
        // freshly-mounted contentEditable with only `<br>` doesn't bind a
        // usable caret on the first click. Non-empty docs leave activeBlockId
        // null until the user clicks (existing behavior).
        const isFreshEmptyDoc =
          m.blocks.length === 1 &&
          m.blocks[0].kind === 'paragraph' &&
          (!m.blocks[0].text || m.blocks[0].text.length === 0) &&
          (!m.blocks[0].runs || m.blocks[0].runs.length === 0);
        if (isFreshEmptyDoc) setActiveBlockId(m.blocks[0].id);
        // Restore desk scroll on the next animation frame — `setLoading(false)`
        // schedules a re-render that mounts the paper, and only after that
        // commit does `scrollHeight` include any content. Setting scrollTop
        // synchronously here would race the paper's first layout pass and
        // be silently clamped to 0. `cancelled` covers the unmount-during-rAF
        // race; a faster tab swap aborts before we touch the dead DOM.
        const remembered = scrollMemory.get(tab.id);
        if (typeof remembered === 'number' && remembered > 0) {
          requestAnimationFrame(() => {
            if (cancelled) return;
            const el = deskScrollRef.current;
            if (el) el.scrollTop = remembered;
          });
        }
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        // R272 — match the then arm's R260 + R257 ordering: setLoading(false)
        // is unconditional (loading is "any parse finished", true regardless
        // of who wins the gen race), gen check then gates setError so a
        // stale parse's throw doesn't overwrite a newer parse's clean
        // resetHistory with a misleading「parse error」 banner. Mirrors
        // R270's catch-arm gen guard in flushPendingSerialize.
        setLoading(false);
        if (myGen !== parseGenRef.current) return;
        setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
      // Capture before tear-down. The ref is still pointing at the live DOM
      // node here — React runs effect cleanups before committing the actual
      // DOM removal — so reading `scrollTop` is safe.
      const el = deskScrollRef.current;
      if (el) scrollMemory.set(tab.id, el.scrollTop);
    };
  }, [tab.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const scheduleSerialize = (m: DocxModel) => {
    pendingModelRef.current = m;
    // Mark dirty *now* so a Ctrl+W within the 400ms debounce window still
    // triggers the unsaved-changes prompt — without this the close path
    // sees stale tab.dirty=false and discards pending edits silently.
    markTabDirty(tab.id);
    if (serializeTimer.current) clearTimeout(serializeTimer.current);
    serializeTimer.current = setTimeout(() => {
      void flushPendingSerialize();
    }, 400);
  };

  /**
   * Monotonic generation counter for in-flight `serializeDocx` calls. Each
   * flush increments and captures its own generation; on completion the
   * captured generation is compared against the latest, and a stale result
   * is dropped on the floor instead of overwriting a newer one. Without
   * this, the following race lost user data:
   *   t=0    User types M1, debounce 400ms.
   *   t=400  Flush A starts, await serializeDocx(M1)…  (large doc, 1500ms)
   *   t=500  User types M2, new debounce 400ms.
   *   t=900  Flush B starts, await serializeDocx(M2)…  (small delta, 100ms)
   *   t=1000 B completes → patchTab(M2 bytes).         ← latest data lands
   *   t=1900 A completes → patchTab(M1 bytes).         ← STALE overwrite!
   * `Packer.toBlob` (docx library) genuinely takes seconds on a long
   * document; the fast typist's M2 keystroke during that await reproduces
   * the race deterministically. The generation guard at the post-await
   * commit point makes the older result a no-op.
   */
  const flushGenRef = useRef(0);

  // R254 — track the bytes the editor last wrote so the external-change
  // re-parse effect below recognises self-induced tab.data updates and
  // skips re-parsing. Same shape as R253's XlsxEditor fix; see effect
  // doc-block at the bottom of this section. Initialize from the FIRST
  // `tab.data` so the re-parse effect skips on initial mount (the
  // useEffect[tab.id] above already handles the first parse with full
  // setup — empty-doc focus, scroll restore, loading flag — that R254's
  // minimal re-parse effect doesn't replicate). useRef's initializer is
  // only consulted on first render; tab swap remounts the editor and
  // re-seeds the ref with the new tab's bytes.
  const lastWrittenBytesRef = useRef<Uint8Array>(tab.data);

  // R257 — monotonic parse-generation counter, see PptxEditor sibling
  // for full doc-block. parseDocx is async (docx library is Promise-
  // based for Packer.toBlob and parse); the [tab.id] effect's parse
  // window can race R254's tab.data-driven re-parse when AI Apply
  // fires during initial load. Newer parse wins via gen check.
  const parseGenRef = useRef(0);

  // R267 — gate the [model] auto-serialize effect (line ~836) against
  // resetHistory-induced model swaps. The auto-serialize effect was
  // designed to push undo/redo's model swap back to tab.data via
  // scheduleSerialize, with `if (loading) return;` intended to skip the
  // initial-parse path. But the [tab.id] effect commits setLoading(false)
  // AND undoApi.resetHistory(m) in the SAME React-18-batched microtask,
  // so by the time [model] runs, loading is already false — the gate is
  // open, scheduleSerialize fires for the freshly-parsed model, markTabDirty
  // flips dirty=true without any user edit, and 400ms later patchTab
  // writes re-serialized bytes back over the original disk bytes (byte-
  // different even when functionally equivalent). R254's external re-parse
  // path (AI Apply) hits the same shape: resetHistory(AI_model) → auto-
  // serialize → patchTab overwrites the AI's authoritative bytes with
  // re-serialized ones. Set this ref true immediately before each
  // resetHistory; the [model] effect consumes it once. User edits
  // (updateBlock / insertAfter / table ops / scheduleSerialize inline
  // calls) don't touch this ref so they schedule normally; undo / redo
  // go through undoApi.undo/redo (setState directly, no resetHistory)
  // so they also schedule normally — only the bytes→model sync paths
  // are skipped, which is precisely the intended invariant.
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
      const bytes = await serializeDocx(m);
      // Race guard — see flushGenRef doc-block. Drop stale result if a
      // newer flush has started while we awaited.
      if (myGen !== flushGenRef.current) return;
      // R254 — record the bytes we're about to patch so the re-parse
      // effect below recognises this as self-induced. Set BEFORE patchTab
      // (Zustand setState fires the React subscription synchronously).
      lastWrittenBytesRef.current = bytes;
      patchTab(tab.id, { data: bytes });
    } catch (err) {
      // R270 — gen-check the catch arm too. Without this, a stale flush's
      // serializer throw (e.g., transient docx-library issue mid-serialize
      // of an OLD model after the user typed more) overwrites a newer
      // flush's success with a false error banner. flushGenRef already
      // gates the SUCCESS path against post-newer-flush commits; symmetry
      // wants the failure path to follow suit so an obsoleted flush's
      // error is treated as noise and dropped. A real persistent
      // serialize bug surfaces on the NEXT flush at the latest gen, so
      // we don't silence anything actually important.
      if (myGen !== flushGenRef.current) return;
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  // R254 — re-parse tab.data when it changes externally (AI Apply on a
  // binary_replace op, undo / redo via App.tsx handleUndo / handleRedo).
  // Without this, the editor's local model stays at the OLD parse forever
  // (the parse effect at line 255 only fires on tab.id change), so:
  //   1. Display drift: AI replaces paragraph 3 → tab.data is NEW bytes,
  //      but RichBlock list renders from `model.blocks` (OLD) so the
  //      paragraph still shows its prior text. Looks like AI didn't apply.
  //   2. Silent revert (data loss): user types in any block → setModel →
  //      scheduleSerialize → flushPendingSerialize → serializeDocx(model_OLD
  //      with the user's typing on top) → bytes derived from OLD model →
  //      patchTab overwrites tab.data → AI's edit is GONE.
  // Same shape as R253's XlsxEditor fix; see that doc-block for the full
  // identity-equality rationale (writeBack passes the same Uint8Array
  // object to patchTab so reference-equality skips re-parse correctly;
  // applyChangeset / undoChangeset construct fresh Uint8Array so identity
  // diverges → re-parse triggered).
  //
  // Skip the resetHistory call when we're skipping the re-parse — the
  // local in-editor undo stack must NOT be wiped on every keystroke
  // round-trip. Only when an external mutation lands do we resetHistory
  // to anchor the new "AI-applied" state as the local-undo baseline;
  // workspace-level Ctrl+Z is the path to undo the AI Apply itself.
  useEffect(() => {
    if (tab.data === lastWrittenBytesRef.current) return;
    let cancelled = false;
    // R257 — bump generation; see parseGenRef doc-block.
    const myGen = ++parseGenRef.current;
    void parseDocx(tab.data)
      .then((m) => {
        if (cancelled) return;
        if (myGen !== parseGenRef.current) return; // newer parse won
        // R267 — same skip-arm as the [tab.id] initial parse path. AI Apply's
        // AI_BYTES are authoritative; we don't want the [model] auto-
        // serialize to overwrite them with the re-serialized output.
        skipNextScheduleRef.current = true;
        undoApi.resetHistory(m);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        // R272 — gen-check the catch arm. cancelled alone covers
        //「tab.data changed again before this parse finished」 (cleanup
        // re-fire sets it), but it does NOT cover the cross-effect race
        // where [tab.id]'s parse and this [tab.data] parse run in
        // parallel, the OTHER one wins the gen race, and ours throws
        // late. Without this guard, the loser's setError clobbers a
        // banner that should stay clean because the winner already
        // resolved the model. Same pattern as R270 / the [tab.id]
        // sibling.
        if (myGen !== parseGenRef.current) return;
        setError(err instanceof Error ? err.message : String(err));
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.data]);

  // Register with the global editor-flush registry so the save flow can
  // force any pending re-serialize before snapshotting tab bytes.
  useEffect(() => {
    return registerEditorFlush(flushPendingSerialize);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  // Force a final flush on unmount so tab close / workspace swap doesn't
  // strand an in-flight edit in the debounce timer.
  useEffect(() => {
    return () => {
      if (serializeTimer.current) {
        void flushPendingSerialize();
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const updateBlock = (id: string, patch: Partial<DocxBlock>) => {
    setModel((prev) => {
      if (!prev) return prev;
      const blocks = prev.blocks.map((b) => {
        if (b.id !== id) return b;
        const merged = { ...b, ...patch };
        // Two distinct mutation shapes for paragraph bodies:
        //   - { runs: [...] }  comes from RichBlock (Phase F-2). Runs are
        //     the source of truth; we derive `text` from them so callers
        //     reading block.text (search index, F&R) still see the visible
        //     content.
        //   - { text: '...' }  comes from F&R replacements and table cells.
        //     A bare text patch can't carry per-run styling, so we drop
        //     `runs` to avoid the serializer emitting per-run text that no
        //     longer matches the visible block.
        if (Object.prototype.hasOwnProperty.call(patch, 'runs')) {
          const next = patch.runs;
          merged.text = next ? next.map((r) => r.text).join('') : '';
        } else if (Object.prototype.hasOwnProperty.call(patch, 'text') && b.runs) {
          delete (merged as { runs?: unknown }).runs;
        }
        return merged;
      });
      const next = { ...prev, blocks };
      scheduleSerialize(next);
      return next;
    });
  };

  const insertAfter = (id: string, fresh?: DocxBlock) => {
    setModel((prev) => {
      if (!prev) return prev;
      const idx = prev.blocks.findIndex((b) => b.id === id);
      if (idx < 0) return prev;
      const block: DocxBlock = fresh ?? { id: genId(), kind: 'paragraph', text: '' };
      const blocks = [...prev.blocks.slice(0, idx + 1), block, ...prev.blocks.slice(idx + 1)];
      const next = { ...prev, blocks };
      scheduleSerialize(next);
      return next;
    });
  };

  /**
   * Splice many blocks after the given id. Used by the "插入 Markdown" path —
   * a single markdown paste typically yields several blocks (heading +
   * paragraph + list + table, etc.) and calling `insertAfter` in a loop would
   * either need a running anchor cursor or mis-order the blocks. Doing the
   * splice as one setModel keeps the operation as a single undo-history entry
   * — Ctrl+Z reverses the whole insert, not one paragraph at a time.
   *
   * If `id` is null OR no longer in the model (race with a delete), the
   * blocks are appended at the end. Empty `fresh` is a no-op.
   */
  const insertBlocksAfter = (id: string | null, fresh: DocxBlock[]) => {
    if (fresh.length === 0) return;
    setModel((prev) => {
      if (!prev) return prev;
      const idx = id ? prev.blocks.findIndex((b) => b.id === id) : -1;
      const insertAt = idx < 0 ? prev.blocks.length : idx + 1;
      const blocks = [
        ...prev.blocks.slice(0, insertAt),
        ...fresh,
        ...prev.blocks.slice(insertAt),
      ];
      const next = { ...prev, blocks };
      scheduleSerialize(next);
      return next;
    });
  };

  const insertTableAfter = (id: string) => {
    insertAfter(id, {
      id: genId(),
      kind: 'table',
      text: '',
      rows: [
        // R409 — i18n: seeded header cells become document content
        [tImp('欄位 1', 'Column 1'), tImp('欄位 2', 'Column 2'), tImp('欄位 3', 'Column 3')],
        ['', '', ''],
        ['', '', ''],
      ],
    });
  };

  /**
   * Pop a file picker, decode the chosen image's natural size, and append a
   * new `image` block after `id`. Caps the on-page width at the page's
   * usable width (page width minus left/right margins) so a large source
   * image doesn't overflow the printable area — preserves aspect ratio.
   *
   * Async: we await the natural-size decode before mutating the model so
   * the block's `widthPx` / `heightPx` are correct on the very first paint
   * (avoids the layout-jiggle that "size 0 → real size on next render"
   * would cause). Errors in image decode just no-op the insert; the user
   * can retry.
   */
  const insertImageBlockFromFile = async (file: File, afterId: string) => {
    let mediaType: DocxImage['mediaType'];
    if (file.type === 'image/jpeg') mediaType = 'jpg';
    else if (file.type === 'image/gif') mediaType = 'gif';
    else if (file.type === 'image/bmp') mediaType = 'bmp';
    else mediaType = 'png';
    // `file.arrayBuffer()` can reject when the OS revokes the File handle
    // (clipboard image whose source page navigated away, OS picker file
    // that was deleted/locked between selection and read, oversized blob
    // that hits an OOM). Without this guard the rejection propagated as
    // an unhandled promise from the three call sites (paste handler at
    // line 814, drag-drop handler at ~1336, toolbar button at ~1358),
    // each of which does `void insertImageBlockFromFile(...)`. The user
    // saw "nothing happened" with no explanation. PptxEditor's
    // `insertPictureFile` already does this (Round 8 fix); bringing
    // DocxEditor up to the same baseline. Note: `readNaturalSize` is
    // *not* the failure path — its `img.onerror` resolves with default
    // dimensions, which the surrounding comment around line 451 calls
    // out as intentional "decode-error → keep going with placeholder
    // size". We only need to catch the byte-read step.
    let buf: Uint8Array;
    try {
      buf = new Uint8Array(await file.arrayBuffer());
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(tImp(`插入圖片失敗：${msg}`, `Failed to insert image: ${msg}`), 'error');
      return;
    }
    const dataUrl = `data:${mimeForKind(mediaType)};base64,${uint8ToBase64(buf)}`;
    const natural = await readNaturalSize(dataUrl);
    // Cap at usable page width. PageSize is in twips (1/20 pt); pixels assume
    // 96 dpi (the convention `docx` uses). 1440 twips = 1 inch = 96 px.
    let widthPx = natural.width;
    let heightPx = natural.height;
    if (model) {
      const usableTwips = model.pageSize.w - model.pageMargins.left - model.pageMargins.right;
      const maxWidthPx = Math.max(96, Math.round((usableTwips / 1440) * 96));
      if (widthPx > maxWidthPx) {
        const scale = maxWidthPx / widthPx;
        widthPx = maxWidthPx;
        heightPx = Math.max(1, Math.round(heightPx * scale));
      }
    }
    insertAfter(afterId, {
      id: genId(),
      kind: 'image',
      text: '',
      align: 'center',
      image: { data: buf, mediaType, widthPx, heightPx, dataUrl },
    });
  };

  const insertImageAfter = async (id: string) => {
    const file = await pickImageFile();
    if (!file) return;
    await insertImageBlockFromFile(file, id);
  };

  const removeBlock = (id: string) => {
    setModel((prev) => {
      if (!prev) return prev;
      if (prev.blocks.length <= 1) {
        // Last block — empty it instead of silently no-op'ing. The user
        // clicked trash because they wanted this content gone; if we just
        // ignored the click they'd have no idea why nothing happened.
        // Replacing with a fresh empty paragraph satisfies "delete this"
        // while keeping the ≥1-block invariant the renderer assumes.
        const fresh: DocxBlock = { id: genId(), kind: 'paragraph', text: '' };
        const next = { ...prev, blocks: [fresh] };
        scheduleSerialize(next);
        return next;
      }
      const blocks = prev.blocks.filter((b) => b.id !== id);
      const next = { ...prev, blocks };
      scheduleSerialize(next);
      return next;
    });
    // Drop any free-position entry too — orphaned positions just bloat the
    // map and confuse a future reset-all action.
    setPositions((prev) => {
      if (!(id in prev)) return prev;
      const { [id]: _drop, ...rest } = prev;
      return rest;
    });
  };

  /**
   * Bulk variant of `removeBlock` for the marquee-selection delete path.
   * One `setModel` call → one undo entry that restores all blocks at once,
   * which is what users expect after pressing Delete on a multi-selection.
   * Positions are pruned in the same beat so a re-paste of the deleted
   * blocks doesn't inherit ghost coordinates from the trashed copies.
   */
  const removeBlocks = (ids: string[]) => {
    if (ids.length === 0) return;
    const idSet = new Set(ids);
    setModel((prev) => {
      if (!prev) return prev;
      const remaining = prev.blocks.filter((b) => !idSet.has(b.id));
      if (remaining.length === 0) {
        // Same ≥1-block invariant as `removeBlock`. Replace with a single
        // empty paragraph so the page sheet still has something to render.
        const fresh: DocxBlock = { id: genId(), kind: 'paragraph', text: '' };
        const next = { ...prev, blocks: [fresh] };
        scheduleSerialize(next);
        return next;
      }
      const next = { ...prev, blocks: remaining };
      scheduleSerialize(next);
      return next;
    });
    setPositions((prev) => {
      let changed = false;
      const next: typeof prev = { ...prev };
      for (const id of ids) {
        if (id in next) {
          delete next[id];
          changed = true;
        }
      }
      return changed ? next : prev;
    });
  };

  const cycleKind = (id: string) => {
    setModel((prev) => {
      if (!prev) return prev;
      const blocks = prev.blocks.map((b) => {
        if (b.id !== id) return b;
        const i = KIND_CYCLE.indexOf(b.kind);
        const nextKind = KIND_CYCLE[(i + 1) % KIND_CYCLE.length];
        return { ...b, kind: nextKind };
      });
      const next = { ...prev, blocks };
      scheduleSerialize(next);
      return next;
    });
  };

  // ── HOOKS ABOVE EARLY RETURNS ─────────────────────────────────────────
  // React requires identical hook count between renders. The first render
  // has `loading=true` / `model=null`, so any hook placed below the early
  // returns would be skipped on the first render and added on the second
  // — that's React error #310. All hooks (`useFormatShortcuts`,
  // `useUndoShortcuts`, `useEffect`, `useState(findOpen)`, `useMemo` for
  // findSegments) must run unconditionally on every render, even before
  // the model is parsed. They guard their bodies with `if (!model) return`
  // where needed.

  const activeBlock = model?.blocks.find((b) => b.id === activeBlockId) ?? null;

  /**
   * Document outline — the ordered list of heading blocks (h1–h6). We
   * recompute on every model change rather than caching with a useMemo
   * because (a) the dependency is the same `blocks` array we already
   * iterate elsewhere and (b) the work is O(n) over a typical doc of
   * hundreds of blocks — well below the threshold where memoisation pays
   * back its bookkeeping cost.
   */
  const outline = useMemo(() => {
    if (!model) return [] as Array<{ id: string; level: number; text: string; index: number }>;
    const items: Array<{ id: string; level: number; text: string; index: number }> = [];
    model.blocks.forEach((b, idx) => {
      const m = b.kind.match(/^heading([1-6])$/);
      if (m) {
        items.push({
          id: b.id,
          level: Number(m[1]),
          // Empty heading still gets an entry so the user sees structure;
          // we just label it "(無標題)" rather than rendering an empty row.
          text: b.text.trim() || tImp('(無標題)', '(Untitled)'),
          index: idx,
        });
      }
    });
    return items;
  }, [model]);

  /**
   * Which outline entry contains the caret? Find the heading whose source
   * position is the latest one ≤ activeBlock's position. If active is itself
   * a heading we get an exact match; otherwise we get the heading the user
   * is "under". -1 means caret is before any heading. Mirrors the same
   * pattern MarkdownEditor uses for OutlinePanel `activeIdx`.
   */
  const activeOutlineIdx = useMemo(() => {
    if (!model || outline.length === 0 || !activeBlockId) return -1;
    const activePos = model.blocks.findIndex((b) => b.id === activeBlockId);
    if (activePos < 0) return -1;
    let active = -1;
    for (let i = 0; i < outline.length; i += 1) {
      if (outline[i].index <= activePos) active = i;
      else break;
    }
    return active;
  }, [model, outline, activeBlockId]);

  /**
   * Jump to a heading by id: select it (so the FormatToolbar reflects it)
   * and scroll its DOM node into view. We rely on the existing
   * `data-block-id` attribute already wired up by RichBlock / TableBlock.
   * `block: 'start'` puts the heading at the top of the viewport which is
   * the standard "I clicked navigate, take me there" convention; `nearest`
   * would do nothing if the heading was already partly visible.
   */
  const jumpToHeading = (id: string) => {
    setActiveBlockId(id);
    // Defer to next frame so React commits setActiveBlockId before we read
    // the DOM (the active block triggers a className change that doesn't
    // affect layout, but keeps the scroll target predictable).
    requestAnimationFrame(() => {
      const el = document.querySelector(`[data-block-id="${id}"]`) as HTMLElement | null;
      el?.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
  };

  /**
   * B/I/U toggle. Two modes:
   *   - Selection mode: the user has highlighted text inside a RichBlock.
   *     We splice runs over the [start, end) char range and patch the
   *     block's runs[]. The block's `style` flag is left alone — it only
   *     governs blocks without a per-run breakdown.
   *   - Block mode: caret-only or focus elsewhere. Toggle the block-level
   *     style as before. If the block currently has a runs[] (carried from
   *     parse), we *also* apply the toggle to every run so the visual
   *     stays consistent with the new block style.
   */
  const toggleStyle = (key: 'bold' | 'italic' | 'underline' | 'strikethrough') => {
    if (!activeBlock) return;
    const richEl = document.querySelector(
      `[data-rich-block="1"][data-block-id="${activeBlock.id}"]`,
    ) as HTMLElement | null;
    if (richEl && richEl.contains(document.activeElement)) {
      const range = getCharRange(richEl);
      if (range && range.end > range.start) {
        const baseRuns =
          activeBlock.runs && activeBlock.runs.length > 0
            ? activeBlock.runs
            : [{ text: activeBlock.text, style: undefined }];
        const newRuns = applyStyleToRange(baseRuns, range.start, range.end, key);
        updateBlock(activeBlock.id, { runs: newRuns });
        return;
      }
    }
    // Block-level fallback (caret without selection, table cells, etc.)
    const cur = activeBlock.style ?? {};
    const next: DocxBlockStyle = { ...cur, [key]: !cur[key] };
    updateBlock(activeBlock.id, { style: normalizeStyle(next) });
  };

  // Ctrl/Cmd+B / +I / +U — mirrors the toolbar buttons. Scoped to this
  // editor's root so the shortcut does not fire when typing in the markdown
  // preview, the tab bar, or another file's editor.
  useFormatShortcuts({
    rootSelector: '[data-docx-editor-root]',
    isActive: () => !!activeBlock,
    toggle: (k) => toggleStyle(k),
  });

  // Ctrl/Cmd+Shift+X — strikethrough. Wired locally because the shared
  // useFormatShortcuts hook deliberately rejects Shift combos. Ref keeps the
  // listener stable while always seeing the latest toggleStyle closure.
  const toggleStyleRef = useRef(toggleStyle);
  toggleStyleRef.current = toggleStyle;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || !e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'x') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-docx-editor-root]')) return;
      e.preventDefault();
      toggleStyleRef.current('strikethrough');
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // Ctrl/Cmd+Z / +Y for editor-local undo/redo. After undo/redo we still need
  // to push the new model bytes through serializeDocx → patchTab so the file
  // on disk follows; scheduleSerialize is fire-and-forget so wrap each call.
  useUndoShortcuts({
    rootSelector: '[data-docx-editor-root]',
    undo: () => {
      undoApi.undo();
      // Read the post-undo model from React's setState callback path is
      // awkward; the next render will have it via `model`, and we trigger a
      // serialize on that render via the effect below.
    },
    redo: () => undoApi.redo(),
  });

  // Whenever `model` swaps via undo/redo (or any other path), push it to
  // disk. Skip while loading to avoid re-serializing the parse output.
  // R267 — `loading` alone is insufficient: setLoading(false) and
  // resetHistory's setState are React-18-batched into the same commit, so
  // by the time this effect runs after a parse, loading is already false.
  // skipNextScheduleRef is set true immediately before each resetHistory
  // (initial [tab.id] parse + R254 [tab.data] external re-parse) and
  // consumed here. See ref doc-block for the full bug shape.
  useEffect(() => {
    if (!model || loading) return;
    if (skipNextScheduleRef.current) {
      skipNextScheduleRef.current = false;
      return;
    }
    scheduleSerialize(model);
    // scheduleSerialize is stable enough — it only reads patchTab via closure.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [model]);

  // Ctrl/Cmd+F: open Find & Replace and (re)focus the query input. Scoped
  // (like the format shortcuts) to this editor's root so it doesn't fire
  // from other tabs. We deliberately don't toggle-close on Ctrl+F — Esc
  // closes, matching VS Code / Chrome / Word. `findFocusNonce` lets the
  // dialog re-focus its input on every Ctrl+F, even when it was already
  // open, so users who clicked into the document to navigate a match
  // can hit Ctrl+F to come back to the query field.
  const [findOpen, setFindOpen] = useState(false);
  const [findFocusNonce, setFindFocusNonce] = useState(0);
  // Inline link-edit dialog (Round 24) — replaces window.prompt for the
  // toolbar's link toggle. Snapshots the target block id at open time so
  // commits don't drift if the user clicks elsewhere while the dialog is up.
  const [linkDialog, setLinkDialog] = useState<{
    blockId: string;
    defaultUrl: string;
  } | null>(null);
  // Markdown-insert dialog — opened from the toolbar's "插入 Markdown" button.
  // `anchorBlockId` snapshots the active block at open time so a stray click
  // into the document while the user is typing in the dialog can't relocate
  // the insertion point (same anchor pattern as MarkdownEditor's link dialog).
  // null `anchorBlockId` falls back to "append at end".
  const [mdInsertDialog, setMdInsertDialog] = useState<{
    anchorBlockId: string | null;
  } | null>(null);
  // HTML-insert dialog — sibling of mdInsertDialog. Paste an HTML snippet,
  // get back proper Word blocks via html-to-docx.ts. Same anchor-snapshot
  // semantics so the splice point stays put while the user types.
  const [htmlInsertDialog, setHtmlInsertDialog] = useState<{
    anchorBlockId: string | null;
  } | null>(null);
  // R350 — Ctrl+G "Go to paragraph N" state declared up-front so the
  // exclusive-open helper below can reference both Find and Goto setters.
  // Original layout had `gotoOpen` declared between the two useEffects,
  // which worked because the cross-references resolve at callback-fire
  // time (post-mount), but interleaving makes the dialog-mutex logic
  // harder to follow.
  const [gotoOpen, setGotoOpen] = useState(false);
  const [gotoFocusNonce, setGotoFocusNonce] = useState(0);
  /**
   * R350 — exclusive open: close every dialog except the one named.
   *
   * Background: DocxEditor renders FIVE floating dialogs at
   * `absolute top-3 right-3` (linkDialog / mdInsertDialog /
   * htmlInsertDialog / FindReplaceDialog / GoToDialog). Each is gated
   * on its own piece of React state; nothing in the render tree
   * prevents two from being open at once. Before this round only the
   * Find/Goto pair had a mutex (Ctrl+F closed Goto, Ctrl+G closed
   * Find, and the toolbar openers mirrored that) — every other pair
   * (link + mdInsert, mdInsert + htmlInsert, link + find, etc.) could
   * stack:
   *   • User opens 插入 Markdown dialog
   *   • Without closing it, clicks 插入 HTML on the toolbar
   *   • Both dialogs render at the SAME absolute position, the more-
   *     recently mounted one on top (DOM source order at z-30); the
   *     hidden one steals keystrokes that bubble to its onKeyDown.
   *   • User hits Esc to dismiss → only the top one closes; the one
   *     underneath is now visible but had its focus stolen on mount,
   *     so Ctrl+Enter / 插入 button might not behave as expected.
   *
   * The selection-hint at line ~1751 already documents「they all anchor
   * to the same corner and stacking them on top of each other reads as
   * a layout glitch」 — it correctly hides itself when ANY of the five
   * is open, but the dialogs themselves never inherited the mutex. This
   * helper does — every dialog-open path calls
   * `closeOtherDialogs('xxx')` first, so opening one is atomically
   * "show this one, close the rest" the way Find/Goto have always
   * done bilaterally.
   *
   * Pattern matches editors with the same multi-dialog problem space
   * (VS Code's command palette / quickPick / inputBox all preempt each
   * other via the same "show implies dismiss others" rule).
   */
  const closeOtherDialogs = useCallback(
    (except: 'link' | 'md' | 'html' | 'find' | 'goto') => {
      if (except !== 'link') setLinkDialog(null);
      if (except !== 'md') setMdInsertDialog(null);
      if (except !== 'html') setHtmlInsertDialog(null);
      if (except !== 'find') setFindOpen(false);
      if (except !== 'goto') setGotoOpen(false);
    },
    [],
  );
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'f') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-docx-editor-root]')) return;
      e.preventDefault();
      // R350 — close every other dialog, not just Goto. Was the
      // original Find/Goto-only mutex from Round 79; extended here to
      // cover link + md-insert + html-insert too.
      closeOtherDialogs('find');
      setFindOpen(true);
      setFindFocusNonce((n) => n + 1);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [closeOtherDialogs]);

  // Ctrl+G — "Go to paragraph N" (Round 78). Word's native Ctrl+G goes to
  // page N, but we don't have a clear page-break abstraction (one continuous
  // scroll), so block index is the most natural N — also aligns with the
  // Navigation Pane (Round 76 reuse via heading order). Scoping mirrors the
  // Ctrl+F binding above so it doesn't fire from other tabs.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.shiftKey || e.altKey) return;
      if (e.key.toLowerCase() !== 'g') return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae?.closest?.('[data-docx-editor-root]')) return;
      e.preventDefault();
      // R350 — close every other dialog, not just Find.
      closeOtherDialogs('goto');
      setGotoOpen(true);
      setGotoFocusNonce((n) => n + 1);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [closeOtherDialogs]);

  // Ctrl+V image paste: when the clipboard carries an image file (e.g. user
  // copied a screenshot), insert it as a new image block after the active
  // block. Mirrors the Excel/XlsxEditor paste flow (Round 71). Scoped to the
  // editor root so we don't hijack pastes from other tabs or the markdown
  // preview. Skip when the paste target is a contenteditable / input — that
  // path may carry text the user actually wants pasted as text.
  useEffect(() => {
    const root = document.querySelector('[data-docx-editor-root]');
    if (!root) return;
    const onPaste = (ev: Event) => {
      const e = ev as ClipboardEvent;
      const items = e.clipboardData?.items;
      if (!items) return;
      // R91 — paste-side parallel to the R85 drop-side fix at the onDrop
      // handler near line 880. See MarkdownEditor.tsx (~line 600) for the
      // canonical doc-comment. Without the sawNonImageFile flag, pasting a
      // PDF / .docx / etc. silently no-ops; the same constraint already
      // surfaces a warning on drop, so paste should match.
      let sawNonImageFile = false;
      for (let i = 0; i < items.length; i += 1) {
        const it = items[i];
        if (it.kind === 'file' && it.type.startsWith('image/')) {
          const file = it.getAsFile();
          if (!file) continue;
          e.preventDefault();
          const target = activeBlockId ?? model?.blocks[model.blocks.length - 1]?.id;
          if (target) void insertImageBlockFromFile(file, target);
          return;
        }
        if (it.kind === 'file') sawNonImageFile = true;
      }
      if (sawNonImageFile) {
        e.preventDefault();
        notify(tImp('只能貼上圖片檔案', 'Only image files can be pasted'), 'warning');
      }
    };
    root.addEventListener('paste', onPaste);
    return () => root.removeEventListener('paste', onPaste);
    // insertImageBlockFromFile closes over `model`/`insertAfter` which are
    // stable through closures; we deliberately re-bind when the active block
    // changes so the paste lands at the right anchor.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeBlockId, model]);

  // Flatten searchable text across paragraphs/headings/lists AND table
  // cells. Table-cell segment ids use a composite "blockId:r:c" form so
  // `onUpdateSegment` can rebuild the right row entry.
  //
  // Labels follow the same locator-style convention as PPTX ("Slide N · Box N")
  // and XLSX ("Sheet1 · A1") — they include a 1-based block index plus a short
  // human-readable kind hint, instead of leaking raw discriminants like
  // "paragraph"/"heading2" into the FindReplace match counter.
  // Returns [] while model is still loading; the find dialog is still
  // mounted lazily by `findOpen`, so the empty-list state is harmless.
  const findSegments = useMemo<SearchSegment[]>(() => {
    if (!model) return [];
    const out: SearchSegment[] = [];
    const kindLabel = (k: DocxBlockKind): string => {
      if (k === 'paragraph') return 'Para';
      if (k === 'bullet') return 'Bullet';
      if (k === 'numbered') return 'Num';
      if (k === 'image') return 'Image';
      if (k === 'table') return 'Table';
      const m = /^heading([1-6])$/.exec(k);
      return m ? `H${m[1]}` : k;
    };
    model.blocks.forEach((b, bi) => {
      if (b.kind === 'table') {
        const rows = b.rows ?? [];
        rows.forEach((row, r) => {
          row.forEach((cellText, c) => {
            out.push({
              id: `${b.id}:${r}:${c}`,
              text: cellText,
              label: `Block ${bi + 1} · Table R${r + 1}C${c + 1}`,
            });
          });
        });
      } else {
        out.push({
          id: b.id,
          text: b.text ?? '',
          label: `Block ${bi + 1} · ${kindLabel(b.kind)}`,
        });
      }
    });
    return out;
  }, [model]);

  // Marquee-selection keyboard ops:
  //   - Esc                — dismiss the selection
  //   - Delete / Backspace — bulk-delete every selected block in one undo
  //                          step (only when no editable surface owns focus,
  //                          so normal in-block Delete still works)
  // Scoped to document so it fires even when no block is focused — right
  // after a marquee, focus is on <body>, not on any of the selected blocks.
  // We bail when a modal dialog is open so we don't steal its Esc.
  // MUST stay above the EARLY RETURNS below: the early-return path skips
  // every hook below it, and inconsistent hook counts across renders is
  // the React error #310 we hit on first deploy.
  useEffect(() => {
    if (selectedBlockIds.size === 0) return;
    const onKey = (e: KeyboardEvent) => {
      if (linkDialog || mdInsertDialog || htmlInsertDialog || findOpen || gotoOpen) return;
      if (e.key === 'Escape') {
        setSelectedBlockIds(new Set());
        return;
      }
      if (e.key !== 'Delete' && e.key !== 'Backspace') return;
      // Don't intercept when the user is typing in a block / cell / input —
      // they want the normal forward-/back-delete-character behaviour.
      const ae = document.activeElement as HTMLElement | null;
      if (ae) {
        const tag = ae.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA' || ae.isContentEditable) return;
      }
      e.preventDefault();
      const ids = Array.from(selectedBlockIds);
      setSelectedBlockIds(new Set());
      removeBlocks(ids);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
    // removeBlocks is stable across renders (defined in the same closure)
    // and reading a stale capture would just re-run the same setModel call,
    // so we deliberately omit it from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedBlockIds, linkDialog, mdInsertDialog, htmlInsertDialog, findOpen, gotoOpen]);

  // ── EARLY RETURNS ─────────────────────────────────────────────────────
  // All hooks have run by this point, so the loading / error / null-model
  // guards can short-circuit rendering without changing hook count between
  // renders. Anything below this point only runs once `model` exists.
  if (loading) {
    return (
      <div className="h-full w-full flex items-center justify-center text-muted-foreground text-sm">
        {t('正在解析 docx…', 'Parsing docx…')}
      </div>
    );
  }
  if (error || !model) {
    return (
      <div className="h-full w-full flex items-center justify-center text-destructive text-sm p-8 text-center">
        {t('無法解析 docx：', 'Failed to parse docx: ')}{error ?? t('未知錯誤', 'Unknown error')}
      </div>
    );
  }

  /** Apply a F&R replacement. Composite ids ("blockId:r:c") rewrite a
   * specific table cell while keeping the rest of the row/column intact;
   * plain ids just patch the block's text. */
  const applyFindReplace = (id: string, newText: string): void => {
    const parts = id.split(':');
    if (parts.length === 3) {
      const [blockId, rs, cs] = parts;
      const r = Number(rs);
      const c = Number(cs);
      const block = model.blocks.find((b) => b.id === blockId);
      if (!block || block.kind !== 'table' || !block.rows) return;
      const newRows = block.rows.map((row, rr) =>
        rr === r ? row.map((cell, cc) => (cc === c ? newText : cell)) : row,
      );
      updateBlock(blockId, { rows: newRows });
    } else {
      updateBlock(id, { text: newText } as Partial<DocxBlock>);
    }
  };

  const setAlign = (align: DocxAlign) => {
    if (!activeBlock) return;
    updateBlock(activeBlock.id, {
      align: activeBlock.align === align ? undefined : align,
    });
  };

  /**
   * Restore focus + caret to the currently active RichBlock. Used after the
   * user touches a control that inevitably steals focus from contentEditable
   * — native `<select>` and `<input type="color">` open browser-managed UI
   * that shifts the document's active element, and there's no way to suppress
   * that with `preventDefault`. Best we can do is hand focus back once the
   * value commits so the user keeps typing where they left off.
   */
  /**
   * Block-focus shim: every contentEditable / cell-input onFocus routes
   * through this so we can clear the marquee multi-selection in one place.
   * Rule of thumb: focusing a block to edit it is incompatible with
   * "selection mode" — the rings would visually persist across blocks the
   * user is no longer manipulating.
   */
  const handleBlockFocus = (id: string) => {
    setActiveBlockId(id);
    setSelectedBlockIds((prev) => (prev.size === 0 ? prev : new Set()));
  };

  const refocusActiveBlock = () => {
    if (!activeBlockId) return;
    // Defer to the next frame so the browser-managed dropdown / picker has
    // fully closed and released focus before we steal it back.
    requestAnimationFrame(() => {
      const el = document.querySelector(
        `[data-rich-block="1"][data-block-id="${activeBlockId}"]`,
      ) as HTMLElement | null;
      el?.focus();
    });
  };

  const setColor = (hex: string) => {
    if (!activeBlock) return;
    const cleaned = hex.replace(/^#/, '').toUpperCase();
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, color: cleaned || undefined }) });
    refocusActiveBlock();
  };

  const clearColor = () => {
    if (!activeBlock) return;
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, color: undefined }) });
    refocusActiveBlock();
  };

  const setFontSize = (pt: number | undefined) => {
    if (!activeBlock) return;
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, fontSize: pt }) });
    refocusActiveBlock();
  };

  const setFontFamily = (name: string | undefined) => {
    if (!activeBlock) return;
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, fontFamily: name }) });
    refocusActiveBlock();
  };

  const setKind = (kind: DocxBlockKind) => {
    if (!activeBlock) return;
    updateBlock(activeBlock.id, { kind });
    refocusActiveBlock();
  };

  const setHighlight = (color: DocxHighlightColor | undefined) => {
    if (!activeBlock) return;
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, highlight: color }) });
    refocusActiveBlock();
  };

  const setLineSpacing = (mult: number | undefined) => {
    if (!activeBlock) return;
    const cur = activeBlock.style ?? {};
    updateBlock(activeBlock.id, { style: normalizeStyle({ ...cur, lineSpacing: mult }) });
    refocusActiveBlock();
  };

  const togglePageBreak = () => {
    if (!activeBlock) return;
    updateBlock(activeBlock.id, {
      pageBreakBefore: activeBlock.pageBreakBefore ? undefined : true,
    });
  };

  /**
   * Toggle the active block's hyperlink. Opens an inline floating dialog
   * (`LinkEditDialog`) with the current value pre-filled — committing an
   * empty URL clears the link, any value updates it, Esc cancels.
   *
   * Round 22 replaced the markdown editor's equivalent `window.prompt` with
   * a floating dialog; this round brings DOCX in line so the app no longer
   * pops native modal prompts in core workflows. We snapshot the active
   * block's id at open time so committing goes to the intended block even
   * if the user clicks into another block while the dialog is showing.
   */
  const setOrToggleLink = () => {
    if (!activeBlock) return;
    // R350 — preempt the four sibling dialogs (md / html / find / goto)
    // before opening link. Without this, opening 插入連結 while find or
    // goto is open stacks the dialogs on top of each other at the same
    // `top-3 right-3` anchor.
    closeOtherDialogs('link');
    setLinkDialog({ blockId: activeBlock.id, defaultUrl: activeBlock.link ?? '' });
  };

  const setPageSize = (size: DocxPageSize) => {
    setModel((prev) => {
      if (!prev) return prev;
      const next: DocxModel = { ...prev, pageSize: size };
      scheduleSerialize(next);
      return next;
    });
  };

  const togglePageOrientation = () => {
    setModel((prev) => {
      if (!prev) return prev;
      const next: DocxModel = { ...prev, pageSize: { w: prev.pageSize.h, h: prev.pageSize.w } };
      scheduleSerialize(next);
      return next;
    });
  };

  const setPageMargins = (m: DocxPageMargins) => {
    setModel((prev) => {
      if (!prev) return prev;
      const next: DocxModel = { ...prev, pageMargins: m };
      scheduleSerialize(next);
      return next;
    });
  };

  /**
   * Free-position drag: pointerdown on the grip captures the block's current
   * box (in twips, relative to the page-sheet top-left including margins) and
   * installs window-level pointermove / pointerup listeners. The block is
   * lifted out of the document flow on first move — we record the in-flow
   * bounding rect once so the block doesn't visually jump on lift.
   *
   * No reorder, no docx persistence — positions are pure editor-state.
   */
  const resetBlockPosition = (id: string) => {
    setPositions((prev) => {
      if (!prev[id]) return prev;
      const { [id]: _drop, ...rest } = prev;
      return rest;
    });
  };

  const startBlockMove = (id: string, e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();

    // Multi-block move: dragging the grip on a marquee-selected block
    // translates every selected block by the same delta, preserving their
    // relative arrangement. Dragging the grip on a non-selected block
    // clears any stale selection and falls back to single-block move so
    // visual rings don't linger on blocks that aren't actually moving.
    const isMulti = selectedBlockIds.has(id) && selectedBlockIds.size > 1;
    const movingIds = isMulti ? Array.from(selectedBlockIds) : [id];
    if (!isMulti && selectedBlockIds.size > 0) {
      setSelectedBlockIds(new Set());
    }

    // Snapshot the start position of every moving block once at drag-start.
    // Reading positions during pointermove would race with React's batched
    // state updates and cause drift on fast drags.
    const starts: Record<string, { x: number; y: number; w: number }> = {};
    for (const bid of movingIds) {
      const stored = positions[bid];
      if (stored) {
        starts[bid] = { x: stored.xTwip, y: stored.yTwip, w: stored.wTwip };
        continue;
      }
      const blockEl = document.querySelector(
        `[data-block-id="${bid}"]`,
      ) as HTMLElement | null;
      const pageEl = blockEl?.closest('[data-page-content]') as HTMLElement | null;
      if (blockEl && pageEl && pxPerTwip > 0) {
        const blockRect = blockEl.getBoundingClientRect();
        const pageRect = pageEl.getBoundingClientRect();
        starts[bid] = {
          x: (blockRect.left - pageRect.left) / pxPerTwip,
          y: (blockRect.top - pageRect.top) / pxPerTwip,
          w: blockRect.width / pxPerTwip,
        };
      } else {
        starts[bid] = { x: 0, y: 0, w: 6000 };
      }
    }

    const pointerStartX = e.clientX;
    const pointerStartY = e.clientY;
    const scale = pxPerTwip;

    const onMove = (ev: PointerEvent) => {
      if (scale <= 0) return;
      const dxTwip = (ev.clientX - pointerStartX) / scale;
      const dyTwip = (ev.clientY - pointerStartY) / scale;
      setPositions((prev) => {
        const next = { ...prev };
        for (const bid of movingIds) {
          const s = starts[bid];
          next[bid] = {
            xTwip: s.x + dxTwip,
            yTwip: s.y + dyTwip,
            wTwip: s.w,
          };
        }
        return next;
      });
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  /**
   * Marquee rubber-band: pointerdown on the page background (i.e. not on a
   * block) starts drawing a selection rectangle in page-content-relative
   * pixels. On release we hit-test each block's bounding rect against the
   * final marquee; intersecting blocks become the selection. A pure click
   * (no movement) clears any prior selection and is otherwise a no-op so
   * users can dismiss the multi-selection by clicking the page margin.
   */
  const handlePagePointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    if (e.button !== 0) return;
    const target = e.target as HTMLElement;
    // Pointerdown that lands on a block (or anything inside one) is the
    // user's normal "click-to-edit" path — the block's own handlers take
    // over and we mustn't preventDefault, otherwise the caret never lands.
    if (target.closest('[data-block-id]')) return;
    // Defensive: if the marquee overlay itself is the target (e.g. a
    // stale frame), don't recursively start another marquee.
    if (target.closest('[data-marquee-overlay]')) return;

    const pageEl = e.currentTarget;
    const pageRect = pageEl.getBoundingClientRect();
    const startX = e.clientX - pageRect.left;
    const startY = e.clientY - pageRect.top;

    e.preventDefault();
    // Click on empty area always dismisses any prior multi-selection — even
    // if the user never ends up dragging — because that's what "click outside
    // to deselect" should do. We do NOT clear `activeBlockId` here, though;
    // a stray click on the page margin shouldn't lose the user's editing
    // focus. That clear is deferred to the first confirmed drag-move below.
    setSelectedBlockIds((prev) => (prev.size === 0 ? prev : new Set()));
    setMarquee({ left: startX, top: startY, width: 0, height: 0 });

    let moved = false;
    const onMove = (ev: PointerEvent) => {
      const curX = ev.clientX - pageRect.left;
      const curY = ev.clientY - pageRect.top;
      // Treat sub-3px movement as a click — avoids spurious 1-pixel drags
      // from imprecise mice clearing the user's selection unexpectedly.
      if (!moved && Math.abs(curX - startX) < 3 && Math.abs(curY - startY) < 3) {
        return;
      }
      if (!moved) {
        moved = true;
        // First confirmed drag-move: now we know this is a marquee, not a
        // stray click. Drop the editing focus so the active ring doesn't
        // visually compete with the rubber-band as it crosses blocks.
        setActiveBlockId(null);
      }
      setMarquee({
        left: Math.min(startX, curX),
        top: Math.min(startY, curY),
        width: Math.abs(curX - startX),
        height: Math.abs(curY - startY),
      });
    };
    const onUp = (ev: PointerEvent) => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      setMarquee(null);
      if (!moved) return;
      const curX = ev.clientX - pageRect.left;
      const curY = ev.clientY - pageRect.top;
      const left = Math.min(startX, curX);
      const top = Math.min(startY, curY);
      const right = Math.max(startX, curX);
      const bottom = Math.max(startY, curY);

      // Hit-test top-level blocks only — `model.blocks` membership filters
      // out nested data-block-id lookalikes (e.g. table cells, image
      // chrome). Coords are page-content-relative on both sides.
      const ids = new Set(model?.blocks.map((b) => b.id) ?? []);
      const hits = new Set<string>();
      const candidates = pageEl.querySelectorAll<HTMLElement>('[data-block-id]');
      for (const el of Array.from(candidates)) {
        const id = el.dataset.blockId ?? '';
        if (!ids.has(id)) continue;
        const r = el.getBoundingClientRect();
        const rl = r.left - pageRect.left;
        const rt = r.top - pageRect.top;
        const rr = rl + r.width;
        const rb = rt + r.height;
        if (rr < left || rl > right || rb < top || rt > bottom) continue;
        hits.add(id);
      }
      setSelectedBlockIds(hits);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  // OS file-drag predicate. `dataTransfer.types` is populated during the drag
  // (unlike `.files`, which is empty until drop for security). "Files" is the
  // canonical type when File Explorer / Finder is the source.
  const isFileDrag = (e: React.DragEvent): boolean => {
    return Array.from(e.dataTransfer.types).includes('Files');
  };

  // Find the top-level block under (or closest to) the drop point. Word
  // documents are linear so we only need to consider clientY: rank blocks by
  // (a) "rect contains Y" first, then (b) closest-above when Y is past the
  // last block, then (c) closest-below when Y is above all blocks. We filter
  // querySelectorAll results against `model.blocks` to skip nested
  // `data-block-id` lookalikes (table cells, image overlays). Returns null
  // when the page hasn't laid out yet — caller falls back to active/last.
  const blockAtPoint = (clientY: number): string | null => {
    if (!model) return null;
    const ids = new Set(model.blocks.map((b) => b.id));
    const candidates = Array.from(
      document.querySelectorAll<HTMLElement>('[data-block-id]'),
    ).filter((el) => ids.has(el.dataset.blockId ?? ''));
    if (candidates.length === 0) return null;
    let bestContains: { id: string; top: number } | null = null;
    let bestAbove: { id: string; bottom: number } | null = null;
    let bestBelow: { id: string; top: number } | null = null;
    for (const el of candidates) {
      const id = el.dataset.blockId;
      if (!id) continue;
      const r = el.getBoundingClientRect();
      if (clientY >= r.top && clientY <= r.bottom) {
        // Inside a block: pick the topmost (most outer) one if multiple
        // candidates overlap — guards against nested matches even after the
        // ids-set filter (defensive).
        if (!bestContains || r.top < bestContains.top) {
          bestContains = { id, top: r.top };
        }
      } else if (r.bottom < clientY) {
        if (!bestAbove || r.bottom > bestAbove.bottom) {
          bestAbove = { id, bottom: r.bottom };
        }
      } else {
        if (!bestBelow || r.top < bestBelow.top) {
          bestBelow = { id, top: r.top };
        }
      }
    }
    if (bestContains) return bestContains.id;
    if (bestAbove) return bestAbove.id; // drop below last block → append after it
    if (bestBelow) {
      // Drop above first block → insert before it by anchoring after the
      // *previous* block, which doesn't exist; fall back to first block's id
      // and accept that the image lands as the second block. Better than
      // dropping at the cursor's far-away active block.
      return bestBelow.id;
    }
    return null;
  };

  return (
    <div
      data-docx-editor-root
      className="h-full w-full flex flex-col bg-background relative"
      onDragEnter={(e) => {
        if (!isFileDrag(e)) return;
        e.preventDefault();
        dragDepthRef.current += 1;
        setDraggingFile(true);
      }}
      onDragOver={(e) => {
        if (!isFileDrag(e)) return;
        // preventDefault is what makes the editor a valid drop target —
        // without it, `drop` never fires and the OS bounces the file back.
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
          // R85 silent-swallow fix: previously a non-image drop (PDF / .docx /
          // anything else) silently dismissed the overlay with zero feedback,
          // leaving users to wonder "did the editor accept it? am I in the
          // wrong place?" — the dragOver preventDefault advertised the editor
          // as a valid drop target, then drop refused without explanation.
          // PptxEditor.tsx:880 + MarkdownEditor.tsx:628 carry the same fix.
          // Tier + vocabulary mirror FileExplorer.tsx:700 (`不支援的檔案類型`)
          // — same warning level, same "tell the user what's wrong" stance.
          notify(tImp('只能拖入圖片檔案', 'Only image files can be dropped'), 'warning');
          return;
        }
        // Drop precision (Round 74): anchor to the block under the cursor
        // rather than the active block — users dragging from File Explorer
        // expect "where I drop is where it lands" (matches Excel's
        // cellAtPoint behaviour from Round 70). Falls back to the active
        // block (then last block) when point-resolution fails — e.g. drop
        // landed on the page margin or the page hasn't laid out yet.
        const target =
          blockAtPoint(e.clientY) ?? activeBlockId ?? model.blocks[model.blocks.length - 1]?.id;
        if (target) void insertImageBlockFromFile(imageFile, target);
      }}
    >
      {/* R414 — toolbar stays visible in preview mode. Preview only strips
          the page-level editing chrome (grips, gutters, rings); the blocks
          themselves remain editable (see RichBlock), so every toolbar
          action still works while inspecting the print appearance. */}
      {(
        <FormatToolbar
          block={activeBlock}
          previewMode={previewMode}
          pageSize={model.pageSize}
          pageMargins={model.pageMargins}
          navOpen={navOpen}
          onToggleNav={() => setNavOpen((v) => !v)}
          onToggleStyle={toggleStyle}
          onSetKind={setKind}
          onSetAlign={setAlign}
          onSetColor={setColor}
          onClearColor={clearColor}
          onSetFontSize={setFontSize}
          onSetFontFamily={setFontFamily}
          onSetHighlight={setHighlight}
          onSetLineSpacing={setLineSpacing}
          onTogglePageBreak={togglePageBreak}
          onToggleLink={setOrToggleLink}
          onInsertTable={() => activeBlock && insertTableAfter(activeBlock.id)}
          onInsertImage={() => {
            // Allow inserting at end-of-document when no block is active.
            const target = activeBlock?.id ?? model.blocks[model.blocks.length - 1]?.id;
            if (target) void insertImageAfter(target);
          }}
          onInsertMarkdown={() => {
            // Snapshot the active block id at open time (or null = append).
            // The dialog's commit handler reads this back so a stray click
            // into the document while typing doesn't relocate the splice.
            // R350 — close every sibling dialog before opening.
            closeOtherDialogs('md');
            setMdInsertDialog({ anchorBlockId: activeBlock?.id ?? null });
          }}
          onInsertHtml={() => {
            // Same anchor-snapshot semantics as onInsertMarkdown above.
            // R350 — close every sibling dialog before opening.
            closeOtherDialogs('html');
            setHtmlInsertDialog({ anchorBlockId: activeBlock?.id ?? null });
          }}
          onSetPageSize={setPageSize}
          onTogglePageOrientation={togglePageOrientation}
          onSetPageMargins={setPageMargins}
          onTogglePreview={() => setPreviewMode((v) => !v)}
          onOpenFind={() => {
            // Mirror the Ctrl+F handler — close any conflicting dialog and
            // refocus the query input via the nonce so repeat clicks behave
            // like repeat Ctrl+F presses.
            // R350 — extended from Find/Goto-only mutex to close all 4
            // siblings (link / md-insert / html-insert / goto). Toolbar
            // and Ctrl+F path now share the same closeOtherDialogs helper.
            closeOtherDialogs('find');
            setFindOpen(true);
            setFindFocusNonce((n) => n + 1);
          }}
          onOpenGoto={() => {
            // Symmetric to onOpenFind — Ctrl+G handler at line 785-798 closes
            // Find on open, so the toolbar entry must do the same to behave
            // identically whether the user reaches GoTo by keyboard or click.
            // The focus-nonce bump matches the Ctrl+G keymap so a repeat
            // click on this button refocuses+selects the input even when the
            // dialog is already open (mirrors VS Code's "Go to Line" behaviour
            // documented in GoToDialog.tsx:8-11).
            // R350 — same closeOtherDialogs helper as the Ctrl+G path.
            closeOtherDialogs('goto');
            setGotoOpen(true);
            setGotoFocusNonce((n) => n + 1);
          }}
        />
      )}
      <div className="relative flex-1 min-h-0 flex">
        {/* Word's Navigation Pane — toggled via the toolbar's outline icon
            or persisted via localStorage. Lives outside the scrolling page-
            sheet container so it stays put while the document scrolls. */}
        {/* R414 — nav pane no longer hidden by preview: jumping to a heading
            while inspecting print appearance is a legitimate flow, and the
            toolbar's nav toggle staying functional in preview requires the
            panel to actually appear. */}
        {navOpen && (
          <DocxNavPanel
            tabId={tab.id}
            outline={outline}
            activeIdx={activeOutlineIdx}
            onJump={jumpToHeading}
          />
        )}
        {/* Floating exit pill — anchored to the editor flex container (not
            window) so it stays bounded inside the editor surface and never
            collides with the AI panel sitting to its right. `absolute` keeps
            it pinned to the top-right of the editor while scroll happens
            inside the inner page-sheet container. */}
        {previewMode && (
          <button
            type="button"
            onClick={() => setPreviewMode(false)}
            title={tImp('退出預覽 (Esc)', 'Exit preview (Esc)')}
            className="absolute top-3 right-3 z-40 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-background/95 border border-border shadow-md text-xs hover:bg-secondary transition-colors"
          >
            <Eye className="h-3.5 w-3.5" />
            <span>{t('預覽中 — 點擊或按 Esc 結束', 'Previewing — click or press Esc to exit')}</span>
          </button>
        )}
        {/* The desk surrounding the white paper. Pinned to zinc-200 always —
            the work-canvas wrapper (EditorSurface) re-pins the CSS palette
            but Tailwind's `dark:` variant is class-based not var-based, so
            we drop it here to keep the desk paper-friendly even in dark
            theme. */}
        <div
          ref={deskScrollRef}
          className="relative flex-1 min-h-0 overflow-auto bg-zinc-200"
          onPointerDown={(e) => {
            // Click outside the page (in the gray desk gutter) clears the
            // marquee multi-selection — matches the OS-level convention
            // "click empty area to deselect". Pointerdowns that land on
            // the page or inside one of our floating dialogs are owned by
            // those handlers and must not be intercepted here.
            if (selectedBlockIds.size === 0) return;
            const t = e.target as HTMLElement;
            if (t.closest('[data-page-content]')) return;
            if (t.closest('[role="dialog"]')) return;
            // Don't clear if they tapped the floating hint itself — its
            // own buttons (取消 / 刪除) need to fire on the same click.
            if (t.closest('[data-selection-hint]')) return;
            setSelectedBlockIds(new Set());
          }}
        >
        <FindReplaceDialog
          open={findOpen}
          focusNonce={findFocusNonce}
          onClose={() => setFindOpen(false)}
          segments={findSegments}
          onUpdateSegment={applyFindReplace}
          // Composite "blockId:r:c" ids strip the suffix so we still
          // surface the parent table block (which the editor can scroll
          // to), even if we can't yet focus the specific cell.
          onLocateSegment={(id) => {
            const blockId = id.split(':')[0];
            setActiveBlockId(blockId);
            // Scroll the matched block into view. Without this, clicking
            // a Find result for a paragraph that sits below the desk's
            // scrollport just moved the active-ring offscreen — Find Next
            // felt like a no-op. rAF defers until React commits the
            // activeBlockId state so the ring class lands on the right
            // node before we measure. `block: 'nearest'` mirrors xlsx's
            // locate path: blocks already on screen don't jiggle, only
            // out-of-view hits scroll. We deliberately do NOT use
            // jumpToHeading's `block: 'start'` here — Find→Locate is
            // mid-document navigation, not a TOC jump, so dragging the
            // hit to the very top would be more disorienting than
            // helpful (loses the surrounding paragraphs the user might
            // be reading for context).
            requestAnimationFrame(() => {
              const el = deskScrollRef.current?.querySelector<HTMLElement>(
                `[data-block-id="${blockId}"]`,
              );
              el?.scrollIntoView({ block: 'nearest' });
            });
          }}
        />
        <GoToDialog
          open={gotoOpen}
          focusNonce={gotoFocusNonce}
          onClose={() => setGotoOpen(false)}
          max={model.blocks.length}
          label={t('跳到第幾段？', 'Go to which paragraph?')}
          onJump={(oneBased) => {
            const target = model.blocks[oneBased - 1];
            // Reuse the heading-jump path so highlight + smooth-scroll behave
            // identically whether the user clicked a heading in the Navigation
            // Pane (Round 76) or typed a paragraph number in this dialog.
            if (target) jumpToHeading(target.id);
          }}
        />
        {linkDialog && (
          <LinkEditDialog
            defaultUrl={linkDialog.defaultUrl}
            onClose={() => {
              setLinkDialog(null);
              refocusActiveBlock();
            }}
            onCommit={(url) => {
              const id = linkDialog.blockId;
              if (!url) {
                updateBlock(id, { link: undefined });
              } else {
                // Auto-prefix bare URLs so users can paste "google.com" without
                // scheme. Allow file:, mailto:, http(s): through unchanged.
                const normalized = /^[a-z][a-z0-9+.-]*:/i.test(url) ? url : `https://${url}`;
                updateBlock(id, { link: normalized });
              }
              setLinkDialog(null);
              refocusActiveBlock();
            }}
          />
        )}
        {selectedBlockIds.size > 0 &&
          !previewMode &&
          // Hide while any of our top-3 right-3 dialogs is open — they all
          // anchor to the same corner and stacking them on top of each
          // other reads as a layout glitch. The hint reappears as soon as
          // the dialog closes, so users who pressed Ctrl+F by accident
          // don't permanently lose their selection-mode UI.
          !linkDialog &&
          !mdInsertDialog &&
          !htmlInsertDialog &&
          !findOpen &&
          !gotoOpen && (
            // Floating selection-mode hint. Sticks to the top-right of the
            // scrolling editor area (not the page sheet) so it stays visible
            // even when the user has scrolled past the top of the document.
            // The "取消" button mirrors what Esc does — present both because
            // the keyboard shortcut isn't discoverable from a fresh marquee.
            <div
              data-selection-hint
              role="status"
              aria-live="polite"
              className="absolute top-3 right-3 z-30 flex items-center gap-2 rounded-md border bg-background/95 backdrop-blur shadow-sm px-2.5 py-1 text-xs"
            >
              <span className="text-muted-foreground">
                {/* R409 — i18n: text split around the styled count span */}
                {t('已選取 ', '')}<span className="font-medium text-foreground">{selectedBlockIds.size}</span>{t(' 個段落 · 拖曳握把整體移動 · Delete 可一併刪除', ' paragraphs selected · drag a handle to move them together · Delete removes them all')}
              </span>
              <button
                type="button"
                onClick={() => {
                  const ids = Array.from(selectedBlockIds);
                  setSelectedBlockIds(new Set());
                  removeBlocks(ids);
                }}
                className="px-1.5 py-0.5 rounded text-destructive hover:text-destructive hover:bg-destructive/10"
                title={tImp('刪除全部選取的段落 (Delete)', 'Delete all selected paragraphs (Delete)')}
              >
                {t('刪除', 'Delete')}
              </button>
              <button
                type="button"
                onClick={() => setSelectedBlockIds(new Set())}
                className="px-1.5 py-0.5 rounded text-muted-foreground hover:text-foreground hover:bg-secondary"
                title={tImp('取消選取 (Esc)', 'Clear selection (Esc)')}
              >
                {t('取消', 'Cancel')}
              </button>
            </div>
          )}
        {mdInsertDialog && (
          <MarkdownInsertDialog
            onClose={() => {
              setMdInsertDialog(null);
              refocusActiveBlock();
            }}
            onCommit={(source) => {
              const fresh = markdownToDocxBlocks(source, genId);
              if (fresh.length === 0) {
                // R341 — same silent-close bug R334 closed for HtmlInsert
                // Dialog. Comment used to claim "we should never reach
                // here with empty input" — true for source.trim() === ''
                // (dialog's own pre-check stops that), wrong for non-empty
                // markdown that LEXES to zero docx blocks. Realistic
                // triggers:
                //   • `---\n---\n---` (only horizontal rules; marked emits
                //     `hr` tokens and markdown-to-docx skips them per
                //     line 294-298 — "hr has no equivalent in the block
                //     model").
                //   • `<!-- todo -->` alone (block-level html token with
                //     `.text.trim()` non-empty actually DOES emit, so
                //     this case doesn't fire here, but mentioned for
                //     completeness vs the html dialog).
                //   • Pure `<space>` tokens between non-existent content
                //     (`\n\n\n` only — though trim() catches it before
                //     lexer).
                // Return an inline error so the dialog stays open with
                // a red-bordered message; mirrors the post-R335 protocol
                // where the dialog displays whatever caller returns.
                return tImp(
                  '解析後沒有任何可插入的內容（Markdown 可能只含 horizontal rules / 空白）',
                  'Nothing to insert after parsing (Markdown may contain only horizontal rules / whitespace)',
                );
              }
              insertBlocksAfter(mdInsertDialog.anchorBlockId, fresh);
              setMdInsertDialog(null);
              // Activate the first inserted block so subsequent edits land
              // in the right spot — matches the post-paste UX where the
              // caret naturally moves to the new content.
              setActiveBlockId(fresh[0].id);
              return undefined;
            }}
          />
        )}
        {htmlInsertDialog && (
          <HtmlInsertDialog
            onClose={() => {
              setHtmlInsertDialog(null);
              refocusActiveBlock();
            }}
            onCommit={(source) => {
              const { blocks: fresh, parseError } = htmlToDocxBlocks(source, genId);
              if (parseError) {
                // HTML parse failure surfaces inside the dialog (the dialog
                // shows the message inline so the user can fix the snippet);
                // bail without splicing or closing.
                // R335 — caller now owns the full user-facing message
                // (was prefixed dialog-side, but that prefix didn't fit
                // the R334 no-content path). Prefix「HTML 解析失敗：」
                // here so it still reads naturally for the genuine parse-
                // error case.
                return tImp(`HTML 解析失敗：${parseError}`, `HTML parse failed: ${parseError}`);
              }
              if (fresh.length === 0) {
                // R334 — surface as an inline error and KEEP THE DIALOG OPEN.
                // The previous "Defensive no-op — the dialog should keep
                // itself open on empty input via its own inline-hint path"
                // comment was wrong: the dialog's own empty-input check
                // (HtmlInsertDialog.submit at line ~3944) only catches
                // `source.trim() === ''` (raw input empty). The fresh.length
                // === 0 case is the OUTPUT being empty after htmlToDocx
                // Blocks ran — happens for legal-but-content-free input
                // like `<script>alert(1)</script>` (R317 skips script
                // entirely → zero blocks), `<!-- a comment -->` (DOMParser
                // emits a Comment node, walkBlock has no branch for it,
                // unknown-wrapper path's `el.childNodes.forEach` skips it
                // as non-element-non-text), or `<div></div>` (unknown
                // wrapper, recurses into empty children, produces 0 blocks).
                // The dialog closing silently with no insertion is a worse
                // UX than the parse-error case: the user clicked 插入,
                // saw the dialog disappear, and assumed *something*
                // landed in the document — but the active block didn't
                // change and no new block appeared. Returning a string
                // (per the same contract parseError uses) keeps the
                // dialog open with the message painted inline, so the
                // user can either edit the source to add real content
                // or hit Esc to cancel knowingly.
                return tImp(
                  '解析後沒有任何可插入的內容（HTML 只含 script / style / 註解 / 空標籤）',
                  'Nothing to insert after parsing (HTML contains only script / style / comments / empty tags)',
                );
              }
              insertBlocksAfter(htmlInsertDialog.anchorBlockId, fresh);
              setHtmlInsertDialog(null);
              setActiveBlockId(fresh[0].id);
              return undefined;
            }}
          />
        )}
        <PageSheet
          pageSize={model.pageSize}
          pageMargins={model.pageMargins}
          onPxPerTwipChange={setPxPerTwip}
          onPagePointerDown={previewMode ? undefined : handlePagePointerDown}
        >
          <div className="space-y-2">
            {model.blocks.map((b) =>
              b.kind === 'table' ? (
                <TableBlock
                  key={b.id}
                  block={b}
                  active={b.id === activeBlockId}
                  selected={selectedBlockIds.has(b.id)}
                  position={positions[b.id]}
                  pxPerTwip={pxPerTwip}
                  previewMode={previewMode}
                  onFocus={() => handleBlockFocus(b.id)}
                  onChangeRows={(rows) => updateBlock(b.id, { rows })}
                  onInsertAfter={() => insertAfter(b.id)}
                  onRemove={() => removeBlock(b.id)}
                  onStartMove={(e) => startBlockMove(b.id, e)}
                  onResetPosition={() => resetBlockPosition(b.id)}
                />
              ) : b.kind === 'image' ? (
                <ImageBlockRow
                  key={b.id}
                  block={b}
                  active={b.id === activeBlockId}
                  selected={selectedBlockIds.has(b.id)}
                  position={positions[b.id]}
                  pxPerTwip={pxPerTwip}
                  previewMode={previewMode}
                  onFocus={() => handleBlockFocus(b.id)}
                  onResize={(w, h) =>
                    updateBlock(b.id, {
                      image: b.image
                        ? { ...b.image, widthPx: w, heightPx: h }
                        : b.image,
                    })
                  }
                  onInsertAfter={() => insertAfter(b.id)}
                  onRemove={() => removeBlock(b.id)}
                  onStartMove={(e) => startBlockMove(b.id, e)}
                  onResetPosition={() => resetBlockPosition(b.id)}
                />
              ) : (
                <BlockRow
                  key={b.id}
                  block={b}
                  active={b.id === activeBlockId}
                  selected={selectedBlockIds.has(b.id)}
                  position={positions[b.id]}
                  pxPerTwip={pxPerTwip}
                  previewMode={previewMode}
                  onFocus={() => handleBlockFocus(b.id)}
                  onChangeRuns={(runs) => updateBlock(b.id, { runs })}
                  onCycleKind={() => cycleKind(b.id)}
                  onInsertAfter={() => insertAfter(b.id)}
                  onRemove={() => removeBlock(b.id)}
                  onStartMove={(e) => startBlockMove(b.id, e)}
                  onResetPosition={() => resetBlockPosition(b.id)}
                />
              ),
            )}
          </div>
          {marquee && (
            <div
              data-marquee-overlay
              className="absolute pointer-events-none border border-primary/70 bg-primary/10 z-30"
              style={{
                left: marquee.left,
                top: marquee.top,
                width: marquee.width,
                height: marquee.height,
              }}
            />
          )}
        </PageSheet>
        </div>
      </div>
      {draggingFile && (
        // Subtle drop-target indicator. `pointer-events-none` is critical —
        // without it, the overlay would steal dragenter/leave from descendants
        // and the depth counter would never settle. Pure visual cue, no DOM
        // interaction.
        // R105 — overlay copy now discloses cursor-precision drop, mirroring
        // XlsxEditor.tsx:3188「放開以將圖片插入到游標所在儲存格」. The drop
        // logic at line 1358-1366 above explicitly anchors to `blockAtPoint
        // (e.clientY)` (with active-block / last-block fallbacks), and its
        // own doc-comment names the design goal verbatim: "users dragging
        // from File Explorer expect 'where I drop is where it lands'
        // (matches Excel's cellAtPoint behaviour from Round 70)". So the
        // *behaviour* aligned with XlsxEditor in R74, but this overlay
        // never picked up the parallel vocabulary — DocxEditor was the
        // silent middle in a 3-way comparison: PptxEditor.tsx:1147 says
        //「放開即插入圖片到目前投影片」(slide-level), XlsxEditor says
        //「游標所在儲存格」(cell-level), Doc said just「放開即插入圖片」
        // with zero hint about *where*. Keeping「即」(not XlsxEditor's
        //「以將」) preserves Doc/Ppt's shared connector while picking up
        // the Excel side's cursor-precision noun phrase「游標所在X」 with
        // 段落 substituted for 儲存格. */}
        <div className="pointer-events-none absolute inset-0 bg-primary/5 border-2 border-dashed border-primary/40 z-50 flex items-center justify-center">
          <div className="px-3 py-1.5 rounded bg-background/90 text-xs text-primary border border-primary/40 shadow">
            {/* R409 — i18n */}
            {t('放開即插入圖片到游標所在段落', "Release to insert the image at the cursor's paragraph")}
          </div>
        </div>
      )}
    </div>
  );
}

/**
 * Word's Navigation Pane equivalent. Flat list of headings indented by
 * level, click-to-jump, current heading auto-highlighted and auto-scrolled
 * into view. Mirrors MarkdownEditor's OutlinePanel structurally so the two
 * editors stay visually consistent (intentional — users shouldn't have to
 * relearn navigation across formats).
 */
function DocxNavPanel({
  tabId,
  outline,
  activeIdx,
  onJump,
}: {
  tabId: string;
  outline: Array<{ id: string; level: number; text: string; index: number }>;
  activeIdx: number;
  onJump: (id: string) => void;
}): JSX.Element {
  const t = useT(); // R409 — i18n for header + empty-outline hint
  // The `<aside>` is the actual scroll container (overflow-auto); the
  // `<ul>` inside it has no overflow style. We need a separate ref on
  // the aside to read/write its scrollTop for memory — `listRef` would
  // give us the inner list, whose scrollTop is always 0.
  const asideRef = useRef<HTMLElement>(null);
  const listRef = useRef<HTMLUListElement>(null);
  // Auto-scroll the active heading into view in the panel itself when the
  // caret moves into a different section. `block: 'nearest'` means we don't
  // jiggle the panel when the active row is already visible.
  useEffect(() => {
    if (activeIdx < 0) return;
    const el = listRef.current?.querySelector<HTMLElement>(
      `[data-nav-idx="${activeIdx}"]`,
    );
    el?.scrollIntoView({ block: 'nearest' });
  }, [activeIdx]);
  // Outline-pane scroll memory across tab swaps. Restore via
  // `queueMicrotask` so we run AFTER the activeIdx effect above (which
  // declares first and synchronously scrollIntoViews the active row) —
  // microtask flush comes after the current render's effect chain, so
  // the remembered offset wins. Without this ordering trick, mounting on
  // a tab whose active heading is near the top would snap the pane to
  // the top, even though the user had scrolled to read H30-H40 before
  // tab-swapping away.
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

  // Arrow-key navigation among headings. Tab cycles through the buttons via
  // the browser default, but in long documents that's verbose — ↑/↓ lets
  // the user scan headings without leaving the panel. Enter is handled
  // natively by the focused <button>. Mirrors the equivalent OutlinePanel
  // wiring in MarkdownEditor.
  const focusEntry = (i: number) => {
    const el = listRef.current?.querySelector<HTMLButtonElement>(
      `[data-nav-idx="${i}"]`,
    );
    el?.focus();
  };
  const onKeyNav = (e: React.KeyboardEvent<HTMLElement>) => {
    if (outline.length === 0) return;
    const focused = document.activeElement as HTMLElement | null;
    const ownIdx = focused?.dataset?.navIdx
      ? Number(focused.dataset.navIdx)
      : Math.max(0, activeIdx);
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      focusEntry(Math.min(outline.length - 1, ownIdx + 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      focusEntry(Math.max(0, ownIdx - 1));
    } else if (e.key === 'Home') {
      e.preventDefault();
      focusEntry(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      focusEntry(outline.length - 1);
    }
  };
  return (
    <aside
      ref={asideRef}
      onKeyDown={onKeyNav}
      className="w-60 shrink-0 border-r bg-secondary/20 overflow-auto text-xs"
    >
      {/* Keyboard-nav hint mirrored across all three outline panels this
          round (PptxEditor PptxNavPanel header, MarkdownEditor OutlinePanel
          header). onKeyNav at line 1742-1761 handles ↑/↓/Home/End but
          nothing in the UI told users the arrow keys work — see fuller
          rationale at PptxEditor.tsx near "投影片大綱". */}
      <div
        className="px-3 py-2 text-[10px] uppercase tracking-wider text-muted-foreground font-medium border-b"
        title={tImp('↑/↓ 切換 · Home/End 跳到首/末', '↑/↓ to switch · Home/End to jump to first/last')}
      >
        {t('導覽窗格', 'Navigation pane')}
      </div>
      {outline.length === 0 ? (
        <div className="px-3 py-2 text-[11px] text-muted-foreground italic">
          {t('（沒有標題。將段落改成「標題 1/2/3…」即可建立大綱）', '(No headings. Set a paragraph to Heading 1/2/3… to build an outline)')}
        </div>
      ) : (
        <ul ref={listRef} className="py-1">
          {outline.map((e, i) => {
            const isActive = i === activeIdx;
            return (
              <li key={e.id}>
                <button
                  type="button"
                  data-nav-idx={i}
                  onClick={() => onJump(e.id)}
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
    </aside>
  );
}

/**
 * Renders content inside a Word-style "page" — white sheet with shadow whose
 * width and padding mirror the document's `<w:pgSz>` / `<w:pgMar>`. This is
 * a *visual* boundary only: content longer than one page just keeps flowing
 * (the editor doesn't paginate). 96 DPI ≈ 1/15 px-per-twip gives A4 ≈ 794px,
 * matching Word's 100% zoom. On narrow viewports we scale down proportionally
 * so the page still fits without horizontal scroll.
 */
function PageSheet({
  pageSize,
  pageMargins,
  onPxPerTwipChange,
  onPagePointerDown,
  children,
}: {
  pageSize: { w: number; h: number };
  pageMargins: { top: number; right: number; bottom: number; left: number };
  onPxPerTwipChange?: (v: number) => void;
  /** Pointer-down on the page background (margin / inter-block gaps) — used
   *  by DocxEditor to start a marquee multi-selection. Undefined in
   *  preview mode so the page can't be marqueed when blocks aren't editable. */
  onPagePointerDown?: (e: React.PointerEvent<HTMLDivElement>) => void;
  children: React.ReactNode;
}): JSX.Element {
  const outerRef = useRef<HTMLDivElement | null>(null);
  const [pxPerTwip, setPxPerTwip] = useState(1 / 15);
  useEffect(() => {
    const el = outerRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      const avail = el.clientWidth - 32; // px-4 gutter on each side
      if (avail <= 0) return;
      const ideal = 1 / 15;
      const fit = avail / pageSize.w;
      const next = Math.min(ideal, fit);
      setPxPerTwip(next);
      onPxPerTwipChange?.(next);
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, [pageSize.w, onPxPerTwipChange]);

  const widthPx = Math.max(200, Math.round(pageSize.w * pxPerTwip));
  const minHeightPx = Math.max(200, Math.round(pageSize.h * pxPerTwip));
  const padTop = Math.round(pageMargins.top * pxPerTwip);
  const padRight = Math.round(pageMargins.right * pxPerTwip);
  const padBottom = Math.round(pageMargins.bottom * pxPerTwip);
  const padLeft = Math.round(pageMargins.left * pxPerTwip);

  return (
    <div ref={outerRef} className="py-8 px-4">
      <div
        data-page-content
        onPointerDown={onPagePointerDown}
        className="mx-auto bg-background border border-border shadow-lg rounded-sm relative"
        style={{
          width: widthPx,
          minHeight: minHeightPx,
          paddingTop: padTop,
          paddingRight: padRight,
          paddingBottom: padBottom,
          paddingLeft: padLeft,
        }}
      >
        {children}
      </div>
    </div>
  );
}

/** Common Word font sizes in points. */
const DOCX_FONT_SIZES_PT = [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

function FormatToolbar({
  block,
  pageSize,
  pageMargins,
  navOpen,
  onToggleNav,
  onToggleStyle,
  onSetKind,
  onSetAlign,
  onSetColor,
  onClearColor,
  onSetFontSize,
  onSetFontFamily,
  onSetHighlight,
  onSetLineSpacing,
  onTogglePageBreak,
  onToggleLink,
  onInsertTable,
  onInsertImage,
  onInsertMarkdown,
  onInsertHtml,
  onSetPageSize,
  onTogglePageOrientation,
  onSetPageMargins,
  onTogglePreview,
  onOpenFind,
  onOpenGoto,
  previewMode,
}: {
  block: DocxBlock | null;
  pageSize: DocxPageSize;
  pageMargins: DocxPageMargins;
  navOpen: boolean;
  /** R414 — the toolbar persists through preview mode; the Eye button shows
   *  pressed state and the tooltip flips between enter / exit. */
  previewMode: boolean;
  onToggleNav: () => void;
  onToggleStyle: (k: 'bold' | 'italic' | 'underline' | 'strikethrough') => void;
  onSetKind: (k: DocxBlockKind) => void;
  onSetAlign: (a: DocxAlign) => void;
  onSetColor: (hex: string) => void;
  onClearColor: () => void;
  onSetFontSize: (pt: number | undefined) => void;
  onSetFontFamily: (name: string | undefined) => void;
  onSetHighlight: (color: DocxHighlightColor | undefined) => void;
  onSetLineSpacing: (mult: number | undefined) => void;
  onTogglePageBreak: () => void;
  onToggleLink: () => void;
  onInsertTable: () => void;
  onInsertImage: () => void;
  /** Open the markdown-paste dialog. Convert + splice happens in the parent
   *  via insertBlocksAfter so undo entries are coherent. */
  onInsertMarkdown: () => void;
  /** Open the HTML-paste dialog. Same convert + splice flow as Markdown via
   *  html-to-docx.ts. */
  onInsertHtml: () => void;
  onSetPageSize: (s: DocxPageSize) => void;
  onTogglePageOrientation: () => void;
  onSetPageMargins: (m: DocxPageMargins) => void;
  onTogglePreview: () => void;
  onOpenFind: () => void;
  onOpenGoto: () => void;
}): JSX.Element {
  const t = useT();
  const disabled = !block;
  const style = block?.style ?? {};
  const tDisabled = t('請先點選一個段落', 'Select a paragraph first');
  // Kind switching / page-break flags only apply to paragraph-family blocks
  // — tables and images keep their structural kind.
  const isParaBlock = !!block && block.kind !== 'table' && block.kind !== 'image';
  const styleOpt = block ? STYLE_OPTIONS.find((o) => o.kind === block.kind) : undefined;
  // Wrapper no longer dims the whole bar when no block is active. Previously
  // we set `opacity-50` on the container, but several controls
  // (nav-pane toggle / image insert / page settings / find / print preview)
  // intentionally stay enabled regardless — and they were getting visually
  // greyed-out alongside the truly-disabled buttons, contradicting the
  // intent. Per-control disabled styling lives on `ToolbarBtn` and the
  // select / color picker (`opacity-50` already applied there individually).
  return (
    <div className="flex items-center gap-0.5 px-2 py-1 border-b bg-secondary/30">
      {/* Navigation pane toggle — first item so it's adjacent to the panel
          it controls (the pane appears immediately to the left when open).
          Stays enabled even when no block is active because the user might
          *be* opening the pane to find something to click on. */}
      <ToolbarBtn
        active={navOpen}
        disabled={false}
        title={navOpen ? t('隱藏導覽窗格', 'Hide navigation pane') : t('顯示導覽窗格 — 文件大綱', 'Show navigation pane — document outline')}
        onClick={onToggleNav}
      >
        <List className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      {/* Style dropdown — names the active block's kind and switches it.
          Same Radix dropdown pattern as PageSettingsMenu below; disabled for
          table / image blocks whose kind is structural, not typographic. */}
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button
            type="button"
            disabled={!isParaBlock}
            title={
              disabled
                ? tDisabled
                : !isParaBlock
                  ? t('表格 / 圖片區塊無法切換段落樣式', 'Tables / images cannot switch paragraph style')
                  : t('段落樣式', 'Paragraph style')
            }
            aria-label={t('段落樣式', 'Paragraph style')}
            className={cn(
              'inline-flex items-center gap-1 h-7 px-2 rounded text-xs min-w-[5.5rem] justify-between',
              'text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors',
              !isParaBlock && 'opacity-50 cursor-not-allowed',
            )}
          >
            <span className="truncate">
              {block
                ? styleOpt
                  ? t(styleOpt.zh, styleOpt.en)
                  : KIND_LABEL[block.kind]
                : t('樣式', 'Style')}
            </span>
            <span className="text-[10px] text-muted-foreground/70">▾</span>
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="min-w-[10rem]">
          <DropdownMenuLabel>{t('段落樣式', 'Paragraph style')}</DropdownMenuLabel>
          {STYLE_OPTIONS.map((o) => (
            <DropdownMenuItem
              key={o.kind}
              onSelect={() => onSetKind(o.kind)}
              className={cn(
                o.kind === 'heading1' && 'text-base font-bold',
                o.kind === 'heading2' && 'font-bold',
                o.kind === 'heading3' && 'font-semibold',
                block?.kind === o.kind && 'bg-accent text-accent-foreground',
              )}
            >
              {t(o.zh, o.en)}
            </DropdownMenuItem>
          ))}
        </DropdownMenuContent>
      </DropdownMenu>
      <Divider />
      {/* Tooltips spell out the keyboard shortcuts the editor binds (see
          useFormatShortcuts at toggleStyle / Ctrl+F at setFindOpen / Ctrl+G
          at setGotoOpen). Without these hints, users have no visual way to
          discover the bindings — MarkdownToolbar has had them since day one
          and this brings the Word toolbar to parity.

          R87 — disabled-state tooltip override. When `disabled` (no block
          selected) the static labels "粗體 (Ctrl+B)" etc. are misleading:
          the button is greyed out, click does nothing, hover still says
          「粗體 (Ctrl+B)」 with zero hint as to *why*. Mirrors the Insert
          Table button at line ~2150 below — its tooltip already flips to
          「請先點選一個段落」when disabled, and the comment block above it
          explicitly defends the pattern (insert-at-selection button whose
          label assumes a selection exists). Extending the same flip to the
          inline-style / link / align buttons closes the only remaining set
          of disable-but-don't-explain holes in this toolbar. */}
      <ToolbarBtn active={!!style.bold} disabled={disabled} title={disabled ? tDisabled : t('粗體 (Ctrl+B)', 'Bold (Ctrl+B)')} onClick={() => onToggleStyle('bold')}>
        <Bold className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.italic} disabled={disabled} title={disabled ? tDisabled : t('斜體 (Ctrl+I)', 'Italic (Ctrl+I)')} onClick={() => onToggleStyle('italic')}>
        <Italic className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.underline} disabled={disabled} title={disabled ? tDisabled : t('底線 (Ctrl+U)', 'Underline (Ctrl+U)')} onClick={() => onToggleStyle('underline')}>
        <Underline className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={!!style.strikethrough} disabled={disabled} title={disabled ? tDisabled : t('刪除線 (Ctrl+Shift+X)', 'Strikethrough (Ctrl+Shift+X)')} onClick={() => onToggleStyle('strikethrough')}>
        <Strikethrough className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        active={!!block?.link}
        disabled={disabled}
        title={disabled ? tDisabled : block?.link ? t(`編輯連結：${block.link}`, `Edit link: ${block.link}`) : t('插入連結', 'Insert link')}
        onClick={onToggleLink}
      >
        <LinkIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn active={block?.align === 'left'} disabled={disabled} title={disabled ? tDisabled : t('靠左對齊', 'Align left')} onClick={() => onSetAlign('left')}>
        <AlignLeft className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={block?.align === 'center'} disabled={disabled} title={disabled ? tDisabled : t('置中對齊', 'Align center')} onClick={() => onSetAlign('center')}>
        <AlignCenter className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={block?.align === 'right'} disabled={disabled} title={disabled ? tDisabled : t('靠右對齊', 'Align right')} onClick={() => onSetAlign('right')}>
        <AlignRight className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn active={block?.align === 'justify'} disabled={disabled} title={disabled ? tDisabled : t('兩端對齊', 'Justify')} onClick={() => onSetAlign('justify')}>
        <AlignJustify className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Line spacing — paragraph-level like alignment, so it lives in the
          same cluster. 預設 clears the key (serializer falls back to the
          document default). */}
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button
            type="button"
            disabled={disabled}
            title={disabled ? tDisabled : t('行距', 'Line spacing')}
            aria-label={t('行距', 'Line spacing')}
            aria-pressed={!!style.lineSpacing}
            className={cn(
              'h-7 w-7 inline-flex items-center justify-center rounded transition-colors',
              style.lineSpacing
                ? 'bg-primary/20 text-primary'
                : 'text-muted-foreground hover:text-foreground hover:bg-secondary',
              disabled && 'cursor-not-allowed opacity-50',
            )}
          >
            <AlignVerticalSpaceAround className="h-3.5 w-3.5" />
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="min-w-[8rem]">
          <DropdownMenuLabel>{t('行距', 'Line spacing')}</DropdownMenuLabel>
          {LINE_SPACING_OPTIONS.map((o) => (
            <DropdownMenuItem
              key={o.v}
              onSelect={() => onSetLineSpacing(o.v)}
              className={cn(style.lineSpacing === o.v && 'bg-accent text-accent-foreground')}
            >
              {t(o.zh, o.en)}
            </DropdownMenuItem>
          ))}
          <DropdownMenuSeparator />
          <DropdownMenuItem
            onSelect={() => onSetLineSpacing(undefined)}
            className={cn(!style.lineSpacing && 'bg-accent text-accent-foreground')}
          >
            {t('預設', 'Default')}
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      <Divider />
      {/* R98 — extend R97's disabled-state tooltip flip from PptxEditor /
          XlsxEditor to this third sibling. R87 already wired the flip onto
          B/I/U/Link/Align (lines 2027-2055) of this same toolbar with the
          same `disabled` flag and the same 「請先點選一個段落」 wording;
          R97 closed the colour-picker / font / size gap in the other two
          editors but stopped before this one. Hovering a greyed Palette /
          字型 select / 字級 select with no block selected returned the
          action label («文字顏色» / «字型» / «字級 (pt)») on a dead control
          — exactly the gap R87 closed for the action buttons two rows up.
          Wording reused verbatim from R87 keeps this toolbar speaking one
          voice across every disabled-when-no-block control. */}
      <span className="relative inline-flex items-center">
        <label
          title={disabled ? tDisabled : t('文字顏色', 'Text color')}
          // Don't blur the active block on label-mousedown. The native color
          // input *will* still grab focus when its picker opens (browser-
          // managed), but `setColor` calls `refocusActiveBlock` to put it
          // back when the picker closes.
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
            title={disabled ? tDisabled : t('清除文字顏色', 'Clear text color')}
            onMouseDown={(e) => e.preventDefault()}
            onClick={onClearColor}
            disabled={disabled}
            className="ml-px text-muted-foreground hover:text-destructive text-[10px] leading-none px-0.5"
          >
            ×
          </button>
        ) : null}
      </span>
      {/* Highlight palette — six named docx highlight colors + clear. The
          color bar under the icon mirrors the Palette swatch convention. */}
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button
            type="button"
            disabled={disabled}
            title={disabled ? tDisabled : t('螢光標示', 'Text highlight')}
            aria-label={t('螢光標示', 'Text highlight')}
            aria-pressed={!!style.highlight}
            className={cn(
              'h-7 w-7 inline-flex items-center justify-center rounded transition-colors',
              style.highlight
                ? 'bg-primary/20 text-primary'
                : 'text-muted-foreground hover:text-foreground hover:bg-secondary',
              disabled && 'cursor-not-allowed opacity-50',
            )}
          >
            <span className="relative">
              <Highlighter className="h-3.5 w-3.5" />
              <span
                className="absolute -bottom-0.5 left-0 right-0 h-0.5 rounded"
                style={{
                  background: style.highlight ? HIGHLIGHT_CSS[style.highlight] : 'transparent',
                }}
              />
            </span>
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="min-w-[9rem]">
          <DropdownMenuLabel>{t('螢光標示', 'Highlight')}</DropdownMenuLabel>
          {HIGHLIGHT_COLORS.map((h) => (
            <DropdownMenuItem
              key={h.id}
              onSelect={() => onSetHighlight(h.id)}
              className={cn(style.highlight === h.id && 'bg-accent text-accent-foreground')}
            >
              <span
                className="h-3 w-3 rounded-sm border border-border shrink-0"
                style={{ background: h.css }}
              />
              {t(h.zh, h.en)}
            </DropdownMenuItem>
          ))}
          <DropdownMenuSeparator />
          <DropdownMenuItem onSelect={() => onSetHighlight(undefined)}>
            {t('清除', 'Clear')}
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
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
        title={disabled ? tDisabled : t('字型', 'Font')}
        style={style.fontFamily ? { fontFamily: style.fontFamily } : undefined}
      >
        <option value="">{t('預設字型', 'Default font')}</option>
        {FONT_FAMILIES.map((f) => (
          <option key={f.value} value={f.value} style={{ fontFamily: f.value }}>
            {f.label}
          </option>
        ))}
      </select>
      <Type className="h-3.5 w-3.5 text-muted-foreground ml-1" />
      <select
        disabled={disabled}
        value={style.fontSize ?? ''}
        onChange={(e) => {
          const v = e.target.value;
          onSetFontSize(v === '' ? undefined : Number(v));
        }}
        className={cn(
          'h-7 text-xs rounded border border-border bg-background px-1.5',
          disabled && 'cursor-not-allowed',
        )}
        title={disabled ? tDisabled : t('字級 (pt)', 'Font size (pt)')}
      >
        <option value="">{t('預設', 'Default')}</option>
        {DOCX_FONT_SIZES_PT.map((s) => (
          <option key={s} value={s}>{s} pt</option>
        ))}
      </select>
      <Divider />
      {/* State-aware tooltip — the static title references「此段」(deictic
          for the active paragraph), but when `disabled` there is no active
          paragraph to point at. Mirrors the R51 XlsxEditor insert row/col
          tooltips at line 1883-1937 (same pattern: insert-at-selection
          button whose label assumes a selection exists). The image / markdown
          inserts below also use a state-aware tooltip (see R102 comment) but
          they stay enabled and disclose their end-of-document fallback rather
          than blocking the click. */}
      <ToolbarBtn
        disabled={disabled}
        title={disabled ? tDisabled : t('在此段後插入表格', 'Insert table after this paragraph')}
        onClick={onInsertTable}
      >
        <Rows className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* R102 — image / markdown inserts stay enabled because the parent
          (line 1385-1395) falls back to splicing after the last block when
          no paragraph is active. Previously the static tooltips「插入圖片」/
          「插入 Markdown 內容」hid this entirely: a user with nothing
          selected clicked the image button, the image appeared at the end
          of the document, and the toolbar never told them where it would
          land. Same-file three-way insert-cluster mismatch — the disabled-
          aware sibling at line 2161 explicitly explains its constraint
          (「請先點選一個段落」), while these two always-enabled peers stayed
          silent on theirs. The smoking gun is the author's own defensive
          comment that previously admitted the fallback in code but never
          surfaced it in the UI. State-aware tooltip pattern reused verbatim
          from XlsxEditor.tsx:2043-2057 (R99) where the same friction was
          fixed for the spreadsheet image-insert button. */}
      <ToolbarBtn
        disabled={false}
        title={disabled
          ? t('於文件結尾插入圖片(PNG / JPG / GIF / BMP)', 'Insert image at end of document (PNG / JPG / GIF / BMP)')
          : t('在此段後插入圖片(PNG / JPG / GIF / BMP)', 'Insert image after this paragraph (PNG / JPG / GIF / BMP)')}
        onClick={onInsertImage}
      >
        <ImageIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Markdown insert — paste a markdown snippet, get back proper Word
          blocks (headings, lists, tables, B/I runs). Same end-of-document
          fallback as image insert; tooltip discloses both states (R102). */}
      <ToolbarBtn
        disabled={false}
        title={disabled
          ? t('於文件結尾插入 Markdown 內容', 'Insert Markdown at end of document')
          : t('在此段後插入 Markdown 內容', 'Insert Markdown after this paragraph')}
        onClick={onInsertMarkdown}
      >
        <FileCode className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* HTML insert — same shape as Markdown insert; html-to-docx.ts parses
          via DOMParser into the same DocxBlock model so undo / find / paragraph-
          level styling all work identically afterward. */}
      <ToolbarBtn
        disabled={false}
        title={disabled
          ? t('於文件結尾插入 HTML 內容', 'Insert HTML at end of document')
          : t('在此段後插入 HTML 內容', 'Insert HTML after this paragraph')}
        onClick={onInsertHtml}
      >
        <Code2 className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Page-break-before toggle — flags the active block so the serializer
          emits a page break ahead of it. Paragraph-family blocks only. */}
      <ToolbarBtn
        active={!!block?.pageBreakBefore}
        disabled={!isParaBlock}
        title={
          disabled
            ? tDisabled
            : !isParaBlock
              ? t('表格 / 圖片區塊不支援分頁符', 'Tables / images do not support page breaks')
              : block?.pageBreakBefore
                ? t('移除此段前的分頁符', 'Remove page break before this paragraph')
                : t('在此段前插入分頁符', 'Insert page break before this paragraph')
        }
        onClick={onTogglePageBreak}
      >
        <SeparatorHorizontal className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <PageSettingsMenu
        pageSize={pageSize}
        pageMargins={pageMargins}
        onSetPageSize={onSetPageSize}
        onTogglePageOrientation={onTogglePageOrientation}
        onSetPageMargins={onSetPageMargins}
      />
      {/* Find & Replace — Ctrl+F is already bound (line ~578), but the
          binding is invisible to mouse-first users. Surfacing it on the
          toolbar matches MarkdownToolbar and gives users a discoverable way
          to open it. Stays enabled regardless of `disabled`: search across
          the whole document is meaningful even with no active block. */}
      <ToolbarBtn disabled={false} title={t('尋找與取代 (Ctrl+F)', 'Find & Replace (Ctrl+F)')} onClick={onOpenFind}>
        <Search className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Go to paragraph — Ctrl+G is bound at line 778-799 but had zero UI
          surface, the lone outlier among the four editors:
            • PptxEditor surfaces it on the "第 N / M 張投影片" indicator
              (PptxEditor.tsx:1024)
            • XlsxEditor surfaces it on the Name Box tooltip
              (XlsxEditor.tsx:2299)
            • MarkdownEditor's Ctrl+G triggers CM6's native Go-to-line panel
              (which is its own visible UI when summoned)
          The Word toolbar already advertises every other shortcut it binds
          (Ctrl+B/I/U on bold/italic/underline, Ctrl+F on Find above) — this
          button completes that pattern and brings DocxEditor to discoverability
          parity. Pilcrow (¶) is the standard typographic mark for paragraph,
          which matches the dialog's "跳到第幾段？" label (line 1477). Stays
          enabled regardless of `disabled` — jumping to a paragraph is exactly
          how a user with no active block reaches one. */}
      <ToolbarBtn disabled={false} title={t('跳至段落… (Ctrl+G)', 'Go to paragraph… (Ctrl+G)')} onClick={onOpenGoto}>
        <Pilcrow className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Print-preview toggle — hides the page-level editing affordances
          (block grips, margin gutters, rings) so the user can sanity-check
          what the page actually looks like before exporting. R414: the
          toolbar itself and the nav pane stay visible — blocks remain
          editable in preview, so formatting actions keep working. Esc or
          a second click returns to edit mode; the floating pill too. */}
      <ToolbarBtn
        disabled={false}
        active={previewMode}
        title={
          previewMode
            ? t('退出預覽（Esc）', 'Exit preview (Esc)')
            : t('預覽列印外觀（隱藏編輯外框，Esc 返回）', 'Print preview (hide editor chrome, Esc to return)')
        }
        onClick={onTogglePreview}
      >
        <Eye className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Selection status pushed to the right edge so the (left-side)
          formatting controls stay visually anchored regardless of the
          status string's length. `ml-auto` reserves a flexible gutter
          between the action buttons and the readout. */}
      <span className="ml-auto text-[11px] text-muted-foreground whitespace-nowrap">
        {block ? t(`已選：${KIND_LABEL_ZH[block.kind]}`, `Selected: ${KIND_LABEL_EN[block.kind]}`) : t('點選一段文字以開始編輯', 'Click a paragraph to start editing')}
      </span>
    </div>
  );
}

/**
 * Toolbar dropdown for page-size / orientation / margins. Operates directly
 * on the model's `pageSize` + `pageMargins`, both stored in twips (twentieths
 * of a point) so they map 1:1 to the OOXML `<w:pgSz>` and `<w:pgMar>` units.
 */
function PageSettingsMenu({
  pageSize,
  pageMargins,
  onSetPageSize,
  onTogglePageOrientation,
  onSetPageMargins,
}: {
  pageSize: DocxPageSize;
  pageMargins: DocxPageMargins;
  onSetPageSize: (s: DocxPageSize) => void;
  onTogglePageOrientation: () => void;
  onSetPageMargins: (m: DocxPageMargins) => void;
}): JSX.Element {
  const t = useT();
  const isLandscape = pageSize.w > pageSize.h;
  // For preset matching we normalise so portrait/landscape of the same paper
  // both highlight the same row.
  const portrait =
    pageSize.w <= pageSize.h ? pageSize : { w: pageSize.h, h: pageSize.w };
  const matchedPreset = PAGE_PRESETS.find(
    (p) => p.w === portrait.w && p.h === portrait.h,
  );
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button
          type="button"
          title={t('頁面大小 / 方向 / 邊界', 'Page size / orientation / margins')}
          className="inline-flex items-center gap-1 h-7 px-2 rounded text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
        >
          <FileText className="h-3.5 w-3.5" />
          <span>{matchedPreset ? matchedPreset.label.split(' ')[0] : t('自訂', 'Custom')}</span>
          <span className="text-[10px] text-muted-foreground/70">
            {isLandscape ? t('橫', 'L') : t('直', 'P')}
          </span>
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="min-w-[14rem]">
        <DropdownMenuLabel>{t('紙張大小', 'Paper size')}</DropdownMenuLabel>
        {PAGE_PRESETS.map((p) => {
          const selected = matchedPreset?.id === p.id;
          return (
            <DropdownMenuItem
              key={p.id}
              onSelect={() => onSetPageSize(isLandscape ? { w: p.h, h: p.w } : { w: p.w, h: p.h })}
              className={cn(selected && 'bg-accent text-accent-foreground')}
            >
              {p.label}
            </DropdownMenuItem>
          );
        })}
        <DropdownMenuSeparator />
        <DropdownMenuLabel>{t('方向', 'Orientation')}</DropdownMenuLabel>
        <DropdownMenuItem onSelect={onTogglePageOrientation}>
          {isLandscape ? t('改為直向', 'Switch to portrait') : t('改為橫向', 'Switch to landscape')}
        </DropdownMenuItem>
        <DropdownMenuSeparator />
        <DropdownMenuLabel>{t('邊界', 'Margins')}</DropdownMenuLabel>
        {MARGIN_PRESETS.map((m) => {
          const selected = pageMargins.top === m.v && pageMargins.bottom === m.v
            && pageMargins.left === m.v && pageMargins.right === m.v;
          return (
            <DropdownMenuItem
              key={m.id}
              onSelect={() =>
                onSetPageMargins({
                  top: m.v,
                  right: m.v,
                  bottom: m.v,
                  left: m.v,
                  header: pageMargins.header,
                  footer: pageMargins.footer,
                })
              }
              className={cn(selected && 'bg-accent text-accent-foreground')}
            >
              {t(m.labelZh, m.labelEn)}
            </DropdownMenuItem>
          );
        })}
      </DropdownMenuContent>
    </DropdownMenu>
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
      // ToolbarBtn definitions in MarkdownToolbar / PptxEditor / XlsxEditor
      // (this round). Format buttons in this editor's toolbar (粗體 / 斜體 /
      // 底線 line 2043-2050、對齊四聯 line 2061-2071、連結 line 2053) all
      // pass `active={!!style.X}` / `active={block?.align === 'X'}`; before
      // this fix, SR users heard them as plain action buttons regardless of
      // whether the active selection was currently 粗體 — the visual
      // 「bg-primary/20 text-primary」 active-highlight was a sighted-only
      // signal. `aria-pressed={undefined}` renders no attribute, so action-
      // only callsites (the toolbar's 復原 / 重做 / 插入圖片 etc. that don't
      // pass `active`) keep clean action-button semantics.
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

function normalizeStyle(s: DocxBlockStyle): DocxBlockStyle | undefined {
  const out: DocxBlockStyle = {};
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  if (s.strikethrough) out.strikethrough = true;
  if (s.color) out.color = s.color;
  if (s.fontSize) out.fontSize = s.fontSize;
  if (s.fontFamily) out.fontFamily = s.fontFamily;
  if (s.highlight) out.highlight = s.highlight;
  if (s.lineSpacing) out.lineSpacing = s.lineSpacing;
  return Object.keys(out).length === 0 ? undefined : out;
}

function BlockRow({
  block,
  active,
  selected,
  position,
  pxPerTwip,
  previewMode,
  onFocus,
  onChangeRuns,
  onCycleKind,
  onInsertAfter,
  onRemove,
  onStartMove,
  onResetPosition,
}: {
  block: DocxBlock;
  active: boolean;
  /** Part of a marquee multi-selection — drives a stronger ring so the
   *  user can see the whole group that will move when they drag a grip. */
  selected: boolean;
  position: { xTwip: number; yTwip: number; wTwip: number } | undefined;
  pxPerTwip: number;
  previewMode: boolean;
  onFocus: () => void;
  /** New runs from RichBlock. Parent normalizes block.text via updateBlock. */
  onChangeRuns: (runs: DocxRun[]) => void;
  onCycleKind: () => void;
  onInsertAfter: () => void;
  onRemove: () => void;
  onStartMove: (e: React.PointerEvent) => void;
  onResetPosition: () => void;
}): JSX.Element {
  const t = useT(); // R409 — i18n for the empty-paragraph placeholder
  const textClass = useMemo(() => kindTextClass(block.kind), [block.kind]);
  const inlineStyle = useMemo<React.CSSProperties>(() => {
    const s = block.style ?? {};
    const isLink = !!block.link;
    // Hyperlink blocks force the canonical Word link visual (blue +
    // underline) so the user immediately sees that the run will export as
    // a hyperlink. Block-level color/underline still apply for non-links.
    const deco: string[] = [];
    if (s.underline || isLink) deco.push('underline');
    if (s.strikethrough) deco.push('line-through');
    return {
      fontWeight: s.bold ? 'bold' : undefined,
      fontStyle: s.italic ? 'italic' : undefined,
      textDecoration: deco.length > 0 ? deco.join(' ') : undefined,
      color: isLink ? '#0563C1' : s.color ? `#${s.color}` : undefined,
      textAlign: block.align,
      // pt → px (rough 1.333× factor for screen rendering); we don't try to be
      // exact since Word's render is very different anyway.
      fontSize: s.fontSize ? `${Math.round(s.fontSize * 1.333)}px` : undefined,
      fontFamily: withEmojiFallback(s.fontFamily),
      backgroundColor: s.highlight ? HIGHLIGHT_CSS[s.highlight] : undefined,
      lineHeight: s.lineSpacing,
    };
  }, [block.style, block.align, block.link]);
  const floatStyle = position
    ? {
        position: 'absolute' as const,
        left: position.xTwip * pxPerTwip,
        top: position.yTwip * pxPerTwip,
        width: position.wTwip * pxPerTwip,
      }
    : undefined;
  return (
    <>
      {/* Page-break marker — edit-mode affordance only; preview shows the
          page exactly as it would print (the editor doesn't paginate). */}
      {block.pageBreakBefore && !previewMode && (
        <div className="flex items-center gap-2 text-[10px] text-muted-foreground/70 select-none">
          <div className="flex-1 border-t border-dashed border-muted-foreground/40" />
          <span>{t('分頁', 'Page break')}</span>
          <div className="flex-1 border-t border-dashed border-muted-foreground/40" />
        </div>
      )}
    <div
      data-block-id={block.id}
      style={floatStyle}
      className={cn(
        // R413 — `relative` anchors the margin-gutter controls below. The
        // grip + kind chip used to be IN-FLOW flex children, pushing the
        // text ~90px right of where print/preview renders it; they now
        // live in an absolutely-positioned gutter in the page's left
        // margin so the text column starts exactly at the page margin in
        // both edit and preview modes (WYSIWYG alignment).
        'group relative flex gap-1 items-start rounded transition-colors',
        active && !previewMode && 'ring-1 ring-primary/30',
        position && !previewMode && 'bg-background/80 backdrop-blur-[1px] ring-1 ring-primary/20 px-1 z-10',
        // Marquee selection wins over the focus ring — when the user is in
        // multi-select mode, the unified blue/2px ring is the load-bearing
        // visual cue for "drag any grip moves this whole set together".
        selected && !previewMode && 'ring-2 ring-primary/70 bg-primary/5',
      )}
    >
      {!previewMode && (
        // R413 — hover/focus-revealed margin gutter (Notion-style). `invisible`
        // (not just opacity-0) so the hidden buttons can't swallow pointer
        // events meant for the page-margin marquee. Revealed while the row is
        // hovered, focused (caret inside), active, or marquee-selected.
        <div
          className={cn(
            'absolute right-full top-0 mr-1.5 flex items-start gap-1',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            'group-focus-within:visible group-focus-within:opacity-100',
            (active || selected) && 'visible opacity-100',
          )}
        >
        <button
          type="button"
          onPointerDown={onStartMove}
          // Suppress focus-shift on the compatibility mousedown that follows
          // pointerdown — otherwise tapping the grip without dragging blurs
          // the active RichBlock and loses the user's caret.
          onMouseDown={(e) => e.preventDefault()}
          title={tImp('拖曳到任意位置', 'Drag to any position')}
          className="mt-1 shrink-0 p-1 rounded text-muted-foreground/50 hover:text-foreground hover:bg-secondary cursor-grab active:cursor-grabbing touch-none"
        >
          <GripVertical className="h-3.5 w-3.5" />
        </button>
      {(() => {
        // R111 — disclose the cycle order + state-aware "next" target.
        // Prior shape was `title="點擊切換樣式"` — discloses the gesture
        // but nothing about the rotation, leaving three opaque facets:
        //
        //   1. Cycle order hidden. KIND_CYCLE at line 97-104 rotates
        //      paragraph → H1 → H2 → H3 → • 列表 → 1. 列表 → wrap.
        //      A user wanting H3 from a paragraph has no clue it's 3
        //      clicks; from • 列表 it would be 5 (or wrap to 4).
        //
        //   2. Next-target hidden. The codebase already establishes
        //      state-aware action tooltips on click-to-act buttons
        //      whose effect varies by state — XlsxEditor.tsx:2055-2058
        //      (R99) flips「插入圖片（錨定於 ${selectionAddr}）」 vs
        //      「未選取儲存格時錨定於 A1」; DocxEditor.tsx:2152-2199
        //      (R102) flips「在此段後插入圖片」vs「於文件結尾」;
        //      PptxEditor.tsx:1080-1119 (R107) names「目前投影片」.
        //      A click-to-rotate selector should name what it'll
        //      become on this click, not just say "it'll change".
        //
        //   3. Capability boundary undisclosed. KIND_LABEL at line
        //      106-118 has 11 entries; KIND_CYCLE at line 97-104 has
        //      only 6. BlockRow renders for non-table / non-image
        //      blocks (model.blocks.map at line 1621-1677 splits by
        //      kind), so out-of-cycle kinds reaching this button are
        //      heading4-6. A block imported with kind=heading4 shows
        //      "H4" as the visible label, but cycleKind at line 579
        //      computes `KIND_CYCLE[(indexOf(b.kind) + 1) % length]`
        //      → indexOf returns -1 → next = paragraph, dropping the
        //      heading semantics. The state-aware tooltip below makes
        //      this knowable on hover (the H4-block user sees「點擊
        //      改為「段落」」 and isn't surprised when the heading
        //      disappears).
        //
        // Wording combines the verb-first state-aware shape used by
        // R99/R102/R107 with the option-list-on-hover shape used by
        // TabBar.tsx:485 / R107 / R108 / R109. Cycle list uses → (U+2192)
        // arrows rather than the ` / ` separator from TabBar/R108/R109,
        // because here the order matters — separator choice has to
        // signal that paragraph→H1 is forward, not arbitrary. */
        const cycleIdx = KIND_CYCLE.indexOf(block.kind);
        const nextKind = KIND_CYCLE[(cycleIdx + 1) % KIND_CYCLE.length];
        return (
          <button
            type="button"
            onClick={onCycleKind}
            onMouseDown={(e) => e.preventDefault()}
            className="mt-1 shrink-0 text-xs px-1.5 py-0.5 rounded bg-secondary text-muted-foreground hover:text-foreground hover:bg-secondary/80 w-16 text-center"
            title={tImp(
              `點擊改為「${KIND_LABEL_ZH[nextKind]}」（循環：段落 → H1 → H2 → H3 → • 列表 → 1. 列表 → 段落）`,
              `Click to change to "${KIND_LABEL_EN[nextKind]}" (cycle: Paragraph → H1 → H2 → H3 → • List → 1. List → Paragraph)`,
            )}
          >
            {KIND_LABEL[block.kind]}
          </button>
        );
      })()}
        </div>
      )}
      <div className="flex-1 min-w-0 relative">
        <RichBlock
          blockId={block.id}
          runs={
            block.runs && block.runs.length > 0
              ? block.runs
              : [{ text: block.text }]
          }
          textClass={textClass}
          inlineStyle={inlineStyle}
          placeholder={block.kind === 'paragraph' ? t('輸入文字…', 'Type text…') : ''}
          active={active}
          previewMode={previewMode}
          onFocus={onFocus}
          onChange={onChangeRuns}
          onEnter={onInsertAfter}
        />
      </div>
      {!previewMode && (
        // R413 — moved out of the flex flow into the right page margin so the
        // text column spans the full printable width (matches print/preview).
        <div
          className={cn(
            'absolute left-full top-0 ml-1 flex flex-col gap-1 mt-1',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            'group-focus-within:visible group-focus-within:opacity-100',
            active && 'visible opacity-100',
          )}
        >
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onInsertAfter}
            className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground"
            title={tImp('在後方插入段落', 'Insert paragraph below')}
          >
            <Plus className="h-3 w-3" />
          </button>
          {position ? (
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={onResetPosition}
              className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground text-[10px] leading-none"
              title={tImp('回到原本流式位置', 'Restore to original flow position')}
            >
              ↺
            </button>
          ) : null}
          {/* R114 — name what gets deleted, mirroring the two sibling
              delete buttons in this same parent's three-way split.
              model.blocks.map at line 1621-1677 dispatches table-kind
              blocks to TableBlock, image-kind blocks to ImageBlockRow,
              and everything else (paragraph / H1-6 / bullet / numbered)
              to BlockRow — so this delete button is the third leg of a
              structurally-identical trio:

                • TableBlock delete (line 2917)    "刪除整個表格" ✅
                • ImageBlockRow delete (line 3187) "刪除圖片"     ✅
                • BlockRow delete (this button)    "刪除"          ← was lone holdout

              All three buttons are byte-identical in markup
              (`<Trash2>` icon + `onClick={onRemove}` + same className +
              same onMouseDown preventDefault), placed at the bottom of
              the same hover-revealed action cluster. Two name the
              object; this one didn't. Same in-file 2-vs-1 framing R107
              / R108 / R109 / R110 / R112 used to close their respective
              same-row holdouts.

              Wording mirrors「刪除整個表格」 verbatim in structure
              (「刪除整個」+ object). All six BlockRow kinds are <w:p>
              paragraphs in DOCX (the cycle button R111 just polished
              labels them as 段落/H1/H2/H3/• 列表/1. 列表 in user-facing
              text, but all share the same paragraph block primitive),
              and DocxEditor.tsx:1947 already establishes「整段」 as the
              user-facing vocabulary for "the whole paragraph block":
              「對齊 / 字色 / 字型仍以整段為單位」. The cycle button
              right next to this delete button shows the current style
              (e.g. "H1") so the user can read both together: "this is
              styled as H1, and 刪除 removes the whole paragraph block".
              Static rather than state-aware (unlike R111's cycle button
              which IS state-dependent — the next-kind target genuinely
              varies) because the action target is identical across
              kinds: always remove the current block. Static parity with
              the table/image siblings is the stronger precedent here. */}
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onRemove}
            className="p-1 rounded hover:bg-destructive/20 text-muted-foreground hover:text-destructive"
            title={tImp('刪除整個段落', 'Delete entire paragraph')}
          >
            <Trash2 className="h-3 w-3" />
          </button>
        </div>
      )}
    </div>
    </>
  );
}

function kindTextClass(kind: DocxBlockKind): string {
  switch (kind) {
    case 'heading1':
      return 'text-2xl font-bold';
    case 'heading2':
      return 'text-xl font-bold';
    case 'heading3':
      return 'text-lg font-semibold';
    case 'heading4':
      return 'text-base font-semibold';
    case 'heading5':
      return 'text-sm font-semibold';
    case 'heading6':
      return 'text-xs font-semibold';
    case 'bullet':
      return 'text-sm pl-4';
    case 'numbered':
      return 'text-sm pl-4';
    case 'table':
    case 'paragraph':
    default:
      return 'text-sm';
  }
}

function TableBlock({
  block,
  active,
  selected,
  position,
  pxPerTwip,
  previewMode,
  onFocus,
  onChangeRows,
  onInsertAfter,
  onRemove,
  onStartMove,
  onResetPosition,
}: {
  block: DocxBlock;
  active: boolean;
  selected: boolean;
  position: { xTwip: number; yTwip: number; wTwip: number } | undefined;
  pxPerTwip: number;
  previewMode: boolean;
  onFocus: () => void;
  onChangeRows: (rows: string[][]) => void;
  onInsertAfter: () => void;
  onRemove: () => void;
  onStartMove: (e: React.PointerEvent) => void;
  onResetPosition: () => void;
}): JSX.Element {
  const t = useT(); // R409 — i18n for row/column controls
  const rows = block.rows ?? [['']];
  const cols = rows[0]?.length ?? 1;

  const setCell = (r: number, c: number, value: string) => {
    const next = rows.map((row, ri) =>
      ri === r ? row.map((cell, ci) => (ci === c ? value : cell)) : row,
    );
    onChangeRows(next);
  };
  /**
   * Multi-cell paste — when the clipboard carries TSV (Excel/Google Sheets
   * convention: tabs between columns, newlines between rows) AND spans more
   * than one cell, tile it across adjacent cells starting at the focused
   * one. Without this, copying a 3×2 selection from Excel and pasting into
   * a Word table dumped the whole `"A\tB\tC\nD\tE\tF"` string into a
   * single cell — visibly corrupted, and the user had to manually re-split.
   *
   * Single-line paste (no \t, no \n) defers to the browser's native input
   * paste so undo / IME composition / autocorrect still work normally —
   * we only intervene for the actual tabular case.
   *
   * Auto-extends rows / cols when the paste exceeds the current table size
   * so a 5-column paste into a 3-column table fills out, not truncates.
   */
  const onCellPaste = (e: React.ClipboardEvent<HTMLInputElement>, r: number, c: number) => {
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (!text || (!text.includes('\t') && !text.includes('\n'))) return;
    // Strip a single trailing newline that Excel always appends — without
    // this we'd add an empty row at the bottom of every paste.
    const trimmed = text.replace(/\r\n/g, '\n').replace(/\n$/, '');
    const grid = trimmed.split('\n').map((line) => line.split('\t'));
    if (grid.length === 0) return;
    e.preventDefault();
    const pasteRows = grid.length;
    // R371 — compute pasteCols via reduce. `Math.max(...arr)` spread blows
    // V8's argument-count limit (~65K) for very large pastes — clipboard
    // payloads from「複製整欄 Excel 資料」routinely exceed that for log /
    // export dumps. Realistic trigger: user copies a 100K-row data export
    // from Excel and pastes into a Word table cell here; pasteRows
    // calculation is fine (array.length is O(1)), but pasteCols's spread
    // would throw RangeError, the catch handler isn't present, and the
    // paste silently fails OR the renderer crashes. Same R328 / R370 fix
    // shape; reduce has no spread, handles any array size.
    const pasteCols = grid.reduce<number>((m, row) => Math.max(m, row.length), 0);
    const newRowCount = Math.max(rows.length, r + pasteRows);
    const newColCount = Math.max(cols, c + pasteCols);
    const next: string[][] = [];
    for (let ri = 0; ri < newRowCount; ri += 1) {
      const out: string[] = [];
      for (let ci = 0; ci < newColCount; ci += 1) {
        const dr = ri - r;
        const dc = ci - c;
        if (dr >= 0 && dr < pasteRows && dc >= 0 && dc < pasteCols) {
          out.push(grid[dr][dc] ?? '');
        } else {
          out.push(rows[ri]?.[ci] ?? '');
        }
      }
      next.push(out);
    }
    onChangeRows(next);
  };
  const addRow = () => onChangeRows([...rows, new Array(cols).fill('')]);
  const removeRow = (r: number) => {
    if (rows.length <= 1) return;
    onChangeRows(rows.filter((_, ri) => ri !== r));
  };
  const addCol = () => onChangeRows(rows.map((row) => [...row, '']));
  const removeCol = (c: number) => {
    if (cols <= 1) return;
    onChangeRows(rows.map((row) => row.filter((_, ci) => ci !== c)));
  };

  const floatStyle = position
    ? {
        position: 'absolute' as const,
        left: position.xTwip * pxPerTwip,
        top: position.yTwip * pxPerTwip,
        width: position.wTwip * pxPerTwip,
      }
    : undefined;
  /**
   * Tab / Shift+Tab cycles through cells in row-major order. At the last
   * cell of the last row, Tab adds a fresh row and lands on its first cell
   * (matches Word's table behavior — the user can keep typing without
   * lifting their hands).
   */
  // Synchronous focus — the destination cell already exists in the DOM
  // because we're moving inside the current rows. Previously this was
  // wrapped in requestAnimationFrame, which let rapid Tab presses race
  // each other: a typist hitting Tab three times in <16 ms would fire
  // three handlers from the SAME source cell (focus hadn't moved yet),
  // all scheduling rAFs to focus cell N+1. They'd converge on cell N+1
  // and the user would land two cells short of where they expected.
  const focusCellSync = (r: number, c: number) => {
    const el = document.querySelector(
      `[data-table-cell="${block.id}:${r}:${c}"]`,
    ) as HTMLInputElement | null;
    if (!el) return;
    el.focus();
    el.select();
  };
  // Deferred focus — used only when we just appended a row. The new cell
  // doesn't exist in the DOM yet (React hasn't re-rendered), so a sync
  // querySelector returns null. rAF waits for the next paint to sweep the
  // updated rows in. Same need for paste auto-extend below.
  const focusCellDeferred = (r: number, c: number) => {
    requestAnimationFrame(() => focusCellSync(r, c));
  };
  const onCellKeyDown = (e: React.KeyboardEvent<HTMLInputElement>, r: number, c: number) => {
    // R235 — IME composition guard, mirrors the R231 / R232 / R233 / R234
    // family. Korean Hangul IMEs use Tab to cycle through candidate
    // characters during composition (Enter confirms the selection); some
    // Japanese IMEs also bind Tab inside the candidate window. Without
    // this guard, table-cell Tab override (preventDefault + focus next
    // cell) hijacks the IME's Tab, breaking candidate navigation and
    // forcing the user to commit whatever raw jamo / kana they were
    // mid-composing into the previous cell. The cell's `<input>` is a
    // CJK typing surface for any Word table content in Chinese /
    // Japanese / Korean, so the guard belongs at the top of the
    // handler.
    if (e.nativeEvent.isComposing) return;
    if (e.key === 'Tab') {
      e.preventDefault();
      const total = rows.length * cols;
      const linear = r * cols + c;
      if (e.shiftKey) {
        if (linear === 0) return;
        const prev = linear - 1;
        focusCellSync(Math.floor(prev / cols), prev % cols);
      } else {
        if (linear === total - 1) {
          // Append a new row and land on its first cell.
          onChangeRows([...rows, new Array(cols).fill('')]);
          focusCellDeferred(rows.length, 0);
          return;
        }
        const next = linear + 1;
        focusCellSync(Math.floor(next / cols), next % cols);
      }
    }
  };

  return (
    <div
      data-block-id={block.id}
      style={floatStyle}
      className={cn(
        // R413 — same margin-gutter layout as BlockRow: controls float in
        // the page margins so the table spans the true printable width.
        'group relative flex gap-1 items-start rounded transition-colors',
        active && !previewMode && 'ring-1 ring-primary/30',
        position && !previewMode && 'bg-background/80 backdrop-blur-[1px] ring-1 ring-primary/20 px-1 z-10',
        selected && !previewMode && 'ring-2 ring-primary/70 bg-primary/5',
      )}
    >
      {!previewMode && (
        <div
          className={cn(
            'absolute right-full top-0 mr-1.5 flex items-start gap-1',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            'group-focus-within:visible group-focus-within:opacity-100',
            (active || selected) && 'visible opacity-100',
          )}
        >
          <button
            type="button"
            onPointerDown={onStartMove}
            onMouseDown={(e) => e.preventDefault()}
            title={tImp('拖曳到任意位置', 'Drag to any position')}
            className="mt-1 shrink-0 p-1 rounded text-muted-foreground/50 hover:text-foreground hover:bg-secondary cursor-grab active:cursor-grabbing touch-none"
          >
            <GripVertical className="h-3.5 w-3.5" />
          </button>
          <span className="mt-1 shrink-0 text-xs px-1.5 py-0.5 rounded bg-secondary text-muted-foreground w-16 text-center">
            {KIND_LABEL.table}
          </span>
        </div>
      )}
      <div className="flex-1 min-w-0">
        <table className="border-collapse w-full text-xs">
          <tbody>
            {rows.map((row, r) => (
              <tr key={r} className="group/row">
                {row.map((cell, c) => (
                  <td key={c} className="border border-border p-0 bg-background">
                    <input
                      type="text"
                      data-table-cell={`${block.id}:${r}:${c}`}
                      value={cell}
                      // Mark active on focus so the toolbar's "selected"
                      // status reflects the table block while the user is
                      // editing a cell. Select-all on focus mirrors Excel/
                      // Word behavior — entering a cell highlights its text
                      // so the user can immediately overtype. Preview mode
                      // intentionally leaves the cell editable — only the
                      // grip / kind tag / +/- chrome around it disappears.
                      onFocus={(e) => {
                        onFocus();
                        e.currentTarget.select();
                      }}
                      onKeyDown={(e) => onCellKeyDown(e, r, c)}
                      onPaste={(e) => onCellPaste(e, r, c)}
                      onChange={(e) => setCell(r, c, e.target.value)}
                      className="w-full px-2 py-1 bg-transparent outline-none focus:bg-primary/10"
                    />
                  </td>
                ))}
                {!previewMode && (
                  <td className="opacity-0 group-hover/row:opacity-100 pl-1">
                    {/* Disable + explain at the 1-row floor: removeRow at
                        line 2528 early-returns when `rows.length <= 1`,
                        so clicking on a single-row table silently does
                        nothing — destructive intent + zero feedback is
                        the same failure mode XlsxEditor's row delete had
                        before R49. Mirror that fix here so the table's
                        "刪最右欄" sibling at line 2686 (which already
                        sets `disabled={cols <= 1}` but had no tooltip)
                        and this trash icon both follow the same pattern:
                        disable at boundary, tooltip explains why. */}
                    <button
                      type="button"
                      disabled={rows.length <= 1}
                      title={rows.length <= 1 ? t('至少要保留一列', 'At least one row must remain') : t('刪除此列', 'Delete this row')}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => removeRow(r)}
                      className="p-0.5 rounded text-muted-foreground hover:text-destructive disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:text-muted-foreground"
                    >
                      <Trash2 className="h-3 w-3" />
                    </button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
        {!previewMode && (
          <div className="flex gap-2 mt-1 text-[11px] text-muted-foreground">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={addRow}
              className="hover:text-foreground inline-flex items-center gap-1"
            >
              <Rows className="h-3 w-3" /> {t('新增列', 'Add row')}
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={addCol}
              className="hover:text-foreground inline-flex items-center gap-1"
            >
              <Columns className="h-3 w-3" /> {t('新增欄', 'Add column')}
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => removeCol(cols - 1)}
              disabled={cols <= 1}
              title={cols <= 1 ? t('至少要保留一欄', 'At least one column must remain') : t('刪除最右側的欄', 'Delete the rightmost column')}
              className="hover:text-foreground inline-flex items-center gap-1 disabled:opacity-50"
            >
              {t('刪最右欄', 'Delete last column')}
            </button>
          </div>
        )}
      </div>
      {!previewMode && (
        // R413 — right-margin action gutter, mirrors BlockRow.
        <div
          className={cn(
            'absolute left-full top-0 ml-1 flex flex-col gap-1 mt-1',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            'group-focus-within:visible group-focus-within:opacity-100',
            active && 'visible opacity-100',
          )}
        >
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onInsertAfter}
            className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground"
            title={tImp('在後方插入段落', 'Insert paragraph below')}
          >
            <Plus className="h-3 w-3" />
          </button>
          {position ? (
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={onResetPosition}
              className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground text-[10px] leading-none"
              title={tImp('回到原本流式位置', 'Restore to original flow position')}
            >
              ↺
            </button>
          ) : null}
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onRemove}
            className="p-1 rounded hover:bg-destructive/20 text-muted-foreground hover:text-destructive"
            title={tImp('刪除整個表格', 'Delete entire table')}
          >
            <Trash2 className="h-3 w-3" />
          </button>
        </div>
      )}
    </div>
  );
}

// ── image insertion helpers ─────────────────────────────────────────────

/**
 * Open a hidden `<input type=file>` and resolve to the chosen image File
 * or null if the user cancels. Picker dialog is fire-and-forget; we don't
 * keep the input element around.
 */
function pickImageFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png,image/jpeg,image/gif,image/bmp';
    let resolved = false;
    input.onchange = () => {
      resolved = true;
      resolve(input.files?.[0] ?? null);
    };
    // The browser doesn't fire 'cancel' reliably across all platforms; the
    // focus trick below approximates "the dialog closed without a pick" so
    // we don't leak unresolved promises.
    window.addEventListener(
      'focus',
      () => {
        setTimeout(() => {
          if (!resolved) resolve(null);
        }, 300);
      },
      { once: true },
    );
    input.click();
  });
}

function mimeForKind(kind: DocxImage['mediaType']): string {
  switch (kind) {
    case 'jpg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    default:
      return 'image/png';
  }
}

/**
 * Encode a Uint8Array as base64 without busting the JS call-stack on large
 * images. We chunk to ~32k characters per `String.fromCharCode.apply` call
 * because some engines refuse arrays beyond ~64k. `btoa` handles the rest.
 */
function uint8ToBase64(bytes: Uint8Array): string {
  const CHUNK = 0x8000;
  let s = '';
  for (let i = 0; i < bytes.length; i += CHUNK) {
    s += String.fromCharCode.apply(
      null,
      Array.from(bytes.subarray(i, Math.min(i + CHUNK, bytes.length))),
    );
  }
  return btoa(s);
}

/**
 * Decode a `data:` URL with an off-screen `Image()` and return the natural
 * pixel dimensions. Falls back to a square 4-inch placeholder (384px ≈
 * 4" at 96 dpi) if decode fails so the insert never silently no-ops.
 */
function readNaturalSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () =>
      resolve({ width: img.naturalWidth || 384, height: img.naturalHeight || 384 });
    img.onerror = () => resolve({ width: 384, height: 384 });
    img.src = dataUrl;
  });
}

/**
 * Render an inline image block. Mirrors TableBlock's outer chrome (drag
 * handle, kind label, position-anchor controls) so users get the same
 * affordances they're already used to. Click-and-drag the bottom-right
 * corner to resize while preserving aspect ratio (Word/Pages convention);
 * Shift-drag is reserved for free-aspect resize per the same convention.
 */
function ImageBlockRow({
  block,
  active,
  selected,
  position,
  pxPerTwip,
  previewMode,
  onFocus,
  onResize,
  onInsertAfter,
  onRemove,
  onStartMove,
  onResetPosition,
}: {
  block: DocxBlock;
  active: boolean;
  selected: boolean;
  position: { xTwip: number; yTwip: number; wTwip: number } | undefined;
  pxPerTwip: number;
  previewMode: boolean;
  onFocus: () => void;
  onResize: (widthPx: number, heightPx: number) => void;
  onInsertAfter: () => void;
  onRemove: () => void;
  onStartMove: (e: React.PointerEvent) => void;
  onResetPosition: () => void;
}): JSX.Element {
  const t = useT(); // R409 — i18n for the missing-image fallback
  const img = block.image;
  // Live drag-resize state. Committed via onResize on pointerup.
  const [previewSize, setPreviewSize] = useState<{ w: number; h: number } | null>(null);

  if (!img) {
    return (
      <div data-block-id={block.id} className="text-xs text-muted-foreground italic px-2 py-1">
        {t('（圖片資料缺失）', '(Image data missing)')}
      </div>
    );
  }

  const startResize = (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startY = e.clientY;
    const startW = img.widthPx;
    const startH = img.heightPx;
    const aspect = startH / Math.max(1, startW);
    const onMoveDoc = (ev: PointerEvent) => {
      const dx = ev.clientX - startX;
      const dy = ev.clientY - startY;
      // Default: aspect-locked from the dominant axis. Shift releases the
      // lock, matching Word/Pages "free resize" modifier convention.
      let w = Math.max(32, startW + dx);
      let h = Math.max(32, startH + dy);
      if (!ev.shiftKey) {
        // Lock to whichever axis the user moved more; the other axis is
        // derived from `aspect`. This feels more natural than always
        // using width because dragging straight down should grow height.
        if (Math.abs(dx) >= Math.abs(dy)) {
          h = Math.max(32, Math.round(w * aspect));
        } else {
          w = Math.max(32, Math.round(h / Math.max(0.0001, aspect)));
        }
      }
      setPreviewSize({ w, h });
    };
    const onUpDoc = () => {
      window.removeEventListener('pointermove', onMoveDoc);
      window.removeEventListener('pointerup', onUpDoc);
      setPreviewSize((cur) => {
        if (cur) onResize(cur.w, cur.h);
        return null;
      });
    };
    window.addEventListener('pointermove', onMoveDoc);
    window.addEventListener('pointerup', onUpDoc);
  };

  const floatStyle = position
    ? {
        position: 'absolute' as const,
        left: position.xTwip * pxPerTwip,
        top: position.yTwip * pxPerTwip,
        width: position.wTwip * pxPerTwip,
      }
    : undefined;

  // Honour paragraph alignment for the image's horizontal placement —
  // mirrors how Word renders an inline image inside a centered paragraph.
  const justify =
    block.align === 'center'
      ? 'justify-center'
      : block.align === 'right'
        ? 'justify-end'
        : 'justify-start';

  const liveW = previewSize?.w ?? img.widthPx;
  const liveH = previewSize?.h ?? img.heightPx;

  return (
    <div
      data-block-id={block.id}
      style={floatStyle}
      onClick={onFocus}
      className={cn(
        // R413 — `relative` anchors the margin gutters (see BlockRow).
        'group relative flex gap-1 items-start rounded transition-colors',
        active && !previewMode && 'ring-1 ring-primary/30',
        position && !previewMode && 'bg-background/80 backdrop-blur-[1px] ring-1 ring-primary/20 px-1 z-10',
        selected && !previewMode && 'ring-2 ring-primary/70 bg-primary/5',
      )}
    >
      {!previewMode && (
        // R413 — margin gutter, mirrors BlockRow/TableBlock.
        <div
          className={cn(
            'absolute right-full top-0 mr-1.5 flex items-start gap-1',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            'group-focus-within:visible group-focus-within:opacity-100',
            (active || selected) && 'visible opacity-100',
          )}
        >
          <button
            type="button"
            onPointerDown={onStartMove}
            onMouseDown={(e) => e.preventDefault()}
            title={tImp('拖曳到任意位置', 'Drag to any position')}
            className="mt-1 shrink-0 p-1 rounded text-muted-foreground/50 hover:text-foreground hover:bg-secondary cursor-grab active:cursor-grabbing touch-none"
          >
            <GripVertical className="h-3.5 w-3.5" />
          </button>
          <span className="mt-1 shrink-0 text-xs px-1.5 py-0.5 rounded bg-secondary text-muted-foreground w-16 text-center">
            {KIND_LABEL.image}
          </span>
        </div>
      )}
      <div className={cn('flex-1 min-w-0 flex items-start', justify)}>
        <div className="relative inline-block" style={{ width: liveW, height: liveH }}>
          <img
            src={img.dataUrl}
            alt=""
            draggable={false}
            className={cn(
              'block w-full h-full object-contain rounded-sm border border-transparent',
              !previewMode && 'group-hover:border-border',
            )}
          />
          {/* Live size badge during resize so the user has precise
              feedback before releasing — same idea as PptxEditor's frame
              readout. 96 dpi = 1in / 96 px ≈ 0.75 pt per pixel. */}
          {previewSize && (
            <div className="absolute top-0 left-0 px-1.5 py-0.5 text-[10px] font-mono rounded-br bg-zinc-900/90 text-zinc-50 pointer-events-none z-10">
              {liveW} × {liveH} px
            </div>
          )}
          {/* Bottom-right resize handle. Aspect-locked by default; Shift
              for free aspect — matches Word / Pages / Keynote convention. */}
          {!previewMode && (
            <span
              role="presentation"
              onPointerDown={startResize}
              title={tImp('拖曳調整大小 · Shift 解除等比例', 'Drag to resize · Shift to unlock aspect ratio')}
              className="absolute -bottom-1 -right-1 h-2.5 w-2.5 rounded-sm bg-background border border-primary cursor-nwse-resize opacity-0 group-hover:opacity-100"
            />
          )}
        </div>
      </div>
      {!previewMode && (
        // R413 — right-margin action gutter, mirrors BlockRow/TableBlock.
        // The cluster keeps its horizontal layout but is now revealed on
        // hover/active like its siblings (it sits in the page margin, so
        // an always-visible cluster would float oddly next to the page).
        <div
          className={cn(
            'absolute left-full top-0 ml-1 mt-1 flex items-center gap-0.5',
            'opacity-0 invisible transition-opacity',
            'group-hover:visible group-hover:opacity-100',
            (active || selected) && 'visible opacity-100',
          )}
        >
          {/* R115 — match the two sibling block-action clusters that
              also live in this same parent's three-way kind split:

                • BlockRow at line 2566-2574 (paragraph / H1-6 / list)
                  carries `<Plus>` + `title="在後方插入段落"`
                • TableBlock at line 2928-2936 (table)
                  carries `<Plus>` + `title="在後方插入段落"`
                • ImageBlockRow (this cluster) was the lone holdout —
                  the prop was wired by the parent dispatcher at line
                  1655 (`onInsertAfter={() => insertAfter(b.id)}`)
                  AND declared on the props type at line 3069
                  (`onInsertAfter: () => void`), but the destructuring
                  silently omitted it and no Plus button rendered.
                  TypeScript flagged it (`'onInsertAfter' is declared
                  but its value is never read.`) once destructured —
                  i.e. the diagnostics tooling already knew the prop
                  was orphaned, just nobody had wired the consumer.

              User impact: a doc ending with an image stranded the
              user. Both paragraph and table siblings have a Plus
              button on hover for "continue with another paragraph
              after this thing"; an image at the document tail had
              no equivalent affordance. The cycle button R111 just
              polished and the delete button R114 just normalised
              both rely on the assumption that all three block
              clusters present the same 3-action shape — this Plus
              button restores that contract.

              insertAfter at line 389-400 always creates a fresh
              `kind: 'paragraph'` block regardless of the source
              block's kind (no `fresh` arg passed), so the tooltip
              「在後方插入段落」 reused verbatim from the two siblings
              describes what the user will get.

              Layout note (updated R413): the cluster keeps its
              horizontal `flex items-center gap-0.5` arrangement but now
              lives in the right page-margin gutter and is hover/active-
              revealed like BlockRow/TableBlock's vertical clusters.
              Plus button slots in as the first item, matching the
              sibling order Plus → ↺ → 🗑️. */}
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onInsertAfter}
            className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground"
            title={tImp('在後方插入段落', 'Insert paragraph below')}
          >
            <Plus className="h-3 w-3" />
          </button>
          {position ? (
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={onResetPosition}
              className="p-1 rounded hover:bg-secondary text-muted-foreground hover:text-foreground text-[10px] leading-none"
              title={tImp('回到原本流式位置', 'Restore to original flow position')}
            >
              ↺
            </button>
          ) : null}
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={onRemove}
            className="p-1 rounded hover:bg-destructive/20 text-muted-foreground hover:text-destructive"
            title={tImp('刪除圖片', 'Delete image')}
          >
            <Trash2 className="h-3 w-3" />
          </button>
        </div>
      )}
    </div>
  );
}

/**
 * Inline floating link-edit dialog for the DOCX block-level link toggle.
 * Mirrors the Round 22 markdown LinkInsertDialog: pinned to upper-right of
 * the editor surface (matches Find / Go-to placement), Esc cancels, Enter
 * commits, empty URL clears the link. URL field auto-focuses + selects on
 * mount so the user can immediately overtype an existing value or just
 * start typing for a fresh one. The "清除連結" footer button is shown only
 * when there's an existing link to remove — for a fresh insert it would
 * just be a confusing no-op.
 */
function LinkEditDialog({
  defaultUrl,
  onClose,
  onCommit,
}: {
  defaultUrl: string;
  onClose: () => void;
  onCommit: (trimmedUrl: string) => void;
}) {
  const [url, setUrl] = useState(defaultUrl);
  const [urlError, setUrlError] = useState(false);
  const urlRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  // Insert vs edit mode is inferred from defaultUrl. In edit mode an empty
  // submit legitimately means "clear the link" (and there's a dedicated
  // 清除連結 button for that path); in insert mode an empty submit is a
  // no-op that silently closes the dialog — guard it so the user sees why.
  const isInsert = !defaultUrl;
  // Focus restore on close — same contract as GoToDialog / LinkInsertDialog
  // (Markdown). Without this, an Esc-cancel left focus on <body>: the next
  // Ctrl+S / Ctrl+Z / arrow-key didn't reach the editor until the user
  // clicked back in. Only restore when nothing else has claimed focus
  // post-close — onCommit may legitimately move focus into the editor as
  // the link transaction lands.
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
      urlRef.current?.focus();
      urlRef.current?.select();
    });
  }, []);
  // Document-level Esc — closes even when focus has wandered out of the
  // dialog (e.g. user clicked back into the editor to verify what link
  // they're editing). Mirrors GoToDialog / FindReplaceDialog. Focus
  // inside the dialog is handled by the inline onKeyDown below; focus
  // in a different editable surface defers to that surface so we don't
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
    if (!trimmed && isInsert) {
      setUrlError(true);
      urlRef.current?.focus();
      return;
    }
    onCommit(trimmed);
  };
  return (
    <div
      ref={dialogRef}
      role="dialog"
      aria-label={tImp('編輯連結', 'Edit link')}
      className="absolute top-3 right-3 z-30 w-[320px] rounded-md border bg-background shadow-lg p-3 text-sm space-y-2"
      onKeyDown={(e) => {
        // R234 — IME composition guard, mirrors R231 / R232 / R233. The
        // 連結 dialog has a label-text <input> below; CJK users typing a
        // Chinese link label「點擊看詳情」 via bopomofo / pinyin trigger
        // candidate confirmation with bare Enter, which bubbles here and
        // fires submit() against the raw IME buffer instead of the chosen
        // CJK glyphs. Same isComposing short-circuit pattern.
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
      <div className="text-xs font-medium">{defaultUrl ? tImp('編輯連結', 'Edit link') : tImp('插入連結', 'Insert link')}</div>
      <label className="block">
        <span className="block text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
          {tImp('連結網址', 'Link URL')}
        </span>
        <input
          ref={urlRef}
          type="text"
          value={url}
          onChange={(e) => {
            setUrl(e.target.value);
            if (urlError) setUrlError(false);
          }}
          aria-invalid={urlError ? true : undefined}
          placeholder="https://example.com"
          className={cn(
            'w-full px-2 py-1 text-xs border rounded bg-background outline-none focus:ring-1',
            urlError
              ? 'border-destructive/60 focus:ring-destructive'
              : 'focus:ring-primary',
          )}
        />
        {urlError && (
          <div className="mt-1 text-[10px] text-destructive">
            {tImp('請先填入網址再插入', 'Enter a URL before inserting')}
          </div>
        )}
      </label>
      <div className="flex items-center justify-between pt-1">
        <span className="text-[10px] text-muted-foreground">{tImp('Esc 取消 · Enter', 'Esc to cancel · Enter to')} {defaultUrl ? tImp('套用', 'apply') : tImp('插入', 'insert')}</span>
        <div className="flex items-center gap-1">
          {defaultUrl && (
            <button
              type="button"
              onClick={() => onCommit('')}
              className="px-2 py-1 text-xs rounded border hover:bg-secondary"
              title={tImp('移除目前連結', 'Remove current link')}
            >
              {tImp('清除連結', 'Clear link')}
            </button>
          )}
          {/* Sibling-shortcut-in-tooltip parity: every primary-action button
              in the app advertises its keystroke on the button itself —
              AIPanel.tsx:1162「送出 (Enter)」, SettingsDialog.tsx:258「儲存
              API key (Enter)」, DiffPreview.tsx:89「套用變更 (Ctrl+Enter)」,
              GoToDialog.tsx:191「跳轉 (Enter)」(the R47 fix), FindReplaceDialog
              replace button. The button is the most stable surface (footer
              hints can flip on error states — see R47's GoToDialog comment),
              so the keystroke belongs here regardless of the footer line at
              3197. The internal smoking gun: the secondary 清除連結 button at
              line 3204 already carries `title="移除目前連結"` — this primary
              button being unmarked while its sibling explains itself was a
              within-dialog inconsistency. Tooltip mirrors the visible label
              swap (套用 / 插入) so edit-vs-insert mode reads identically. */}
          <button
            type="button"
            onClick={submit}
            title={defaultUrl ? tImp('套用變更 (Enter)', 'Apply changes (Enter)') : tImp('插入連結 (Enter)', 'Insert link (Enter)')}
            className="px-2 py-1 text-xs rounded bg-primary text-primary-foreground hover:bg-primary/90"
          >
            {defaultUrl ? tImp('套用', 'Apply') : tImp('插入', 'Insert')}
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * Insert-Markdown dialog — paste/type markdown, convert to DocxBlock[] on commit.
 *
 * The conversion is a one-shot: once inserted, the blocks become regular Word
 * paragraphs/lists/tables that the user can edit with the normal toolbar.
 * Round-trip back to markdown is not supported (the docx adapter doesn't carry
 * a markdown-source field), so this is intentionally a one-way ingestion.
 *
 * Keyboard:
 *   - Esc       — cancel
 *   - Ctrl+Enter — commit (Enter alone inserts a newline, which is what users
 *     expect from a multi-line markdown editor)
 */
function MarkdownInsertDialog({
  onClose,
  onCommit,
}: {
  onClose: () => void;
  // R341 — mirror HtmlInsertDialog's contract: caller can return a
  // human-readable error string to keep the dialog open + display the
  // message inline. Previously typed `(source) => void`, which forced
  // every error path (parse-empty, R341's fresh-empty) to either
  // silently close or call a toast outside the dialog's own UI
  // channel. Returning `undefined` still means "OK, close".
  onCommit: (source: string) => string | undefined;
}) {
  const t = useT(); // R409 — i18n for label + sample placeholder
  const [source, setSource] = useState('');
  // R341 — error state widened from boolean to `string | null` so the
  // inline message text comes from the caller (or the local empty-
  // input check), not a hardcoded literal. Same shape HtmlInsertDialog
  // adopted post-R335.
  const [error, setError] = useState<string | null>(null);
  const taRef = useRef<HTMLTextAreaElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  // Focus restore on close — same contract as LinkEditDialog above /
  // GoToDialog / FindReplaceDialog. Esc-cancel without this stranded the
  // user on <body> after they closed the dialog, breaking the next
  // Ctrl+S / arrow-key until they clicked back into the document.
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
    queueMicrotask(() => taRef.current?.focus());
  }, []);

  // Document-level Esc fallback — same pattern as LinkEditDialog. Lets users
  // dismiss the dialog with Esc even if focus has wandered (e.g. they clicked
  // into the editor to copy something out and forgot to come back).
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
    const trimmed = source.trim();
    if (!trimmed) {
      // R341 — was `setError(true)` with a hardcoded literal at the
      // display site; now sets the message text directly so all error
      // paths route through one display channel.
      setError(tImp('請先填入 Markdown 內容再插入', 'Enter Markdown content before inserting'));
      taRef.current?.focus();
      return;
    }
    const err = onCommit(source);
    if (err) {
      setError(err);
      taRef.current?.focus();
    }
  };

  return (
    <div
      ref={dialogRef}
      role="dialog"
      aria-label={tImp('插入 Markdown', 'Insert Markdown')}
      className="absolute top-3 right-3 z-30 w-[460px] rounded-md border bg-background shadow-lg p-3 text-sm space-y-2"
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
          e.stopPropagation();
          e.preventDefault();
          submit();
        }
      }}
    >
      <div className="text-xs font-medium">{tImp('插入 Markdown 內容', 'Insert Markdown content')}</div>
      <label className="block">
        <span className="block text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
          {t('Markdown 原始碼', 'Markdown source')}
        </span>
        <textarea
          ref={taRef}
          value={source}
          onChange={(e) => {
            setSource(e.target.value);
            // R341 — clear the message on any keystroke; same dismiss-
            // on-edit pattern HtmlInsertDialog uses (line 3984).
            if (error) setError(null);
          }}
          aria-invalid={error ? true : undefined}
          placeholder={t('# 標題\n\n- 項目一\n- 項目二\n\n**粗體** 與 _斜體_ 也可以。', '# Heading\n\n- Item one\n- Item two\n\n**Bold** and _italic_ work too.')}
          rows={10}
          className={cn(
            'w-full px-2 py-1.5 text-xs font-mono border rounded bg-background outline-none focus:ring-1 resize-y',
            error
              ? 'border-destructive/60 focus:ring-destructive'
              : 'focus:ring-primary',
          )}
        />
        {error && (
          // R341 — display the caller-provided / locally-set message
          // verbatim. Was a hardcoded literal「請先填入 Markdown 內容
          // 再插入」 that didn't fit the fresh-empty fall-through.
          <div className="mt-1 text-[10px] text-destructive">{error}</div>
        )}
      </label>
      <div className="flex items-center justify-between pt-1">
        <span className="text-[10px] text-muted-foreground">
          {tImp('Esc 取消 · Ctrl+Enter 插入', 'Esc to cancel · Ctrl+Enter to insert')}
        </span>
        <div className="flex items-center gap-1">
          {/* R106 — sibling-shortcut-in-tooltip parity for the secondary
              cancel button. Same-file precedent at line 1584 already
              attaches `title="取消選取 (Esc)"` to a 取消 button bound to
              the exact same gesture (the comment there at line 1557
              spells out "The 取消 button mirrors what Esc does — present
              both because the keyboard shortcut isn't discoverable from
              a fresh marquee", which applies word-for-word to a fresh
              insert dialog too). Within this same dialog the primary
              插入 button at line 3469 already advertises its keystroke
              「(Ctrl+Enter)」 per R58 — leaving the 取消 sibling beside
              it as the lone within-dialog asymmetry: footer hint at
              line 3442 names "Esc 取消" but the button itself was
              silent, so a hover-to-discover user got the shortcut on
              one button and not the other. Wording mirrors the L1584
              sibling's verb-first shape ("取消選取 (Esc)" → "取消插入
              (Esc)"), naming what's being cancelled rather than echoing
              the bare label — same pattern FindReplaceDialog.tsx:580
              now uses with「關閉尋找與取代 (Esc)」 (R128 caught that
              this footnote's claim was originally false: the line it
              referenced read bare 「關閉 (Esc)」 until R128 brought it
              into the same verb-then-scope shape this comment claims). */}
          <button
            type="button"
            onClick={onClose}
            title={tImp('取消插入 (Esc)', 'Cancel insertion (Esc)')}
            className="px-2 py-1 text-xs rounded border hover:bg-secondary"
          >
            {tImp('取消', 'Cancel')}
          </button>
          {/* Ctrl+Enter (not bare Enter) — the keymap at line 3320 binds it
              that way because the textarea above is multi-line and bare Enter
              must reach the field as a literal newline (the docstring at line
              3244-3247 spells this out). The tooltip therefore advertises
              「(Ctrl+Enter)」, mirroring DiffPreview.tsx:89 which is the only
              other primary action in the app that needs the Ctrl modifier
              for the same reason (multi-line context).
              R58 cleanup completion: I had claimed in that round that the
              two link dialogs were the lone primary-action buttons without
              shortcut tooltips — this MarkdownInsertDialog button was a
              third instance that I missed, surfaced this round by grepping
              every `bg-primary text-primary-foreground` button in the
              components folder. With this fix the claim from R58 holds
              across the whole codebase. */}
          <button
            type="button"
            onClick={submit}
            title={tImp('插入 Markdown 內容 (Ctrl+Enter)', 'Insert Markdown content (Ctrl+Enter)')}
            className="px-2 py-1 text-xs rounded bg-primary text-primary-foreground hover:bg-primary/90"
          >
            {tImp('插入', 'Insert')}
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * HtmlInsertDialog — paste an HTML snippet, hand back the source to the
 * parent for conversion (via html-to-docx.ts) + splice. Sibling of
 * MarkdownInsertDialog with the same Esc/Ctrl+Enter semantics, focus
 * restore, and an inline parser-error display path (returned by onCommit
 * so we keep the dialog open and tell the user what's wrong with their
 * HTML rather than swallowing the failure).
 */
function HtmlInsertDialog({
  onClose,
  onCommit,
}: {
  onClose: () => void;
  /** Returns undefined on success, or a string parseError to display
   *  inline (in which case the dialog stays open for the user to fix). */
  onCommit: (source: string) => string | undefined;
}) {
  const t = useT(); // R409 — i18n for label + sample placeholder
  const [source, setSource] = useState('');
  const [error, setError] = useState<string | null>(null);
  const taRef = useRef<HTMLTextAreaElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
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
    queueMicrotask(() => taRef.current?.focus());
  }, []);
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
    const trimmed = source.trim();
    if (!trimmed) {
      setError(tImp('請先填入 HTML 內容再插入', 'Enter HTML content before inserting'));
      taRef.current?.focus();
      return;
    }
    // R335 — the returned string is now whatever the caller wants displayed
    // verbatim, no dialog-side prefixing. R334 introduced a second non-OK
    // return shape (`fresh.length === 0`, parse succeeded but produced
    // nothing to insert) and the previous hardcoded「HTML 解析失敗：」
    // prefix was a misnomer for that path — parsing genuinely succeeded;
    // it's the post-walk output that's empty. Letting the caller decide
    // the wording (parseError vs. no-content) keeps the two error
    // categories accurately labelled while still routing through one
    // dialog-level display channel. Variable renamed `parseError → err`
    // for the same reason — it's now a general onCommit error, not
    // specifically a parse error.
    const err = onCommit(source);
    if (err) {
      setError(err);
      taRef.current?.focus();
    }
  };
  return (
    <div
      ref={dialogRef}
      role="dialog"
      aria-label={tImp('插入 HTML', 'Insert HTML')}
      className="absolute top-3 right-3 z-30 w-[460px] rounded-md border bg-background shadow-lg p-3 text-sm space-y-2"
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
          e.stopPropagation();
          e.preventDefault();
          submit();
        }
      }}
    >
      <div className="text-xs font-medium">{tImp('插入 HTML 內容', 'Insert HTML content')}</div>
      <label className="block">
        <span className="block text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
          {t('HTML 原始碼', 'HTML source')}
        </span>
        <textarea
          ref={taRef}
          value={source}
          onChange={(e) => {
            setSource(e.target.value);
            if (error) setError(null);
          }}
          aria-invalid={error ? true : undefined}
          placeholder={t('<h2>標題</h2>\n<p><strong>粗體</strong>與<em>斜體</em>。</p>\n<ul>\n  <li>項目一</li>\n  <li>項目二</li>\n</ul>', '<h2>Heading</h2>\n<p><strong>Bold</strong> and <em>italic</em>.</p>\n<ul>\n  <li>Item one</li>\n  <li>Item two</li>\n</ul>')}
          rows={10}
          className={cn(
            'w-full px-2 py-1.5 text-xs font-mono border rounded bg-background outline-none focus:ring-1 resize-y',
            error
              ? 'border-destructive/60 focus:ring-destructive'
              : 'focus:ring-primary',
          )}
        />
        {error && (
          <div className="mt-1 text-[10px] text-destructive">{error}</div>
        )}
      </label>
      <div className="flex items-center justify-between pt-1">
        <span className="text-[10px] text-muted-foreground">
          {tImp('Esc 取消 · Ctrl+Enter 插入', 'Esc to cancel · Ctrl+Enter to insert')}
        </span>
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={onClose}
            title={tImp('取消插入 (Esc)', 'Cancel insertion (Esc)')}
            className="px-2 py-1 text-xs rounded border hover:bg-secondary"
          >
            {tImp('取消', 'Cancel')}
          </button>
          <button
            type="button"
            onClick={submit}
            title={tImp('插入 HTML 內容 (Ctrl+Enter)', 'Insert HTML content (Ctrl+Enter)')}
            className="px-2 py-1 text-xs rounded bg-primary text-primary-foreground hover:bg-primary/90"
          >
            {tImp('插入', 'Insert')}
          </button>
        </div>
      </div>
    </div>
  );
}
