/**
 * Workspace store — owns the open `.gd`, the tab list, dirty state, and the
 * currently active tab. Editor components read from here; AI ChangeSet apply
 * mutates here.
 */

import { create } from 'zustand';
import { v4 as uuid } from 'uuid';
import type { Manifest, TabDescriptor, TabType } from '../types/manifest';
import { tabFromDescriptor, type Tab } from '../types/tab';
import type { FsFileContent, OpenedWorkspace } from '../types/ipc';
import type { ChangeSet } from '../types/changeset';
import { useAI } from './ai';
import { cancelInflight, pushSyntheticToolResult } from '../ai/orchestrator';

interface SelectionInfo {
  tabId: string;
  /** Free-form serialized snippet shown in the AI Context badge. Truncated /
   *  decorated for display. NOT what gets sent to the model. */
  preview: string;
  /** Full raw selection text — this is what AIPanel forwards as
   *  `selectionContext` so the model sees the complete user selection,
   *  not the 60-char preview. Empty string when the host editor has no
   *  meaningful flat-text representation (e.g. a spreadsheet range —
   *  those should encode their own structured payload). */
  text: string;
  /** Machine-readable payload sent to AI as context. */
  payload: unknown;
}

/**
 * Async-save lifecycle state, surfaced in the toolbar and status bar so
 * users can see whether a save is in flight, finished, or errored.
 *
 *   idle    — no save has happened in this session (or the indicator was
 *             cleared after a successful one timed out).
 *   saving  — handleSave/handleSaveAs is awaiting the IPC round-trip.
 *   success — finished cleanly; we keep this state for ~2s so the user
 *             gets a brief check-icon confirmation, then drift to idle.
 *   error   — the last save threw. UI shows the error message tooltip.
 */
export type SaveState = 'idle' | 'saving' | 'success' | 'error';

interface WorkspaceState {
  filePath: string;
  manifest: Manifest;
  tabs: Tab[];
  activeTabId: string | null;
  selection: SelectionInfo | null;
  /** True if any tab is dirty or manifest itself changed. */
  dirty: boolean;
  /** Stable id used for SQLite history scoping. */
  workspaceId: string | null;
  /**
   * R388 — monotonic counter of "true workspace swaps" (open / new), distinct
   * from `workspaceId`. `workspaceId` is the SQLite history scoping key and
   * intentionally changes on Save-As (R385) so undo / conversation rows
   * follow the new filePath identity; but from the USER's perspective Save-As
   * isn't a workspace swap — they're still in the same buffer, same
   * conversation, same composing flow. Subscribers that wipe session-only UI
   * state (AIPanel's draftMemory / chatScrollMemory) must distinguish the
   * two: wipe only on a real swap (sessionEpoch bumped) and preserve state
   * on a Save-As (sessionEpoch unchanged). Bumped by loadFromOpened and
   * newWorkspace; left alone everywhere else including performSave's
   * R385 setState. Null sentinel mirrors workspaceId's null-initial
   * convention so subscribers can use the same "skip on first open from
   * null" idiom without inventing a parallel epoch-0 sentinel.
   */
  sessionEpoch: number | null;
  /** Current state of the async save pipeline. */
  saveState: SaveState;
  /** Epoch-ms timestamp of the last successful save, null if never. */
  lastSavedAt: number | null;
  /** Last save error message, populated when saveState === 'error'. */
  saveError: string | null;
  /**
   * Per-keystroke edit pulse. Bumped (`Date.now()`) by the high-frequency
   * mutations that flip dirty=true mid-typing — `patchTab`, `markTabDirty`.
   * The auto-save effect in App.tsx subscribes to this so its setTimeout
   * debounce *actually* re-arms on every keystroke; without it, `dirty`
   * stays referentially true between save cycles and the timer fires
   * `autoSaveMs` after the FIRST keystroke instead of the last (popping a
   * mid-typing save flash that the comment claimed wouldn't happen). Reset
   * to 0 on workspace swap so a stale pulse doesn't trigger an immediate
   * save against the freshly-loaded `.gd`.
   */
  lastEditAt: number;
  /**
   * Transient "已匯出 foo.md" flash for the StatusBar. Export goes through
   * the OS save-as dialog, so the user knows *they* triggered it, but they
   * don't see where the file landed — previously the success path only hit
   * console.log. We surface the basename for ~5s; an `at` timestamp lets
   * the auto-clear timeout no-op when a newer export has already replaced
   * it (so back-to-back exports don't clobber each other early).
   */
  exportFlash: { fileName: string; filePath?: string; at: number } | null;
  /**
   * LIFO stack of recently-closed tabs. Bounded so an aggressive close-all
   * doesn't grow forever. Reset whenever the workspace itself swaps so a
   * stale tab from a different .gd can't re-appear after Open.
   */
  recentlyClosedTabs: Tab[];
  /**
   * Session-only redo stack for AI changesets. Persistent undo lives in
   * SQLite (cross-session); redo only makes sense within a single editing
   * session because applying a *new* AI change branches the timeline and
   * invalidates the redo. Bounded indirectly by the same 50-step undo cap.
   */
  aiRedoStack: ChangeSet[];
  /**
   * Mirror of "is there a row in the persistent undo SQLite stack for this
   * workspaceId?". Used to disable the toolbar's Undo button so its
   * affordance matches Redo (which already greys out via `aiRedoStack.length`).
   * The actual stack lives in the main process — App.tsx owns the IPC
   * round-trip that refreshes this value on workspace swap / after each
   * undo/redo, and AIPanel optimistically flips it true on apply.
   */
  canUndo: boolean;

  pushAiRedo(cs: ChangeSet): void;
  popAiRedo(): ChangeSet | null;
  clearAiRedo(): void;
  setCanUndo(b: boolean): void;
  loadFromOpened(o: OpenedWorkspace): void;
  newWorkspace(): void;
  setActiveTab(id: string): void;
  addTab(type: TabType, name?: string): Tab;
  /**
   * Open an external file (read by the file-explorer pane via `fs:readFile`)
   * as a new tab. Returns the new tab id, or null if the extension isn't a
   * type Gen Doc can edit. Activates the new tab on success. When
   * `sourcePath` is provided, it is recorded on the tab so a subsequent
   * click on the same file activates the existing tab instead of opening
   * a duplicate.
   */
  openExternalFile(content: FsFileContent, sourcePath?: string): string | null;
  removeTab(id: string): void;
  /** Pop the last-closed tab and re-insert it. Returns the restored tab id. */
  reopenClosedTab(): string | null;
  renameTab(id: string, name: string): void;
  reorderTabs(orderedIds: string[]): void;
  patchTab(id: string, patch: Partial<Tab>): void;
  /**
   * Flip a tab's `dirty` flag (and the workspace-level flag) to true. No-op
   * when both are already dirty so calling this on every keystroke doesn't
   * burn re-renders. Used by editors that debounce their heavy serialize
   * work — they need the dirty bit set immediately so a Ctrl+W within the
   * debounce window still triggers the unsaved-changes prompt.
   */
  markTabDirty(id: string): void;
  markDirty(dirty: boolean): void;
  setSelection(info: SelectionInfo | null): void;
  serializeForSave(): { manifest: Manifest; payloads: { descriptor: TabDescriptor; bytes: Uint8Array }[] };
  setFilePath(path: string): void;
  setSaveState(state: SaveState, error?: string | null): void;
  /** Flash a transient "已匯出" indicator in the StatusBar for ~5s. The
   *  optional `filePath` is the full absolute destination — surfaced as the
   *  tooltip on hover so users can confirm *where* the file landed without
   *  reopening the OS save dialog. The visible text shows just the basename
   *  to keep the StatusBar slot compact. */
  flashExport(fileName: string, filePath?: string): void;
}

const NEW_TAB_DEFAULTS: Record<TabType, { name: string; ext: string; bytes: () => Uint8Array }> = {
  markdown: { name: 'Untitled.md', ext: 'md', bytes: () => new TextEncoder().encode('# 新文件\n\n') },
  html: {
    name: 'Untitled.html',
    ext: 'html',
    bytes: () =>
      new TextEncoder().encode(
        '<!DOCTYPE html>\n<html lang="zh-Hant">\n<head>\n  <meta charset="UTF-8">\n  <title>新文件</title>\n</head>\n<body>\n  <h1>新文件</h1>\n  <p></p>\n</body>\n</html>\n',
      ),
  },
  docx: { name: 'Untitled.docx', ext: 'docx', bytes: () => new Uint8Array(0) },
  xlsx: { name: 'Untitled.xlsx', ext: 'xlsx', bytes: () => new Uint8Array(0) },
  pptx: { name: 'Untitled.pptx', ext: 'pptx', bytes: () => new Uint8Array(0) },
};

/**
 * Return a name that doesn't collide with any existing tab's name. If
 * `baseName` is free, returns it as-is. Otherwise inserts " (N)" before the
 * extension — "Untitled.md" → "Untitled (2).md" — the same convention
 * Windows Explorer / VS Code use for duplicate filenames. Without this,
 * clicking the TabBar's `+` dropdown three times in a row produced three
 * tabs all literally labelled "Untitled.md" with identical tooltips, since
 * the TabBar's same-basename dir-suffix disambiguator only kicks in for
 * tabs with a `sourcePath` (fresh blanks have none).
 */
function uniqueDefaultName(baseName: string, tabs: Tab[]): string {
  const taken = new Set(tabs.map((t) => t.name));
  if (!taken.has(baseName)) return baseName;
  // Split at the LAST dot so multi-dot stems ("notes.draft.md") still get
  // the suffix in the right place. Hidden-file names with no real extension
  // (".gitignore", lastIndexOf returns 0) fall through to append-at-end so
  // we don't produce nonsense like " (2).gitignore".
  const dot = baseName.lastIndexOf('.');
  const stem = dot > 0 ? baseName.slice(0, dot) : baseName;
  const ext = dot > 0 ? baseName.slice(dot) : '';
  for (let n = 2; n < 1000; n++) {
    const candidate = `${stem} (${n})${ext}`;
    if (!taken.has(candidate)) return candidate;
  }
  // Pathological — 1000 tabs with the same default name. Fall back to the
  // raw base; the duplicate label is far less harmful than refusing to add.
  return baseName;
}

// R258 — workspaceId must be unique per workspace within a session.
// Original `workspaceIdFromTitle` hashed only manifest.title, which gave the
// SAME id to:
//   • Multiple unsaved workspaces (all titled「未命名」 by emptyManifest).
//   • Saved workspaces with coincidentally identical titles (Report1.gd /
//     Report2.gd both with manifest.title = "Report" — title isn't tied to
//     filename, and the title field is rarely user-edited so this collision
//     happens by default for files saved from the same template).
//
// Three cascading failures:
//   1. Cross-workspace undo bleed. SQLite undo_entries are scoped by
//      workspace_id; two workspaces sharing an id share an undo stack.
//      popUndo(wsId) inside handleUndo returns the most recent row across
//      BOTH workspaces' edits. R173's `useWorkspace.getState().workspaceId
//      !== workspaceId` check passes (same id) so the foreign cs is
//      applied; undoChangeset's `t.id === op.tabId` filter doesn't match
//      (foreign cs has the OTHER workspace's tab uuids), it visually
//      no-ops, but `setState({ tabs, dirty: true })` still flips the
//      false-positive dirty flag.
//   2. Workspace-swap detection collapse. R178 / R179 / R180 / R236-R243
//      / R248 / R250 all gate on `workspaceId` changing across the swap;
//      when two workspaces share an id, swap is invisible to these
//      guards — OLD turn's tool_results land in NEW workspace's chat,
//      `ai.error` setError leaks across, etc. The entire family of
//      workspace-pinning fixes assumed the id was unique.
//   3. Cross-workspace ensureConversation dedupe collision. R160's
//      `pendingCreate.wsId === wsId` check would treat two unsaved
//      workspaces as the same — the second's first message would
//      reuse the first's createConversation in-flight promise.
//
// Fix: a saved workspace's id derives from its absolute filePath
// (genuinely unique on disk per session); an unsaved workspace's id
// uses a monotonic session-counter so each newWorkspace call yields
// a fresh id. Title is no longer the basis for identity — it's
// purely a display-layer concern. The `f-` / `u-` prefix keeps the
// two id classes visually distinguishable in DevTools / SQLite logs.
let unsavedWorkspaceCounter = 0;
// R385 — exported so App.tsx's performSave can recompute workspaceId after a
// Save-As that changed the filePath. See callsite for the orphan-undo
// rationale.
export function workspaceIdFor(filePath: string): string {
  if (filePath) {
    let h = 0;
    for (let i = 0; i < filePath.length; i++) {
      h = (h * 31 + filePath.charCodeAt(i)) | 0;
    }
    return `f-${(h >>> 0).toString(36)}`;
  }
  return `u-${++unsavedWorkspaceCounter}`;
}

function emptyManifest(title = '未命名'): Manifest {
  const now = new Date().toISOString();
  return {
    version: '1.0',
    title,
    createdAt: now,
    modifiedAt: now,
    tabs: [],
    settings: { embedChatHistory: false, defaultModel: 'claude-sonnet-4-6' },
    metadata: { appVersion: '1.0.0' },
  };
}

export const useWorkspace = create<WorkspaceState>((set, get) => ({
  filePath: '',
  manifest: emptyManifest(),
  tabs: [],
  activeTabId: null,
  selection: null,
  dirty: false,
  workspaceId: null,
  // R388 — see WorkspaceState.sessionEpoch doc-block.
  sessionEpoch: null,
  saveState: 'idle',
  lastSavedAt: null,
  saveError: null,
  lastEditAt: 0,
  exportFlash: null,
  recentlyClosedTabs: [],
  aiRedoStack: [],
  canUndo: false,

  pushAiRedo(cs) {
    set((s) => ({ aiRedoStack: [...s.aiRedoStack, cs].slice(-50) }));
  },
  popAiRedo() {
    const stack = get().aiRedoStack;
    if (stack.length === 0) return null;
    const cs = stack[stack.length - 1];
    set({ aiRedoStack: stack.slice(0, -1) });
    return cs;
  },
  clearAiRedo() {
    set({ aiRedoStack: [] });
  },
  setCanUndo(b) {
    set({ canUndo: b });
  },

  loadFromOpened(o) {
    // R396 — delegate to tabFromDescriptor. Original hand-rolled construction
    // duplicated types/tab.ts's tabFromDescriptor and used `as Tab` cast for
    // the docx/xlsx/pptx fallback — silently accepting unknown types without
    // the exhaustiveness check tabFromDescriptor's switch + `_exhaustive:
    // never` already enforces. R376's validateManifest in main catches bad
    // types BEFORE loadFromOpened ever runs in practice today, so the two
    // paths are behaviorally equivalent on the current code surface — but:
    //   • The R347 doc-block at changeset-apply.ts:60-63 already presumed
    //     loadFromOpened uses tabFromDescriptor ("keeps the construction
    //     shape in lockstep with `loadFromOpened`'s sibling path") — the
    //     comment was inaccurate, the code lagged behind.
    //   • Adding a new tab type (a hypothetical 'csv' / 'json' / 'pdf'
    //     v2 future) requires updating BOTH places; missing one silently
    //     emits structurally invalid Tabs through the cast. The R347 bug
    //     report (HtmlTab construction with phantom `data` field) was
    //     exactly this shape — fixed in changeset-apply.ts but left
    //     dormant here.
    //   • Any future code path that constructs a Tab without going
    //     through R376's validator (programmatic addTab variant,
    //     clipboard-paste of a serialized payload, test fixture)
    //     bypasses the safety net; tabFromDescriptor's throw is the
    //     load-bearing exhaustiveness guarantee.
    // Single source of truth, same end state (dirty: false → workspace-
    // level dirty: false on next setState), same behavior on the
    // currently-supported five types.
    const tabs: Tab[] = o.tabs.map((p) => tabFromDescriptor(p.descriptor, p.bytes));
    tabs.sort((a, b) => a.order - b.order);
    set({
      filePath: o.filePath,
      manifest: o.manifest,
      tabs,
      activeTabId: tabs[0]?.id ?? null,
      dirty: false,
      // R258 — derive id from filePath (unique on disk) instead of title
      // (collides when two .gd files share manifest.title). See
      // workspaceIdFor doc-block for the full collision-failure trace.
      workspaceId: workspaceIdFor(o.filePath),
      // R388 — bump sessionEpoch so AIPanel's draftMemory / chatScrollMemory
      // wipe-on-swap subscribers fire for THIS open (real swap) but NOT for
      // any subsequent Save-As within the same session (which only updates
      // workspaceId, not sessionEpoch). See WorkspaceState.sessionEpoch
      // doc-block.
      sessionEpoch: (get().sessionEpoch ?? 0) + 1,
      selection: null,
      recentlyClosedTabs: [],
      aiRedoStack: [],
      lastEditAt: 0,
      // Provisionally clear; App.tsx watches `workspaceId` and re-queries
      // the persistent stack — if this .gd has prior undo rows it'll flip
      // back to true within a tick.
      canUndo: false,
      // R185 — reset workspace-scoped UI status so OLD workspace's last
      // save / export feedback doesn't leak into NEW workspace's StatusBar.
      // Without this, opening a different .gd shows「X 秒前儲存」(referring
      // to OLD's save time) and「已匯出 foo.md」(OLD's export) even though
      // NEW was never saved or exported by the user. saveState/lastSavedAt/
      // saveError carry over because Zustand's `set` is a partial merge —
      // any field not specified retains its previous value. The four
      // workspace-scoped status fields are wiped here so NEW's StatusBar
      // accurately reflects "freshly opened, never modified by user".
      saveState: 'idle',
      lastSavedAt: null,
      saveError: null,
      exportFlash: null,
    });
    // Wipe the AI session — chat history and pending changes from the
    // previous workspace reference tab IDs that no longer exist (Apply
    // would crash). Each .gd is its own conversation context.
    //
    // Critically, cancel any in-flight chat turn first. `clear()` resets
    // the AI store's streaming/messages, but the orchestrator's `inflight`
    // handle and the IPC chunk listener live outside the store — left
    // alone they keep dispatching chunks into the freshly-cleared store,
    // which then `finalizeStreaming` into the *new* workspace as a brand
    // new conversation row. Cancelling here unsubscribes the listener and
    // null-clears the handle before clear() runs.
    cancelInflight({ persistPartial: false });
    useAI.getState().clear();
  },

  newWorkspace() {
    const m = emptyManifest();
    set({
      filePath: '',
      manifest: m,
      tabs: [],
      activeTabId: null,
      dirty: false,
      // R258 — counter-based id for unsaved workspaces; previously every
      // newWorkspace produced the same id (hash of "未命名"), so two
      // sequential `Ctrl+N`s collided. See workspaceIdFor doc-block.
      workspaceId: workspaceIdFor(''),
      // R388 — same sessionEpoch bump as loadFromOpened above; see
      // WorkspaceState.sessionEpoch doc-block for the Save-As distinction.
      sessionEpoch: (get().sessionEpoch ?? 0) + 1,
      selection: null,
      recentlyClosedTabs: [],
      aiRedoStack: [],
      lastEditAt: 0,
      // Fresh workspace has no persistent undo entries by definition.
      canUndo: false,
      // R185 — same status reset as loadFromOpened above. A user creating
      // a new project shouldn't see OLD workspace's save / export status
      // bleed into the empty workspace.
      saveState: 'idle',
      lastSavedAt: null,
      saveError: null,
      exportFlash: null,
    });
    // Same rationale as in loadFromOpened: kill the inflight stream before
    // clearing the AI store, otherwise a turn started in the previous
    // workspace lands in this fresh one as an orphaned message.
    cancelInflight({ persistPartial: false });
    useAI.getState().clear();
  },

  setActiveTab(id) {
    // Clear the AI Context badge on tab switch. Editors are unmounted/
    // remounted on tab change (EditorBoundary `key={active.id}` in
    // EditorSurface) so the previous tab's selection no longer corresponds
    // to anything visible. Without this, switching A→B→A makes the badge
    // re-appear with stale text the editor doesn't actually have selected,
    // and the next AI send forwards that ghost selection as context.
    // No-op when clicking the already-active tab so re-clicks don't blow
    // away an in-progress selection.
    set((s) => (s.activeTabId === id ? s : { activeTabId: id, selection: null }));
  },

  addTab(type, name) {
    const def = NEW_TAB_DEFAULTS[type];
    const id = uuid();
    // Caller-supplied names are trusted as-is — only auto-number the default
    // path so a programmatic `addTab(type, 'foo.md')` doesn't get silently
    // mangled to 'foo (2).md'. The TabBar `+` dropdown / EditorSurface quick-
    // start buttons (the only callers today) both go through the default.
    const finalName =
      name !== undefined ? name : uniqueDefaultName(def.name, get().tabs);
    const fileName = `${id}.${def.ext}`;
    const order = get().tabs.length;
    const tab: Tab =
      type === 'markdown' || type === 'html'
        ? {
            id,
            name: finalName,
            type,
            file: `doc/${fileName}`,
            order,
            dirty: true,
            content: new TextDecoder().decode(def.bytes()),
          }
        : ({
            id,
            name: finalName,
            type,
            file: `doc/${fileName}`,
            order,
            dirty: true,
            data: def.bytes(),
          } as Tab);
    // R306 — clear selection on activeTabId change, mirroring setActiveTab's
    // invariant. Without this, a user with a live selection in tab A who
    // hits `+` to open a new tab leaves selection.tabId=A stranded while
    // activeTabId moves to the new tab. AIPanel's `activeSelection` filter
    // hides the badge initially, but the moment activeTabId comes BACK to
    // A (via removeTab's neighbor-fallback when the new tab is closed),
    // the filter passes and the badge re-appears — even though A's
    // editor was unmounted/remounted in between (EditorBoundary `key=
    // {active.id}` at EditorSurface.tsx:45) and the CodeMirror instance
    // has no actual selection. Sending then forwards stale text as
    // `selectionContext` to the model. Same fix shape at
    // openExternalFile / reopenClosedTab / removeTab fallback below.
    set((s) => ({ tabs: [...s.tabs, tab], activeTabId: id, selection: null, dirty: true }));
    return tab;
  },

  openExternalFile(content, sourcePath) {
    // Map common extensions to our four supported tab types. The .gd archive
    // case is handled by `loadFromOpened` (separate flow), not here.
    const extToType: Record<string, TabType> = {
      md: 'markdown',
      markdown: 'markdown',
      txt: 'markdown',
      html: 'html',
      htm: 'html',
      docx: 'docx',
      xlsx: 'xlsx',
      pptx: 'pptx',
    };
    const type = extToType[content.ext];
    if (!type) return null;
    // Activate the existing tab if this exact file is already open — avoids
    // building up duplicate copies as the user re-clicks files in the
    // explorer or re-drops them. Skip the dedupe when sourcePath is unknown
    // (e.g. drag-drop without a resolvable path) so we still open *something*.
    if (sourcePath) {
      const existing = get().tabs.find((t) => t.sourcePath === sourcePath);
      if (existing) {
        // R306 — match setActiveTab's "clear selection on cross-tab switch"
        // invariant. Skip the clear when the user re-opens the file they're
        // already on (existing.id === current activeTabId), so an in-progress
        // selection in the same tab isn't blown away by a redundant drop.
        set((s) =>
          s.activeTabId === existing.id ? s : { activeTabId: existing.id, selection: null },
        );
        return existing.id;
      }
    }
    const id = uuid();
    const order = get().tabs.length;
    const fileExt = type === 'markdown' ? 'md' : type === 'html' ? 'html' : type;
    // `file` here is the would-be archive path *if* this workspace were saved
    // as a .gd. Until then the actual on-disk source is tracked separately
    // (by the file explorer's own state); we don't try to round-trip back to
    // the source path on save — this is "import a copy" semantics.
    const file = `doc/${id}.${fileExt}`;
    const tab: Tab =
      type === 'markdown' || type === 'html'
        ? {
            id,
            name: content.name,
            type,
            file,
            order,
            dirty: true,
            sourcePath,
            content: new TextDecoder().decode(content.bytes),
          }
        : ({
            id,
            name: content.name,
            type,
            file,
            order,
            dirty: true,
            sourcePath,
            data: content.bytes,
          } as Tab);
    // R306 — same selection-clear as addTab; the new external file becomes
    // active and any prior selection from the previously-active tab is now
    // stale relative to the freshly-mounted editor.
    set((s) => ({ tabs: [...s.tabs, tab], activeTabId: id, selection: null, dirty: true }));
    return id;
  },

  removeTab(id) {
    // Evict any AI PendingChange that targets this tab — leaving it queued
    // would mean: (a) the DiffPreview keeps showing for a tab the user can
    // no longer see, (b) clicking Apply silently no-ops because
    // `applyChangeset` filters ops by `t.id === op.tabId`, and (c) the
    // assistant's `tool_use` block is now orphaned, so the *next* user
    // prompt round-trips to Anthropic with "tool_use ids were found
    // without tool_result blocks" → 400. Push a synthetic error
    // tool_result for each orphan so the API invariant holds and the model
    // sees why nothing happened. We exclude `tab_create` because its
    // tabId names a *new* tab to be created, not one that just closed —
    // a uuid collision with the freshly-closed id is effectively zero.
    const aiState = useAI.getState();
    const orphaned = aiState.pending.filter((p) =>
      p.changeset.ops.some((op) => op.type !== 'tab_create' && op.tabId === id),
    );
    if (orphaned.length > 0) {
      // R243 — capture workspaceId before the void pushSyntheticToolResult
      // calls so the .catch handler below doesn't leak OLD workspace's error
      // text into NEW's `ai.error` banner across a workspace swap. Same
      // shape as R236 (handleAssistantToolCalls outer catch), R237
      // (cancelInflight partial-persist), R238 (runTurn onDone catch),
      // R239 (AIPanel.send), R240 (AIPanel onApply / onReject) — this
      // closes the renderer-side catch sibling that R204 originally
      // added. Trigger: user × tab T in workspace A while a PendingChange
      // for T exists, the void IPC fires, user immediately Ctrl+O to
      // workspace B before persistMessage resolves; sqlite-wal contention
      // during the swap window rejects the append; the .catch's
      // setError reads LIVE useAI (now NEW after clear()) and writes
      // OLD's exception text into NEW's banner. The pushMessage
      // inside pushSyntheticToolResult is sync (already ran into OLD's
      // ai.messages, then clear() wiped it on swap — that's an existing
      // dropped-result concern, not addressed here); only the
      // user-visible banner-pinning invariant is what R243 closes.
      const removeWorkspaceId = useWorkspace.getState().workspaceId;
      for (const p of orphaned) {
        // R204 — attach .catch so a persistMessage rejection inside
        // pushSyntheticToolResult doesn't escape as unhandledrejection.
        // Same rationale as the AIPanel onModify callsite. The synthetic
        // tool_result is already in the live ai.messages array (pushMessage
        // is sync); only the IPC persist can fail. On reject, surface via
        // setError so the user sees a toast rather than crashing the
        // renderer mid-tab-close.
        void pushSyntheticToolResult(
          p.toolUseId,
          'Target tab was closed before the user could apply this change.',
          true,
        ).catch((err) => {
          // R243 — workspace guard, see capture above.
          if (useWorkspace.getState().workspaceId !== removeWorkspaceId) return;
          useAI.getState().setError(err instanceof Error ? err.message : String(err));
        });
      }
      useAI.setState({
        pending: aiState.pending.filter((p) => !orphaned.includes(p)),
      });
    }
    set((s) => {
      const removedIdx = s.tabs.findIndex((t) => t.id === id);
      // R208 — no-op when the tab doesn't exist in the live store. Without
      // this, a stale removeTab call (the canonical race: closeMany awaits
      // its dirty-confirm dialog while the user fires Ctrl+O; swap clears
      // the store's tabs; the for-loop runs `removeTab(id)` for OLD ids
      // against NEW workspace's tab list) still falls through to the set
      // below, which unconditionally writes `dirty: true`. NEW workspace
      // was clean before the stale call; the spurious dirty flip then
      // triggers an auto-save flash, pops 「尚未儲存」on the next close,
      // and confuses users about whether their just-loaded .gd is dirty.
      // Same defence applies to React StrictMode-driven double-invocation
      // of any setState chain that calls removeTab — the second invocation
      // sees an already-removed id and would otherwise re-write dirty=true.
      // Returning the input state is the standard zustand no-op pattern;
      // it skips the subscriber notification entirely.
      if (removedIdx < 0) return s;
      const removed = s.tabs[removedIdx];
      const next = s.tabs.filter((t) => t.id !== id).map((t, i) => ({ ...t, order: i }));
      // When the active tab is closed, fall onto the right neighbor (the
      // tab that slides into the closed tab's position) rather than jumping
      // back to tabs[0]. Same convention as VS Code, Chrome, Firefox — and
      // crucial when working with long tab lists, where landing on the first
      // tab feels like teleporting away from your context. Fall back left
      // if we closed the rightmost tab, then null if nothing's left.
      let activeTabId = s.activeTabId;
      if (s.activeTabId === id) {
        if (removedIdx < next.length) activeTabId = next[removedIdx].id;
        else if (next.length > 0) activeTabId = next[next.length - 1].id;
        else activeTabId = null;
      }
      // R202 — drop a stale selection that points at the just-removed tab.
      // The AIPanel `activeSelection` derive at line ~100 already filters
      // selections whose tabId doesn't match activeTabId, so the visible
      // badge stays clean — but the underlying `s.selection` record sits
      // in the store referencing a tab id no other code can lookup, and
      // any future code that consumes `s.selection` directly (without the
      // tabId filter) would silently see the stale text. Clearing here
      // keeps the store's invariant consistent: every non-null selection
      // points at a live tab.
      // R306 — also clear when the FALLBACK activeTabId path fires (active
      // tab was closed, neighbor takes over). Same hazard as the addTab /
      // openExternalFile / reopenClosedTab fixes in this round: a stale
      // selection from BEFORE this round-trip can match the new fallback
      // activeTabId (typical trigger — user selects in A, opens C via +,
      // closes C → fallback to A → selection.tabId still A from the
      // pre-addTab state). With R306 on addTab the stranded state can't
      // form in that exact path, but other code paths (binary editor's
      // future setSelection, programmatic openTab outside the store
      // helpers) could still leave selection unaligned with activeTabId.
      // Belt-and-suspenders: any activeTabId change inside removeTab
      // resets selection if it doesn't match the new active tab, leaving
      // setActiveTab's invariant intact across every store mutation.
      const selectionRaw =
        s.selection && s.selection.tabId === id ? null : s.selection;
      const selection =
        s.activeTabId !== activeTabId && selectionRaw && selectionRaw.tabId !== activeTabId
          ? null
          : selectionRaw;
      // Cap the closed-tab history at 10 to bound memory — closing a 20-tab
      // workspace shouldn't squat on every byte forever.
      const recentlyClosedTabs = [removed, ...s.recentlyClosedTabs].slice(0, 10);
      return { tabs: next, activeTabId, selection, dirty: true, recentlyClosedTabs };
    });
  },

  reopenClosedTab() {
    const s = get();
    if (s.recentlyClosedTabs.length === 0) return null;
    const [restored, ...rest] = s.recentlyClosedTabs;
    // Re-key the order onto the end of the current list. The original `id`
    // is preserved so chat-history references and undo entries pointing at
    // this tab keep working — closing didn't actually invalidate them.
    //
    // Preserve the original tab.dirty rather than forcing true: if the user
    // closed a clean tab (bytes matched disk) and immediately reopens it,
    // the bytes are still identical to disk, so the dirty dot would lie.
    // The workspace-level dirty below is still correct — the tab list
    // changed, so the .gd container needs re-saving — but per-tab dirty
    // should reflect actual content drift from disk.
    const order = s.tabs.length;
    const re = { ...restored, order } as Tab;
    // R306 — same selection-clear as addTab / openExternalFile. The
    // restored tab becomes active; any selection still in the store
    // belongs to the previously-active tab whose editor just unmounted
    // via the EditorBoundary key change at EditorSurface.tsx:45.
    set({
      tabs: [...s.tabs, re],
      activeTabId: re.id,
      selection: null,
      dirty: true,
      recentlyClosedTabs: rest,
    });
    return re.id;
  },

  renameTab(id, name) {
    set((s) => {
      // The TabBar rename input commits its current value on blur
      // (`renameTab(id, e.target.value || tab.name)`), so a no-op rename —
      // double-clicking a tab and then clicking elsewhere — would otherwise
      // silently mark the workspace dirty and kick off an auto-save / unsaved
      // changes prompt. Bail out when the name is unchanged.
      const cur = s.tabs.find((t) => t.id === id);
      if (!cur) return s;
      // Whitespace-only names render as a blank tab — visually indistinguishable
      // from a UI glitch, and `e.target.value || tab.name` in TabBar only
      // catches the empty-string case, not a string of spaces. Trim and reject
      // the empty result so the rename silently no-ops, mirroring how OS file
      // managers refuse to commit blank names.
      const trimmed = name.trim();
      if (trimmed === '' || trimmed === cur.name) return s;
      return {
        tabs: s.tabs.map((t) => (t.id === id ? { ...t, name: trimmed, dirty: true } : t)),
        dirty: true,
      };
    });
  },

  reorderTabs(orderedIds) {
    set((s) => {
      // R184 — preserve tabs that exist in the live store but weren't in
      // `orderedIds`. The caller (TabBar drag-drop handler) builds
      // `orderedIds` from a React-closure `tabs` snapshot at drag start;
      // if a new tab spawns during the drag (AI `tab_create` Apply, an
      // openExternalFile drop, or the future-proof programmatic addTab),
      // the snapshot doesn't include it and the previous filter-out-nulls
      // implementation would silently delete it from the workspace. Walk
      // orderedIds first to honour the user's drag intent for the tabs
      // they DID see, then append any leftovers from the live store at
      // the end in their pre-drag order — the alternative (snapping
      // them to the front, or inserting at the drag-drop position) would
      // surprise the user more than "new tab landed at the end".
      const map = new Map(s.tabs.map((t) => [t.id, t]));
      const next: Tab[] = [];
      for (const id of orderedIds) {
        const t = map.get(id);
        if (!t) continue;
        next.push({ ...t, order: next.length });
        map.delete(id);
      }
      for (const t of s.tabs) {
        if (map.has(t.id)) {
          next.push({ ...t, order: next.length });
        }
      }
      return { tabs: next, dirty: true };
    });
  },

  patchTab(id, patch) {
    set((s) => {
      // Guard against stale-id calls. XlsxEditor's unmount cleanup fires a
      // fire-and-forget writeBack whose promise can resolve *after* the user
      // has already swapped to a different workspace (loadFromOpened replaces
      // the tabs array wholesale). Without this check the late patchTab would
      // mark the brand-new workspace as dirty even though nothing in it
      // actually changed, leaving an unwarranted "未儲存" indicator on a fresh
      // open. Same risk applies to any other async writer that captures a tab
      // id and races with workspace lifecycle.
      let found = false;
      const tabs = s.tabs.map((t) => {
        if (t.id !== id) return t;
        found = true;
        return { ...t, ...patch, dirty: true } as Tab;
      });
      if (!found) return s;
      return { tabs, dirty: true, lastEditAt: Date.now() };
    });
  },

  markTabDirty(id) {
    set((s) => {
      const t = s.tabs.find((x) => x.id === id);
      // No-op when already dirty *and* we already pulsed within the same
      // millisecond. The auto-save debounce relies on `lastEditAt` ticking
      // forward on every keystroke, so we deliberately do NOT bail just
      // because `dirty` is already true — that was the original bug. We
      // still skip when nothing genuinely changed (no matching tab, or the
      // exact same Date.now() tick) to avoid burning re-renders on truly
      // duplicate calls within one frame.
      if (!t) return s;
      const now = Date.now();
      if (t.dirty && s.dirty && s.lastEditAt === now) return s;
      return {
        tabs: t.dirty ? s.tabs : s.tabs.map((x) => (x.id === id ? { ...x, dirty: true } : x)),
        dirty: true,
        lastEditAt: now,
      };
    });
  },

  markDirty(dirty) {
    set({ dirty });
  },

  setSelection(info) {
    set({ selection: info });
  },

  setFilePath(p) {
    set({ filePath: p });
  },

  setSaveState(state, error) {
    if (state === 'success') {
      set({ saveState: 'success', lastSavedAt: Date.now(), saveError: null });
    } else if (state === 'error') {
      set({ saveState: 'error', saveError: error ?? '未知錯誤' });
    } else {
      set({ saveState: state, saveError: null });
    }
  },

  flashExport(fileName, filePath) {
    const at = Date.now();
    set({ exportFlash: { fileName, filePath, at } });
    // Auto-clear after 5s. Guarded by the `at` timestamp so a fresher export
    // landing within the window doesn't get nuked by an earlier timer.
    setTimeout(() => {
      set((s) => (s.exportFlash?.at === at ? { exportFlash: null } : {}));
    }, 5000);
  },

  serializeForSave() {
    const s = get();
    const tabs = s.tabs;
    const descriptors: TabDescriptor[] = tabs.map((t, i) => ({
      id: t.id,
      name: t.name,
      type: t.type,
      file: t.file,
      order: i,
    }));
    const payloads = tabs.map((t, i) => {
      const desc = descriptors[i];
      const bytes =
        t.type === 'markdown' || t.type === 'html'
          ? new TextEncoder().encode(t.content)
          : t.data;
      return { descriptor: desc, bytes };
    });
    const manifest: Manifest = {
      ...s.manifest,
      tabs: descriptors,
      modifiedAt: new Date().toISOString(),
    };
    return { manifest, payloads };
  },
}));
