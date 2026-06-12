/**
 * IDE-style file explorer pane (left side of the app shell).
 *
 * UX (mirrors VS Code's Explorer):
 *   - Header: folder name (truncated) + "Open Folder" button + collapse toggle.
 *   - Tree: folders before files, alpha-sorted; click ▸ / ▾ to expand/collapse,
 *     click file to open into a new tab. Lazy load children on first expand.
 *   - Loading spinner per directory while its listing is in-flight.
 *   - Refresh button rescans the active subtree.
 *
 * Out of scope (intentional, MVP): rename/delete/new-file inline, drag-drop
 * reorder, multi-root workspaces, watch-for-changes (we re-list manually on
 * refresh). Adding any of these means committing to a write surface in `fs:*`.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  FileText,
  FolderClosed,
  FolderOpen,
  Loader2,
  RefreshCw,
  FolderInput,
  PanelLeftClose,
  FileSpreadsheet,
  Presentation,
  FileCode2,
} from 'lucide-react';
import type { FsEntry } from '../types/ipc';
import { useWorkspace } from '../store/workspace';
import { notify } from '../store/toast';
import { cn } from '../lib/utils';
import { bumpLoadGen, currentLoadGen } from '../lib/workspace-load-gen';
import { tryEnterOpen, exitOpen } from '../lib/workspace-open-busy';
import { useT, tImp } from '../lib/i18n';

interface Props {
  /** Currently opened folder (absolute path). Null = no folder picked yet. */
  rootPath: string | null;
  /** Called when user picks a new folder via the in-pane button. */
  onPickFolder: () => void;
  /** Called when user clicks the collapse arrow. */
  onCollapse: () => void;
}

/**
 * Map common file extensions to a small icon + colour. Anything not in the
 * map falls back to a generic file icon. We deliberately don't try to be
 * exhaustive — VS Code icons would be a 200-rule lookup; the goal here is
 * just "the user can tell at a glance which file is which type".
 */
function iconForFile(name: string): JSX.Element {
  const dot = name.lastIndexOf('.');
  const ext = dot > 0 ? name.slice(dot + 1).toLowerCase() : '';
  switch (ext) {
    case 'md':
    case 'markdown':
    case 'txt':
      return <FileText className="h-3.5 w-3.5 text-blue-500 shrink-0" />;
    case 'docx':
      return <FileText className="h-3.5 w-3.5 text-sky-600 shrink-0" />;
    case 'xlsx':
      return <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600 shrink-0" />;
    case 'pptx':
      return <Presentation className="h-3.5 w-3.5 text-orange-500 shrink-0" />;
    case 'gd':
      return <FileCode2 className="h-3.5 w-3.5 text-violet-600 shrink-0" />;
    default:
      return <FileText className="h-3.5 w-3.5 text-muted-foreground shrink-0" />;
  }
}

/**
 * Extensions we know how to open as a tab. Used to dim unsupported files.
 *
 * R337 — `html` / `htm` added so the file explorer is consistent with the
 * other two entry points that load external HTML files:
 *   • `workspace.openExternalFile` (workspace.ts:447-456) maps `html` / `htm`
 *     → tab.type `'html'` and constructs an HtmlTab.
 *   • App.tsx OS drag-drop handler (line ~1726) calls openExternalFile for
 *     every dropped file, so dragging an .html into the window opens it
 *     fine.
 * Without the explorer entry, the user sees .html files listed but greyed
 * out / unclickable, while the SAME files dropped onto the window open
 * cleanly — an inconsistent affordance that suggests .html is unsupported
 * when it actually isn't. Missed when HTML was added as a first-class tab
 * type; this set is the lone gate-keeper for the explorer's click path.
 */
const OPENABLE_EXTS = new Set([
  'md', 'markdown', 'txt',
  'html', 'htm',
  'docx', 'xlsx', 'pptx', 'gd',
]);

function isOpenable(name: string): boolean {
  const dot = name.lastIndexOf('.');
  if (dot <= 0) return false;
  return OPENABLE_EXTS.has(name.slice(dot + 1).toLowerCase());
}

/** Last path segment of a folder path, OS-agnostic. */
function basename(p: string): string {
  const i = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
  return i >= 0 ? p.slice(i + 1) : p;
}

/**
 * Module-level memory of which folders the user had expanded, keyed by
 * `rootPath`. The FileExplorer itself unmounts whenever the explorer panel
 * collapses (Ctrl+B / focus mode / drag the splitter to zero), and React
 * loses every `useState` along with it — so a user who drilled down into
 * `src/components/icons/...`, hid the panel to focus on editing, and
 * brought it back lands at the bare root with every subfolder collapsed.
 * Stash the set on every change, seed the next mount from it.
 *
 * Keying by rootPath means switching folders doesn't pollute the original
 * folder's tree state: each folder remembers its own expansion shape, and
 * coming back to a folder restores it. Map persists for the renderer-
 * process lifetime; cleared on app reload — same volatility band as the
 * other "session-only" memories (slideMemory, sheetMemory, viewMemory).
 *
 * Stale entries (folders deleted on disk while the panel was hidden) are
 * self-healing: `loadDirectory` catches the listing error, drops the
 * directory from `entries`, and collapses it back to ▸. The user just
 * sees the dead branches not expand on remount, no banner for ones we
 * proactively probed.
 */
const expandedMemory = new Map<string, string[]>();
/**
 * Tree-pane scroll position keyed by rootPath, complementing
 * `expandedMemory`. Without it, Ctrl+B-collapse-then-reopen of a deep tree
 * dumps the user back at the top of the listing — they could see they
 * were "200px down somewhere" but not where exactly. Restored on a rAF
 * loop because `loadDirectory` is async: the scrollHeight only reaches
 * the remembered offset once the relevant subtrees have re-listed.
 */
const treeScrollMemory = new Map<string, number>();

export function FileExplorer({ rootPath, onPickFolder, onCollapse }: Props): JSX.Element {
  // Per-directory cached listing. Keyed by absolute path; null means "loading".
  const [entries, setEntries] = useState<Map<string, FsEntry[] | null>>(new Map());
  // Seed expanded set from `expandedMemory` so a re-mount of the explorer
  // panel (after Ctrl+B collapse / focus mode / pane resize to zero) keeps
  // the user where they were. Falls back to the root-only set on first
  // ever visit to a folder. The init function only runs once per mount —
  // subsequent rootPath changes are handled by the effect below.
  const [expanded, setExpanded] = useState<Set<string>>(() => {
    if (!rootPath) return new Set();
    const remembered = expandedMemory.get(rootPath);
    return new Set(remembered ?? [rootPath]);
  });
  const [error, setError] = useState<string | null>(null);
  // R255 — generation counter that invalidates in-flight loadDirectory
  // completions when the user picks a different folder. Without this,
  // a stale `fs.listDirectory` IPC for OLD rootPath that resolves AFTER
  // the rootPath useEffect cleared `entries` lands its result via
  // `setEntries(prev => prev.set(OLD_dirPath, list))` — adding a path
  // from the abandoned folder back into NEW's `entries` Map. Concrete
  // symptoms:
  //   • Memory leak: every cross-folder swap that races a slow listing
  //     (large folder, slow disk, network share) bloats `entries` with
  //     stale paths that are never displayed (Tree only renders entries
  //     keyed by current rootPath subtree) but are also never freed.
  //   • Refresh loops over ghosts: refreshAll iterates `entries.keys()`
  //     and fires loadDirectory for each — including the stale OLD
  //     paths. Wasted IPC, and any of them failing pops an error
  //     banner naming a path the user no longer sees in the tree
  //     (the doc-block at line ~138-141 already calls this exact
  //     "previously-mounted network drive that's since gone offline"
  //     symptom — but that fix only addressed the entries-clear-on-
  //     swap, not the in-flight-completion-after-swap follow-up race).
  //   • Stale errors: if OLD's listDirectory rejects AFTER swap, the
  //     catch at line ~250 setError's a 「無法讀取 OLD_path」 banner
  //     for a folder that's no longer the user's context.
  // Bumped at the top of the rootPath useEffect; captured at
  // loadDirectory entry; checked before every setState write inside
  // loadDirectory's async body. Same shape as workspace-load-gen.ts
  // (R168) for workspace-replacement IPCs and pingEpochRef
  // (SettingsDialog.tsx:74) for the AI ping.
  const rootGenRef = useRef(0);
  // Auto-load the root listing whenever the user picks a folder. Reset the
  // entries cache + last error too — without that, every previously-visited
  // folder's listing leaks into the new session: 重新整理 then iterates over
  // those dead paths and calls loadDirectory on each, which both wastes IPC
  // and pops error banners about paths the user can no longer see in the
  // tree (most painfully, a previously-mounted network drive that's since
  // gone offline). Memory also grows unbounded across folder swaps.
  //
  // For `expanded`: seed from `expandedMemory[rootPath]` if present (so
  // returning to a folder restores its tree shape), otherwise just the
  // root. Trigger loadDirectory for every initially-expanded path so the
  // children render immediately rather than waiting for the user to
  // re-toggle each level.
  useEffect(() => {
    if (!rootPath) return;
    rootGenRef.current++;
    setEntries(new Map());
    setError(null);
    const remembered = expandedMemory.get(rootPath);
    const initial = remembered && remembered.length > 0 ? remembered : [rootPath];
    setExpanded(new Set(initial));
    for (const p of initial) {
      void loadDirectory(p);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rootPath]);

  // Persist the expanded set as it changes. Writing on every toggle (rather
  // than only on unmount) means the memory is also durable across an
  // unrelated component error / boundary catch — we don't drop the user's
  // navigation just because some other surface threw.
  useEffect(() => {
    if (!rootPath) return;
    expandedMemory.set(rootPath, Array.from(expanded));
  }, [rootPath, expanded]);

  // Scroll-position memory for the tree's overflow-auto wrapper. Capture
  // on cleanup (unmount or rootPath change), restore on the next mount.
  // Restoration runs on a rAF loop bounded at 500ms because remembered
  // offsets often exceed the initial scrollHeight — we have to wait for
  // the async loadDirectory calls above to populate the relevant subtrees
  // before the scroller can accept the offset. Once scrollHeight has
  // grown past the target we apply it and stop polling; if it never
  // grows that far (folders pruned on disk) we just bail at 500ms.
  const treeScrollRef = useRef<HTMLDivElement | null>(null);
  // R266 — restoration-completed flag. Cleanup only persists scrollTop if
  // we either had no target to restore OR the rAF tick actually applied
  // it. Without this gate, a rapid folder swap during the loadDirectory
  // window writes back the partial-load scrollTop (typically 0) and
  // permanently clobbers the previously remembered offset.
  const treeScrollRestoredRef = useRef(false);
  useEffect(() => {
    if (!rootPath) return;
    const target = treeScrollMemory.get(rootPath) ?? 0;
    let cancelled = false;
    treeScrollRestoredRef.current = target <= 0;
    if (target > 0) {
      const start = performance.now();
      const tick = () => {
        if (cancelled) return;
        const el = treeScrollRef.current;
        if (!el) return;
        const max = el.scrollHeight - el.clientHeight;
        if (max >= target) {
          el.scrollTop = target;
          treeScrollRestoredRef.current = true;
          return;
        }
        if (performance.now() - start > 500) return;
        requestAnimationFrame(tick);
      };
      requestAnimationFrame(tick);
    }
    return () => {
      cancelled = true;
      const el = treeScrollRef.current;
      if (!el) return;
      if (!treeScrollRestoredRef.current) return;
      treeScrollMemory.set(rootPath, el.scrollTop);
    };
  }, [rootPath]);

  // Bring the active tab's row into view when the active tab changes via
  // anything OTHER than a click on this panel — Ctrl+Tab / Ctrl+1..9 /
  // TabBar click on a tab whose file lives further down the tree. Without
  // this the file got highlighted in primary tint but the highlight was
  // off-screen, exactly the symptom the TabBar's own auto-scroll fixes
  // (TabBar.tsx:75-81). We mirror that pattern but add a transition guard:
  // skipping the initial mount run is essential because `treeScrollMemory`
  // restores `scrollTop` via the rAF loop above, and scrolling to the
  // active row on mount would override the user's last-known browse
  // position. Clicking a row (the common path) is a no-op too, since the
  // clicked row is already in view and `block: 'nearest'` doesn't move it.
  // CSS.escape protects against paths containing characters that break
  // the attribute selector (parentheses, quotes — Windows lets them in
  // filenames).
  const activeTabPath = useWorkspace((s) => {
    const active = s.tabs.find((t) => t.id === s.activeTabId);
    return active?.sourcePath ?? null;
  });
  const prevActiveTabPathRef = useRef<string | null>(activeTabPath);
  useEffect(() => {
    if (activeTabPath && activeTabPath !== prevActiveTabPathRef.current) {
      const el = treeScrollRef.current?.querySelector<HTMLElement>(
        `[data-file-path="${CSS.escape(activeTabPath)}"]`,
      );
      el?.scrollIntoView({ block: 'nearest' });
    }
    prevActiveTabPathRef.current = activeTabPath;
  }, [activeTabPath]);

  const loadDirectory = useCallback(async (dirPath: string) => {
    // R255 — capture the current rootPath generation. Every setState
    // below is guarded against staleness: if rootPath changed during
    // the IPC, this load belongs to the old folder and its result must
    // not pollute the new folder's state.
    const myGen = rootGenRef.current;
    setEntries((prev) => {
      if (myGen !== rootGenRef.current) return prev;
      const next = new Map(prev);
      next.set(dirPath, null);
      return next;
    });
    try {
      const list = await window.gendoc.fs.listDirectory(dirPath);
      if (myGen !== rootGenRef.current) return;
      setEntries((prev) => {
        const next = new Map(prev);
        next.set(dirPath, list);
        return next;
      });
      setError(null);
    } catch (e) {
      if (myGen !== rootGenRef.current) return;
      // R406 — bilingual error banner. Frozen-at-throw-time copy is fine
      // here (the banner is dismissable and re-firing the IPC produces a
      // fresh string in the user's current locale anyway), so tImp at the
      // setError site is enough — wiring a render-time t() through
      // setError would require restructuring the error state into
      // {path, message} and is overkill for a transient banner.
      setError(tImp(
        `無法讀取 ${dirPath}：${(e as Error).message}`,
        `Failed to read ${dirPath}: ${(e as Error).message}`,
      ));
      setEntries((prev) => {
        const next = new Map(prev);
        next.delete(dirPath);
        return next;
      });
      // Also collapse the failed directory. Without this, the tree leaves
      // the chevron in the ▾ (expanded) position with zero rows beneath —
      // an "expanded but empty" state with no obvious connection to the
      // banner above, and no in-place retry affordance. Collapsing flips
      // the caret back to ▸ and turns re-expansion into the natural retry
      // (toggleDirectory's `entries.has` check is now false, so a fresh
      // load fires).
      setExpanded((prev) => {
        if (!prev.has(dirPath)) return prev;
        const next = new Set(prev);
        next.delete(dirPath);
        return next;
      });
    }
  }, []);

  const toggleDirectory = useCallback(
    (dirPath: string) => {
      setExpanded((prev) => {
        const next = new Set(prev);
        if (next.has(dirPath)) {
          next.delete(dirPath);
        } else {
          next.add(dirPath);
          // Lazy-load on first expand.
          if (!entries.has(dirPath)) void loadDirectory(dirPath);
        }
        return next;
      });
    },
    [entries, loadDirectory],
  );

  const refreshAll = useCallback(() => {
    if (!rootPath) return;
    // Re-load every currently-cached path so subtrees the user already
    // expanded come back fresh, not just the root.
    for (const p of entries.keys()) void loadDirectory(p);
  }, [rootPath, entries, loadDirectory]);

  return (
    <div className="h-full flex flex-col bg-secondary/30 text-xs">
      <Header
        rootPath={rootPath}
        onPickFolder={onPickFolder}
        onRefresh={refreshAll}
        onCollapse={onCollapse}
      />
      <div ref={treeScrollRef} className="flex-1 min-h-0 overflow-auto">
        {!rootPath ? (
          <EmptyState onPickFolder={onPickFolder} />
        ) : (
          <>
            {/* A single sub-directory failing to list (permissions, network
                drive disconnected, deleted mid-session) shouldn't hide the
                rest of the tree the user was navigating. Render the error as
                a dismissible banner above the tree instead — and keep the
                successfully-loaded folders interactable. */}
            {error ? (
              <div className="flex items-start gap-2 p-2 mx-1 mt-1 text-[11px] text-destructive break-words border border-destructive/30 bg-destructive/[0.05] rounded">
                <span className="flex-1">{error}</span>
                {/* R121 — see Toaster.tsx:130-136 for the full rationale.
                    This banner is variant-fixed error (only rendered when
                    `error` is truthy, and styled with destructive colours
                    at the parent <div>), so it can reuse AIPanel.tsx:588's
                    canonical 「關閉錯誤訊息」 verbatim — the same string
                    for the same operation, instead of the bare 關閉 which
                    R119 + R121 both flagged as the outlier vocabulary. */}
                <button
                  type="button"
                  onClick={() => setError(null)}
                  title={tImp('關閉錯誤訊息', 'Dismiss error message')}
                  // R152 — icon-only-button accessible-name parity. Mirrors
                  // AIPanel.tsx:609-610 verbatim (the canonical close-X
                  // doublet) and TabBar.tsx:423-424 / line ~547 (this same
                  // round). The visible character is `×` — assistive tech
                  // either reads it as「乘號」or skips it; an explicit
                  // aria-label closes the gap.
                  aria-label={tImp('關閉錯誤訊息', 'Dismiss error message')}
                  className="opacity-60 hover:opacity-100 shrink-0 leading-none"
                >
                  ×
                </button>
              </div>
            ) : null}
            <Tree
              rootPath={rootPath}
              entries={entries}
              expanded={expanded}
              onToggle={toggleDirectory}
            />
          </>
        )}
      </div>
    </div>
  );
}

function Header({
  rootPath,
  onPickFolder,
  onRefresh,
  onCollapse,
}: {
  rootPath: string | null;
  onPickFolder: () => void;
  onRefresh: () => void;
  onCollapse: () => void;
}): JSX.Element {
  // R405 — bilingual.
  const t = useT();
  const folderName = rootPath ? basename(rootPath) : t('檔案總管', 'File Explorer');
  return (
    <div className="flex items-center h-9 px-2 border-b border-border gap-0.5 shrink-0">
      <span
        className="font-semibold tracking-wide uppercase text-[10px] text-muted-foreground truncate flex-1"
        title={rootPath ?? t('尚未開啟資料夾', 'No folder opened yet')}
      >
        {folderName}
      </span>
      {/* Both action buttons in this header bind to keyboard accelerators
          that are already surfaced on their App.tsx toolbar twins (App.tsx:
          1154 / 1165) but were silent here — same action, different tooltip.
          Users discovering shortcuts via hover would see "(Ctrl+B)" on the
          global toolbar's explorer toggle, then collapse the panel via this
          local button and lose the hint. Open Folder is wired in menu.ts:85
          (Ctrl+K Ctrl+O); Ctrl+B is handled directly in App.tsx:710-718.
          重新整理 has no global accelerator (F5/Ctrl+R are intentionally
          blocked — see menu.ts:151) so its tooltip stays unchanged. */}
      <IconBtn title={t('開啟資料夾… (Ctrl+K Ctrl+O)', 'Open folder… (Ctrl+K Ctrl+O)')} onClick={onPickFolder}>
        <FolderInput className="h-3.5 w-3.5" />
      </IconBtn>
      {rootPath ? (
        <IconBtn title={t('重新整理', 'Refresh')} onClick={onRefresh}>
          <RefreshCw className="h-3.5 w-3.5" />
        </IconBtn>
      ) : null}
      <IconBtn title={t('收合檔案總管 (Ctrl+B)', 'Collapse file explorer (Ctrl+B)')} onClick={onCollapse}>
        <PanelLeftClose className="h-3.5 w-3.5" />
      </IconBtn>
    </div>
  );
}

function IconBtn({
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
      title={title}
      onClick={onClick}
      className="h-6 w-6 inline-flex items-center justify-center rounded text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
    >
      {children}
    </button>
  );
}

function EmptyState({ onPickFolder }: { onPickFolder: () => void }): JSX.Element {
  // R405 — bilingual.
  const t = useT();
  return (
    <div className="p-3 flex flex-col gap-2 text-muted-foreground">
      <p>{t('尚未開啟資料夾。', 'No folder opened yet.')}</p>
      {/* Three surfaces wire to the same handleOpenFolder (App.tsx:578-590,
          bound to menu.ts:86 → 'menu:openFolder', accelerator Ctrl+K Ctrl+O):
          the top-bar Button at App.tsx:1182 and the header IconBtn above at
          line 367 both carry `title="開啟資料夾… (Ctrl+K Ctrl+O)"`, but this
          empty-state button — paradoxically the *first* surface a fresh user
          sees on startup, before any tabs or folder exist — was the only one
          silent on the shortcut. The discoverability cost is highest here:
          users who pick their first folder via this big button never get the
          hover hint and miss the keystroke entirely. Reuse the exact tooltip
          string for cross-surface parity (sibling pattern: every editor's
          Find/Goto buttons share their tooltip across toolbar + dialog). */}
      <button
        type="button"
        onClick={onPickFolder}
        title={t('開啟資料夾… (Ctrl+K Ctrl+O)', 'Open folder… (Ctrl+K Ctrl+O)')}
        className="px-2 py-1 rounded border border-border bg-background hover:bg-secondary text-foreground text-left"
      >
        <FolderInput className="h-3.5 w-3.5 inline mr-1.5 -mt-0.5" />
        {t('開啟資料夾…', 'Open folder…')}
      </button>
      <p className="text-[10px] leading-relaxed pt-1">
        {t(
          '選擇後可以瀏覽 .md / .docx / .xlsx / .pptx / .gd 檔案，點擊即會開啟為新頁籤。',
          'After picking a folder, browse .md / .docx / .xlsx / .pptx / .gd files; click one to open it as a new tab.',
        )}
      </p>
    </div>
  );
}

interface TreeProps {
  rootPath: string;
  entries: Map<string, FsEntry[] | null>;
  expanded: Set<string>;
  onToggle: (path: string) => void;
}

/**
 * Renders the root directory's children, recursing through expanded subdirs.
 * The root itself is implicit — header already shows its name. Indent is
 * proportional to depth (0 = root children → flush left).
 */
function Tree({ rootPath, entries, expanded, onToggle }: TreeProps): JSX.Element {
  const list = entries.get(rootPath);
  if (list === null) return <LoadingRow depth={0} />;
  if (list === undefined) return <></>;
  if (list.length === 0) {
    return <div className="px-3 py-2 text-muted-foreground italic">{tImp('（空資料夾）', '(Empty folder)')}</div>;
  }
  return (
    <div className="py-1">
      {list.map((entry) => (
        <Node
          key={entry.path}
          entry={entry}
          depth={0}
          entries={entries}
          expanded={expanded}
          onToggle={onToggle}
        />
      ))}
    </div>
  );
}

function Node({
  entry,
  depth,
  entries,
  expanded,
  onToggle,
}: {
  entry: FsEntry;
  depth: number;
  entries: Map<string, FsEntry[] | null>;
  expanded: Set<string>;
  onToggle: (path: string) => void;
}): JSX.Element {
  if (entry.isDirectory) {
    const isOpen = expanded.has(entry.path);
    const children = entries.get(entry.path);
    return (
      <>
        <DirRow
          entry={entry}
          depth={depth}
          isOpen={isOpen}
          onToggle={() => onToggle(entry.path)}
        />
        {isOpen ? (
          children === null ? (
            <LoadingRow depth={depth + 1} />
          ) : children ? (
            children.map((child) => (
              <Node
                key={child.path}
                entry={child}
                depth={depth + 1}
                entries={entries}
                expanded={expanded}
                onToggle={onToggle}
              />
            ))
          ) : null
        ) : null}
      </>
    );
  }
  return <FileRow entry={entry} depth={depth} />;
}

function DirRow({
  entry,
  depth,
  isOpen,
  onToggle,
}: {
  entry: FsEntry;
  depth: number;
  isOpen: boolean;
  onToggle: () => void;
}): JSX.Element {
  return (
    <button
      type="button"
      onClick={onToggle}
      className="w-full flex items-center gap-1 px-1 py-0.5 hover:bg-secondary text-foreground"
      style={{ paddingLeft: 4 + depth * 12 }}
      title={entry.path}
    >
      {isOpen ? (
        <ChevronDown className="h-3 w-3 text-muted-foreground shrink-0" />
      ) : (
        <ChevronRight className="h-3 w-3 text-muted-foreground shrink-0" />
      )}
      {isOpen ? (
        <FolderOpen className="h-3.5 w-3.5 text-amber-500 shrink-0" />
      ) : (
        <FolderClosed className="h-3.5 w-3.5 text-amber-500 shrink-0" />
      )}
      <span className="truncate">{entry.name}</span>
    </button>
  );
}

function FileRow({ entry, depth }: { entry: FsEntry; depth: number }): JSX.Element {
  const openable = isOpenable(entry.name);
  const [opening, setOpening] = useState(false);
  const openFile = useOpenFile();
  // R406 — FileRow's title was hardcoded Chinese; the parent FileExplorer
  // sub-components (Header, EmptyState) already subscribe via useT().
  // FileRow is a separately-mounted React component so it needs its own
  // subscription for the row tooltip to flip on locale toggle.
  const t = useT();
  // Surface tab state next to the file so the explorer feels connected to the
  // tab bar instead of just being a "list of paths". Clicking an already-open
  // file already short-circuits to the existing tab (see useOpenFile), but
  // without any visual marker users had no way to tell "this file is open" /
  // "this file is the active tab" from "this file isn't loaded yet". Three
  // states in priority order: active (bg tint + accent bar) > open-elsewhere
  // (subtle dot) > unopened.
  //
  // Path comparison is strict-equal — the same convention as `useOpenFile`'s
  // existing-tab short-circuit and the workspace store. If those ever switch
  // to a normalised compare (case-insensitive on Windows), this should follow
  // suit so the highlight and the click-behaviour stay in sync.
  const isActiveTab = useWorkspace((s) => {
    const active = s.tabs.find((t) => t.id === s.activeTabId);
    return active?.sourcePath === entry.path;
  });
  // R411 — keep the selector pure. The previous form closed over the
  // component-local `isActiveTab` inside the selector, so a store update
  // that flipped both values in one commit could evaluate against the
  // stale closure for one render. Select only store state; derive the
  // combined flag outside.
  const isOpenAnywhere = useWorkspace((s) =>
    s.tabs.some((t) => t.sourcePath === entry.path),
  );
  const isOpenInTab = !isActiveTab && isOpenAnywhere;
  const handleClick = useCallback(async () => {
    if (!openable || opening) return;
    setOpening(true);
    try {
      await openFile(entry.path);
    } finally {
      setOpening(false);
    }
  }, [openable, opening, openFile, entry.path]);

  return (
    <button
      type="button"
      onClick={handleClick}
      // data-attribute lets FileExplorer's auto-scroll-to-active effect
      // locate this row by absolute path. Plain `data-file-path` (not a
      // namespaced one) because the explorer's effect uses a single
      // attribute selector and there is no naming clash inside this panel.
      data-file-path={entry.path}
      // Disable while a previous open is in-flight too — the handler already
      // guards against re-entry, but `disabled` removes hover affordance and
      // makes the suppressed click visually obvious.
      disabled={!openable || opening}
      className={cn(
        'w-full flex items-center gap-1 px-1 py-0.5 text-left',
        openable ? 'text-foreground' : 'text-muted-foreground cursor-not-allowed',
        // Hover affordance kept on every openable row, including the active
        // one — VS Code does the same: the active tint stays, hover just
        // deepens it slightly. No hover at all on active feels broken.
        openable && (isActiveTab ? 'bg-primary/15 hover:bg-primary/20' : 'hover:bg-secondary'),
        isActiveTab && 'text-primary font-medium',
        opening && 'opacity-70 cursor-wait',
      )}
      // Indent extra to line up with the dir's caret + folder icon (caret = 12px,
      // icon = 14px, gap = 4px = ~30px). Files have no caret so we pad equivalently.
      style={{ paddingLeft: 4 + depth * 12 + 16 }}
      title={
        openable
          ? isActiveTab
            ? `${entry.path}${t('（目前頁籤）', ' (current tab)')}`
            : isOpenInTab
              ? `${entry.path}${t('（已在其他頁籤開啟）', ' (open in another tab)')}`
              : entry.path
          : `${entry.path}${t('（不支援的檔案類型）', ' (unsupported file type)')}`
      }
    >
      {opening ? (
        <Loader2 className="h-3.5 w-3.5 animate-spin text-muted-foreground shrink-0" />
      ) : (
        iconForFile(entry.name)
      )}
      <span className="truncate flex-1">{entry.name}</span>
      {/* Open-in-another-tab marker. A small dot is enough — the user just
          needs to know "clicking this jumps to a tab" vs "clicking this
          opens fresh". Skipped on the active row to keep the visual quiet
          (the bg tint + bold already say "this is the active one"). */}
      {isOpenInTab && (
        <span
          aria-hidden="true"
          className="h-1.5 w-1.5 rounded-full bg-primary/60 shrink-0 mr-0.5"
        />
      )}
    </button>
  );
}

function LoadingRow({ depth }: { depth: number }): JSX.Element {
  return (
    <div
      className="flex items-center gap-1 px-1 py-0.5 text-muted-foreground"
      style={{ paddingLeft: 4 + depth * 12 + 16 }}
    >
      <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
      <span>{tImp('載入中…', 'Loading…')}</span>
    </div>
  );
}

/**
 * Hook that returns a function for opening any external file path into a new
 * tab. Routes .gd archives through `workspace.openPath`, everything else
 * through `fs.readFile` + `openExternalFile`. Surfaces failures via alert
 * since the explorer pane has no inline error slot per row.
 */
function useOpenFile(): (filePath: string) => Promise<void> {
  // R177 — drop the dirtyRef + useEffect mirror; read live store at click
  // time instead. The ref pattern was stable-callback friendly but lagged
  // by one render cycle (useEffect runs after commit), so a click landing
  // between an autosave's setState and the next commit would see the
  // ref's stale `dirty=true` and pop the「尚未儲存」 confirm on a workspace
  // that's already saved. Same fix as R175 (App.tsx) + R176 (TabBar).
  // Reading via getState() at call time is the canonical convention this
  // codebase has settled on for confirm-gate dirty checks.
  return useCallback(async (filePath: string) => {
    const lower = filePath.toLowerCase();
    if (lower.endsWith('.gd')) {
      // R199 — sync gate before any IPC. Same shape as handleOpen /
      // handleOpenRecent (App.tsx). Without this, rapid double-click on
      // a .gd file in the explorer pane queues two `app.confirm` IPCs in
      // main; main's `showMessageBoxSync` for the first one blocks main
      // (renderer keeps queuing), and the second confirm dialog stacks
      // on top of the first's open-file dialog after dismissal. Module-
      // level gate cross-cuts every workspace-open entry point including
      // App.tsx's, so a FileExplorer click while a Ctrl+O is in flight
      // also bounces.
      if (!tryEnterOpen()) return;
      // R289 — flatten outer/inner try; see App.handleOpen sibling
      // doc-block. confirm-reject path was unhandled in original nested
      // layout. Same final shape as EditorSurface.onOpenExisting and
      // App.handleOpen after R289.
      try {
        // .gd swap: confirm if there are unsaved edits, then load via the
        // existing workspace pipeline so chat history / undo scope reset cleanly.
        if (
          useWorkspace.getState().dirty &&
          !(await window.gendoc.app.confirm(tImp(
            '目前的變更尚未儲存，確定開啟其他檔案？',
            'There are unsaved changes. Open a different file anyway?',
          )))
        )
          return;
        // R187 — join the workspace-load-gen exclusion class. Without this,
        // a click on a .gd in FileExplorer concurrent with App.tsx's auto-open
        // / handleOpen / handleOpenRecent / drag-drop / handleNew would race
        // — IPC completion order isn't guaranteed, so the older load could
        // overwrite the newer. R168 / R169 covered App.tsx's 5 entries via
        // a useRef counter; R187 promotes that counter to a shared module
        // (`lib/workspace-load-gen.ts`) so this 6th entry — outside App.tsx's
        // component scope — joins the same "latest wins" invariant.
        const myGen = bumpLoadGen();
        const opened = await window.gendoc.workspace.openPath(filePath);
        if (myGen !== currentLoadGen()) return;
        if (opened) useWorkspace.getState().loadFromOpened(opened);
      } catch (e) {
        notify(tImp(`開啟失敗：${(e as Error).message}`, `Failed to open: ${(e as Error).message}`), 'error');
      } finally {
        exitOpen();
      }
      return;
    }
    // Short-circuit before reading bytes when the file is already open —
    // re-clicking a file in the explorer should activate the existing tab
    // instead of opening a duplicate. This is the common case after the
    // user opens something, edits it, and clicks back to the explorer.
    {
      const existing = useWorkspace
        .getState()
        .tabs.find((t) => t.sourcePath === filePath);
      if (existing) {
        useWorkspace.getState().setActiveTab(existing.id);
        return;
      }
    }
    try {
      const content = await window.gendoc.fs.readFile(filePath);
      const id = useWorkspace.getState().openExternalFile(content, filePath);
      if (!id) {
        notify(tImp(`不支援的檔案類型：${content.name}`, `Unsupported file type: ${content.name}`), 'warning');
      }
    } catch (e) {
      // R406 — sibling-leak with the .gd branch above (line ~796) which
      // already routes through tImp. This loose-file path was the lone
      // outlier — same operation (open file → failure → toast), the only
      // difference being which kind of file.
      notify(tImp(
        `開啟失敗：${(e as Error).message}`,
        `Failed to open: ${(e as Error).message}`,
      ), 'error');
    }
  }, []);
}

