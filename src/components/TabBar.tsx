/**
 * Tab bar — top of the editor surface. Spec §5.1.
 *
 * v1.0 reorder is implemented with native HTML5 drag-and-drop instead of
 * pulling in dnd-kit, to keep the bundle lean. Behaviour matches the spec:
 * click to switch, double-click to rename, hover to reveal close, dot when
 * dirty, plus button at the end with a 4-format dropdown.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import { ChevronDown, Download, Plus, X, FileText, FileSpreadsheet, Presentation, FileCode2, Code2 } from 'lucide-react';
import { useWorkspace } from '../store/workspace';
import { notify } from '../store/toast';
import { flushEditors } from '../lib/editor-flush';
import { exitClose, tryEnterClose } from '../lib/tab-close-busy';
import { exitExportTab, tryEnterExportTab } from '../lib/export-tab-busy';
import type { Tab } from '../types/tab';
import type { TabType } from '../types/manifest';
import { cn, clampToViewport } from '../lib/utils';
import { useT, tImp } from '../lib/i18n';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from './ui/dropdown';
import { Button } from './ui/button';

const TYPE_ICON: Record<TabType, React.ComponentType<{ className?: string }>> = {
  markdown: FileCode2,
  html: Code2,
  docx: FileText,
  xlsx: FileSpreadsheet,
  pptx: Presentation,
};

/** Tailwind classes for the per-type accent color of the tab icon. */
const TYPE_COLOR: Record<TabType, string> = {
  markdown: 'text-sky-500',
  html: 'text-rose-500',
  docx: 'text-indigo-500',
  xlsx: 'text-emerald-500',
  pptx: 'text-orange-500',
};

interface ContextMenuState {
  /** Viewport-x in CSS pixels — anchors the floating menu's left edge. */
  x: number;
  /** Viewport-y in CSS pixels — anchors the floating menu's top edge. */
  y: number;
  /** Which tab the right-click landed on. Bookkeeping for the actions. */
  tabId: string;
}

export function TabBar(): JSX.Element {
  const { tabs, activeTabId, setActiveTab, addTab, removeTab, renameTab, reorderTabs } =
    useWorkspace();
  // Drive the "重新開啟剛才關閉的頁籤" menu item's enabled state. We subscribe
  // narrowly to length (not the array) so a no-op re-render only happens when
  // the count actually changes — every removeTab pushes onto the array
  // (workspace.ts:460-462), and reopenClosedTab pops from it (line 470-487).
  const recentlyClosedCount = useWorkspace((s) => s.recentlyClosedTabs.length);
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [dragId, setDragId] = useState<string | null>(null);
  // While a drag is in flight we light up the prospective drop slot. Without
  // this users were dragging a tab into the void — no insertion line, no
  // highlight, only a tab order that snapped after release. We also reset on
  // dragend so a cancelled drag (Esc, drop outside) doesn't leak state.
  const [dragOverId, setDragOverId] = useState<string | null>(null);
  const [ctxMenu, setCtxMenu] = useState<ContextMenuState | null>(null);
  // Escape during rename has to defeat the input's onBlur — when we set
  // `renamingId=null`, React unmounts the input, which fires native blur,
  // which fires our onBlur handler, which would commit whatever partial
  // text the user typed. A ref flag is the cleanest way to short-circuit
  // the commit; setting state on Escape would race the same render cycle.
  const cancelRenameRef = useRef(false);
  // Auto-scroll the active tab into view when activeTabId changes. Without
  // this, switching via Ctrl+Tab / Ctrl+1..9 / context menu in an overflowing
  // tab bar leaves the highlight off-screen — user sees no visible feedback
  // for the keystroke and has to scroll the bar by hand to find what's
  // selected. `inline: 'nearest'` is the key flag: a tab already in view
  // doesn't trigger any scroll, so users dragging or clicking visible tabs
  // see no surprise jump.
  const barRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!activeTabId) return;
    const el = barRef.current?.querySelector<HTMLElement>(
      `[data-tab-id="${activeTabId}"]`,
    );
    el?.scrollIntoView({ inline: 'nearest', block: 'nearest' });
  }, [activeTabId]);

  // Dismiss the context menu on any outside click or Esc — same UX as the
  // OS file manager. We listen on `window` so a click on the editor below
  // also closes us.
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

  /**
   * Close a tab after flushing pending debounced editor serializes. PPTX /
   * DOCX / XLSX editors batch re-serialize behind a 400ms debounce, so a
   * close issued mid-burst would otherwise snapshot *stale* `tab.data` into
   * `recentlyClosedTabs` — Ctrl+Shift+T then restores a pre-edit state and
   * the user's edits silently vanish, even though the unsaved-changes
   * prompt accurately fired (markTabDirty is synchronous). Flushing first
   * lands the latest bytes in the store before the tab is captured.
   */
  const closeTabWithFlush = async (id: string) => {
    await flushEditors();
    removeTab(id);
  };

  /**
   * R220 — single per-tab close-confirm helper, shared across the three
   * close-confirm sites (middle-click, ×, context-menu 關閉). All three
   * previously inlined live-dirty read → optional `app.confirm` →
   * `closeTabWithFlush` and none gated against rapid double-fire on the
   * SAME tab; double-clicking × (which appears on group-hover so two
   * clicks in one hover frame are easy) or mashing the wheel-button
   * mid-scroll queued multiple `app.confirm` IPCs and stacked OS
   * dialogs. R221 — promoted the in-flight Set to
   * `lib/tab-close-busy.ts` so App.tsx's Ctrl+W keymap can claim the
   * same gate; that entry was the one path R220's component-local
   * Set didn't reach, leaving auto-repeat Ctrl+W still able to stack
   * dialogs on a dirty tab.
   */
  const closeWithConfirm = useCallback(
    (tabId: string, tabName: string) => {
      if (!tryEnterClose(tabId)) return;
      void (async () => {
        // R290 — add catch around app.confirm. Original try had only
        // finally; confirm-reject (main IPC anomaly) escaped as unhandled
        // rejection through the void IIFE. Same R288 / R289 flat-try
        // idiom for sibling close paths.
        try {
          const cur = useWorkspace.getState().tabs.find((t) => t.id === tabId);
          if (
            cur?.dirty &&
            !(await window.gendoc.app.confirm(
              `「${tabName}」尚未儲存，確定關閉？`,
            ))
          )
            return;
          void closeTabWithFlush(tabId);
        } catch (err) {
          notify(tImp(`關閉頁籤失敗：${(err as Error).message}`, `Failed to close tab: ${(err as Error).message}`), 'error');
        } finally {
          exitClose(tabId);
        }
      })();
    },
    // closeTabWithFlush is closure-captured (uses removeTab from store,
    // which is stable, and flushEditors module-level). Stable across
    // renders.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [],
  );

  /**
   * Bulk-close helper for the "close others/right/all" menu items. If any of
   * the targets are dirty we ask once for the whole batch — pestering with
   * one confirm per tab would be much worse UX than the per-tab × button.
   * One flush covers every registered editor, so we do it once up front.
   *
   * R126 — confirm-dialog terminal-punctuation harmonization (full-width 「？」).
   * The codebase already had two camps for CJK dialog/prompt strings:
   *   • full-width 「？」 in DocxEditor.tsx:1512 (`跳到第幾段？`),
   *     GoToDialog.tsx:13/14/40 (the shared `跳到第幾項？` label),
   *     PptxEditor.tsx:416/494/510 (3 native-dialog confirms — same
   *     IPC path this file uses), and PptxEditor.tsx:1015.
   *   • half-width 「?」 in this file (4 sites: lines 131, 237, 469, 641),
   *     App.tsx (5 sites: 254, 263, 286, 787, 930), FileExplorer.tsx:679,
   *     XlsxEditor.tsx:679, and main.ts:102 (the OS-level "save before
   *     quit?" detail).
   * The internal-inconsistency smoking gun is sharpest *within* each
   * half-width string: every one of these mixes full-width 「，」 / 「『』」
   * with a trailing half-width 「?」 — i.e. the writer reached for the CJK
   * comma and quote pair without question, then defaulted back to ASCII
   * for the question mark. PptxEditor's three confirms — which call the
   * exact same `window.gendoc.app.confirm` IPC bridge as this file —
   * already finished the job (full-width throughout). R124 just
   * established the same direction for `（）` in the AIPanel placeholder
   * (AIPanel.tsx:1218); the question mark is the same call. Aligning
   * brings every native confirm dialog the user sees to one CJK
   * typography voice without touching any logic.
   */
  const closeMany = async (ids: string[], action: string) => {
    // R176 — live tab dirty read. The closure-bound `tabs` here came from
    // useWorkspace((s) => ({tabs, ...})) at render time; if autosave flipped
    // some `t.dirty` between render and the user picking the menu item,
    // we'd over-count dirty tabs and pop the「N 個頁籤尚未儲存」confirm
    // for a count that's higher than the actual unsaved count. Reading
    // live from the store at click time keeps the prompt honest. Also
    // covers the rarer race where the user creates / closes another tab
    // via shortcut between menu open and click.
    const liveTabs = useWorkspace.getState().tabs;
    const targets = liveTabs.filter((t) => ids.includes(t.id));
    if (targets.length === 0) return;
    const dirtyCount = targets.filter((t) => t.dirty).length;
    // R290 — wrap the IPC + flush + remove sequence in try/catch. Original
    // code had no protection: confirm-reject (rare main IPC anomaly) escaped
    // as unhandled rejection through `void closeMany(...)` at the three
    // context-menu callsites. Batch-close UX is especially bad without
    // notify — the user picks「關閉其他 / 關閉右側 / 全部關閉」, expects
    // a confirm, and gets silence instead. Same R288 / R289 flat-try idiom.
    try {
      if (
        dirtyCount > 0 &&
        !(await window.gendoc.app.confirm(`有 ${dirtyCount} 個頁籤尚未儲存，確定${action}？`))
      ) {
        return;
      }
      await flushEditors();
      for (const id of ids) removeTab(id);
    } catch (err) {
      notify(tImp(`${action}失敗：${(err as Error).message}`, `${action} failed: ${(err as Error).message}`), 'error');
    }
  };

  // Disambiguate same-basename tabs (VSCode-style). Common scenario: opening
  // a/report.md and b/report.md both render as "report.md" with identical
  // tooltips — users have to click each to tell them apart. When the same
  // name appears more than once we append the parent directory of each
  // colliding tab. Tabs without a sourcePath (unsaved blanks, restored from
  // a .gd archive) skip the suffix — there's nothing to disambiguate by.
  const nameCounts = new Map<string, number>();
  for (const t of tabs) nameCounts.set(t.name, (nameCounts.get(t.name) ?? 0) + 1);

  // R405 — bilingual; useT here so this component's strings flip on toggle.
  const t = useT();
  return (
    <div ref={barRef} className="flex items-center h-10 border-b bg-secondary/40 px-1 gap-0.5 overflow-x-auto">
      {tabs.map((tab, idx) => {
        const Icon = TYPE_ICON[tab.type];
        const isActive = tab.id === activeTabId;
        const colliding = (nameCounts.get(tab.name) ?? 0) > 1;
        let dirSuffix: string | null = null;
        if (colliding && tab.sourcePath) {
          const parts = tab.sourcePath.split(/[\\/]/).filter(Boolean);
          if (parts.length >= 2) dirSuffix = parts[parts.length - 2];
        }
        // Surface the Ctrl+1..9 tab-jump shortcut (App.tsx:828-845) on the
        // first nine tabs. Previously the binding was wired but never
        // advertised: only three internal comments mention it (App.tsx:829,
        // TabBar.tsx:73, FileExplorer.tsx:205) and no tooltip / menu / status
        // bar surface ever told the user about it. Sibling navigation
        // shortcuts (Ctrl+W on close button line 389, Ctrl+E on export
        // button line 357, Ctrl+Shift+T in the context menu after R39) all
        // have visible hints — Ctrl+1..9 was the lone orphan in this row.
        // Index-based gating (idx < 9) matches the handler's behaviour
        // exactly: it parses '1'..'9' and indexes tabs[parseInt-1], so the
        // 10th+ tab is not reachable by this binding and showing a hint
        // there would lie. Unlike the conditional Ctrl+W/Ctrl+E hints
        // (which advertise only on the active tab because the keystroke
        // targets *the* active tab), Ctrl+N specifically targets *this*
        // tab regardless of the current active state, so every tab in the
        // first nine carries the hint.
        const jumpHint = idx < 9 ? ` (Ctrl+${idx + 1})` : '';
        return (
          <div
            key={tab.id}
            data-tab-id={tab.id}
            draggable
            onDragStart={(e) => {
              setDragId(tab.id);
              // 'move' tells the OS to render a move cursor (vs. copy/no-drop)
              // and tells the browser this is an in-app rearrange, not a
              // file copy.
              e.dataTransfer.effectAllowed = 'move';
            }}
            onDragOver={(e) => {
              if (!dragId || dragId === tab.id) return;
              e.preventDefault();
              e.dataTransfer.dropEffect = 'move';
              if (dragOverId !== tab.id) setDragOverId(tab.id);
            }}
            onDragLeave={() => {
              // dragleave fires for every nested child the cursor crosses
              // (icon, label span, close button), so naïvely clearing flickers
              // the highlight. Only clear when this tab is actually the one
              // we last marked.
              if (dragOverId === tab.id) setDragOverId(null);
            }}
            onDragEnd={() => {
              // Fires whether the drop succeeded, was cancelled (Esc), or
              // landed outside the bar. Without this both ids leaked across
              // drags.
              setDragId(null);
              setDragOverId(null);
            }}
            onDrop={() => {
              if (!dragId || dragId === tab.id) {
                setDragOverId(null);
                return;
              }
              const ordered = tabs.map((t) => t.id);
              const from = ordered.indexOf(dragId);
              // R215 — bail when `dragId` is no longer in the live tabs
              // array. Most concrete trigger: workspace swap mid-drag.
              // The user starts a drag in workspace A (dragstart sets
              // setDragId(A's tab id)), Ctrl+O loads workspace B before
              // they drop, the TabBar re-renders against B's tabs, the
              // user releases the mouse over a B tab. dragId still
              // holds A's id (useState survives re-render); B's tabs
              // array has none of A's ids, so `from = -1`. The
              // following `ordered.splice(-1, 1)` is a sneaky footgun:
              // Array.splice with a negative index counts from the
              // *end*, so it removes B's LAST tab from `ordered`,
              // then inserts A's stale id at `to` — reorderTabs is
              // called with a corrupted array (missing one B tab,
              // gaining a phantom A id). Inside reorderTabs the R184
              // straggler-preservation loop re-attaches B's last tab
              // at the very end, so end state is "B's tabs reordered
              // weirdly + dirty=true for no reason." Same race shape
              // as R208's stale-id removeTab: the canonical defence
              // is a fail-fast at the entry point. Also covers any
              // race where the tab being dragged was removed by some
              // other code path (AI tool_use tab_delete + Apply
              // landing during the drag) — drop just no-ops the
              // gesture instead of corrupting the order.
              if (from < 0) {
                setDragId(null);
                setDragOverId(null);
                return;
              }
              const to = ordered.indexOf(tab.id);
              ordered.splice(from, 1);
              ordered.splice(to, 0, dragId);
              reorderTabs(ordered);
              setDragId(null);
              setDragOverId(null);
            }}
            onClick={() => setActiveTab(tab.id)}
            // Middle-click closes the tab — universal browser/IDE gesture.
            // `onAuxClick` fires for non-primary buttons; check for button=1
            // (wheel) so a right-click drag doesn't trip it.
            onAuxClick={(e) => {
              if (e.button !== 1) return;
              e.preventDefault();
              // R220 — single helper handles dirty-confirm + per-tab
              // double-fire gate. See `closeWithConfirm` doc-block.
              closeWithConfirm(tab.id, tab.name);
            }}
            onDoubleClick={() => setRenamingId(tab.id)}
            onContextMenu={(e) => {
              e.preventDefault();
              setActiveTab(tab.id);
              setCtxMenu({ x: e.clientX, y: e.clientY, tabId: tab.id });
            }}
            className={cn(
              'group flex items-center gap-1.5 px-3 h-8 rounded-md cursor-pointer text-sm select-none',
              isActive
                ? 'bg-background text-foreground border'
                : 'text-muted-foreground hover:bg-accent/40',
              // Drop target during reorder. Faded source tab + outlined target
              // gives users an unambiguous read of where the tab will land.
              dragId === tab.id && 'opacity-40',
              dragOverId === tab.id && dragId !== tab.id && 'ring-2 ring-primary',
            )}
          >
            <Icon className={cn('h-3.5 w-3.5', TYPE_COLOR[tab.type])} />
            {renamingId === tab.id ? (
              <input
                autoFocus
                defaultValue={tab.name}
                // performSave queries for this attribute and blurs the
                // element before snapshotting tab state — Ctrl+S issued
                // while the rename input still has focus then commits the
                // typed name (via this input's onBlur → renameTab) so the
                // *new* name reaches disk this round, not the old one.
                data-commit-on-save="true"
                // Select all on focus so users can overtype the suggested
                // name immediately. autoFocus + select via setTimeout because
                // autoFocus places the caret at the end by default.
                onFocus={(e) => e.currentTarget.select()}
                onBlur={(e) => {
                  if (cancelRenameRef.current) {
                    cancelRenameRef.current = false;
                    setRenamingId(null);
                    return;
                  }
                  // Match the XlsxEditor sheet-rename pattern
                  // (XlsxEditor.tsx:2413-2445): trim, then notify on empty
                  // submit instead of silently snapping back to the old name.
                  // The previous `e.target.value || tab.name` fallback was
                  // gentler in spirit but the same UX foot-gun the sheet
                  // rename comment calls out: "users see their tab snap back
                  // to the old name and assume the app ate their keystrokes."
                  // No-change submits (same name, including whitespace-trim
                  // matches) stay quiet — the user just clicked away without
                  // editing anything. Tabs are allowed to share names (each
                  // tab has its own id + sourcePath), so we don't dedupe.
                  const trimmed = e.target.value.trim();
                  setRenamingId(null);
                  if (trimmed === tab.name) return;
                  if (!trimmed) {
                    notify(t('頁籤名稱不能為空', 'Tab name cannot be empty'), 'warning');
                    return;
                  }
                  renameTab(tab.id, trimmed);
                }}
                onKeyDown={(e) => {
                  // Enter / Tab both commit (Tab is the natural "I'm done"
                  // gesture in form-y UIs). Escape cancels without writing.
                  // R231 — skip Enter (and Tab, defensively) during IME
                  // composition. CJK input on Mandarin / Japanese / Korean
                  // layouts uses Enter to confirm a candidate from the IME
                  // candidate window. Without this guard, our preventDefault
                  // + blur fires BEFORE the IME has committed the candidate
                  // — the input value is whatever raw bopomofo / pinyin /
                  // kana / hangul jamo the user typed pre-confirmation, NOT
                  // the chosen CJK glyph. onBlur then commits that broken
                  // text via renameTab. Mirrors AIPanel.tsx:1456-1471's
                  // send-Enter handler which already gates on
                  // `!e.nativeEvent.isComposing`. Tab is also skipped during
                  // composition because some IMEs use Tab for candidate
                  // navigation (Korean Hangul) and committing on Tab there
                  // would have the same broken effect.
                  if (e.nativeEvent.isComposing) return;
                  if (e.key === 'Enter' || e.key === 'Tab') {
                    e.preventDefault();
                    (e.target as HTMLInputElement).blur();
                  }
                  if (e.key === 'Escape') {
                    cancelRenameRef.current = true;
                    (e.target as HTMLInputElement).blur();
                  }
                }}
                className="bg-transparent outline-none border-b border-primary w-32"
              />
            ) : (
              // Cap long names so a single tab can't push the rest off-screen.
              // Tooltip prefers the full source path (when available) over
              // the bare name — same-basename collisions only resolve when
              // hover reveals the directory. Trailing 「雙擊重新命名」 hint
              // mirrors XlsxEditor's sheet-tab tooltip at line 2580
              // (`${s.name}（雙擊重新命名）`): same inline-rename gesture
              // (onDoubleClick at line 243 above sets renamingId, identical
              // to XlsxEditor.tsx:2568-2571 startRename), but only the
              // workbook's sheet tabs were advertising it — the workspace
              // tabs left the gesture as a tribal-knowledge feature.
              // The file's own docstring at line 6 documents
              // "double-click to rename" for internal readers; the right-
              // click menu's 重新命名 item (line 470) is the only user-
              // facing surface, but users typically reach for hover before
              // right-click. Hint sits after the (Ctrl+N) jump-shortcut so
              // the most-frequently-used cue stays first.
              //
              // 「中鍵關閉」 added alongside 雙擊重新命名 — same kind of
              // mouse-button gesture on the tab itself, fully wired at
              // onAuxClick line 226-242 and called out in the inline code
              // comment there as "universal browser/IDE gesture", yet zero
              // in-app advertisement. Same R64 / R66-class discoverability
              // gap the Ctrl+Tab fix targeted: a feature 100% implemented,
              // 0% surfaced. The × button at line 418 advertises Ctrl+W,
              // the right-click menu surfaces Ctrl+W / Ctrl+Shift+T, the
              // double-click gesture lives in this same tooltip — middle-
              // click-close was the lone close gesture left as tribal
              // knowledge. Middle-dot 「·」 separator matches the R65
              // PresentationMode tooltips ("→ / Space / PageDown") and the
              // existing R39 ContextItem layout convention for compound
              // hint strings, so two gestures read as a list rather than
              // as one merged phrase.
              //
              // 「右鍵選單」 added in R70 to mirror the same-day R69 fix on
              // XlsxEditor.tsx:2580 (sheet tab tooltip). Both surfaces are
              // tab-like draggable navigation elements with right-click
              // context menus, and once the sheet tab adopted 右鍵選單 the
              // workspace tab became the inconsistent one. The argument for
              // this surface is actually STRONGER than for sheet tabs:
              // XlsxEditor's right-click menu had ONE no-fallback action
              // (複製工作表). This menu (lines 497-579) has THREE:
              //   關閉其他   → no toolbar / no shortcut / right-click only
              //   關閉右側   → no toolbar / no shortcut / right-click only
              //   全部關閉   → no toolbar / no shortcut / right-click only
              // (重新命名 / 重新開啟 / 關閉 all have keyboard or visible-
              //  button fallbacks; the close-cluster bulk operations don't.)
              // Three features locked behind a tribal-knowledge gesture.
              // Middle-dot 「·」 stays consistent with the existing 中鍵關閉
              // cue and with XlsxEditor.tsx:2580's "雙擊重新命名 · 右鍵選單",
              // giving tab-like surfaces across the app a single tooltip
              // dialect (TabBar = 雙擊重新命名 · 中鍵關閉 · 右鍵選單,
              // sheet tab = 雙擊重新命名 · 右鍵選單). Tooltip reads as a
              // list of mouse gestures in increasing obscurity, matching
              // R66's stated rationale for the 「·」 separator.
              // 拖曳排序 added in R75 to back-propagate the R73 PptxEditor
              // SlideRail dialect (`（拖曳排序 · 右鍵選單）`) to the two other
              // draggable tab-like surfaces. Drag is the *only* reorder
              // mechanism here — file-header docstring at line 4 explicitly
              // says "reorder is implemented with native HTML5 drag-and-drop"
              // (no chevron buttons, no Ctrl+Shift+PgUp/PgDn keystroke), so
              // the gesture isn't an alternative path, it's the path. Slot
              // sits before 右鍵選單 to keep menu-as-catch-all last and
              // group mouse-button gestures together (中鍵 → drag → 右鍵).
              // Sibling sheet-tab fix at XlsxEditor.tsx:2601 lands the same
              // string in the same insertion point.
              <span
                className="max-w-[220px] truncate"
                title={`${tab.sourcePath ?? tab.name}${jumpHint}${t(
                  '（雙擊重新命名 · 中鍵關閉 · 拖曳排序 · 右鍵選單）',
                  ' (double-click to rename · middle-click to close · drag to reorder · right-click for menu)',
                )}`}
              >
                {tab.name}
                {dirSuffix && (
                  <span className="ml-1 opacity-50 text-[11px]">— {dirSuffix}</span>
                )}
              </span>
            )}
            {tab.dirty && (
              // 6 px primary-coloured circle — visually marks the tab as dirty
              // (unsaved). Without a tooltip the dot is opaque to first-time
              // users: the parent span only reveals `tab.sourcePath`, the
              // export icon advertises export, but the dot itself never
              // explained itself. `aria-label` covers screen readers; `title`
              // covers hover discovery. Matches the dot-with-tooltip pattern
              // we'd expect from VS Code's modified-indicator.
              <span
                className="h-1.5 w-1.5 rounded-full bg-primary"
                title={t('尚未儲存', 'Unsaved')}
                aria-label={t('尚未儲存', 'Unsaved')}
              />
            )}
            {/* Export disabled for fresh binary tabs (docx/xlsx/pptx start at
                byteLength=0) — exportSingleTab would just pop an alert.
                Showing the button visibly-disabled with an explanatory tooltip
                beats the click-then-get-rejected flow. Markdown tabs always
                exportable; an empty .md is still a valid file. */}
            {(() => {
              const canExport =
                tab.type === 'markdown' || tab.type === 'html' || tab.data.byteLength > 0;
              // R125 — same-operation cross-component disclosure asymmetry.
              // App.tsx:1137-1140's top-toolbar export button discloses the
              // extension (`匯出此頁籤為 .docx (Ctrl+E)`) by deriving it from
              // `activeTabType`; this per-tab ↓ button used to read
              // `匯出此頁籤為單一檔案` — generic stand-in chosen before the
              // App.tsx surface landed extension disclosure. The asymmetry
              // was upside-down: App.tsx only knows the *active* tab's type
              // yet discloses it; this row iterates each tab with `tab.type`
              // (line 411) right there in scope and discarded that
              // information. A user with four tabs (.md / .docx / .xlsx /
              // .pptx) hovering each row's ↓ button read four identical
              // tooltips for four different output formats. Mirroring
              // App.tsx's `exportExt` formula brings both export-trigger
              // surfaces — toolbar and per-tab — to one voice. Both call
              // `exportSingleTab` (line 416 here, App.tsx via menu/handlers)
              // so the operation truly is the same; only the surface differs.
              const exportExt = tab.type === 'markdown' ? '.md' : `.${tab.type}`;
              return (
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    if (canExport) void exportSingleTab(tab);
                  }}
                  disabled={!canExport}
                  className={cn(
                    'rounded h-4 w-4 flex items-center justify-center',
                    canExport
                      ? 'opacity-0 group-hover:opacity-100 hover:bg-accent'
                      : 'opacity-0 group-hover:opacity-40 cursor-not-allowed',
                  )}
                  // Ctrl+E only matches THIS button when the tab is active —
                  // App.tsx:472-473 + menu.ts:106-108 export `activeTabId`,
                  // while clicking ↓ on a non-active tab exports that tab
                  // specifically. Same conditional-shortcut pattern the close
                  // × uses below (line 363) for Ctrl+W; gating on isActive
                  // keeps the hint truthful: it advertises Ctrl+E only on the
                  // tab the keystroke actually targets, not on every tab.
                  //
                  // Disabled-state wording mirrors the post-trigger toasts that
                  // surface for the SAME zero-byte-binary-tab condition via
                  // the other two entry points:
                  //   • TabBar.tsx:810 — `notify('這個頁籤還是空的，請先
                  //     在編輯器中編輯內容再匯出。', 'warning')` (the
                  //     fallback path for non-active tab exports)
                  //   • App.tsx:501 — same string verbatim (the Ctrl+E /
                  //     menu:exportTab path)
                  // The two toasts agree exactly. Previously this tooltip
                  // read `頁籤是空的，請先編輯內容才能匯出` — same fact, but
                  // a different sentence: dropped 「這個」, dropped 「在編
                  // 輯器中」, used 「才能匯出」 instead of 「再匯出」. A
                  // user hovering the disabled button got one phrasing,
                  // then on Ctrl+E saw a structurally-different toast for
                  // the same condition. Aligned to the toasts' fact-and-fix
                  // substring so all three entry points (disabled tooltip
                  // / non-active tab toast / Ctrl+E or menu toast) describe
                  // the situation identically. Trailing 。 dropped because
                  // every other Chinese tooltip in the codebase ends without
                  // a period (toasts keep theirs — that's a per-surface
                  // typography norm, not a content difference).
                  title={
                    canExport
                      ? isActive
                        ? t(`匯出此頁籤為 ${exportExt} (Ctrl+E)`, `Export this tab as ${exportExt} (Ctrl+E)`)
                        : t(`匯出此頁籤為 ${exportExt}`, `Export this tab as ${exportExt}`)
                      : t('這個頁籤還是空的，請先在編輯器中編輯內容再匯出', "This tab is empty — edit content in the editor before exporting")
                  }
                >
                  <Download className="h-3 w-3" />
                </button>
              );
            })()}
            <button
              onClick={(e) => {
                e.stopPropagation();
                // R220 — single helper, see closeWithConfirm doc-block.
                closeWithConfirm(tab.id, tab.name);
              }}
              className="opacity-0 group-hover:opacity-100 hover:bg-destructive/20 rounded h-4 w-4 flex items-center justify-center"
              // Surface the Ctrl+W shortcut only on the *active* tab — that's
              // the one Ctrl+W actually closes (App.tsx:747-784). Showing
              // "(Ctrl+W)" universally would lie for non-active tabs: clicking
              // × on Tab B closes Tab B, but Ctrl+W from anywhere closes the
              // active Tab A — different action. Mirrors the conditional
              // shortcut-hint pattern from DiffPreview's `keyboardActive`
              // prop (DiffPreview.tsx:27-33), which advertises the hotkey only
              // on the card the keyboard handler actually targets.
              //
              // R119 — same-row sibling parity. Previously this read
              // `關閉 (Ctrl+W)` / `關閉` (bare verb) while the export button
              // 4 lines up (line 432-438) reads `匯出此頁籤為單一檔案
              // (Ctrl+E)` / `匯出此頁籤為單一檔案` — same row, same scope
              // ("THIS tab"), only the export side disclosed it. Now matches
              // by adding 此頁籤 to both branches. This also aligns with the
              // project's destructive-scope-button convention established by
              // R116 and earlier rounds:
              //   PptxEditor.tsx:3193   '刪除此投影片'  (slide rail)
              //   PptxEditor.tsx:2191   '刪除此文字框'  (frame X)
              //   XlsxEditor.tsx:2696   '刪除工作表'    (sheet tab)
              // — destructive buttons whose target is implied by hover
              // location *still* spell out the noun. The bare 「關閉」 was
              // the lone tab-close outlier. The compact 中鍵關閉 cue at
              // line 383 stays terse because it lives inside a 4-gesture
              // compound tooltip already scoped to this tab — the
              // standalone × button has no such ambient scope.
              title={isActive ? t('關閉此頁籤 (Ctrl+W)', 'Close this tab (Ctrl+W)') : t('關閉此頁籤', 'Close this tab')}
              // R152 — icon-only-button accessible-name parity. Same-file
              // sibling at line 423-424 (尚未儲存 dirty-dot) already pairs
              // `title` with `aria-label`; the canonical icon-only-close
              // doublet is AIPanel.tsx:609-610 (錯誤訊息 X), where the
              // pattern was first established with the comment「title 為
              // hover 視覺、aria-label 為 SR 朗讀」. This × button is the
              // single most-used close affordance in the app yet was the
              // lone TabBar interactive control without an `aria-label`.
              // SR fallback to `title` exists but is inconsistent across
              // engines (NVDA reads it; VoiceOver / Orca skip it on
              // iconified buttons unless `aria-label` is present); the
              // shortcut parenthetical is dropped here per AIPanel.tsx:775
              // 's convention — aria-label carries the action verb only,
              // shortcut belongs in the visual tooltip.
              aria-label={t('關閉此頁籤', 'Close this tab')}
            >
              <X className="h-3 w-3" />
            </button>
          </div>
        );
      })}
      <DropdownMenu>
        {/* Width is auto so the trailing ChevronDown actually fits next to the
            Plus icon — without that hint, the button looks like a plain
            「新增頁籤」 action and users miss the dropdown of formats. The
            tooltip below mirrors what the dropdown actually offers so
            keyboard / hover users get the same affordance.

            R130 — same-file (and cross-component) vocabulary alignment:
            English `tab` → 頁籤. This file already uses 頁籤 18× in
            user-visible strings (the four R126 confirms「${tab.name}」尚未
            儲存，確定關閉？」, R119 close-X 「關閉此頁籤 (Ctrl+W)」, R125
            export 「匯出此頁籤為 ${exportExt}」, the per-tab title at
            line 405 「（雙擊重新命名 · 中鍵關閉 · 拖曳排序 · 右鍵選單）」
            chain with 頁籤 elsewhere…), and the broader codebase has
            36 occurrences across 8 files. The new-tab tooltip below
            and the PptxEditor empty-state two strings (PptxEditor.tsx:
            807 「這個 pptx 頁籤還是空的」, line 810 「先用 Markdown
            頁籤編寫內容」) were the only renderer-visible strings still
            using English `tab` mixed with Chinese; AI tool descriptions
            in ai/tools/index.ts and ai/dispatcher.ts deliberately keep
            English `tab` because that's the developer-facing JSON
            schema served to the LLM, not user-facing text. Aligning
            here brings every user-visible mention of the concept to
            one term. (The Plus-button's quoted exemplar in this very
            comment used to read 「新增 tab」 — same quiet self-
            inconsistency as R121/R128 caught between cross-references
            and the strings they cited; updated above to keep the
            description honest. The two PptxEditor exemplars listed
            above were similarly preserved in pre-R130 form here in
            this very comment even after R130 had aligned them at
            PptxEditor.tsx:807 / 810 — same self-contradicting-quote
            pattern as R133's PPTX context-menu comment, where a
            comment that narrates the fix kept quoting the pre-fix
            string. Updated to match the post-R130 source.) */}
        <DropdownMenuTrigger asChild>
          <Button
            variant="ghost"
            size="sm"
            className="h-7 px-1.5 ml-1 gap-0.5"
            title={t(
              '新增頁籤 — 點擊選擇格式 (Markdown / HTML / Word / Excel / PowerPoint)',
              'New tab — click to pick a format (Markdown / HTML / Word / Excel / PowerPoint)',
            )}
          >
            <Plus className="h-4 w-4" />
            <ChevronDown className="h-3 w-3 opacity-60" />
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start">
          <DropdownMenuItem onSelect={() => addTab('markdown')}>
            <FileCode2 className="h-4 w-4 text-sky-500" /> Markdown
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => addTab('html')}>
            <Code2 className="h-4 w-4 text-rose-500" /> HTML
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => addTab('docx')}>
            <FileText className="h-4 w-4 text-indigo-500" /> Word
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => addTab('xlsx')}>
            <FileSpreadsheet className="h-4 w-4 text-emerald-500" /> Excel
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => addTab('pptx')}>
            <Presentation className="h-4 w-4 text-orange-500" /> PowerPoint
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>

      {ctxMenu &&
        (() => {
          const idx = tabs.findIndex((t) => t.id === ctxMenu.tabId);
          if (idx < 0) return null;
          const tab = tabs[idx];
          const otherIds = tabs.filter((t) => t.id !== tab.id).map((t) => t.id);
          const rightIds = tabs.slice(idx + 1).map((t) => t.id);
          const allIds = tabs.map((t) => t.id);
          // Stop mousedown from bubbling to the window-level dismiss
          // handler so users can actually click an item.
          // Estimate generous enough to bound the menu's actual rendered
          // size: 5 items + 1 separator at ~28px each ≈ 168px tall; min-w
          // 180px + padding ≈ 200px wide. Caller-side clamp avoids the
          // flicker of measure-and-adjust patterns. See clampToViewport's
          // docstring for the trade-off.
          const pos = clampToViewport(ctxMenu.x, ctxMenu.y, 200, 200);
          return (
            <div
              role="menu"
              onMouseDown={(e) => e.stopPropagation()}
              style={{ left: pos.left, top: pos.top }}
              className="fixed z-50 min-w-[180px] rounded-md border bg-popover text-popover-foreground shadow-md py-1 text-xs"
            >
              <ContextItem
                label={t('重新命名', 'Rename')}
                onClick={() => {
                  setRenamingId(tab.id);
                  setCtxMenu(null);
                }}
              />
              <div className="my-1 h-px bg-border" />
              {/* 重新開啟剛才關閉的頁籤 — surfaces Ctrl+Shift+T, which is bound
                  globally at App.tsx:734-746 and pops the head of
                  `recentlyClosedTabs` (workspace.ts:467-490). The shortcut
                  was previously invisible: the only mention anywhere in the
                  app was an internal comment block (TabBar.tsx:104, the
                  `closeTabWithFlush` doc) — fully implemented recovery
                  feature with zero discoverability for users who don't
                  already know the universal browser/IDE muscle memory.
                  Placing it inside the close-related cluster mirrors VS
                  Code's tab right-click menu ("Reopen Closed Editor" sits
                  alongside the Close family). Disabled when there's nothing
                  to restore — the keystroke itself no-ops in that state
                  (line 469), so the menu item should mirror that truth
                  rather than ghost-fire when clicked. */}
              {/* `disabledReason` (added alongside this round's R82 fix)
                  carries the boundary explanation users get on hover when
                  the menu item ghosts out. Same pattern R80 / R81 applied
                  to PptxEditor's SlideRail context menu and XlsxEditor's
                  sheet-tab context menu — but here the gap was sharper:
                  the close-cluster comment ~200 lines up (TabBar.tsx:355-
                  361) explicitly notes 關閉其他 / 關閉右側 / 全部關閉 are
                  "no toolbar / no shortcut / right-click only … Three
                  features locked behind a tribal-knowledge gesture." When
                  one of these items disables at a boundary the user has
                  ZERO feedback (no hover hint) AND ZERO alternative entry
                  points (no toolbar, no keystroke). Click, nothing
                  happens, no explanation, no fallback. The R80 / R81
                  fixes had a comparatively easier path (their context
                  menus had visible-button siblings already carrying the
                  boundary message) — here, the right-click menu IS the
                  only surface, so closing the gap means the menu has to
                  carry the explanation itself.
                  Wording follows the established app-wide vocabulary:
                  "已經是X" for navigation boundaries (matches PptxEditor's
                  `已經是第一張` / `已經是最後一張` at lines 2828 / 2835),
                  descriptive phrasing for state-of-world boundaries that
                  don't fit the "已經是X" mold. */}
              <ContextItem
                label={t('重新開啟剛才關閉的頁籤', 'Reopen recently closed tab')}
                shortcut="Ctrl+Shift+T"
                disabled={recentlyClosedCount === 0}
                disabledReason={t('目前沒有可重新開啟的頁籤', 'No recently closed tabs to reopen')}
                onClick={() => {
                  useWorkspace.getState().reopenClosedTab();
                  setCtxMenu(null);
                }}
              />
              <div className="my-1 h-px bg-border" />
              {/* `Ctrl+W` always targets the *active* tab, and right-click
                  pre-activates the tab it lands on (line 224 above) — so by
                  the time this menu is open, the right-clicked tab IS the
                  active tab, and 關閉 here does the same thing as Ctrl+W
                  globally. Surfacing the shortcut closes the parity gap with
                  the × button (line 384), which already advertises Ctrl+W on
                  the active tab. The other close items (其他 / 右側 / 全部)
                  have no global keystroke, so we leave their shortcut empty
                  rather than fabricate one. */}
              <ContextItem
                label={t('關閉', 'Close')}
                shortcut="Ctrl+W"
                onClick={() => {
                  // Close the menu first so it doesn't sit on top of the
                  // dialog. R220 — single helper handles live-dirty read
                  // + per-tab double-fire gate.
                  setCtxMenu(null);
                  closeWithConfirm(tab.id, tab.name);
                }}
              />
              <ContextItem
                label={t('關閉其他', 'Close Others')}
                disabled={otherIds.length === 0}
                disabledReason={t('目前只有這一個頁籤', 'Only this tab is open')}
                onClick={() => {
                  void closeMany(otherIds, t('關閉其他', 'Close Others'));
                  setCtxMenu(null);
                }}
              />
              <ContextItem
                label={t('關閉右側', 'Close to the Right')}
                disabled={rightIds.length === 0}
                disabledReason={t('已經是最右側的頁籤', 'Already the rightmost tab')}
                onClick={() => {
                  void closeMany(rightIds, t('關閉右側', 'Close to the Right'));
                  setCtxMenu(null);
                }}
              />
              {/* 全部關閉 is intentionally left without a `disabledReason`:
                  `allIds.length === 0` is unreachable from this menu (the
                  context menu only opens from a tab right-click — line 511
                  reads `tabs[idx]`, so by the time this branch renders
                  `tabs.length >= 1`). The `disabled` prop stays as
                  defensive dead code rather than fabricating a hint for
                  a state users can't reach. */}
              <ContextItem
                label={t('全部關閉', 'Close All')}
                disabled={allIds.length === 0}
                onClick={() => {
                  void closeMany(allIds, t('全部關閉', 'Close All'));
                  setCtxMenu(null);
                }}
              />
            </div>
          );
        })()}
    </div>
  );
}

function ContextItem({
  label,
  onClick,
  disabled,
  disabledReason,
  shortcut,
}: {
  label: string;
  onClick: () => void;
  disabled?: boolean;
  /** Hover hint shown only when the item is disabled — explains the
   *  boundary that ghosted it out so the user has a fallback signal
   *  beyond a silently-greyed click target. Same pattern PptxEditor's
   *  rail-icon delete button uses (PptxEditor.tsx:2845
   *  `title={slideCount <= 1 ? '至少要保留一張投影片' : '刪除'}`) and the
   *  follow-on R80 / R81 fixes wired into the SlideRail and sheet-tab
   *  context menus. The gap is sharper for THIS component because the
   *  close-cluster items (關閉其他 / 關閉右側) are right-click-only —
   *  comment block ~200 lines up at line 354-361 calls them out as
   *  "Three features locked behind a tribal-knowledge gesture." When
   *  they disable, there's no toolbar / no shortcut / no sibling
   *  surface to fall back on. Without a tooltip the user clicks, nothing
   *  happens, no signal whatsoever. Active-state stays empty (rather
   *  than echoing `label`) for the same reason as R80 / R81: the button
   *  already carries the visible label, so an active-state tooltip
   *  would just stack a redundant browser tooltip on top. */
  disabledReason?: string;
  /** Right-aligned keyboard hint (e.g. "Ctrl+W"). VS Code-style: dimmed and
   *  monospaced so the shortcut never visually competes with the label.
   *  Pass undefined when the action has no global shortcut wired up — we
   *  don't fabricate hints for actions that aren't actually bound. */
  shortcut?: string;
}) {
  return (
    <button
      type="button"
      role="menuitem"
      disabled={disabled}
      title={disabled && disabledReason ? disabledReason : undefined}
      onClick={onClick}
      className={cn(
        'w-full flex items-center gap-3 px-3 py-1.5 hover:bg-accent disabled:opacity-40 disabled:pointer-events-none',
      )}
    >
      <span className="flex-1 text-left">{label}</span>
      {shortcut && (
        <span className="text-[10px] font-mono text-muted-foreground tabular-nums">
          {shortcut}
        </span>
      )}
    </button>
  );
}

/**
 * Export a single tab to disk in its native format. Shared by the per-tab
 * download icon here and the toolbar / Ctrl+E entry-point — both bottom out
 * in `window.gendoc.workspace.exportTab`.
 */
// R197 / R225 — gate moved to lib/export-tab-busy.ts so this path AND
// App.tsx::handleExportTab (toolbar / Ctrl+E) share the same in-flight
// invariant. Cross-fire (per-tab download click + toolbar Download click,
// or per-tab + Ctrl+E within the same hover frame) used to queue two
// `workspace.exportTab` IPCs and stack two OS save dialogs because R197's
// component-local useRef and this module-local boolean couldn't see each
// other. The OS save dialog is app-wide modal-ish so a global gate is the
// correct granularity — gating per-tab still allows different-tab
// stacking, which is just as confusing.
async function exportSingleTab(tab: Tab): Promise<void> {
  if (!tryEnterExportTab()) {
    // R322 — surface busy state. Same toast shape / rationale as
    // handleExportTab in App.tsx; toast.ts dedupes the message+variant
    // so rapid double-fire stays quiet.
    notify(tImp('正在匯出中…請等目前的匯出完成再試', 'Exporting… please wait for the current export to finish before retrying'), 'info');
    return;
  }
  try {
  // Flush pending debounced serializes so the export reflects the on-screen
  // state, then re-read the tab from the store — the closure arg is stale
  // once flushers patch new bytes in. Mirrors the Ctrl+E export path.
  // R172 — capture workspaceId pre-IPC so the post-await flashExport
  // doesn't paint「已匯出 …」 onto a workspace the user has since swapped
  // away from. Same shape as R171 (App.tsx handleExportTab) and R172
  // sibling in MarkdownEditor.handleExportPdf.
  const exportWorkspaceId = useWorkspace.getState().workspaceId;
  await flushEditors();
  if (useWorkspace.getState().workspaceId !== exportWorkspaceId) return;
  const fresh = useWorkspace.getState().tabs.find((t) => t.id === tab.id) ?? tab;
  let bytes: Uint8Array;
  let ext: 'md' | 'html' | 'docx' | 'xlsx' | 'pptx';
  if (fresh.type === 'markdown') {
    bytes = new TextEncoder().encode(fresh.content);
    ext = 'md';
  } else if (fresh.type === 'html') {
    bytes = new TextEncoder().encode(fresh.content);
    ext = 'html';
  } else {
    if (fresh.data.byteLength === 0) {
      notify(tImp('這個頁籤還是空的，請先在編輯器中編輯內容再匯出', 'This tab is empty — edit content in the editor before exporting'), 'warning');
      return;
    }
    bytes = fresh.data;
    ext = fresh.type;
  }
  try {
    const res = await window.gendoc.workspace.exportTab({
      ext,
      suggestedName: fresh.name,
      bytes,
    });
    // R172 — workspace guard. Disk write is OLD's bytes (correct), only
    // the in-renderer flash is gated.
    if (useWorkspace.getState().workspaceId !== exportWorkspaceId) return;
    // Mirror App.tsx's Ctrl+E path so the per-tab download icon also yields
    // a visible "已匯出" flash — without this the icon click feels like a
    // no-op once the save dialog closes.
    const fileName = res.filePath.split(/[/\\]/).pop() ?? fresh.name;
    useWorkspace.getState().flashExport(fileName, res.filePath);
  } catch (e) {
    const msg = (e as Error).message;
    if (msg !== 'export_cancelled') notify(tImp(`匯出失敗：${msg}`, `Export failed: ${msg}`), 'error');
  }
  } finally {
    // R225 — release shared gate.
    exitExportTab();
  }
}
