/**
 * The middle pane: tab bar on top, then the active tab's editor renders
 * the appropriate component.
 */

import { useEffect, useState } from 'react';
import {
  Clock,
  Code2,
  FileCode2,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Presentation,
  Sparkles,
} from 'lucide-react';
import { useWorkspace } from '../store/workspace';
import { notify } from '../store/toast';
import { bumpLoadGen, currentLoadGen } from '../lib/workspace-load-gen';
import { exitOpen, tryEnterOpen } from '../lib/workspace-open-busy';
import { isFileMissingError } from '../lib/utils';
import { useT, tImp } from '../lib/i18n';
import type { TabType } from '../types/manifest';
import { DocxEditor } from './DocxEditor';
import { ErrorBoundary } from './ErrorBoundary';
import { HtmlEditor } from './HtmlEditor';
import { MarkdownEditor } from './MarkdownEditor';
import { PptxEditor } from './PptxEditor';
import { TabBar } from './TabBar';
import { XlsxEditor } from './XlsxEditor';

export function EditorSurface(): JSX.Element {
  const { tabs, activeTabId } = useWorkspace();
  const active = tabs.find((t) => t.id === activeTabId);

  // Per-tab key on the boundary so switching tabs always resets a previously
  // caught error — otherwise a broken pptx tab would keep showing the error
  // screen even after you switch to a healthy markdown tab.
  return (
    <div className="flex flex-col h-full">
      <TabBar />
      {/* `work-canvas` re-pins the palette CSS vars to the light values
          inside this subtree so all editor canvases (markdown source +
          preview, docx paper desk, xlsx cells, pptx slides) render on
          white regardless of the user's chosen UI theme. See index.css. */}
      <div className="flex-1 min-h-0 relative work-canvas">
        {!active && <EmptyState />}
        {active ? (
          <ErrorBoundary key={active.id} label={`Tab: ${active.name} (${active.type})`}>
            {active.type === 'markdown' && <MarkdownEditor tab={active} />}
            {active.type === 'html' && <HtmlEditor tab={active} />}
            {active.type === 'docx' && <DocxEditor tab={active} />}
            {active.type === 'xlsx' && <XlsxEditor tab={active} />}
            {active.type === 'pptx' && <PptxEditor tab={active} />}
          </ErrorBoundary>
        ) : null}
      </div>
    </div>
  );
}

interface QuickStart {
  type: TabType;
  label: string;
  hint: string;
  icon: React.ComponentType<{ className?: string }>;
  iconClass: string;
}

const QUICK_STARTS: QuickStart[] = [
  {
    type: 'markdown',
    label: 'Markdown 筆記',
    hint: '純文字 + 段落結構，最快開始的格式',
    icon: FileCode2,
    iconClass: 'text-sky-500 bg-sky-500/10',
  },
  {
    type: 'html',
    label: 'HTML 文件',
    hint: '原始碼 / 對照 / 預覽三模式，可離線匯出 .html',
    icon: Code2,
    iconClass: 'text-rose-500 bg-rose-500/10',
  },
  {
    type: 'docx',
    label: 'Word 文件',
    hint: '段落 / 標題 / 列表結構編輯',
    icon: FileText,
    iconClass: 'text-indigo-500 bg-indigo-500/10',
  },
  {
    // R122 — two converging gaps fixed (parallel to R117's pptx fix
    // 27 lines below):
    //
    // 1. Mixed-language hint. Previously read `多 sheet、cell 值編輯`,
    //    using English `sheet` and `cell` while the other three cards
    //    use full Chinese unit names (Markdown 段落, Docx 段落 / 標題 /
    //    列表, Pptx 投影片 / 文字框 / 圖形). Just as importantly it
    //    contradicted the editor it opens: XlsxEditor.tsx:2696 reads
    //    `刪除工作表`, line 2401 reads `輸入內容或 = 公式後按 Enter`,
    //    lines 2321-2322 read `公式互相循環參照` / `公式語法錯誤` — every
    //    user-facing surface inside the actual Excel editor uses 工作表
    //    / 儲存格 / 公式. The promotion card spoke a different language
    //    than the editor it advertises.
    //
    // 2. Shape parity. R117 deliberately mirrored docx's slash-list
    //    shape onto pptx so the cards "read as a parallel set
    //    describing their editable units" (see comment on the pptx
    //    entry below). xlsx was untouched — comma separator, different
    //    structure, leaving 3 of 4 cards in slash-list shape and xlsx
    //    as the lone holdout. Aligning xlsx onto the same slash-list
    //    shape closes that 3-vs-1 gap.
    //
    // 結構編輯 (used by docx / pptx) is dropped on purpose: Excel's
    // primary edit model is data-centric (cell values + formulas),
    // not structural like docx paragraphs / pptx shapes. Bare `編輯`
    // is the honest verb for the content-level operations in
    // XlsxEditor (CellInput's commit-on-Enter, the formula bar's
    // 輸入內容或 = 公式 path).
    type: 'xlsx',
    label: 'Excel 試算表',
    hint: '工作表 / 儲存格 / 公式 編輯',
    icon: FileSpreadsheet,
    iconClass: 'text-emerald-500 bg-emerald-500/10',
  },
  {
    // R117 — two converging gaps fixed:
    //
    // 1. Label parity. Sibling QUICK_STARTS pair `English format + Chinese
    //    descriptor`: 'Markdown 筆記' / 'Word 文件' / 'Excel 試算表'. The
    //    PowerPoint card was the lone outlier with no Chinese descriptor,
    //    even though the welcome paragraph 17 lines above (line 157) had
    //    already established the four parallel terms「筆記、文件、表格
    //    與簡報」 in the same component. 簡報 is the dominant document-
    //    level term across the codebase (used in pptx-adapter.ts:17 and
    //    65× in PptxEditor.tsx); adopting it here makes the welcome screen
    //    speak with one voice.
    //
    // 2. Hint honesty. Previously read「需從既有 .pptx 載入；可改投影片
    //    文字」— a stale constraint. PptxEditor.tsx:214-221 explicitly
    //    bootstraps zero-byte pptx tabs via createBlankPptx() (pptx-
    //    adapter.ts:176-190 builds a minimal valid one-slide deck with a
    //    title placeholder ready to edit), and the sibling docx / xlsx
    //    adapters do the same (docx-adapter.ts:194-200, xlsx-adapter.ts:
    //    128-137). All four formats are equally usable from blank — the
    //    pptx hint was the one string that never got updated when the
    //    bootstrap landed. A fresh user reading "需從既有 .pptx 載入"
    //    would conclude the card can't actually create a new pptx and
    //    either skip it or go fishing for an existing file, losing a
    //    feature that's been wired all along. The new hint mirrors
    //    docx's `段落 / 標題 / 列表結構編輯` shape verbatim so the four
    //    cards read as a parallel set describing their editable units.
    type: 'pptx',
    label: 'PowerPoint 簡報',
    hint: '投影片 / 文字框 / 圖形結構編輯',
    icon: Presentation,
    iconClass: 'text-orange-500 bg-orange-500/10',
  },
];

function EmptyState() {
  const addTab = useWorkspace((s) => s.addTab);
  // R405 — bilingual via useT(). QUICK_STARTS at module scope keeps the
  // structural metadata (icon, iconClass, type); the human-readable label /
  // hint pair is resolved per-locale at render time via the `labels` map
  // below. Re-renders on locale change flip every card / heading / button.
  const t = useT();
  const quickLabels: Record<TabType, { label: string; hint: string }> = {
    markdown: {
      label: t('Markdown 筆記', 'Markdown notes'),
      hint: t('純文字 + 段落結構，最快開始的格式', 'Plain text + paragraph structure — the fastest format to start with'),
    },
    html: {
      label: t('HTML 文件', 'HTML document'),
      hint: t('原始碼 / 對照 / 預覽三模式，可離線匯出 .html', 'Source / split / preview modes; exportable to a self-contained .html'),
    },
    docx: {
      label: t('Word 文件', 'Word document'),
      hint: t('段落 / 標題 / 列表結構編輯', 'Paragraph / heading / list structural editing'),
    },
    xlsx: {
      label: t('Excel 試算表', 'Excel spreadsheet'),
      hint: t('工作表 / 儲存格 / 公式 編輯', 'Sheet / cell / formula editing'),
    },
    pptx: {
      label: t('PowerPoint 簡報', 'PowerPoint presentation'),
      hint: t('投影片 / 文字框 / 圖形結構編輯', 'Slide / text-box / shape structural editing'),
    },
  };
  // Recent files surfaced inline — Adobe / VS Code / IntelliJ all put recent
  // documents on the start screen so re-opening yesterday's work is one
  // click instead of File → 最近開啟 → submenu. Loaded once on mount; the
  // EmptyState is unmounted as soon as a tab opens, so we don't need to
  // subscribe to changes.
  const [recents, setRecents] = useState<string[]>([]);
  useEffect(() => {
    let alive = true;
    // R245 — `.catch(() => undefined)` swallow so a config.get rejection
    // (loadConfig falling through to fs.writeFile of DEFAULTS, then THAT
    // failing on EACCES / disk full) doesn't surface as
    // unhandledrejection from the EmptyState's mount effect. The Welcome
    // screen's recents preload is best-effort — if it fails, the user
    // still has File → 最近開啟 in the native menu, the EmptyState
    // just shows zero recents. Aligns with the R241 / R242 / R245
    // (handleUndo's recovery push) sweep that established `.catch(()
    // => undefined)` as the canonical "this is fire-and-forget"
    // pattern for non-critical post-event reads / writes.
    void window.gendoc.config
      .get()
      .then((cfg) => {
        if (alive) setRecents(cfg.recentFiles.slice(0, 5));
      })
      .catch(() => undefined);
    return () => {
      alive = false;
    };
  }, []);

  const onOpenExisting = async () => {
    // R205 — sync gate + load-gen, mirroring the App.tsx handleOpen pattern
    // (workspace-open-busy.ts doc-block enumerates the entry points the
    // R199 sweep originally covered: handleOpen / handleOpenRecent / FileExplorer
    // entry mash). The welcome-screen open buttons were left out because
    // EmptyState's onOpenExisting / onOpenRecent live in EditorSurface,
    // not App, and were missed by the original sweep. Without the gate,
    // double-clicking 「開啟 .gd 專案」 fires two `dialog.showOpenDialog`
    // calls in main and stacks two OS pickers on the parent BrowserWindow;
    // without the load-gen check, two recents clicked in quick succession
    // produce two `loadFromOpened` calls whose order depends on which IPC
    // resolves first (not which the user clicked last). Same shape as the
    // doc-block's recommended usage: gate before any await, capture
    // myGen before the IPC, drop on stale, exitOpen in finally.
    //
    // Match the error-handling contract from onOpenRecent below: corrupt
    // zip, EACCES on a locked .gd, or a malformed manifest can all reject
    // the IPC. Without try/catch the empty-state "開啟現有檔案" button
    // silently no-ops on failure — the user clicks the picker, picks a
    // bad file, the dialog closes, and they're left looking at the same
    // empty state with no signal anything happened. Notify aligns the
    // pick-from-dialog path with the recents path next door.
    if (!tryEnterOpen()) return;
    try {
      // R247 — dirty-confirm gate, mirroring the five sibling
      // workspace-replacement entries (handleNew / handleOpen /
      // handleOpenRecent / FileExplorer .gd-open / Ctrl+O via App.tsx
      // keymap), all of which already gate on `useWorkspace.getState()
      // .dirty && !(await app.confirm(...))`. EmptyState's two open
      // buttons were the holdouts. Concrete trigger: user opens
      // workspace A.gd, edits T1 / T2 / T3, closes them all via Ctrl+W
      // (each close prompts per-tab dirty + the user clicks 確定 to
      // close). All tabs gone → EmptyState renders, but workspace-level
      // `dirty=true` because each `removeTab` flips it (workspace.ts:
      // 487-491). At this point Ctrl+S would write an EMPTY workspace
      // back to A.gd, destroying the user's data. Ctrl+Shift+T can
      // recover the closed tabs from `recentlyClosedTabs`. The
      // dirty-confirm here is the user's last warning to recover via
      // Ctrl+Shift+T before swapping workspace; without it, clicking
      // 「開啟現有檔案」 silently swaps to B.gd and the recently-closed
      // stack is wiped by `loadFromOpened` (workspace.ts:263). Same
      // prompt copy as App.tsx handleOpen for cross-entry-point UX
      // parity.
      if (
        useWorkspace.getState().dirty &&
        !(await window.gendoc.app.confirm(tImp(
          '目前的變更尚未儲存，確定開啟其他檔案？',
          'There are unsaved changes. Open a different file anyway?',
        )))
      )
        return;
      const myGen = bumpLoadGen();
      const opened = await window.gendoc.workspace.open();
      if (myGen !== currentLoadGen()) return;
      if (opened) useWorkspace.getState().loadFromOpened(opened);
    } catch (err) {
      notify(tImp(`開啟失敗：${(err as Error).message}`, `Failed to open: ${(err as Error).message}`), 'error');
    } finally {
      exitOpen();
    }
  };

  const onOpenRecent = async (filePath: string) => {
    // R205 — same gate + load-gen as onOpenExisting above. The
    // welcome-screen「最近開啟」list was missed by R199's sweep along with
    // its sibling button. Mashing the same recent row, or clicking row A
    // then row B before A's IPC returns, produced two competing
    // loadFromOpened calls. After the gate, the second click no-ops; the
    // load-gen check is a belt-and-braces second line for the case where
    // the second click landed via a different code path (e.g. native
    // menu's handleOpenRecent in App.tsx) that already passed the gate
    // and incremented the counter — in that case we must drop our own
    // resolve so the user's later click wins.
    if (!tryEnterOpen()) return;
    try {
      // R247 — same dirty-confirm gate as onOpenExisting; see doc-block
      // there for the close-all-tabs / Ctrl+Shift+T recovery rationale.
      if (
        useWorkspace.getState().dirty &&
        !(await window.gendoc.app.confirm(tImp(
          '目前的變更尚未儲存，確定開啟其他檔案？',
          'There are unsaved changes. Open a different file anyway?',
        )))
      )
        return;
      const myGen = bumpLoadGen();
      const opened = await window.gendoc.workspace.openPath(filePath);
      if (myGen !== currentLoadGen()) return;
      if (opened) useWorkspace.getState().loadFromOpened(opened);
    } catch (err) {
      // Most common cause: the file was moved / deleted since it landed in
      // recents. Same prune-on-fail policy as the menu's handleOpenRecent
      // so a second click doesn't repeat the dead-path error.
      // R395 — only prune when the failure is "file is gone from disk"
      // (ENOENT). Other failure classes — corrupted manifest (R376 / R377 /
      // R390 validation throws), EACCES from a transient antivirus lock,
      // app.confirm rejection from rare main IPC anomaly — leave the .gd
      // intact on disk; permanently dropping the recent entry for those
      // would erase user state for a file that's still recoverable. The
      // welcome-screen list (this site) is just as bad as App.tsx's two
      // R392-fixed sites: the user CAN re-add via Ctrl+O afterwards, but
      // discovering the recent entry vanished without explanation is the
      // surprise we're closing. See isFileMissingError doc-block in
      // lib/utils.ts for the asymmetry rationale.
      notify(tImp(`開啟失敗：${(err as Error).message}`, `Failed to open: ${(err as Error).message}`), 'error');
      if (isFileMissingError(err)) {
        try {
          const cfg = await window.gendoc.config.get();
          const next = cfg.recentFiles.filter((p) => p !== filePath);
          await window.gendoc.config.set({ recentFiles: next });
          setRecents(next.slice(0, 5));
        } catch {
          /* config IPC offline — next launch will retry */
        }
      }
    } finally {
      exitOpen();
    }
  };

  return (
    <div className="h-full w-full overflow-auto">
      <div className="max-w-2xl mx-auto px-6 py-12 space-y-8">
        <header className="space-y-2">
          <div className="flex items-center gap-2 text-primary">
            <Sparkles className="h-5 w-5" />
            <h1 className="text-2xl font-semibold text-foreground">
              {t('歡迎使用 Gen Doc', 'Welcome to Gen Doc')}
            </h1>
          </div>
          <p className="text-sm text-muted-foreground">
            {t(
              '一個 .gd 檔可以同時放筆記、HTML、文件、表格與簡報。右側的 AI 對話框會根據你選的內容做結構化編輯。',
              'A single .gd workspace can hold notes, HTML, documents, sheets, and slides. The AI panel on the right makes structured edits to whatever you select.',
            )}
          </p>
        </header>

        <section className="space-y-3">
          <h2 className="text-xs uppercase tracking-wider text-muted-foreground">
            {t('建立新頁籤', 'Create a new tab')}
          </h2>
          <div className="grid grid-cols-2 gap-2">
            {QUICK_STARTS.map(({ type, icon: Icon, iconClass }) => {
              const { label, hint } = quickLabels[type];
              return (
                <button
                  key={type}
                  type="button"
                  onClick={() => addTab(type)}
                  className="group flex items-start gap-3 px-3 py-3 rounded-lg border bg-card hover:border-primary/40 hover:bg-accent/40 transition-colors text-left"
                >
                  <div className={`p-2 rounded-md ${iconClass}`}>
                    <Icon className="h-4 w-4" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-medium">{label}</div>
                    <div className="text-xs text-muted-foreground line-clamp-2">{hint}</div>
                  </div>
                </button>
              );
            })}
          </div>
        </section>

        <section className="space-y-3">
          <h2 className="text-xs uppercase tracking-wider text-muted-foreground">
            {t('或開啟既有檔案', 'Or open an existing file')}
          </h2>
          <button
            type="button"
            onClick={() => void onOpenExisting()}
            className="w-full flex items-center gap-3 px-3 py-3 rounded-lg border bg-card hover:border-primary/40 hover:bg-accent/40 transition-colors text-left"
          >
            <div className="p-2 rounded-md bg-primary/10 text-primary">
              <FolderOpen className="h-4 w-4" />
            </div>
            <div className="flex-1">
              <div className="text-sm font-medium">
                {t('開啟 .gd 專案', 'Open a .gd project')}
              </div>
              <div className="text-xs text-muted-foreground">
                <kbd className="px-1 rounded bg-secondary border text-[10px]">Ctrl/⌘+O</kbd>
              </div>
            </div>
          </button>
          {/* Drag-drop file open is fully wired at App.tsx:858-977 with a
              polished dashed-border overlay (App.tsx:1058-1069) listing all
              five supported types — but the overlay only appears DURING the
              drag, so a fresh user landing on this welcome screen has no
              way to know the gesture exists short of randomly trying it.
              Same R64 (Ctrl+Tab) / R66 (中鍵關閉) pattern: feature 100%
              implemented, 0% advertised at the fresh-user discovery moment.
              Welcome-screen placement is symmetric with the existing
              "開啟 .gd 專案" button right above — same "或開啟既有檔案"
              section, same fresh-user discovery slot. The supported-types
              list mirrors the drop-overlay copy verbatim (App.tsx:1065)
              so the hint and the in-drag overlay phrase the gesture the
              same way (cross-surface tooltip parity, same convention
              R63 used for export tooltips). Single muted line keeps the
              welcome layout tight without competing visual weight against
              the primary Ctrl+O button. */}
          <p className="text-[11px] text-muted-foreground px-1">
            {t(
              '也可將 .gd / .md / .html / .docx / .xlsx / .pptx 檔案拖曳到此視窗開啟',
              'You can also drag .gd / .md / .html / .docx / .xlsx / .pptx files onto this window to open them',
            )}
          </p>
        </section>

        {recents.length > 0 && (
          <section className="space-y-2">
            <h2 className="text-xs uppercase tracking-wider text-muted-foreground">
              {t('最近開啟', 'Recent')}
            </h2>
            <ul className="rounded-lg border bg-card divide-y">
              {recents.map((p) => {
                const i = Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\'));
                const name = i >= 0 ? p.slice(i + 1) : p;
                const dir = i >= 0 ? p.slice(0, i) : '';
                return (
                  <li key={p}>
                    <button
                      type="button"
                      onClick={() => void onOpenRecent(p)}
                      className="w-full flex items-center gap-3 px-3 py-2 text-left hover:bg-accent/40 transition-colors"
                      title={p}
                    >
                      <Clock className="h-3.5 w-3.5 text-muted-foreground shrink-0" />
                      <div className="flex-1 min-w-0">
                        <div className="text-sm truncate">{name}</div>
                        {dir && (
                          <div className="text-[11px] text-muted-foreground truncate">
                            {dir}
                          </div>
                        )}
                      </div>
                    </button>
                  </li>
                );
              })}
            </ul>
          </section>
        )}

        <section className="text-xs text-muted-foreground border-t pt-4 space-y-1.5">
          <div className="font-medium text-foreground">常用快捷鍵</div>
          {/* Ctrl+Tab / Ctrl+Shift+Tab — sequential tab navigation, wired at
              App.tsx:719-733. Universal browser / VS Code / IDE muscle
              memory ("the keystroke users already know"), but until now the
              only mentions in the entire codebase were four internal code
              comments (App.tsx:720, FileExplorer.tsx:205, TabBar.tsx:73,
              MarkdownEditor.tsx:107) — zero user-facing UI surface. Every
              other globally-bound nav shortcut has at least one tooltip /
              menu / footer hint:
                Ctrl+W       → per-tab × tooltip on active tab (TabBar:524)
                Ctrl+Shift+T → context-menu item with shortcut col (TabBar:662)
                Ctrl+1..9    → per-tab tooltip suffix (TabBar:197)
                Ctrl+L       → AI panel header tooltip (AIPanel:707)
                Ctrl+Shift+F → focus-mode toolbar button (App.tsx:1296)
              Ctrl+Tab was the lone orphan. Welcome-footer placement is
              symmetric with its index-nav cousin Ctrl+1–9 right above —
              same fresh-user discovery slot, same explicit "常用快捷鍵"
              header. The hint is only seen at startup before any tab
              opens, but that's also the moment a user is most receptive
              to learning navigation gestures (mirrors R61 EmptyState
              rationale). The 「切換上/下個頁籤」 label uses 上/下個 (vs.
              previous/next) to match the directional vocabulary already
              in use for slide-rail nav (PptxEditor.tsx:1064 Ctrl+PgUp/
              PgDn tooltip says "上/下一張") so the in-app phrasing
              for sequential nav stays consistent. */}
          {/* Ctrl+F — 尋找與取代, wired in ALL FOUR editor surfaces:
                MarkdownEditor.tsx:479 (CM6 built-in openSearchPanel)
                DocxEditor.tsx:739-783 (custom FindReplaceDialog)
                XlsxEditor.tsx:1009-1012 (custom FindReplaceDialog)
                PptxEditor.tsx:680-702 (custom FindReplaceDialog)
              Each editor's own toolbar already advertises it via:
                DocxEditor.tsx:2231      title="尋找與取代 (Ctrl+F)"
                MarkdownToolbar.tsx:233   title={... ? disabledTitle : '尋找與取代 (Ctrl+F)'}
                XlsxEditor.tsx:2089      title="尋找與取代 (Ctrl+F)"
                PptxEditor.tsx:1386      title="尋找與取代 (Ctrl+F)"
              and a code comment at DocxEditor.tsx:2243 explicitly groups
              Ctrl+F with Ctrl+B/I/U as peer toolbar-bound shortcuts — yet
              the welcome footer (the canonical "常用快捷鍵" list users
              read at startup) omitted Ctrl+F entirely. Same R64 (Ctrl+Tab)
              pattern as above: shortcut wired four times over with rich
              dialog UX, but the fresh-user discovery surface stays silent.
              Among universal editor shortcuts Ctrl+F arguably ranks below
              only Ctrl+S in muscle memory (browser / VS Code / Word /
              Excel / PowerPoint all bind it identically), so its absence
              from the footer was the loudest gap. The grouping next to
              Ctrl+S keeps "edit-the-active-tab" shortcuts visually
              clustered before navigation (Ctrl+L → Ctrl+Tab cluster). */}
          {/* Ctrl+Shift+Z — 重做 AI 變更, the orphaned twin of Ctrl+Z above.
              Wired at App.tsx:563-580 (handleRedo), App.tsx:799-818 (keymap
              with same focus-skip rules as Ctrl+Z), menu at App.tsx:608
              ('menu:redo'). The toolbar's Redo2 button at App.tsx:1233-1241
              already advertises the keystroke via title="重做 AI 變更
              (Ctrl+Shift+Z)" — exactly mirroring the Undo2 button's title at
              App.tsx:1229 ("復原 AI 變更 (Ctrl+Z)"). Yet the welcome footer
              listed Ctrl+Z without its redo twin: a fresh user reading the
              "常用快捷鍵" cluster would see undo and assume there's no
              built-in redo (the AI's branching-timeline note at App.tsx:555-
              556 makes redo non-obvious anyway). Same R64 (Ctrl+Tab) / R68
              (Ctrl+F) pattern: shortcut wired, sibling advertised, footer
              omits the pair member. Placement immediately after Ctrl+Z keeps
              the undo/redo cluster visually paired the same way the toolbar
              keeps Undo2 next to Redo2 (App.tsx:1224-1241). */}
          {/* Ctrl+E — 匯出目前頁籤, the last toolbar-button shortcut still
              missing from the fresh-user discovery cluster. Wired in
              menu.ts:105-108 (CmdOrCtrl+E → 'menu:exportTab') and dispatched
              in App.tsx:466-509 / 593 (handleExportTab → workspace.exportTab
              IPC). The toolbar Download button at App.tsx:1201-1209 already
              advertises the keystroke via title=`匯出此頁籤為 ${exportExt}
              (Ctrl+E)` (App.tsx:1126) — exactly mirroring how the Save
              button next to it carries `儲存 (Ctrl+S)` (App.tsx:1163), the
              Settings button comment at App.tsx:1287-1289 even explicitly
              enumerates Ctrl+E in the same peer-cluster as the other
              footer-listed toolbar shortcuts: "新專案 Ctrl+N, 開啟 Ctrl+O,
              儲存 Ctrl+S, 匯出 Ctrl+E, 復原 Ctrl+Z, 專注模式 Ctrl+Shift+F …".
              Yet the welcome footer's "常用快捷鍵" cluster listed Ctrl+S
              and Ctrl+Shift+S (R76) but skipped 匯出 — a fresh user looking
              for "how do I get my .md / .docx / .xlsx / .pptx out of this
              .gd archive?" would see save and save-as and assume export
              hides under another menu somewhere (it's File → 匯出目前頁籤
              in the native menu, but that menu bar is hidden by default on
              Windows unless Alt is held — same rationale that drove R76).
              Same R64 (Ctrl+Tab) / R68 (Ctrl+F) / R71 (Ctrl+Shift+Z) /
              R76 (Ctrl+Shift+S) pattern: shortcut wired, sibling toolbar
              button advertises, welcome footer omits the pair member.
              Placement immediately after Ctrl+Shift+S keeps the
              persist-bytes-to-disk cluster (儲存 / 另存新檔 / 匯出)
              visually grouped before the search/edit cluster (Ctrl+F /
              Ctrl+Z / Ctrl+Shift+Z) — same order as the toolbar itself
              (App.tsx:1188-1209: Save → Save Spinner state → Export Download
              icon → separator → Undo / Redo). Label uses the menu's
              canonical "匯出目前頁籤" wording (menu.ts:105) rather than
              the toolbar's "匯出此頁籤" — the menu name is the one stable
              across all four format types (the toolbar tooltip splices in
              the active tab's `${exportExt}` so it can't be reused as a
              footer label without losing precision). */}
          {/* Ctrl+Shift+S — 另存新檔, the orphaned twin of Ctrl+S above.
              Wired in menu.ts:99-101 (CmdOrCtrl+Shift+S → 'menu:saveAs') and
              dispatched in App.tsx:692-694 (handleSaveAs → workspace.saveAs
              IPC at ipc.ts:89). The toolbar Save button at App.tsx:1162-1163
              already pairs it with the primary 儲存 hint via the
              `saveAsHint = ' · 另存新檔 (Ctrl+Shift+S)'` middle-dot suffix —
              exactly mirroring how the toolbar's Undo2/Redo2 buttons keep
              Ctrl+Z and Ctrl+Shift+Z visually paired (App.tsx:1216 / 1225).
              Yet the welcome footer here listed Ctrl+S without its 另存新檔
              twin, the same gap the Ctrl+Z → Ctrl+Shift+Z addition above
              just closed: a fresh user reading the "常用快捷鍵" cluster
              would see save and assume there's no built-in save-as (the
              native File menu lists it but Windows hides the menu bar by
              default unless Alt is held — the comment at App.tsx:1149-1150
              explicitly calls this out as the reason the toolbar tooltip
              had to carry the hint). Same R64 (Ctrl+Tab) / R68 (Ctrl+F) /
              R71 (Ctrl+Shift+Z) pattern: shortcut wired, sibling advertises,
              welcome footer omits the pair member. Placement immediately
              after Ctrl+S keeps the save/save-as cluster visually paired the
              same way the undo/redo cluster sits above (line 321-322). */}
          {/* R131 — completes R130's `tab` → 頁籤 sweep across the
              renderer. R130 caught TabBar.tsx:572 (the new-tab tooltip)
              and PptxEditor.tsx:807/810 (the empty-state strings) but
              missed five more user-visible sites that the same regex
              search would have found:
                StatusBar.tsx:118       「{tabs.length} 個 tab」
                DiffPreview.tsx:138     「新增 tab：${op.tab.name}」
                DiffPreview.tsx:140     「刪除 tab：${op.tab.name}」
                EditorSurface.tsx:216   「建立新 tab」(welcome heading)
                EditorSurface.tsx:484   「切換 tab」(this Shortcut row)
                EditorSurface.tsx:485   「切換上/下個 tab」(this row +1)
              The cluster below is the loudest of all because it sits in
              the *same row* as `匯出目前頁籤` (line 428) — sighted users
              can read 頁籤 and `tab` side-by-side as they scan the
              shortcut list. (AI dispatcher / tools strings in
              ai/dispatcher.ts and ai/tools/index.ts are LLM-facing JSON
              schemas, deliberately kept English; the embedded comment
              quoting the old string above at line 329 is updated in
              parallel — same self-reference paradox R130 already noted
              between TabBar.tsx:534's quoted exemplar and its real
              button text.) */}
          {/* R149 — Ctrl+O / Ctrl+Shift+F: the welcome footer's "常用快捷鍵"
              cluster missed two more toolbar-button shortcuts. R107's Ctrl+E
              addition (line 371) signed off as "the last toolbar-button
              shortcut still missing from the fresh-user discovery cluster",
              yet the same comment at line 381-382 enumerates the canonical
              peer cluster as "新專案 Ctrl+N, 開啟 Ctrl+O, 儲存 Ctrl+S,
              匯出 Ctrl+E, 復原 Ctrl+Z, 專注模式 Ctrl+Shift+F …" — naming TWO
              siblings (Ctrl+O and Ctrl+Shift+F) that were ALSO absent here.
              The "last" claim self-contradicts the same comment's own peer-
              cluster enumeration. Both shortcuts are fully wired toolbar
              buttons: Ctrl+O at App.tsx:1192 with title="開啟 .gd (Ctrl+O)"
              (menu.ts:80-81 'menu:open' + keymap at App.tsx:711-713),
              Ctrl+Shift+F at App.tsx:1296 with title="進入專注模式 …
              (Ctrl+Shift+F)" (keymap at App.tsx:689-703 — no menu accelerator;
              the in-focus-mode exit FAB at App.tsx:1059 is the second in-app
              advertisement). Both are first-action onboarding gestures: a
              returning user with an existing .gd reaches for Ctrl+O before
              anything else, and Ctrl+Shift+F is the gateway to distraction-
              free reading. Same R64 (Ctrl+Tab) / R68 (Ctrl+F) / R71
              (Ctrl+Shift+Z) / R76 (Ctrl+Shift+S) / R107 (Ctrl+E) pattern:
              shortcut wired, sibling toolbar button advertises, welcome
              footer omits the pair member. Placement mirrors the toolbar
              order (App.tsx:1189-1196 / 1224-1241 / 1292-1299): Ctrl+O next
              to Ctrl+N (workspace-creation/loading cluster), Ctrl+Shift+F
              between Ctrl+Shift+Z and Ctrl+L (chrome-navigation cluster
              matching the peer list at line 316-322). Labels use toolbar
              tooltip wording verbatim ("開啟 .gd" / "進入專注模式") so the
              same Chinese phrase appears in tooltip and footer. */}
          <div className="flex flex-wrap gap-x-4 gap-y-1">
            <Shortcut keys="Ctrl/⌘+N" label="新專案" />
            <Shortcut keys="Ctrl/⌘+O" label="開啟 .gd" />
            <Shortcut keys="Ctrl/⌘+S" label="儲存" />
            <Shortcut keys="Ctrl/⌘+Shift+S" label="另存新檔" />
            <Shortcut keys="Ctrl/⌘+E" label="匯出目前頁籤" />
            <Shortcut keys="Ctrl/⌘+F" label="尋找與取代" />
            <Shortcut keys="Ctrl/⌘+Z" label="復原 AI 變更" />
            <Shortcut keys="Ctrl/⌘+Shift+Z" label="重做 AI 變更" />
            <Shortcut keys="Ctrl/⌘+Shift+F" label="進入專注模式" />
            <Shortcut keys="Ctrl/⌘+L" label="聚焦 AI 對話框" />
            <Shortcut keys="Ctrl/⌘+1–9" label="切換頁籤" />
            <Shortcut keys="Ctrl/⌘+Tab" label="切換上/下個頁籤" />
          </div>
        </section>
      </div>
    </div>
  );
}

function Shortcut({ keys, label }: { keys: string; label: string }) {
  return (
    <span className="inline-flex items-center gap-1.5">
      <kbd className="px-1.5 py-0.5 rounded bg-secondary border text-[10px] font-mono">{keys}</kbd>
      <span>{label}</span>
    </span>
  );
}
