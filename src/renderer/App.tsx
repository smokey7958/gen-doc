/**
 * Top-level renderer shell — left/right split, menu commands wiring,
 * apply-on-open of saved config to Zustand stores.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  AlertCircle,
  Check,
  Download,
  FilePlus2,
  FileDown,
  FocusIcon,
  FolderOpen,
  FolderInput,
  FolderTree,
  Loader2,
  Minimize2,
  Redo2,
  Save,
  Settings as SettingsIcon,
  Undo2,
  Upload,
} from 'lucide-react';
import { EditorSurface } from '../components/EditorSurface';
import { AIPanel } from '../components/AIPanel';
import { BatchExportDialog, exportExtForTab } from '../components/BatchExportDialog';
import { FileExplorer } from '../components/FileExplorer';
import { SettingsDialog } from '../components/SettingsDialog';
import { StatusBar } from '../components/StatusBar';
import { Toaster } from '../components/Toaster';
import { Button } from '../components/ui/button';
import { useAI } from '../store/ai';
import { resolveSupportedModelId } from '../types/ai';
import { useWorkspace, workspaceIdFor } from '../store/workspace';
import { notify } from '../store/toast';
import { applyChangeset, undoChangeset } from '../ai/changeset-apply';
import { flushEditors } from '../lib/editor-flush';
import { bumpLoadGen, currentLoadGen } from '../lib/workspace-load-gen';
import { isFileMissingError } from '../lib/utils';
import { tryEnterOpen, exitOpen } from '../lib/workspace-open-busy';
import { exitClose, tryEnterClose } from '../lib/tab-close-busy';
import { exitExportTab, isExportTabBusy, tryEnterExportTab } from '../lib/export-tab-busy';
import { deserializeChangeset, serializeChangeset } from '../lib/changeset-serialize';

/**
 * Lazily read a number / bool from localStorage with bounds. We use
 * localStorage rather than the main-process config file because these are
 * pure UI prefs — survival across restarts is nice-to-have, not load-bearing,
 * and pushing every drag-resize through IPC would be wasteful.
 */
function readStoredNumber(key: string, fallback: number, min: number, max: number): number {
  const raw = localStorage.getItem(key);
  if (raw === null) return fallback;
  const n = Number(raw);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}
function readStoredBool(key: string, fallback: boolean): boolean {
  const raw = localStorage.getItem(key);
  if (raw === null) return fallback;
  return raw === '1';
}

/**
 * The app is pinned to a light/white-based UI by design — the theme
 * picker is gone and the `.dark` palette in index.css has been removed.
 * We still defensively strip any stale `dark` class that an older build
 * (or a third-party extension) may have left on documentElement, so the
 * page can never paint with the dark palette even via stray DOM.
 */
function ensureLightTheme(): void {
  document.documentElement.classList.remove('dark');
}

/**
 * Drop a path from the persisted recent-files list — used when an open
 * attempt fails because the file has moved or been deleted, so the menu
 * stops offering a dead entry. Best-effort; swallow IPC errors.
 */
async function pruneRecent(filePath: string): Promise<void> {
  try {
    const cfg = await window.gendoc.config.get();
    const next = cfg.recentFiles.filter((p) => p !== filePath);
    if (next.length === cfg.recentFiles.length) return;
    await window.gendoc.config.set({ recentFiles: next });
  } catch {
    /* ignore — the next launch will retry */
  }
}

// R395 — `isFileMissingError` hoisted to lib/utils.ts so EditorSurface's
// onOpenRecent can share the same recent-list eviction gate. See doc-block
// at the helper's definition for the asymmetry rationale.

export function App(): JSX.Element {
  const [hasApiKey, setHasApiKey] = useState(false);
  const [aiPanelWidth, setAiPanelWidth] = useState(() =>
    readStoredNumber('gendoc.aiPanelWidth', 380, 280, 600),
  );
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [batchExportOpen, setBatchExportOpen] = useState(false);
  // R320 — single guarded open path for BatchExportDialog. Both the toolbar
  // FileDown button and the `menu:batchExport` (Ctrl+Shift+E) accelerator
  // route through here so neither can reopen the dialog while an export
  // is already in flight (OS folder picker open, atomicWrites in
  // progress). Without the guard, opening the dialog underneath the OS
  // folder picker stacks two modals; the user's click on the freshly-
  // reopened dialog's confirm button bounces silently off
  // tryEnterExportTab — looks like "the button doesn't work" with no
  // feedback. The notify toast tells the user the system is busy
  // exporting, not broken. Same shape as R209 / R220's "user gesture
  // during in-flight" guard family.
  const openBatchExportDialog = useCallback(() => {
    if (isExportTabBusy()) {
      notify('正在匯出中…請等目前的匯出完成再開啟', 'info');
      return;
    }
    setBatchExportOpen(true);
  }, []);
  // File-explorer pane state. `null` rootPath means "no folder picked yet"
  // (empty-state UI shown). `explorerOpen=false` collapses the pane entirely
  // — width still tracked separately so re-opening restores the user's size.
  const [explorerOpen, setExplorerOpen] = useState(() =>
    readStoredBool('gendoc.explorerOpen', false),
  );
  // Focus / distraction-free mode — hides Toolbar, FileExplorer, AIPanel, and
  // StatusBar so the editor surface fills the whole window. Adobe InDesign's
  // Presentation Mode (W) and Word's Focus Mode (View → Focus) both lean on
  // this for "review the whole document without UI noise" workflows; the user
  // explicitly called this out as Word's most-needed addition. We persist the
  // flag because some users prefer to live in focus mode (especially on
  // smaller laptop screens) and re-toggling on every launch is friction.
  const [focusMode, setFocusMode] = useState(() =>
    readStoredBool('gendoc.focusMode', false),
  );
  const [explorerWidth, setExplorerWidth] = useState(() =>
    readStoredNumber('gendoc.explorerWidth', 260, 180, 560),
  );
  const [explorerRoot, setExplorerRoot] = useState<string | null>(null);
  // OS-level file drag-drop overlay state. We show a translucent prompt the
  // moment the cursor enters the window with files attached, and clear it on
  // drop or when the cursor leaves the window entirely.
  const [dropActive, setDropActive] = useState(false);
  // Auto-save interval pulled from user config. 0 disables. We mirror it into
  // local state so changes from SettingsDialog (which patches main-side config)
  // can be picked up at runtime via a custom event without a full reload.
  const [autoSaveMs, setAutoSaveMs] = useState(0);
  // R187 — load generation counter moved to a module
  // (`lib/workspace-load-gen.ts`) so FileExplorer's own .gd-open path can
  // join the exclusion class. See doc-block there for the full R168 / R169
  // rationale (latest workspace-replacement wins). The ref-based counter
  // that previously lived here couldn't be shared across components without
  // prop drilling or context plumbing — module state is the cleaner shape
  // since the invariant is renderer-wide ("only one workspace is active").

  // Persist UI preferences across restarts. Drag-resize fires many state
  // updates; localStorage writes are synchronous but cheap, no debounce
  // needed at typical drag rates.
  useEffect(() => {
    localStorage.setItem('gendoc.aiPanelWidth', String(aiPanelWidth));
  }, [aiPanelWidth]);
  useEffect(() => {
    localStorage.setItem('gendoc.explorerWidth', String(explorerWidth));
  }, [explorerWidth]);
  useEffect(() => {
    localStorage.setItem('gendoc.explorerOpen', explorerOpen ? '1' : '0');
  }, [explorerOpen]);
  useEffect(() => {
    localStorage.setItem('gendoc.focusMode', focusMode ? '1' : '0');
  }, [focusMode]);
  const ws = useWorkspace();
  const setModel = useAI((s) => s.setModel);

  // Keep the OS window title in sync with the workspace. Prefer the on-disk
  // filename over manifest.title because the title field is rarely set on
  // legacy/auto-saved files — showing "未命名 — Gen Doc" for a clearly named
  // `2026-q1-report.gd` is uninformative when Alt+Tab'ing or picking from
  // the dock. Manifest title still wins when the user explicitly named it.
  useEffect(() => {
    const dirtyMark = ws.dirty ? '● ' : '';
    let label = ws.manifest.title;
    if ((!label || label === '未命名') && ws.filePath) {
      const base = ws.filePath.split(/[\\/]/).pop() ?? ws.filePath;
      label = base.replace(/\.gd$/i, '');
    }
    document.title = `${dirtyMark}${label || '未命名'} — Gen Doc`;
  }, [ws.manifest.title, ws.filePath, ws.dirty]);

  // Mirror dirty state to the main process so the OS close button can prompt
  // before destroying unsaved work. Fire-and-forget; preload uses ipcRenderer.send.
  useEffect(() => {
    window.gendoc.app.setDirty(ws.dirty);
  }, [ws.dirty]);

  // First-load: hydrate config + check API key.
  useEffect(() => {
    void (async () => {
      // R285 — wrap the boot-time IPC pair. Same sibling as R273
      // (SettingsDialog initial-load): config.get's loadConfig awaits
      // fs.mkdir + fs.readFile + fs.writeFile of DEFAULTS, all of which
      // can reject on first launch (disk full, EACCES on `~/.gendoc`,
      // antivirus pinning the config dir, network user-profile offline).
      // hasApiKey similarly walks safeStorage / keytar which can throw
      // on Linux without libsecret or transient OS-level keystore
      // anomalies. Without this guard the void IIFE rejects as
      // unhandledrejection AND every downstream setState (setModel /
      // setAutoSaveMs / setMaxTokens / setTemperature / setPromptCache /
      // setHasApiKey / ensureLightTheme / autoOpen) is skipped — app
      // boots silently in degraded state with hardcoded defaults and
      // hasApiKey=false even when the user actually has a key saved.
      // Notify so the user knows to investigate; the autoOpen path
      // inside has its own try/catch (line ~204) and remains unaffected.
      let cfg: Awaited<ReturnType<typeof window.gendoc.config.get>>;
      try {
        cfg = await window.gendoc.config.get();
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        notify(`載入設定失敗：${msg}（已使用預設值）`, 'error');
        return;
      }
      // R389 — coerce against current SUPPORTED_MODELS so a config.json
      // from an older app version (with a now-removed model id) doesn't
      // hydrate a stale id into the store. See resolveSupportedModelId
      // doc-block in types/ai.ts for the full failure trace.
      setModel(resolveSupportedModelId(cfg.defaultModel));
      setAutoSaveMs(cfg.autoSaveIntervalMs);
      // Mirror per-turn knobs into the AI store. SettingsDialog persists these
      // to disk but the runtime previously read them from hardcoded literals
      // in AIPanel/orchestrator (4096 / 0.3) — so saving 8192 in Settings
      // appeared to "stick" yet every send still capped at 4096. Hydrate here
      // on first load so the live store matches what the user picked.
      useAI.getState().setMaxTokens(cfg.maxTokens);
      useAI.getState().setTemperature(cfg.temperature);
      // Same mirror-into-store pattern for promptCache. Without this, the
      // SettingsDialog checkbox was a complete no-op — see store/ai.ts
      // promptCache doc-comment + main/ai/anthropic.ts withCache short-circuit
      // for the full chain. Hydrating here on first load means the first chat
      // turn after launch already honors the user's choice.
      useAI.getState().setPromptCache(cfg.promptCache);
      // R285 — separate try for hasApiKey. The config-mirror setStates above
      // are already committed by now; if the keystore probe itself throws
      // (libsecret missing on Linux, transient keytar / DPAPI anomaly),
      // hasApiKey stays at its safe-default `false` — which is the right
      // fallback (better to prompt the user for a key than to falsely
      // assume one exists). Notify so the cause is visible.
      try {
        const has = await window.gendoc.config.hasApiKey();
        setHasApiKey(has);
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        notify(`讀取 API key 狀態失敗：${msg}`, 'warning');
      }
      // App is pinned to a light/white-based UI; defensively strip any
      // stale `dark` class an older build may have left on documentElement.
      ensureLightTheme();
      // Auto-reopen the most recent workspace. We do this *after* state is
      // hydrated and only when nothing's loaded yet (workspace is already
      // dirty/has a path → user navigated here intentionally, don't clobber).
      // Failure (file moved/deleted) is silent: we just stay on the empty
      // workspace and the stale entry will scroll off the recent list as the
      // user opens other things.
      if (cfg.autoOpenLastWorkspace && cfg.recentFiles.length > 0) {
        const state = useWorkspace.getState();
        if (!state.filePath && !state.dirty) {
          const target = cfg.recentFiles[0];
          const myGen = bumpLoadGen();
          try {
            const opened = await window.gendoc.workspace.openPath(target);
            // R168 — drop on the floor if a newer load (handleOpen /
            // handleOpenRecent) started while we were awaiting. See
            // workspace-load-gen.ts doc-block at the top of App for full rationale.
            if (myGen !== currentLoadGen()) return;
            if (opened) useWorkspace.getState().loadFromOpened(opened);
          } catch (err) {
            // R392 — only prune when the failure is "file is gone from disk".
            // Other failure classes (corrupted manifest, EACCES from
            // transient antivirus lock, IPC bridge throw) leave the .gd on
            // disk; permanently dropping the recent entry for those would
            // erase user state for a file that's still recoverable. See
            // isFileMissingError doc-block for the full asymmetry rationale.
            if (isFileMissingError(err)) void pruneRecent(target);
          }
        }
      }
    })();
  }, [setModel]);

  // Pick up live config changes (SettingsDialog dispatches this on patch) so
  // auto-save / default-model respond without an app restart.
  useEffect(() => {
    const onChange = async () => {
      const cfg = await window.gendoc.config.get();
      setAutoSaveMs(cfg.autoSaveIntervalMs);
      // Mirror Settings → 預設模型 into the active AI session. Without this,
      // changing the default model in the dialog appears to do nothing — the
      // AIPanel dropdown stays on the old value and the next prompt still
      // hits the previous model. Users naturally read "this dialog is where
      // I switch models"; only honoring the change on next launch is a
      // surprising inconsistency. (If they want a per-session override, the
      // AIPanel selector still wins because they'd touch it *after* this.)
      const aiState = useAI.getState();
      // R389 — same SUPPORTED_MODELS coercion as the first-load hydration
      // above. SettingsDialog only lets the user pick from SUPPORTED_MODELS,
      // so the live patch here is almost always a valid id, but the dialog
      // sets value={config.defaultModel} unconditionally — when the
      // config-on-disk holds a stale id from an older version, the dialog
      // renders an inconsistent picker AND the patch path's `e.target.value`
      // can carry that stale id through if the user closes Settings without
      // touching the dropdown. Belt-and-braces here matches the boot path.
      const resolved = resolveSupportedModelId(cfg.defaultModel);
      if (aiState.model !== resolved) {
        aiState.setModel(resolved);
      }
      // Mirror Settings → maxTokens / temperature live too. Without this,
      // adjusting either slider in the dialog persists to disk but the
      // running session keeps using the previous value until app restart.
      if (aiState.maxTokens !== cfg.maxTokens) aiState.setMaxTokens(cfg.maxTokens);
      if (aiState.temperature !== cfg.temperature) aiState.setTemperature(cfg.temperature);
      // Mirror Settings → 啟用 prompt cache live so toggling the checkbox
      // takes effect on the very next prompt — same reasoning as the
      // maxTokens / temperature mirrors above. Without this branch, the
      // user would have to restart the app for the toggle to bite, and
      // until that restart Anthropic would keep applying / billing cache
      // writes against a session whose owner had explicitly opted out.
      if (aiState.promptCache !== cfg.promptCache) aiState.setPromptCache(cfg.promptCache);
    };
    window.addEventListener('gendoc:configChanged', onChange);
    return () => window.removeEventListener('gendoc:configChanged', onChange);
  }, []);

  // All three "discard unsaved work?" prompts route through the main-process
  // native dialog (`window.gendoc.app.confirm`) instead of renderer
  // `window.confirm()`. Reason: on Windows the Chromium-native confirm
  // returns control to the renderer with the BrowserWindow in a half-focused
  // state — `document.activeElement` is set, mouse clicks register, but
  // `document.hasFocus()` is false and OS keystrokes are not routed to
  // webContents. The symptom was that pressing Ctrl+N then opening a fresh
  // blank Word produced an editable surface that swallowed every keystroke
  // (no caret, no input). The native main-process dialog handles focus
  // restoration cleanly; the IPC handler also re-focuses the window +
  // webContents after dismissal. See `ipcMain.handle(IPC.app.confirm)`.
  const handleNew = useCallback(async () => {
    // R207 — extend R199's `tryEnterOpen` gate to handleNew. Same hazard
    // as the four entry points R199 already covered (handleOpen,
    // handleOpenRecent, FileExplorer useOpenFile, EmptyState's two open
    // buttons closed by R205): the dirty-state confirm calls
    // `window.gendoc.app.confirm` which routes through main's
    // `dialog.showMessageBoxSync`. Auto-repeat Ctrl+N (the OS firing
    // multiple keydowns while the key is held) or two rapid clicks on
    // 檔案 → 新專案 queues two confirm IPCs; main blocks on the first
    // showMessageBoxSync and the second waits in the IPC queue, so when
    // the user dismisses the first dialog the second one immediately
    // pops on top — looks like "the dialog refused to close." The
    // workspace-open-busy doc-block already includes "creating a fresh
    // workspace" in the same exclusion class as load (per R169's reuse
    // of the same gen counter for both), but the sync-gate side of the
    // contract was omitted at this entry point. tryEnterOpen / exitOpen
    // are renderer-wide so this also blocks an in-flight Ctrl+O from
    // racing with Ctrl+N.
    if (!tryEnterOpen()) return;
    // R289 — add catch around app.confirm. Original try had only finally,
    // leaving confirm-reject (rare main IPC anomaly: main unresponsive,
    // bridge re-bind, channel timeout) to escape as unhandled rejection
    // through `void handleNew()` at the menu listener / Ctrl+N path.
    // Same R288 / EditorSurface flat-try idiom.
    try {
      // R175 — read live dirty (matches handleOpenRecent's pattern at line 305).
      // Closure capture would be stale within the React render commit window:
      // an auto-save that completes between this useCallback's render and the
      // user's click flips ws.dirty to false in the store, but `ws.dirty` in
      // the closure stays true until the next render. The user then sees the
      // 「尚未儲存」 confirm on a workspace that was actually just saved.
      if (
        useWorkspace.getState().dirty &&
        !(await window.gendoc.app.confirm('目前的變更尚未儲存，確定建立新專案？'))
      )
        return;
      // R169 — bump loadGen so any in-flight loadFromOpened (auto-open during
      // splash, slow handleOpenRecent, drag-drop) drops on the floor instead
      // of overwriting the fresh blank workspace. Without this, a user who
      // hits Ctrl+N while auto-open is still running on app start would see
      // the empty workspace flash, then watch it get clobbered by recent[0]
      // when auto-open resolves. The bump only affects post-IPC resolves —
      // sync `newWorkspace()` lands first, then any latent load checks its
      // gen against this incremented value, fails, and returns. Same gen
      // counter as R168 since "creating a fresh workspace" is in the same
      // exclusion class as "loading a workspace from disk" (both replace the
      // tabs/manifest wholesale).
      bumpLoadGen();
      useWorkspace.getState().newWorkspace();
    } catch (err) {
      notify(`建立新專案失敗：${(err as Error).message}`, 'error');
    } finally {
      exitOpen();
    }
  }, []);

  const handleOpen = useCallback(async () => {
    // R198 / R199 — sync gate first; see workspace-open-busy.ts doc-block.
    if (!tryEnterOpen()) return;
    // R289 — flatten outer/inner try into single try+catch+finally.
    // Original nested layout had catch only on the inner openPath; if
    // app.confirm rejected (rare main IPC anomaly) the throw escaped the
    // outer try unhandled. Mirrors EditorSurface.onOpenExisting and the
    // R288-fixed drag-drop path. Same notify message handles both confirm
    // and openPath failures.
    try {
      // R175 — same live-dirty read as handleNew above and handleOpenRecent
      // below. Without this, a stale closure can trigger the「尚未儲存」
      // confirm on a freshly auto-saved workspace.
      if (
        useWorkspace.getState().dirty &&
        !(await window.gendoc.app.confirm('目前的變更尚未儲存，確定開啟其他檔案？'))
      )
        return;
      // Mirror the error-handling contract from handleOpenRecent below: a
      // freshly-picked .gd can still fail (corrupt zip, EACCES on a locked
      // file, malformed manifest). Recents path notifies *and* prunes the
      // dead entry; pick-from-dialog path can't prune anything (the user
      // never adopted the path), so we just notify.
      const myGen = bumpLoadGen();
      const opened = await window.gendoc.workspace.open();
      // R168 — newer load (another open / openRecent / auto-open) wins.
      if (myGen !== currentLoadGen()) return;
      if (opened) useWorkspace.getState().loadFromOpened(opened);
    } catch (err) {
      notify(`開啟失敗：${(err as Error).message}`, 'error');
    } finally {
      exitOpen();
    }
  }, []);

  /** Open a specific .gd path from the File → 最近開啟 submenu. */
  const handleOpenRecent = useCallback(
    async (filePath: string) => {
      if (!filePath) return;
      // R198 / R199 — same sync gate as handleOpen. menu:openRecent can
      // fire from rapid clicks on different recent entries; without this,
      // two of them would each pop their own「尚未儲存」 confirm in
      // sequence. Module gate also pre-empts a FileExplorer .gd click
      // happening at the same moment.
      if (!tryEnterOpen()) return;
      // R289 — flatten outer/inner try; see handleOpen sibling doc-block.
      // The pruneRecent side effect for openPath failures still fires
      // inside the merged catch (it's a no-op if `filePath` isn't in
      // recents anyway, so it's safe to call even for confirm-throw cases).
      try {
        if (
          useWorkspace.getState().dirty &&
          !(await window.gendoc.app.confirm('目前的變更尚未儲存，確定開啟其他檔案？'))
        )
          return;
        const myGen = bumpLoadGen();
        const opened = await window.gendoc.workspace.openPath(filePath);
        // R168 — newer load wins. See workspace-load-gen.ts doc-block.
        if (myGen !== currentLoadGen()) return;
        if (opened) useWorkspace.getState().loadFromOpened(opened);
      } catch (err) {
        // The most likely failure is "file moved/deleted since it was added
        // to recents". Tell the user; prune the dead entry ONLY when the
        // file is actually gone (ENOENT) — leave the entry alone for
        // recoverable failures (corrupted manifest the user may want to
        // re-attempt after fixing externally, EACCES from a transient AV
        // lock, app.confirm IPC reject before the user even saw a dialog).
        // R392 — see isFileMissingError doc-block for the
        //「erring toward false negatives is correct because the asymmetry
        // (false-positive prune erases user state forever vs. false-
        // negative prune re-prompts next session)」 rationale.
        notify(`開啟失敗：${(err as Error).message}`, 'error');
        if (isFileMissingError(err)) void pruneRecent(filePath);
      } finally {
        // R198 / R199 — release sync gate.
        exitOpen();
      }
    },
    [],
  );

  const handleClearRecent = useCallback(async () => {
    // R249 — try/catch with notify so a config.set failure (EACCES on
    // `~/.gendoc/config.json`, ENOSPC, network-share user-profile gone
    // offline, antivirus pinning the config dir) doesn't silently swallow.
    // The user explicitly clicked 清除最近開啟 from the File →
    // 最近開啟 submenu — they expect either the recents to clear OR a
    // visible error. Without this, the rejection escapes through the
    // menu listener's `void handleClearRecent()` (App.tsx:887) into
    // unhandledrejection territory, the dispatchEvent never fires (so
    // the submenu doesn't rebuild from main), and the user clicks the
    // 清除最近開啟 menu item again, watches nothing happen, and assumes
    // the menu is broken. Same shape as SettingsDialog's `patch` helper
    // (SettingsDialog.tsx:138-160) which already routes config.set
    // failures through `notify` with a 「儲存設定失敗：…」 prefix; we
    // mirror its 「動作 + 失敗」 pattern here. Verb 「清除最近開啟」
    // matches the menu label verbatim so the toast reads as a direct
    // response to the click.
    try {
      await window.gendoc.config.set({ recentFiles: [] });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`清除最近開啟失敗：${msg}`, 'error');
      return;
    }
    // Menu rebuild happens main-side on open/save; force a refresh now so
    // the submenu visibly empties without waiting for the next round-trip.
    window.dispatchEvent(new CustomEvent('gendoc:configChanged'));
  }, []);

  /**
   * Reentrancy guard for save: while a save is in flight, ignore further
   * Ctrl+S presses / button clicks. The save IPC is fast in the happy path
   * but can stall on slow disks; without this guard a second save races
   * and can clobber lastSavedAt back to a stale value mid-flight.
   */
  const savingRef = useRef(false);

  // Tracks the pending "已儲存 → idle" fade timer so consecutive saves don't
  // step on each other. Without this ref, save A's 2-second fade timer keeps
  // running while save B fires; when timer A wakes it sees state==='success'
  // (set by save B) and flips it to idle prematurely — save B's badge then
  // visibly blinks for whatever fraction of 2s was left on save A's clock.
  const saveFadeTimerRef = useRef<number | null>(null);
  // R197 / R225 — handleExportTab's sync gate moved to lib/export-tab-busy.ts
  // so TabBar's per-tab download icon (`exportSingleTab`) shares the same
  // single-OS-dialog invariant. Was a component-local useRef before R225.
  // R200 — sync gate for handleOpenFolder, see doc-block at site.
  const folderPickInFlightRef = useRef(false);
  // R198 / R199 — sync gate for handleOpen / handleOpenRecent moved to a
  // shared module (`lib/workspace-open-busy.ts`) so FileExplorer's own
  // .gd-open path can claim the same gate. See doc-block there for the
  // full rationale (one user, one dialog stack, complementary to the
  // post-IPC R168 loadGen).

  const performSave = useCallback(
    async (mode: 'save' | 'saveAs') => {
      if (savingRef.current) return;
      savingRef.current = true;
      const ws = useWorkspace.getState();
      // R170 — capture workspaceId before any await so we can detect a
      // workspace swap that happens during the save IPC. Without this guard
      // the save IPC can be in flight (typically 50-300ms but seconds on a
      // slow disk / large .gd), the user fires Ctrl+O / Ctrl+N during the
      // wait, the workspace store gets replaced wholesale by loadFromOpened
      // / newWorkspace, and our post-await `setState({ filePath: res.filePath,
      // dirty: anyDirty, tabs: reconciled })` writes the OLD workspace's save
      // result onto the NEW workspace — most damagingly, `filePath` points
      // at the old .gd so Ctrl+S in the new workspace would silently
      // overwrite the wrong file. The `tabs: reconciled` write is also
      // wrong (snapshot tab ids don't match new workspace's tabs), and
      // `setSaveState('success')` lies about the new workspace having been
      // saved. Capturing workspaceId here and re-reading it after each
      // await lets us drop the post-IPC writes if the workspace identity
      // changed — the disk write itself still happens for the OLD .gd
      // (correct, the user did ask to save it), only the in-memory state
      // updates targeting the NEW workspace are skipped.
      const savedWorkspaceId = ws.workspaceId;
      ws.setSaveState('saving');
      // Commit any "blur-to-write" input that was still focused when the
      // user fired Ctrl+S — currently just the tab-rename input (TabBar).
      // Without this, hitting Ctrl+S mid-rename saves the OLD name: the
      // rename input is uncontrolled (defaultValue) and only calls
      // renameTab in onBlur, so the new name lives nowhere visible to
      // serializeForSave below until something defocuses the field.
      // Querying by data-attribute (rather than blurring whatever happens
      // to be focused) keeps us from stealing focus from Find/Replace or
      // other plain INPUTs where blur has no useful effect.
      const pendingCommit = document.querySelector<HTMLElement>(
        '[data-commit-on-save="true"]:focus',
      );
      pendingCommit?.blur();
      // Force any debounced editor (pptx / docx use a 400ms re-serialize
      // window) to flush pending edits into `tab.data` *before* we
      // snapshot the workspace. Without this, Ctrl+S issued within the
      // debounce window saves stale bytes and the last burst of edits is
      // silently lost.
      await flushEditors();
      // R229 — bail if the workspace swapped during flushEditors. Without
      // this, a slow flushEditors await (debounced re-serialize on a big
      // pptx / docx — typically 100-400ms) plus a Ctrl+O during the await
      // produces a save CORRUPTION: line 460's serializeForSave reads the
      // LIVE store (now NEW workspace's tabs / manifest), but the IPC
      // call below uses `ws.filePath` (captured at line 421, OLD
      // workspace). Net result: NEW workspace's bytes written to OLD
      // workspace's `.gd` path — OLD's content silently overwritten with
      // an unrelated workspace's data. R170's post-IPC workspace check
      // catches the in-memory state pollution but does NOT catch the
      // disk-level cross-write because the IPC has already fired with
      // a mismatched filePath / payload pair. Bailing pre-IPC is the only
      // place to sever the bad pairing. We don't need to reset
      // `saveState` here — R185's loadFromOpened / newWorkspace already
      // wipes saveState back to 'idle' on swap, so NEW's StatusBar isn't
      // stuck on「儲存中…」.
      if (useWorkspace.getState().workspaceId !== savedWorkspaceId) {
        savingRef.current = false;
        return;
      }
      const { manifest, payloads } = useWorkspace.getState().serializeForSave();
      // Snapshot each tab's content/data reference at serialize time so we
      // can tell, after the IPC round-trip resolves, whether the user typed
      // *during* the save. patchTab always produces fresh objects, so a
      // simple === compares "did this tab's bytes change since we sent
      // them to disk". Without this, fast typists lose keystrokes silently:
      // step (1) flush, (2) snapshot bytes, (3) send IPC (takes 50-300 ms
      // on a slow disk / large .gd), (4) user types more during the await,
      // (5) we blindly setState({dirty:false}) — UI shows clean, on-disk
      // bytes don't have the new keystrokes, the next close-prompt shrugs.
      const snapshot = new Map<string, string | Uint8Array>(
        useWorkspace.getState().tabs.map((t) =>
          t.type === 'markdown' || t.type === 'html' ? [t.id, t.content] : [t.id, t.data],
        ),
      );
      try {
        const res =
          mode === 'saveAs'
            ? await window.gendoc.workspace.saveAs({ manifest, tabs: payloads })
            : await window.gendoc.workspace.save({
                filePath: ws.filePath || undefined,
                manifest,
                tabs: payloads,
              });
        // R170 — workspace may have swapped during the save IPC. Skip every
        // post-await state mutation in that case so the new workspace's
        // identity isn't polluted with the old save's filePath / reconciled
        // tabs / saveState.
        if (useWorkspace.getState().workspaceId !== savedWorkspaceId) return;
        // Reconcile dirty against the in-flight snapshot. A tab whose
        // current content/data still === the snapshot was untouched during
        // the IPC and is genuinely clean. Any tab that diverged keeps its
        // dirty flag so the user's mid-save edits aren't lost. Workspace-
        // level dirty mirrors "any tab still dirty" — when every tab is
        // clean we can drop the workspace dot too. (New tabs created
        // during the save — implausible since flushEditors precedes any
        // user action that opens a tab — keep their dirty=true via the
        // `?? t.dirty` fallback.)
        const reconciled = useWorkspace.getState().tabs.map((t) => {
          const baseline = snapshot.get(t.id);
          if (baseline === undefined) return t; // wasn't in snapshot
          const current = t.type === 'markdown' || t.type === 'html' ? t.content : t.data;
          const stillDirty = current !== baseline;
          return { ...t, dirty: stillDirty };
        });
        const anyDirty = reconciled.some((t) => t.dirty);
        // R385 — recompute workspaceId from the new filePath. Save-As can
        // change filePath from `''` (newly created, was `u-N`) to a real
        // path, or from `/old/path.gd` to `/new/path.gd` (rename via
        // Save-As). The pre-save workspaceId no longer matches the new
        // file's identity (workspaceIdFor hashes the filePath at
        // workspace.ts:229). Without this update:
        //   • undo entries pushed AFTER Save-As (in this same session)
        //     remain tagged to the OLD workspaceId (since handleUndo /
        //     onApply read workspaceId from the store)
        //   • if the user later Ctrl+O reloads the saved file via the
        //     file explorer, loadFromOpened recomputes workspaceId =
        //     workspaceIdFor(filePath_of_new_file) = `f-<new-hash>`,
        //     which doesn't match the orphan undo rows still in
        //     SQLite under the OLD id
        //   • result: the undo button greys out on the post-Save-As
        //     session's edits despite those edits clearly being
        //     undoable — confusing because the user just made them
        // Saving via Ctrl+S (not Save-As) on an existing file path
        // also passes through this path with `res.filePath ===
        // currentFilePath`, so workspaceIdFor produces the same id,
        // and the setState is a no-op for workspaceId.
        const nextWorkspaceId = workspaceIdFor(res.filePath);
        // R386 — when Save-As moves us from an unsaved id (`u-N`) or an older
        // saved id (`f-<old-hash>`) to a new saved id (`f-<new-hash>`), the
        // SQLite undo_entries / conversations rows pushed before this save
        // are still tagged to savedWorkspaceId. Without relink, the post-Save
        // workspace queries under nextWorkspaceId, finds nothing, and:
        //   • App.tsx:970 refreshCanUndo flips canUndo=false → Undo button
        //     greys out even though the user clearly has an undoable stack
        //   • AIPanel's conversation list shows the workspace as fresh,
        //     hiding the long AI session the user just ran pre-Save-As
        // R385 fixed the in-memory mapping; this completes the move by
        // re-tagging the persistent rows. The IPC is awaited so the new id is
        // queryable BEFORE we set workspaceId on the store — otherwise the
        // useEffect at line 995 fires refreshCanUndo against the new id while
        // the rows are still under the old id (transient false). The relink
        // helper short-circuits when oldId === newId, so a plain Ctrl+S that
        // didn't change filePath doesn't churn the DB. Re-check the
        // workspace-swap guard after the await to stay consistent with the
        // R170 family — if the user fired Ctrl+O between the save and now,
        // we drop the post-IPC state updates onto the floor; the relink
        // itself is harmless either way (re-tagging in SQLite is correct for
        // both old and new workspaces, they just share no live state).
        if (savedWorkspaceId !== null && savedWorkspaceId !== nextWorkspaceId) {
          try {
            await window.gendoc.undo.relink(savedWorkspaceId, nextWorkspaceId);
          } catch (relinkErr) {
            // DB locked / disk full — log and continue. The state setState
            // below still commits the new id, so future pushes land in the
            // right place; only pre-save undo history is lost (same outcome
            // as without R386). A toast here would just confuse the user
            // mid-save.
            console.error('R386: relinkWorkspaceId failed', relinkErr);
          }
          if (useWorkspace.getState().workspaceId !== savedWorkspaceId) return;
        }
        useWorkspace.setState({
          filePath: res.filePath,
          workspaceId: nextWorkspaceId,
          dirty: anyDirty,
          tabs: reconciled,
        });
        useWorkspace.getState().setSaveState('success');
        // Auto-fade the "已儲存" badge after a couple of seconds so it
        // doesn't permanently camp in the toolbar — the lastSavedAt
        // timestamp is what feeds the persistent "X 秒前儲存" tooltip.
        // Cancel any prior fade timer so back-to-back saves each get the
        // full 2s of badge time, instead of save B inheriting save A's
        // remaining clock.
        if (saveFadeTimerRef.current !== null) {
          window.clearTimeout(saveFadeTimerRef.current);
        }
        saveFadeTimerRef.current = window.setTimeout(() => {
          saveFadeTimerRef.current = null;
          if (useWorkspace.getState().saveState === 'success') {
            useWorkspace.getState().setSaveState('idle');
          }
        }, 2000);
      } catch (e) {
        // R170 — same workspace-swap guard as the success path above. The
        // OLD workspace's save error shouldn't write 'error' / 'idle' onto
        // the NEW workspace's saveState; the OS dialog cancel toast is
        // also moot in the new context.
        if (useWorkspace.getState().workspaceId !== savedWorkspaceId) return;
        const msg = (e as Error).message;
        if (msg === 'save_cancelled') {
          // User dismissed the OS save dialog — back to idle, no error.
          useWorkspace.getState().setSaveState('idle');
        } else {
          useWorkspace.getState().setSaveState('error', msg);
          notify(`儲存失敗：${msg}`, 'error');
        }
      } finally {
        savingRef.current = false;
      }
    },
    [],
  );

  const handleSave = useCallback(() => performSave('save'), [performSave]);
  const handleSaveAs = useCallback(() => performSave('saveAs'), [performSave]);

  // Save-and-quit handshake. Main triggers this when the user picks "Save"
  // in the close prompt; we run our normal save flow and report back via
  // `saveAndQuitResult`. Reporting failure lets main re-arm the prompt
  // instead of tearing the window down with unsaved data — the most
  // important case is the user cancelling a Save-As dialog on an untitled
  // workspace.
  useEffect(() => {
    const onSaveAndQuit = async () => {
      await performSave('save');
      // After the await, performSave has settled state synchronously: dirty
      // flips false on success and stays true on cancel/error.
      const stillDirty = useWorkspace.getState().dirty;
      window.gendoc.app.saveAndQuitResult(!stillDirty);
    };
    const handler = () => void onSaveAndQuit();
    window.addEventListener('menu:saveAndQuit', handler);
    return () => window.removeEventListener('menu:saveAndQuit', handler);
  }, [performSave]);

  // Auto-save: true debounce that re-arms on every keystroke. Fires only
  // when the workspace already has a file path on disk — we never
  // auto-trigger a Save-As dialog because that's a deliberate user action.
  //
  // The dep on `ws.lastEditAt` is what makes this an actual debounce: every
  // patchTab / markTabDirty bumps that pulse, the effect re-runs, the
  // cleanup `clearTimeout`s the previous arming, and a fresh timer starts.
  // Without it the effect would only re-run when `ws.dirty` flipped (i.e.
  // once per save cycle), so the timer set after the FIRST keystroke would
  // win — popping a save flash mid-typing instead of after the burst.
  useEffect(() => {
    if (autoSaveMs <= 0) return;
    if (!ws.dirty) return;
    if (!ws.filePath) return;
    if (ws.saveState === 'saving') return;
    // Don't keep retrying after a save failure — the renderer's catch path
    // pushes a `notify("儲存失敗：…", 'error')` toast, and without this guard
    // auto-save would re-fire every autoSaveMs and re-toast the same error on
    // a perfectly valid failure mode (disk full, network volume offline,
    // EACCES on the .gd).
    // Manual Ctrl+S still works and will transition saveState back through
    // 'saving' → success/error, naturally re-evaluating this effect.
    if (ws.saveState === 'error') return;
    const t = window.setTimeout(() => {
      void handleSave();
    }, autoSaveMs);
    return () => window.clearTimeout(t);
  }, [autoSaveMs, ws.dirty, ws.filePath, ws.saveState, ws.lastEditAt, handleSave]);

  const handleExportTab = useCallback(async () => {
    // R197 / R225 — shared sync gate via lib/export-tab-busy.ts. R197
    // gated rapid Ctrl+E auto-repeat with a component-local
    // `exportInFlightRef`; R225 promoted the gate to a module so this
    // path AND TabBar's `exportSingleTab` (per-tab download icon) can
    // see each other. Cross-fire (toolbar Download click + per-tab
    // Download click in the same hover frame, or Ctrl+E + per-tab)
    // would otherwise queue two `workspace.exportTab` IPCs and stack
    // two OS save dialogs on the BrowserWindow — same dialog
    // stack-up R207 / R209 / R220 / R221 closed for other entry
    // classes. The OS save dialog is app-wide modal-ish (one save
    // can land at a time anyway), so a single global gate is the
    // correct granularity.
    if (!tryEnterExportTab()) {
      // R322 — surface busy state instead of silent no-op. Same toast shape
      // as R320's openBatchExportDialog: an export is genuinely in flight
      // (OS save / folder picker open in the background, or another
      // export's atomicWrite mid-stream); telling the user 「正在匯出中」
      // beats「按了沒反應」 the previous silent path produced. Toast.ts
      // already dedupes by message+variant, so a true rapid-double-fire
      // (sub-100ms) just refreshes one info toast rather than spamming.
      notify('正在匯出中…請等目前的匯出完成再試', 'info');
      return;
    }
    try {
    // Flush pending debounced serializes from pptx/docx/xlsx editors first —
    // otherwise Ctrl+E captures bytes up to 400 ms behind the latest edit.
    // Same contract as `performSave`; re-read the tab afterwards because
    // flushers patch `data` into the store.
    // R171 — capture workspaceId before any await so the post-IPC flash
    // doesn't bleed into a swapped workspace's StatusBar. Same shape as
    // R170 for performSave: the disk write itself is correct (the user's
    // Ctrl+E targeted the OLD active tab and the bytes are already
    // captured before the export IPC), but the「已匯出 foo.md」 flash at
    // the end is renderer state — it should only land on the workspace
    // the user was looking at when they exported.
    const exportWorkspaceId = useWorkspace.getState().workspaceId;
    await flushEditors();
    const state = useWorkspace.getState();
    if (state.workspaceId !== exportWorkspaceId) return;
    const tab = state.tabs.find((t) => t.id === state.activeTabId);
    if (!tab) {
      notify('沒有開啟的頁籤可以匯出', 'warning');
      return;
    }
    let bytes: Uint8Array;
    let ext: 'md' | 'html' | 'docx' | 'xlsx' | 'pptx';
    if (tab.type === 'markdown') {
      bytes = new TextEncoder().encode(tab.content);
      ext = 'md';
    } else if (tab.type === 'html') {
      bytes = new TextEncoder().encode(tab.content);
      ext = 'html';
    } else {
      // docx / xlsx / pptx tabs carry raw bytes. A brand-new (zero-byte) tab
      // hasn't been initialized yet — ask the user to open & edit it first
      // rather than writing a 0-byte file that Office can't open.
      if (tab.data.byteLength === 0) {
        notify('這個頁籤還是空的，請先在編輯器中編輯內容再匯出', 'warning');
        return;
      }
      bytes = tab.data;
      ext = tab.type;
    }
    try {
      const res = await window.gendoc.workspace.exportTab({
        ext,
        suggestedName: tab.name,
        bytes,
      });
      // R171 — same workspace guard. The OS save dialog can be open for
      // seconds; if the user picks a destination then immediately Ctrl+O's
      // a different .gd, our flashExport would otherwise paint「已匯出
      // foo.md」 on the new workspace's StatusBar. The toast (failure
      // path) targets the global Toaster (not workspace-scoped) so we
      // leave it firing regardless — failure feedback should always
      // surface, and 匯出失敗：… is unambiguous out of context.
      if (useWorkspace.getState().workspaceId !== exportWorkspaceId) return;
      // Surface a 5s "已匯出 foo.md" flash in the StatusBar — the OS save
      // dialog is gone by the time control returns, and the user otherwise
      // has no visible signal that the bytes actually landed on disk.
      const fileName = res.filePath.split(/[/\\]/).pop() ?? tab.name;
      useWorkspace.getState().flashExport(fileName, res.filePath);
    } catch (e) {
      const msg = (e as Error).message;
      if (msg !== 'export_cancelled') notify(`匯出失敗：${msg}`, 'error');
    }
    } finally {
      // R225 — release shared gate.
      exitExportTab();
    }
  }, []);

  /**
   * Batch export — accept the list of tab ids the BatchExportDialog
   * collected, flush debounced editors, build per-tab byte payloads, and
   * hand a single ExportTabsRequest to main. Main pops one folder picker
   * (vs. one Save-As per tab — would be unusable for 5+ files), writes
   * each tab into the chosen folder with collision-aware naming
   * (see invokeExportTabs in main/ipc.ts), and reports per-file success /
   * failure back.
   *
   * Same R197/R225 export-tab-busy gate as the single-tab path so a
   * Ctrl+E + batch click + per-tab × Download click all serialize through
   * one OS dialog at a time (stacking two OS dialogs on the same
   * BrowserWindow is broken at the system level).
   *
   * Workspace-swap guard mirrors handleExportTab — capture workspaceId
   * before flushEditors / IPC, drop StatusBar flash if the user has
   * since swapped (the disk write itself still completes on the bytes we
   * captured at click time; only the renderer-side feedback is workspace-
   * scoped).
   */
  const handleBatchExport = useCallback(async (selectedIds: string[]) => {
    if (selectedIds.length === 0) return;
    if (!tryEnterExportTab()) {
      // R322 — same busy-toast as handleExportTab. With R320 already
      // gating the dialog OPEN trigger on busy, reaching this branch is
      // rare (would require confirm to fire while busy — possible if the
      // dialog was opened before busy=true was claimed by a sibling
      // export). Defensive feedback keeps the silent-no-op out.
      notify('正在匯出中…請等目前的匯出完成再試', 'info');
      return;
    }
    try {
      // Close the BatchExportDialog BEFORE flushEditors / IPC so the OS
      // folder picker doesn't pop with our Radix overlay still dimming
      // the window underneath. The user already committed by clicking
      // 匯出.
      setBatchExportOpen(false);
      const exportWorkspaceId = useWorkspace.getState().workspaceId;
      await flushEditors();
      const state = useWorkspace.getState();
      if (state.workspaceId !== exportWorkspaceId) return;
      const selectedSet = new Set(selectedIds);
      // Preserve the workspace's tab order in the batch (not the order the
      // user happened to check the boxes), so the resulting filenames'
      // natural alphabetical/numerical sort matches the user's mental
      // model of "left-to-right in the tab strip".
      const tabsToExport = state.tabs.filter((t) => selectedSet.has(t.id));
      if (tabsToExport.length === 0) return;
      const payloads = tabsToExport
        .map((tab) => {
          if (tab.type === 'markdown' || tab.type === 'html') {
            return {
              ext: exportExtForTab(tab.type),
              suggestedName: tab.name,
              bytes: new TextEncoder().encode(tab.content),
            };
          }
          if (tab.data.byteLength === 0) {
            // Defensive — the dialog already greys these out. Skip silently
            // here so a rare race (user opened a blank tab between dialog
            // mount and confirm) doesn't drop a 0-byte file on disk.
            return null;
          }
          return {
            ext: exportExtForTab(tab.type),
            suggestedName: tab.name,
            bytes: tab.data,
          };
        })
        .filter((p): p is NonNullable<typeof p> => p !== null);
      if (payloads.length === 0) {
        notify('沒有可匯出的內容（所有勾選的頁籤都還沒有編輯）', 'warning');
        return;
      }
      try {
        const res = await window.gendoc.workspace.exportTabs({ tabs: payloads });
        // Workspace swap during the folder picker / writes — same guard
        // shape as handleExportTab. The disk writes already landed; we
        // just don't want to flash「已匯出 N 個檔案」 onto NEW
        // workspace's StatusBar.
        if (useWorkspace.getState().workspaceId !== exportWorkspaceId) return;
        if (!res) {
          // User cancelled the folder picker — silent no-op (same UX as
          // single-tab export_cancelled).
          return;
        }
        const successCount = res.filePaths.length;
        const failures = res.failures;
        if (failures.length === 0) {
          notify(`已匯出 ${successCount} 個檔案到資料夾`, 'success');
        } else {
          // R323 / R339 — shared dedupe + name-list formatting for both the
          // all-failed and partial-failed branches.
          //
          // R323 added error-message dedupe (`訊息 (× N)`) for the all-
          // failed case: permission / disk-full / antivirus-lock failures
          // hit every file with the same syscall error string; without
          // dedupe the toast was N copies of EACCES that the user had to
          // visually verify were all identical before acting.
          //
          // R339 extends the same dedupe to the partial-failed branch.
          // Previously the partial-failed toast was「匯出 3 個成功，2 個
          // 失敗（foo、bar）」— named WHICH files failed but never WHY.
          // The user had no way to learn the underlying error without
          // re-attempting the export, which often hit the same lock (AV
          // scan, OneDrive sync) and reproduced the same failure with
          // the same opaque toast. The all-failed branch already showed
          // errors; the partial-failed branch was the asymmetric outlier
          // — same data shape (failures[]), same need for the dedupe
          // treatment. The toast variant stays 'warning' for partial
          // (some progress was made) vs 'error' for all-failed (nothing
          // landed), so users can still distinguish severity at a
          // glance.
          const counts = new Map<string, number>();
          for (const f of failures) counts.set(f.error, (counts.get(f.error) ?? 0) + 1);
          const errorLines = Array.from(counts.entries())
            .map(([msg, n]) => (n > 1 ? `${msg} (× ${n})` : msg))
            .join('\n');
          const failedNames = failures
            .map((f) => payloads[f.index]?.suggestedName ?? '(未知)')
            .join('、');
          if (successCount === 0) {
            notify(
              `匯出全部失敗（${failedNames}）：\n${errorLines}`,
              'error',
            );
          } else {
            notify(
              `匯出 ${successCount} 個成功，${failures.length} 個失敗（${failedNames}）：\n${errorLines}`,
              'warning',
            );
          }
        }
      } catch (e) {
        const msg = (e as Error).message;
        notify(`批次匯出失敗：${msg}`, 'error');
      }
    } finally {
      exitExportTab();
    }
  }, []);

  // Re-sync `canUndo` from the persistent SQLite stack. Called on workspace
  // swap and after each undo/redo round-trip so the toolbar's Undo button
  // can disable itself when the stack is empty (mirroring Redo's affordance,
  // see Toolbar `canRedo` wiring). list(.., 1) is cheap — limit clamps it
  // to a single row read.
  const refreshCanUndo = useCallback(async () => {
    const wsId = useWorkspace.getState().workspaceId;
    try {
      const rows = await window.gendoc.undo.list(wsId, 1);
      // R167 — same workspace-swap-during-IPC hazard R166 closed for the
      // undo/redo apply paths and R160 closed for createConversation. The
      // useEffect at line 543-545 re-fires refreshCanUndo when workspaceId
      // changes, so a swap typically queues a fresh call for the new
      // workspace; the race is when the OLD list IPC resolves AFTER the
      // NEW one, so its setCanUndo(rows_old) silently overrides the
      // correct rows_new. Drop the stale write on the floor instead — the
      // useEffect already guarantees a fresh call for the current workspace
      // is in flight, so canUndo will settle to the correct value either
      // way.
      if (useWorkspace.getState().workspaceId !== wsId) return;
      useWorkspace.getState().setCanUndo(rows.length > 0);
    } catch {
      // Storage offline / pre-init — leave whatever value we had.
    }
  }, []);

  // Run once on first mount and any time the workspace identity changes
  // (open, new, switch). The `setCanUndo(false)` in load/new is intentionally
  // pessimistic; this effect re-arms the button if the new workspaceId has
  // existing entries from a prior session.
  useEffect(() => {
    void refreshCanUndo();
  }, [ws.workspaceId, refreshCanUndo]);

  const handleUndo = useCallback(async () => {
    // R166 — read workspaceId BEFORE the pop IPC, then re-read tabs AFTER it.
    // A rapid double-Ctrl+Z would otherwise capture wsState ONCE up-front and
    // both handlers would compute the inverse from the same stale tabs
    // snapshot: handler 1 pops cs_latest, undoes from tabs_v0, setStates;
    // handler 2 pops cs_latest-1 (each pop is its own SQLite DELETE so both
    // get a row), undoes from THE SAME tabs_v0 closure (it never saw
    // handler 1's setState), and the second setState overrides the first.
    // Net: cs_latest's reversal is silently dropped — the user sees only
    // one of two undos visibly applied even though both rows are off the
    // persistent stack and both are on aiRedo. Reading tabs *after* the
    // await closes this — handler 2's undo composes on top of handler 1's
    // already-committed setState.
    const workspaceId = useWorkspace.getState().workspaceId;
    // R271 — wrap the pop IPC. A DB failure (sqlite-wal lock during an OS
    // backup snap, antivirus pinning chat.sqlite, disk full) on the
    // user-triggered undo path would otherwise propagate as an
    // unhandledrejection through `void handleUndo()` at App.tsx:1227 with
    // zero visible cue — the user mashes Ctrl+Z, sees nothing happen, and
    // assumes the shortcut is broken. R245's `.catch(() => undefined)` on
    // the swap-detected re-push (line ~804) is fire-and-forget by design;
    // this entry-point IPC is the opposite — its result IS the user's
    // intent, so notify is the right surface (matching R249's
    // handleClearRecent pattern). No state has been touched yet at this
    // point so a bare `return` after the toast is sufficient — workspace
    // tabs unmodified, aiRedoStack unmodified, persistent stack state
    // ambiguous (the DELETE may or may not have committed) but
    // refreshCanUndo on the next render will reconcile.
    let row: Awaited<ReturnType<typeof window.gendoc.undo.pop>>;
    try {
      row = await window.gendoc.undo.pop(workspaceId);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`復原失敗：${msg}`, 'error');
      void refreshCanUndo();
      return;
    }
    if (!row) return;
    // R216 — Uint8Array-aware deserialization. Naive JSON.parse would
    // turn binary_replace ops' before/after into plain `{0:N,1:N,...}`
    // objects (the symmetric corruption to JSON.stringify's mangle of
    // Uint8Array), and `undoChangeset` would then assign that plain
    // object to `tab.data`, leaving the xlsx/docx/pptx in a state the
    // editors can't parse. See lib/changeset-serialize.ts doc-block.
    // R394 — wrap deserialize in try/catch. `deserializeChangeset` is a
    // thin wrapper around `JSON.parse` (changeset-serialize.ts:71); if the
    // SQLite row's changeset_json is malformed (cosmic-bit-flip on disk,
    // manual edit of chat.sqlite, schema drift from a future version
    // emitting an incompatible op shape), JSON.parse throws SyntaxError.
    // Without this guard the throw escapes through the `void handleUndo()`
    // callsite at App.tsx:1668 (Ctrl+Z keymap) / 1357 (menu:undo) as an
    // unhandledrejection — DevTools shows a stack trace but the user just
    // sees「Ctrl+Z 怎麼按都沒反應」 with no toast, and the row's already
    // gone from SQLite (the pop IPC succeeded line 1128). Mashing Ctrl+Z
    // would loop on each subsequent row until they hit a clean one — a
    // long silent stall.
    //
    // Surface via toast (same channel as R271's pop-IPC catch above) and
    // refreshCanUndo so the toolbar's Undo button state reflects the
    // post-pop persistent stack. We intentionally DO NOT re-push the
    // corrupted row: re-pushing would just hit the same parse failure on
    // the next Ctrl+Z, creating an infinite「按一次就 toast」 loop. The
    // pop is a permanent step (matches R271's bare-`return` posture after
    // pop-IPC failure for the same reason — once the row is gone from
    // SQLite, the cleanest recovery is to let it stay gone).
    let cs;
    try {
      cs = deserializeChangeset(row.changesetJson);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`復原失敗：歷史紀錄損毀（${msg}）`, 'error');
      void refreshCanUndo();
      return;
    }
    // R173 — workspace-swap-during-pop guard. The pop already drained one
    // row from the OLD workspace's persistent undo stack; if we're now in
    // a different workspace, applying that cs to the NEW tabs is wrong on
    // two axes: (a) the cs's tab ids don't exist in NEW tabs so
    // undoChangeset is a no-op visually but still setStates dirty=true,
    // (b) `pushAiRedo(cs)` polluts NEW's redo stack with an unrelated
    // changeset that, if redone, would silently mark NEW as dirty without
    // any visible effect. Push the row back to the OLD workspace's
    // persistent stack so when the user swaps back, the undo entry is
    // still there — losing user-visible undo history is the worst
    // failure mode here. Best-effort fire-and-forget; if the re-push
    // fails the entry is lost but app state stays consistent.
    if (useWorkspace.getState().workspaceId !== workspaceId) {
      // R245 — explicit `.catch(() => undefined)` swallow so a re-push DB
      // failure (sqlite-wal lock, disk full, IPC reject) doesn't surface
      // as unhandledrejection. The doc-block above already commits to
      // "best-effort fire-and-forget; if the re-push fails the entry is
      // lost but app state stays consistent", but the bare `void`
      // didn't actually silence the rejection — Electron / Node would
      // still log a warning to DevTools (and on stricter handler
      // configurations could crash the renderer process). Aligns with
      // the R241 / R242 sweep that established `.catch(() => undefined)`
      // as the canonical "I really mean fire-and-forget" pattern for
      // post-success housekeeping IPCs (pushRecent in 4 IPC handlers,
      // rebuildAppMenu in IPC.config.set). Same shape, same intent —
      // closing the inconsistency at the lone holdout in the renderer.
      void window.gendoc.undo
        .push({
          changesetJson: row.changesetJson,
          workspaceId,
        })
        .catch(() => undefined);
      return;
    }
    const wsState = useWorkspace.getState();
    const { tabs } = undoChangeset(wsState.tabs, cs);
    // R250 — see AIPanel.onApply doc-block. Undo of a tab_create
    // op REMOVES the tab; if that tab was active, activeTabId is
    // now stale.
    const liveTabsBefore = wsState.tabs;
    useWorkspace.setState((s) => {
      const stillValid = tabs.some((t) => t.id === s.activeTabId);
      // R276 — mirror R202's selection invariant for the changeset-apply
      // path. Undo of a tab_create op removes the tab; if that tab held
      // the selection, `s.selection.tabId` is now orphan. Reconcile in
      // the same atomic setState as activeTabId. See AIPanel.onApply
      // R276 doc-block for the full invariant rationale.
      const selStillValid =
        !s.selection || tabs.some((t) => t.id === s.selection!.tabId);
      const selPatch = selStillValid ? null : { selection: null };
      if (stillValid) return { tabs, dirty: true, ...selPatch };
      const removedIdx = liveTabsBefore.findIndex((t) => t.id === s.activeTabId);
      const fallback =
        removedIdx >= 0 && removedIdx < tabs.length
          ? tabs[removedIdx].id
          : tabs[tabs.length - 1]?.id ?? null;
      return { tabs, activeTabId: fallback, dirty: true, ...selPatch };
    });
    // Push to the session redo stack so Ctrl+Shift+Z (or 編輯 → 重做) can
    // re-apply. Persistent redo isn't worth the complexity — branching
    // history across sessions makes the semantics murky.
    wsState.pushAiRedo(cs);
    // The pop may have just emptied the stack, or there may still be more.
    // Re-query so the Undo button's disabled state stays honest.
    void refreshCanUndo();
  }, [refreshCanUndo]);

  const handleRedo = useCallback(async () => {
    // R166 — same hazard shape as handleUndo above. `popAiRedo()` is sync
    // (Zustand setState), so the actual race window is between this sync
    // pop+apply+setState and the later `await window.gendoc.undo.push`:
    // a second Ctrl+Shift+Z firing before push resolves would popAiRedo
    // its own cs, applyChangeset onto the freshly-read live tabs (good),
    // and push its own row — but BOTH push IPCs land on the same
    // persistent stack. That's actually fine for the persistent stack
    // (each row is independent), so the race here is narrower than undo.
    // We still re-read live state on each entry so cross-handler tab
    // composition stays consistent.
    const cs = useWorkspace.getState().popAiRedo();
    if (!cs) return;
    const wsState = useWorkspace.getState();
    const redoWorkspaceId = wsState.workspaceId;
    const { tabs } = applyChangeset(wsState.tabs, cs);
    // R250 — see AIPanel.onApply doc-block. Redo of a tab_delete
    // op re-removes the tab; if that tab was active, activeTabId
    // is stale.
    const liveTabsBeforeRedo = wsState.tabs;
    useWorkspace.setState((s) => {
      const stillValid = tabs.some((t) => t.id === s.activeTabId);
      // R276 — mirror R202's selection invariant. Redo of a tab_delete
      // op re-removes the tab; if that tab held the selection, orphan.
      // Same shape as handleUndo / AIPanel.onApply siblings this round.
      const selStillValid =
        !s.selection || tabs.some((t) => t.id === s.selection!.tabId);
      const selPatch = selStillValid ? null : { selection: null };
      if (stillValid) return { tabs, dirty: true, ...selPatch };
      const removedIdx = liveTabsBeforeRedo.findIndex((t) => t.id === s.activeTabId);
      const fallback =
        removedIdx >= 0 && removedIdx < tabs.length
          ? tabs[removedIdx].id
          : tabs[tabs.length - 1]?.id ?? null;
      return { tabs, activeTabId: fallback, dirty: true, ...selPatch };
    });
    // Re-push to the persistent undo stack so the user can immediately undo
    // again. Without this round-trip the redo would be a one-shot.
    // R271 — wrap the push IPC. The visual apply above already happened
    // (workspace tabs mutated, popAiRedo consumed cs), so an unhandled
    // reject here would leave the redo applied with the undo trail gone
    // for that step AND drop cs irretrievably. Restore cs to aiRedoStack
    // on failure so the user's next Ctrl+Shift+Z retries the push (the
    // apply is idempotent — applyChangeset on already-applied md_text /
    // binary_replace / tab_create / tab_delete is a content no-op). Skip
    // the post-push setCanUndo and refresh instead so the toolbar reflects
    // the actual persistent stack state. Sibling of R249 / R245 family.
    try {
      await window.gendoc.undo.push({
        // R216 — Uint8Array-aware serializer; same rationale as the apply
        // path in AIPanel.onApply. Without this, redo of a binary_replace
        // round-trips through SQLite as JSON and corrupts the bytes on the
        // next undo.
        changesetJson: serializeChangeset(cs),
        workspaceId: redoWorkspaceId,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      notify(`重做失敗：${msg}`, 'error');
      // Workspace swap during the failed push window: don't pollute NEW's
      // redo stack with OLD's cs. Same shape as the post-push setCanUndo
      // guard a few lines down.
      if (useWorkspace.getState().workspaceId === redoWorkspaceId) {
        useWorkspace.getState().pushAiRedo(cs);
      }
      void refreshCanUndo();
      return;
    }
    // R181 — same workspace-swap-during-IPC guard as AIPanel.onApply (R180).
    // The push above correctly targets OLD's stack via the captured
    // `redoWorkspaceId`, but the optimistic `setCanUndo(true)` below reads
    // the live store; a swap during the push IPC would write OLD's
    // "redo just landed" optimism onto NEW's toolbar — NEW's actual
    // canUndo is governed by its own SQLite stack and is corrected by
    // R167's refreshCanUndo on workspaceId change, but the brief
    // misleading-button window between this stray write and refreshCanUndo's
    // resolution is exactly what R180 closes for onApply. Same fix shape
    // here keeps the toolbar honest at every moment.
    if (useWorkspace.getState().workspaceId !== redoWorkspaceId) return;
    // The push guarantees ≥1 row exists for this workspaceId — set
    // optimistically rather than re-querying. (If a competing tab clears
    // the stack between the push and our read, the next handleUndo's
    // refreshCanUndo will reconcile anyway.)
    useWorkspace.getState().setCanUndo(true);
    // R271 — refreshCanUndo is in deps because the catch arm above uses it.
    // It's a stable useCallback with `[]` deps so the dependency cycle is
    // a no-op at runtime; declaring it keeps react-hooks/exhaustive-deps
    // honest and matches handleUndo's sibling pattern.
  }, [refreshCanUndo]);

  const handleFocusAI = useCallback(() => {
    window.dispatchEvent(new Event('gendoc:focusAI'));
  }, []);

  /**
   * Pick a folder via the OS dialog and pin it as the file explorer's root.
   * Auto-opens the pane on success (keeps the menu / button contract simple
   * — invoking "Open Folder" should *show* you the folder, not just store it).
   */
  const handleOpenFolder = useCallback(async () => {
    // R200 — same sync gate shape as R197 / R198. `fs.pickDirectory` opens
    // an OS folder picker via `dialog.showOpenDialog` which doesn't block
    // main (async); rapid double-fire of Ctrl+K Ctrl+O / FileExplorer's
    // 「開啟資料夾…」/ menu item would queue two pickDirectory IPCs and
    // Electron stacks two OS folder pickers on screen — confusing for
    // users who can't tell which keystroke spawned which. Module-level
    // gate isn't shared with workspace-open-busy because the semantic is
    // distinct (file explorer root vs workspace replacement) and cross-
    // gating would block legitimate "open workspace then open folder"
    // sequences. Component-local ref is enough for this single entry.
    if (folderPickInFlightRef.current) return;
    folderPickInFlightRef.current = true;
    // R294 — add catch around pickDirectory. Original try had only finally;
    // pickDirectory reject (rare main IPC anomaly: bridge re-bind, channel
    // timeout, dialog system unavailable) escaped as unhandled rejection
    // through `void handleOpenFolder()` at the menu / button callsites.
    // Same R288 / R289 / R290 / R292 / R293 flat-try idiom for sibling
    // IPC-protected entry points.
    try {
      const picked = await window.gendoc.fs.pickDirectory();
      if (!picked) return;
      setExplorerRoot(picked);
      setExplorerOpen(true);
    } catch (err) {
      notify(`開啟資料夾失敗：${(err as Error).message}`, 'error');
    } finally {
      folderPickInFlightRef.current = false;
    }
  }, []);

  // Wire menu → renderer commands.
  useEffect(() => {
    const channels: Array<[string, (e: CustomEvent) => void | Promise<void>]> = [
      ['menu:newProject', () => handleNew()],
      ['menu:open', () => void handleOpen()],
      ['menu:openFolder', () => void handleOpenFolder()],
      ['menu:save', () => void handleSave()],
      ['menu:saveAs', () => void handleSaveAs()],
      ['menu:exportTab', () => void handleExportTab()],
      ['menu:batchExport', () => openBatchExportDialog()],
      ['menu:undo', () => void handleUndo()],
      ['menu:redo', () => void handleRedo()],
      ['menu:focusAI', () => handleFocusAI()],
      ['menu:openSettings', () => setSettingsOpen(true)],
      // detail is the args array forwarded by preload (see MENU_CHANNELS).
      // Recent file paths arrive as `detail[0]`.
      [
        'menu:openRecent',
        (e) => {
          const detail = (e.detail as unknown[]) ?? [];
          const filePath = typeof detail[0] === 'string' ? detail[0] : '';
          void handleOpenRecent(filePath);
        },
      ],
      ['menu:clearRecent', () => void handleClearRecent()],
    ];
    // The contextBridge doesn't currently re-expose ipcRenderer.on; menu
    // commands therefore round-trip through `window.dispatchEvent` from the
    // preload bootstrap. For now we listen on the window directly.
    const listeners = channels.map(([channel, fn]) => {
      const handler = (e: Event) => void fn(e as CustomEvent);
      window.addEventListener(channel, handler);
      return [channel, handler] as const;
    });
    return () => listeners.forEach(([c, h]) => window.removeEventListener(c, h));
  }, [
    handleNew,
    handleOpen,
    handleOpenFolder,
    handleSave,
    handleSaveAs,
    handleExportTab,
    openBatchExportDialog,
    handleUndo,
    handleRedo,
    handleFocusAI,
    handleOpenRecent,
    handleClearRecent,
  ]);

  // Keyboard shortcuts (spec §5.5.2).
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      // Esc exits focus mode regardless of focus context — the user might be
      // mid-edit and want their panels back without having to leave the
      // editor first. We DO bail out of editors that own a meaningful Esc
      // gesture (CodeMirror search panel, our floating dialogs, PptxEditor
      // presentation mode) by checking whether *any* live overlay is
      // rendered first; otherwise Esc would dismiss the overlay AND drop
      // us out of focus mode in a single keystroke — confusing whiplash.
      if (e.key === 'Escape' && focusMode && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
        // Originally only `.cm-search` was checked, but Find / GoTo / both
        // LinkDialogs (role="dialog") and the PptxEditor PresentationMode
        // (data-pptx-presenting) all own Esc too. We use document-level
        // queries (not `target.closest`) because the user may have focus
        // on body or the editor surface while one of these overlays is
        // visible (e.g. clicked back into the editor mid-rename, then
        // pressed Esc to cancel — the overlay still exists, just isn't
        // the focus ancestor).
        const overlayPresent =
          !!document.querySelector('.cm-search') ||
          !!document.querySelector('[role="dialog"]') ||
          !!document.querySelector('[data-pptx-presenting="true"]');
        // Skip when the user is inside an INPUT/TEXTAREA — the input's
        // local Esc usually has its own meaning (cancel sheet rename in
        // XlsxEditor SheetTabs, revert FormulaBar edit, etc.) and these
        // surfaces remain visible in focus mode (the editor body itself
        // isn't hidden, only the chrome around it). Without this skip,
        // pressing Esc to cancel a sheet rename also exits focus mode
        // in the same keystroke — same whiplash the overlay check is
        // meant to prevent. Mirrors the Ctrl+Shift+F skip below. CM /
        // contentEditable surfaces deliberately NOT skipped: Esc has no
        // intrinsic meaning there, so the focus-mode exit is welcome.
        const tag = (e.target as HTMLElement | null)?.tagName;
        const inTextField = tag === 'INPUT' || tag === 'TEXTAREA';
        if (!overlayPresent && !inTextField) {
          e.preventDefault();
          setFocusMode(false);
          return;
        }
      }
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      // R218 — reject any combo that includes Alt. The realistic trigger is
      // AltGr on European keyboards, which the OS reports as Ctrl+Alt held
      // simultaneously: typing AltGr+E to insert「€」on a German layout, or
      // AltGr+S for「ß」shortcut overlays, currently fires handleExportTab /
      // handleSave because every branch below only checks `mod = ctrlKey ||
      // metaKey`. The user intends to type a character; instead they get an
      // OS save dialog and their character never lands in the editor.
      // Same hazard for Cmd+Alt combos on macOS that aren't Gen Doc
      // shortcuts (e.g. Cmd+Alt+S is "Save As" in some apps but here would
      // hit the bare `'s'` branch and fire plain Save). Tightening the
      // top-level gate is the smallest fix that covers every branch in this
      // chain at once; the per-branch `e.shiftKey` checks already in place
      // (Ctrl+Shift+S → saveAs, Ctrl+Shift+F → focus mode, Ctrl+Shift+T →
      // reopen) continue to gate Shift correctly. Mirrors R217's same-day
      // tightening of AIPanel's keyboard handler — there the issue was
      // Shift+Enter / Shift+M; here it's the AltGr leak. Both reach the
      // same destination: only fire the shortcut when the modifier
      // combination matches what the tooltip / menu accelerator advertises.
      if (e.altKey) return;
      if (e.key.toLowerCase() === 'f' && e.shiftKey) {
        // Ctrl+Shift+F toggles focus mode (Adobe InDesign W / Word "Focus"
        // analogue). Skip when the active surface owns Ctrl+F-family
        // shortcuts (Find/Replace dialog) — Ctrl+Shift+F there typically
        // means "find in files" or "replace all", and we don't want to
        // hijack what the editor's local handler may bind. We only check
        // simple INPUT/TEXTAREA — Find dialogs use those — to avoid blocking
        // the gesture from CodeMirror / PptxEditor where there's no
        // conflicting binding.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;
        e.preventDefault();
        setFocusMode((v) => !v);
        return;
      }
      if (e.key.toLowerCase() === 's' && e.shiftKey) {
        e.preventDefault();
        void handleSaveAs();
      } else if (e.key.toLowerCase() === 's') {
        e.preventDefault();
        void handleSave();
      } else if (e.key.toLowerCase() === 'o' && !e.shiftKey) {
        // R373 — `!e.shiftKey` guard added. Same R330 issue pattern: Ctrl+
        // Shift+O is MarkdownEditor's CodeMirror keymap binding for
        // ordered list (MarkdownEditor.tsx:503 `Mod-Shift-o → setLinePrefix
        // (v, '1. ')`). CM6's keymap fires first, adds `1. ` to the line,
        // and bubbles the event up; this listener used to fire too —
        // `handleOpen()` popped the OS file picker on top of the user's
        // bullet-insertion gesture. The user wanted ordered list, got
        // ordered list AND a file dialog.
        e.preventDefault();
        void handleOpen();
      } else if (e.key.toLowerCase() === 'n') {
        e.preventDefault();
        handleNew();
      } else if (e.key.toLowerCase() === 'l' && !e.shiftKey) {
        // R373 — same `!e.shiftKey` guard. Ctrl+Shift+L is MarkdownEditor's
        // CodeMirror keymap binding for unordered list (MarkdownEditor.tsx
        // :497 `Mod-Shift-l → setLinePrefix(v, '- ')`). Pre-fix flow: user
        // types Ctrl+Shift+L → CM6 inserts `- ` → event bubbles → this
        // listener fires `handleFocusAI()` → AI panel grabs focus → caret
        // leaves the editor mid-list-creation. Visible regression: typing
        // a bullet pulls focus away.
        e.preventDefault();
        handleFocusAI();
      } else if (e.key.toLowerCase() === 'e' && e.shiftKey) {
        // R330 — Ctrl+Shift+E is the documented batch-export accelerator (see
        // menu.ts:111). Without this branch, the keystroke flows through to
        // the bare `'e'` clause below (which has no shiftKey check) and
        // fires `handleExportTab` — the SINGLE-tab export — even though the
        // menu accelerator also fires `menu:batchExport`. In Electron 33 menu
        // accelerators do NOT suppress the renderer keydown event by default;
        // the BrowserWindow's `keydown` listener still receives the keystroke
        // alongside the menu click handler. Net result of the missing branch:
        // pressing Ctrl+Shift+E opens the BatchExportDialog AND, concurrently,
        // the OS save dialog from handleExportTab → tryEnterExportTab claims
        // the export-busy gate before BatchExportDialog's confirm runs, so
        // the user later clicking 匯出 in the batch dialog bounces off
        // `tryEnterExportTab() === false` and surfaces R322's busy toast —
        // visible regression from R225/R322's "no dialog stacking" invariant.
        //
        // Defensive renderer-side handling is also correct independent of
        // Electron's accelerator semantics: the IPC `menu:batchExport` path
        // and this in-renderer keydown path should be redundant, not
        // additive. Mirrors the existing `'s'` / `Shift+'s'` split a few
        // lines up — there Ctrl+Shift+S → handleSaveAs and bare Ctrl+S →
        // handleSave; the missing equivalent for 'e' was the lone bug.
        e.preventDefault();
        openBatchExportDialog();
      } else if (e.key.toLowerCase() === 'e' && !e.shiftKey) {
        e.preventDefault();
        void handleExportTab();
      } else if (e.key.toLowerCase() === 'b' && !e.shiftKey) {
        // Ctrl+B toggles the file-explorer pane — matches VS Code. We don't
        // want to fight Markdown's Ctrl+B (bold), so only handle when the
        // active element isn't an editable text field.
        const tag = (e.target as HTMLElement | null)?.tagName;
        const editable = (e.target as HTMLElement | null)?.isContentEditable;
        if (tag === 'INPUT' || tag === 'TEXTAREA' || editable) return;
        e.preventDefault();
        setExplorerOpen((v) => !v);
      } else if (e.key === 'Tab') {
        // Ctrl+Tab / Ctrl+Shift+Tab cycles tabs — same direction convention
        // as VS Code / browsers (forward = next, Shift = previous). We
        // intercept unconditionally because no editor inside us treats
        // Ctrl+Tab as a meaningful keystroke (Tab without Ctrl is still
        // free for indentation inside CodeMirror / textareas).
        const state = useWorkspace.getState();
        const tabs = state.tabs;
        if (tabs.length === 0) return;
        e.preventDefault();
        const cur = tabs.findIndex((t) => t.id === state.activeTabId);
        const start = cur < 0 ? 0 : cur;
        const delta = e.shiftKey ? -1 : 1;
        const nextIdx = (start + delta + tabs.length) % tabs.length;
        state.setActiveTab(tabs[nextIdx].id);
      } else if (e.key.toLowerCase() === 't' && e.shiftKey) {
        // Ctrl+Shift+T — reopen the most recently closed tab. Universal
        // browser/IDE gesture and a common "I closed the wrong one" recovery
        // path; the most likely focus context is exactly the editor pane the
        // user just closed something near, so we intentionally do NOT skip
        // CodeMirror / contentEditable here. The only surfaces we still
        // skip are plain INPUT/TEXTAREA fields (search box, AI prompt) where
        // the user is mid-typing — they can move focus to recover.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;
        e.preventDefault();
        useWorkspace.getState().reopenClosedTab();
      } else if (e.key.toLowerCase() === 'w') {
        // Ctrl+W closes the active tab — matches editors users already know
        // (VS Code / browsers). If there are no tabs, do nothing rather
        // than swallowing the keystroke. Mirror the X-button confirm so a
        // hot-key user can't accidentally drop unsaved work without prompt.
        // Skip when focus is in a plain INPUT/TEXTAREA (tab-rename field,
        // AI prompt textarea, file-search box) — closing the active tab
        // while the user is mid-typing in an unrelated text field is the
        // opposite of what they want, and would also discard the draft.
        // Same focus-skip rule as Ctrl+Shift+T above. CodeMirror /
        // contentEditable editors are deliberately NOT skipped: closing the
        // current document while editing it matches VS Code's Ctrl+W.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;
        const state = useWorkspace.getState();
        const id = state.activeTabId;
        if (!id) return;
        e.preventDefault();
        // Async path so we can route through the main-process native confirm
        // (avoids the Windows `window.confirm()` focus-bug — see notes near
        // handleNew). The keydown handler itself can't be async, so we
        // fire-and-forget the promise.
        // R177 — read the tab's dirty state INSIDE the async closure so a
        // confirm prompt that materialises after the autosave-completed
        // window doesn't lie about the state. Same shape as R176's TabBar
        // close-confirm sites; Ctrl+W's race window is the few hundred ms
        // between key press and the native confirm dialog popping —
        // autosave that completes inside that window flips tab.dirty
        // false, but a closure-bound `tab` from the keydown firing tick
        // would still claim「尚未儲存」.
        // R221 — claim the shared per-tab close gate (lib/tab-close-busy.ts)
        // so auto-repeat Ctrl+W on a dirty tab doesn't stack confirm
        // dialogs. Same gate the three TabBar close paths use, so a
        // mouse-then-keyboard cross-fire (× click + Ctrl+W on the same
        // tab) also blocks at the second action.
        if (!tryEnterClose(id)) return;
        const closeTab = async () => {
          // R290 — add catch around app.confirm. Original try had only
          // finally; confirm-reject (rare main IPC anomaly) escaped as
          // unhandled rejection through `void closeTab()`. Same R288 /
          // R289 flat-try idiom for sibling close paths.
          try {
            const cur = useWorkspace.getState().tabs.find((t) => t.id === id);
            if (!cur) return;
            if (
              cur.dirty &&
              !(await window.gendoc.app.confirm(`「${cur.name}」尚未儲存，確定關閉？`))
            )
              return;
            // Flush pending debounced serializes (PPTX/DOCX/XLSX) before snapshot
            // so the at-close `tab.data` lands in `recentlyClosedTabs` — Ctrl+W
            // followed by Ctrl+Shift+T should round-trip the actual editor state,
            // not the pre-debounce bytes.
            await flushEditors();
            useWorkspace.getState().removeTab(id);
          } catch (err) {
            notify(`關閉頁籤失敗：${(err as Error).message}`, 'error');
          } finally {
            exitClose(id);
          }
        };
        void closeTab();
      } else if (e.key.toLowerCase() === 'z' && e.shiftKey) {
        // Workspace-level redo (Ctrl+Shift+Z). Same focus-skip rules as
        // Ctrl+Z below — editors that own their own undo/redo stack must
        // get the keystroke instead of the workspace handler. The menu
        // accelerator is intentionally registered with
        // `registerAccelerator: false` (see menu.ts) so this is the sole
        // owner of the keyboard binding; without this branch, redo from
        // outside any editor would silently do nothing.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        // R224 — include `data-xlsx-editor-root` in the editor-skip
        // list. XlsxEditor wraps its whole editor in
        // `data-xlsx-editor-root` (XlsxEditor.tsx:1481) and binds
        // useUndoShortcuts at that selector, mirroring the pptx /
        // docx pattern — but App.tsx's inEditor probe was only
        // updated for pptx + docx and silently left xlsx out, so
        // any keydown event whose path bypasses XlsxEditor's
        // capture-phase listener (formula-bar focus transitions,
        // edge cases where target ≠ activeElement, future Z key
        // sources) would fall through to the workspace undo /
        // redo. The capture listener IS the primary defence and
        // works in normal flow, but this bubble-phase guard is
        // documented as belt-and-braces for the cm-editor case
        // ("CodeMirror's keymap doesn't stopPropagation, so the
        // keydown still bubbles here") — that same defensive
        // posture should cover xlsx symmetrically. Pptx + docx
        // get it; xlsx is the lone outlier this round closes.
        const inEditor =
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          target?.isContentEditable ||
          !!target?.closest?.('.cm-editor') ||
          !!target?.closest?.('[data-pptx-editor-root]') ||
          !!target?.closest?.('[data-docx-editor-root]') ||
          !!target?.closest?.('[data-xlsx-editor-root]');
        if (!inEditor) {
          e.preventDefault();
          void handleRedo();
        }
      } else if (e.key.toLowerCase() === 'z' && !e.shiftKey) {
        // Workspace-level undo (rolls back the last applied AI changeset).
        // Skip when focus is inside any editor that owns its own undo stack —
        // otherwise Ctrl+Z in the markdown editor undoes the CM6 history AND
        // pops a separate AI changeset off the persistent stack, which is
        // surprising and corrupts the redo branch users expect to see.
        // Pptx/Docx/Xlsx already block this via their useUndoShortcuts
        // capture-phase listener (R224 added xlsx to this guard), but
        // CodeMirror's keymap doesn't stopPropagation, so the keydown
        // still bubbles here.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        // R224 — same xlsx addition as the redo branch above.
        const inEditor =
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          target?.isContentEditable ||
          !!target?.closest?.('.cm-editor') ||
          !!target?.closest?.('[data-pptx-editor-root]') ||
          !!target?.closest?.('[data-docx-editor-root]') ||
          !!target?.closest?.('[data-xlsx-editor-root]');
        if (!inEditor) {
          e.preventDefault();
          void handleUndo();
        }
      } else if (/^[1-9]$/.test(e.key)) {
        // Ctrl+1..9 jumps to tab N. Guard against editors that bind the
        // same combo (Markdown's Mod-1/2/3 → headings; CodeMirror swallows
        // them but the keydown still bubbles to us). Skip when focus is
        // inside an editable surface — the user expects the editor's
        // mapping to win there.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        const inEditor =
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          target?.isContentEditable ||
          !!target?.closest?.('.cm-editor');
        if (inEditor) return;
        e.preventDefault();
        const idx = parseInt(e.key, 10) - 1;
        const tabs = useWorkspace.getState().tabs;
        if (tabs[idx]) useWorkspace.getState().setActiveTab(tabs[idx].id);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [handleSave, handleSaveAs, handleOpen, handleNew, handleFocusAI, handleExportTab, openBatchExportDialog, handleUndo, handleRedo, focusMode]);

  // OS file drag-drop. Electron 32+ removed `File.path`, so we resolve each
  // dropped File through the preload's `webUtils.getPathForFile` shim and
  // route it through the same open-file pipeline the explorer uses (.gd →
  // workspace.openPath; everything else → fs.readFile + openExternalFile).
  // We track dragenter/leave with a depth counter because dragover fires on
  // every child element and a naive bool flicker on/off during traversal.
  useEffect(() => {
    let depth = 0;
    // Watchdog: if the user drags out of the window via a path that doesn't
    // fire a final `dragleave` (Electron quirk on some platforms), the depth
    // counter never decrements and the overlay sticks. Each `dragover` resets
    // this timer; if no event arrives for ~120ms the drag has clearly left
    // and we force-clear. Re-armed only while a drag is active.
    let stuckTimer: ReturnType<typeof setTimeout> | null = null;
    const armWatchdog = () => {
      if (stuckTimer) clearTimeout(stuckTimer);
      stuckTimer = setTimeout(() => {
        depth = 0;
        setDropActive(false);
      }, 120);
    };
    const disarmWatchdog = () => {
      if (stuckTimer) {
        clearTimeout(stuckTimer);
        stuckTimer = null;
      }
    };
    const hasFiles = (e: DragEvent) =>
      Array.from(e.dataTransfer?.types ?? []).includes('Files');
    const onEnter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth += 1;
      setDropActive(true);
      armWatchdog();
    };
    const onOver = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      // Must preventDefault to allow the drop event to fire.
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
      armWatchdog();
    };
    const onLeave = () => {
      depth = Math.max(0, depth - 1);
      if (depth === 0) {
        setDropActive(false);
        disarmWatchdog();
      }
    };
    const onDrop = async (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth = 0;
      setDropActive(false);
      disarmWatchdog();
      const files = Array.from(e.dataTransfer?.files ?? []);
      if (files.length === 0) return;
      // First check for a .gd among the drop — that takes the whole-workspace
      // path and we should only open one. If there are multiple .gd files,
      // open the first; the remaining loose files are then ignored to avoid
      // mixing two distinct workspaces' tabs.
      const gd = files.find((f) => f.name.toLowerCase().endsWith('.gd'));
      if (gd) {
        // R209 — extend R199's `tryEnterOpen` gate to the drag-drop .gd
        // path. R187 already enrolled this site in the load-gen exclusion
        // class (line 1177) but the sync-gate side was missed when R199
        // swept handleOpen / handleOpenRecent / FileExplorer; R205 closed
        // EmptyState; R207 closed handleNew; this is the last entry point
        // that opens an `app.confirm` dialog without first claiming the
        // gate. Realistic stack-up: user drags a .gd over the window and
        // simultaneously hits Ctrl+N (keyboard still receives keystrokes
        // during a drag) — handleNew's confirm pops, then the drop
        // resolves and pushes a SECOND confirm on top with no sign of
        // which gesture triggered which. Same shape applies to a Ctrl+O
        // / drop overlap. The gate is renderer-wide so any in-flight
        // workspace-replacement entry blocks all the others, matching
        // R199 doc-block's "single dialog stack" guarantee.
        if (!tryEnterOpen()) return;
        // R288 — collapse the nested try blocks into a single try/catch/
        // finally. Original layout had outer try (with only `finally`,
        // no catch) + inner try/catch around openPath. The outer-without-
        // catch left `app.confirm` (dirty-confirm IPC) unprotected: if main
        // rejected (process unresponsive, IPC bridge re-bind), the throw
        // escaped both blocks unhandled — `finally` released the gate but
        // the rejection leaked out of `onDrop`'s async body, where the
        // event listener doesn't await. Sibling handleOpen / handleOpenRecent
        // (lines ~319-344) use a single flat try; drag-drop's nested shape
        // was the lone outlier. Merge under one catch so confirm-throw,
        // openPath-throw, getPathForFile-throw all surface as the same
        // toast and the workspace-open gate is always released.
        try {
          if (
            useWorkspace.getState().dirty &&
            !(await window.gendoc.app.confirm('目前的變更尚未儲存，確定開啟其他檔案？'))
          )
            return;
          const myGen = bumpLoadGen();
          const path = window.gendoc.webUtils.getPathForFile(gd);
          const opened = await window.gendoc.workspace.openPath(path);
          // R168 — newer load wins. See workspace-load-gen.ts doc-block.
          if (myGen !== currentLoadGen()) return;
          if (opened) useWorkspace.getState().loadFromOpened(opened);
        } catch (err) {
          notify(`開啟失敗：${(err as Error).message}`, 'error');
        } finally {
          exitOpen();
        }
        return;
      }
      // Open each non-.gd file as a new tab. Sequential rather than parallel
      // so the resulting tab order matches the drop order. Pass the disk
      // path through so re-dropping the same file activates the existing
      // tab instead of stacking duplicates.
      // Failures (folders, unsupported types, read errors) accumulate so the
      // user sees a single summary at the end instead of N blocking alerts —
      // dropping a directory of 20 mixed files used to require 20 OK clicks.
      // R248 — capture workspaceId at drop start so the per-file
      // openExternalFile / setActiveTab calls don't leak into a workspace
      // the user has since swapped to. The drop's `for...of` loop is
      // sequential and each `await window.gendoc.fs.readFile(path)` is
      // a swap window: a slow first file (large pptx, network drive)
      // can hold the loop for seconds, during which the user might
      // Ctrl+O / Ctrl+N / drop a .gd to swap workspace. Without this
      // guard, the remaining iterations:
      //   • read live store via `useWorkspace.getState().tabs.find(...)`
      //     for the dedupe — finds nothing in NEW (different tab set)
      //     and proceeds to openExternalFile.
      //   • call `openExternalFile(content, path)` — which mutates LIVE
      //     workspace store, adding tabs to NEW that the user dropped
      //     onto OLD. Same applies to `setActiveTab(existing.id)` in
      //     the dedupe-hit branch — activates a tab id that may exist
      //     in NEW by coincidence (uuid collision impossible) or no-
      //     ops (R208 stale-id guard) but in either case targets the
      //     wrong workspace.
      // Same shape as R170 (performSave), R171 (handleExportTab), R178/
      // R179 (handleAssistantToolCalls / continueAfterToolResult): drop
      // is a multi-step async user gesture that must pin to its
      // origin workspace. Aborting the whole loop on swap is correct
      // — partial drops (3 of 5 files added) would silently mismatch
      // the user's "drop these N files together" intent; bailing
      // wholesale leaves OLD with whatever already landed and NEW
      // untouched.
      const dropWorkspaceId = useWorkspace.getState().workspaceId;
      const failures: string[] = [];
      for (const file of files) {
        if (useWorkspace.getState().workspaceId !== dropWorkspaceId) break;
        try {
          const path = window.gendoc.webUtils.getPathForFile(file);
          if (!path) {
            failures.push(`${file.name || '(unknown)'}：無法取得路徑（資料夾或非檔案）`);
            continue;
          }
          const existing = useWorkspace
            .getState()
            .tabs.find((t) => t.sourcePath === path);
          if (existing) {
            useWorkspace.getState().setActiveTab(existing.id);
            continue;
          }
          const content = await window.gendoc.fs.readFile(path);
          // R248 — re-check after the readFile await; swap could land here.
          if (useWorkspace.getState().workspaceId !== dropWorkspaceId) break;
          const id = useWorkspace.getState().openExternalFile(content, path);
          if (!id) {
            failures.push(`${content.name}：不支援的檔案類型`);
          }
        } catch (err) {
          failures.push(`${file.name}：${(err as Error).message}`);
        }
      }
      // R248 — only surface the failure summary on the originating workspace.
      // Cross-workspace toast is acceptable in principle (Toaster is
      // global, see AIPanel/handleExportTab toast precedent) but a「N 個
      // 檔案無法開啟」 message in NEW workspace, listing OLD's drop
      // contents, would confuse — NEW didn't receive any drop. Skip if
      // workspace already swapped.
      if (useWorkspace.getState().workspaceId !== dropWorkspaceId) return;
      if (failures.length === 1) {
        notify(`開啟失敗：${failures[0]}`, 'error');
      } else if (failures.length > 1) {
        notify(`${failures.length} 個檔案無法開啟：\n\n${failures.join('\n')}`, 'error');
      }
    };
    window.addEventListener('dragenter', onEnter);
    window.addEventListener('dragover', onOver);
    window.addEventListener('dragleave', onLeave);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragenter', onEnter);
      window.removeEventListener('dragover', onOver);
      window.removeEventListener('dragleave', onLeave);
      window.removeEventListener('drop', onDrop);
      disarmWatchdog();
    };
  }, []);

  return (
    <div className="flex h-screen w-screen overflow-hidden">
      {/* Focus mode hides chrome (Toolbar / FileExplorer / AIPanel /
          StatusBar) but keeps the editor surface mounted with the same
          props — we don't unmount because that would tear down all editor
          local state (PptxEditor activeIdx, CodeMirror selection, undo
          history) and re-running parsePptx / re-mounting CM6 on every
          F-toggle would be both slow and lossy. We just don't render the
          chrome around it. */}
      {!focusMode && explorerOpen ? (
        <aside
          className="relative shrink-0 border-r border-border"
          style={{ width: explorerWidth }}
        >
          <FileExplorer
            rootPath={explorerRoot}
            onPickFolder={() => void handleOpenFolder()}
            onCollapse={() => setExplorerOpen(false)}
          />
          <ExplorerResizeHandle width={explorerWidth} onChange={setExplorerWidth} />
        </aside>
      ) : null}
      <main className="flex-1 min-w-0 flex flex-col">
        {!focusMode && (
          <Toolbar
            dirty={ws.dirty}
            title={ws.manifest.title}
            filePath={ws.filePath}
            activeTabType={ws.tabs.find((t) => t.id === ws.activeTabId)?.type}
            saveState={ws.saveState}
            saveError={ws.saveError}
            explorerOpen={explorerOpen}
            canRedo={ws.aiRedoStack.length > 0}
            canUndo={ws.canUndo}
            onToggleExplorer={() => setExplorerOpen((v) => !v)}
            onOpenFolder={() => void handleOpenFolder()}
            onNew={handleNew}
            onOpen={() => void handleOpen()}
            onSave={() => void handleSave()}
            onExportTab={() => void handleExportTab()}
            onBatchExport={openBatchExportDialog}
            hasAnyTab={ws.tabs.length > 0}
            onUndo={() => void handleUndo()}
            onRedo={() => void handleRedo()}
            onSettings={() => setSettingsOpen(true)}
            onEnterFocus={() => setFocusMode(true)}
          />
        )}
        <div className="flex-1 min-h-0">
          <EditorSurface />
        </div>
        {!focusMode && <StatusBar />}
      </main>
      {!focusMode && (
        <AIPanel
          width={aiPanelWidth}
          onWidthChange={setAiPanelWidth}
          hasApiKey={hasApiKey}
          onOpenSettings={() => setSettingsOpen(true)}
        />
      )}
      {focusMode && (
        // Floating exit pill — minimal but discoverable. Top-right corner so
        // it doesn't fight with the EditorSurface's own header (TabBar lives
        // there for some editors). Translucent and only fully opaque on
        // hover so it doesn't camp visually over content the user is reading.
        <button
          type="button"
          onClick={() => setFocusMode(false)}
          title="離開專注模式 (Esc 或 Ctrl+Shift+F)"
          className="fixed top-3 right-3 z-50 flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-background/70 hover:bg-background border border-border/60 hover:border-border shadow-sm text-[11px] text-muted-foreground hover:text-foreground backdrop-blur-sm transition-all"
        >
          <Minimize2 className="h-3 w-3" />
          離開專注模式
        </button>
      )}
      <SettingsDialog
        open={settingsOpen}
        onOpenChange={setSettingsOpen}
        onApiKeyChange={setHasApiKey}
      />
      <BatchExportDialog
        open={batchExportOpen}
        onOpenChange={setBatchExportOpen}
        tabs={ws.tabs}
        onConfirm={(ids) => void handleBatchExport(ids)}
      />
      {dropActive ? (
        // pointer-events-none so the underlying drop events still fire on
        // window — the overlay is purely visual feedback.
        <div className="fixed inset-0 z-[100] pointer-events-none flex items-center justify-center bg-primary/10 backdrop-blur-[2px]">
          <div className="flex flex-col items-center gap-3 px-8 py-6 rounded-lg border-2 border-dashed border-primary bg-background/95 shadow-lg">
            <Upload className="h-8 w-8 text-primary" />
            <div className="text-sm font-medium">
              將檔案放入以開啟（.gd / .md / .html / .docx / .xlsx / .pptx）
            </div>
          </div>
        </div>
      ) : null}
      {/* Global toast viewport — replaces blocking `alert()` popups for
          fire-and-forget notifications (open/save/export failures, "no
          tab to export", etc.). Mounted once at the App root so any code
          path can fire `notify(...)` without prop drilling. */}
      <Toaster />
    </div>
  );
}

function Toolbar({
  dirty,
  title,
  filePath,
  activeTabType,
  saveState,
  saveError,
  explorerOpen,
  canRedo,
  canUndo,
  onToggleExplorer,
  onOpenFolder,
  onNew,
  onOpen,
  onSave,
  onExportTab,
  onBatchExport,
  hasAnyTab,
  onUndo,
  onRedo,
  onSettings,
  onEnterFocus,
}: {
  dirty: boolean;
  title: string;
  filePath: string;
  activeTabType: 'markdown' | 'html' | 'docx' | 'xlsx' | 'pptx' | undefined;
  saveState: 'idle' | 'saving' | 'success' | 'error';
  saveError: string | null;
  explorerOpen: boolean;
  canRedo: boolean;
  canUndo: boolean;
  onToggleExplorer: () => void;
  onOpenFolder: () => void;
  onNew: () => void;
  onOpen: () => void;
  onSave: () => void;
  onExportTab: () => void;
  /** Open the batch-export checkbox dialog. */
  onBatchExport: () => void;
  /** Workspace has at least one tab (used to grey-out batch export when empty). */
  hasAnyTab: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onSettings: () => void;
  onEnterFocus: () => void;
}) {
  // Show only the file name + parent dir, not the full absolute path. Keeps
  // the bar from getting clobbered by deep paths and the full path is still
  // available in the tooltip / status bar.
  const shortPath = filePath ? shortenPath(filePath) : '未儲存';
  const exportExt = activeTabType === 'markdown' ? '.md' : activeTabType ? `.${activeTabType}` : '';
  const exportTitle = activeTabType
    ? `匯出此頁籤為 ${exportExt} (Ctrl+E)`
    : '匯出此頁籤（先開啟一個頁籤）';
  // Compose the Save button's icon + title from saveState. Saving wins
  // over dirty (we're literally saving), success briefly flashes a check,
  // error keeps the icon but tints + tooltips the message.
  let saveIcon: JSX.Element;
  let saveTitle: string;
  let saveDisabled = false;
  if (saveState === 'saving') {
    saveIcon = <Loader2 className="h-4 w-4 animate-spin" />;
    saveTitle = '儲存中…';
    saveDisabled = true;
  } else if (saveState === 'success') {
    saveIcon = <Check className="h-4 w-4 text-emerald-500" />;
    saveTitle = '已儲存';
  } else if (saveState === 'error') {
    saveIcon = <AlertCircle className="h-4 w-4 text-destructive" />;
    saveTitle = `儲存失敗：${saveError ?? '未知錯誤'}`;
  } else {
    saveIcon = <Save className="h-4 w-4" />;
    // Surface 另存新檔 (Ctrl+Shift+S) alongside the primary 儲存 hint.
    // The keystroke is fully wired (App.tsx:692-694 → handleSaveAs → IPC
    // workspace.saveAs at ipc.ts:89) and the native File menu lists it
    // (menu.ts:99-101), but the native menu is hidden-by-default on Windows
    // unless Alt is held down — the only in-app surface advertising it was
    // therefore invisible to the typical user. The Save button is the
    // natural discovery point for the related "save under a different name"
    // action: a user looking at this button to confirm Ctrl+S is also the
    // user most likely to ask "and how do I save as?". Same sibling-
    // shortcut-in-tooltip pattern as the find-replace navigation buttons
    // (FindReplaceDialog.tsx:620 / 632 — "上一個 (Shift+Enter / Shift+F3)"
    // pairs the primary key with its mirror), and the StatusBar's export
    // tooltip pattern. Only emitted in the idle branch — saving / success /
    // error tooltips report transient state, where extra shortcut text
    // would dilute the live status read. The middle-dot "·" mirrors the
    // existing R39 ContextItem layout convention for compound hint strings.
    const saveAsHint = ' · 另存新檔 (Ctrl+Shift+S)';
    saveTitle = (dirty ? '儲存（有未儲存變更） (Ctrl+S)' : '儲存 (Ctrl+S)') + saveAsHint;
  }
  return (
    <div className="flex items-center h-11 px-2 border-b gap-0.5 bg-secondary/20">
      <Button
        size="icon"
        variant={explorerOpen ? 'secondary' : 'ghost'}
        onClick={onToggleExplorer}
        title={explorerOpen ? '收合檔案總管 (Ctrl+B)' : '展開檔案總管 (Ctrl+B)'}
        // R153 — toggle-state SR exposure. The `variant` flips secondary /
        // ghost to give sighted users a visible「目前展開」 cue, but SR users
        // had no equivalent. Same `aria-pressed` pattern landed this round
        // on the AIPanel Wrench button (line ~787) and the four ToolbarBtn
        // definitions in MarkdownToolbar / DocxEditor / PptxEditor / Xlsx
        // Editor — this is the only icon-only toggle in the top toolbar
        // that carries persistent state (Save / Undo / Redo / Export are
        // action buttons whose `disabled` already conveys their state).
        aria-pressed={explorerOpen}
      >
        <FolderTree className="h-4 w-4" />
      </Button>
      <div className="w-px h-5 bg-border mx-1" />
      <Button size="icon" variant="ghost" onClick={onNew} title="新專案 (Ctrl+N)">
        <FilePlus2 className="h-4 w-4" />
      </Button>
      <Button size="icon" variant="ghost" onClick={onOpen} title="開啟 .gd (Ctrl+O)">
        <FolderOpen className="h-4 w-4" />
      </Button>
      <Button size="icon" variant="ghost" onClick={onOpenFolder} title="開啟資料夾… (Ctrl+K Ctrl+O)">
        <FolderInput className="h-4 w-4" />
      </Button>
      {/* Save button reflects async state: spinner while saving, check on
          success, alert on error. The dirty-dot only shows in idle so it
          doesn't fight with the saving spinner. */}
      <Button
        size="icon"
        variant="ghost"
        onClick={onSave}
        disabled={saveDisabled}
        title={saveTitle}
        className="relative"
      >
        {saveIcon}
        {saveState === 'idle' && dirty && (
          <span className="absolute top-1 right-1 h-1.5 w-1.5 rounded-full bg-amber-500 ring-2 ring-secondary/20" />
        )}
      </Button>
      <Button
        size="icon"
        variant="ghost"
        onClick={onExportTab}
        disabled={!activeTabType}
        title={exportTitle}
      >
        <Download className="h-4 w-4" />
      </Button>
      {/* Batch export — open dialog to checkbox-select multiple tabs, then
          pick a destination folder and write each in its native format.
          Sibling of the single-tab Download button to its left. Disabled
          when the workspace has no tabs at all (nothing to pick from). */}
      <Button
        size="icon"
        variant="ghost"
        onClick={onBatchExport}
        disabled={!hasAnyTab}
        // R340 — advertise the keyboard accelerator inline. Every other
        // toolbar button this surface renders includes its shortcut in
        // parens — 新專案 (Ctrl+N) at line 1992, 開啟 .gd (Ctrl+O) at
        // 1995, 開啟資料夾… (Ctrl+K Ctrl+O) at 1998, 匯出此頁籤為 .md
        // (Ctrl+E) at line 1932-1933's exportTitle, Undo (Ctrl+Z) /
        // Redo (Ctrl+Y) further right, and AI 對話框 (Ctrl+L) on the
        // right side. The batch-export icon was the lone outlier
        // advertising only the function and not the shortcut — even
        // though the same accelerator IS bound (menu.ts:111
        // `CmdOrCtrl+Shift+E`, App.tsx:1349-1353's R330 keydown
        // handler) and named in the native menu「批次匯出多個頁籤…
        // Ctrl+Shift+E」. Toolbar tooltip is where most users would
        // discover keyboard shortcuts in this app (the menu bar is
        // not auto-shown on Windows / Linux unless they press Alt),
        // so omitting it here costs visibility on a real
        // productivity affordance. Empty-tabs branch stays as-is —
        // shortcut still works the same way, but advertising it
        // when the button is greyed out is more noise than signal.
        title={
          hasAnyTab
            ? '批次匯出多個頁籤到資料夾 (Ctrl+Shift+E)'
            : '沒有可匯出的頁籤'
        }
      >
        <FileDown className="h-4 w-4" />
      </Button>
      <div className="w-px h-5 bg-border mx-1" />
      <Button
        size="icon"
        variant="ghost"
        onClick={onUndo}
        disabled={!canUndo}
        title={canUndo ? '復原 AI 變更 (Ctrl+Z)' : '沒有可復原的變更'}
      >
        <Undo2 className="h-4 w-4" />
      </Button>
      <Button
        size="icon"
        variant="ghost"
        onClick={onRedo}
        disabled={!canRedo}
        title={canRedo ? '重做 AI 變更 (Ctrl+Shift+Z)' : '沒有可重做的變更'}
      >
        <Redo2 className="h-4 w-4" />
      </Button>

      <div className="flex-1 flex flex-col items-center min-w-0 px-3">
        <div className="flex items-center gap-1.5 max-w-full">
          {dirty && (
            // Same hover/screen-reader affordance the TabBar's per-tab dirty
            // dot got: a bare 6 px amber circle is opaque to first-time users
            // who can see "something is colour-coded near my project name"
            // but not what it means. The Save button to the left already
            // surfaces the same info via its tooltip, but this dot lives in
            // the title cluster — the natural place a user looks to confirm
            // "is the workspace I'm reading the latest version?" — so the
            // affordance belongs here too. Pattern lifted verbatim from
            // TabBar.tsx:299-303 so the same dot in two places reads the
            // same way.
            <span
              className="h-1.5 w-1.5 rounded-full bg-amber-500 shrink-0"
              title="尚未儲存"
              aria-label="尚未儲存"
            />
          )}
          {/* Truncate-with-no-tooltip was a quiet inconsistency in this
              two-row title cluster: the filePath row immediately below has
              `title={filePath || '尚未儲存到磁碟'}` (line ~1253) so a user
              hovering a clipped path sees the full string, but the project
              title above it — same `truncate` class, same flex container
              that narrows on small windows / long titles / many toolbar
              buttons — exposed nothing on hover. The TabBar's tab-name span
              at TabBar.tsx:334 has carried this affordance from the start
              (`title={tab.sourcePath ?? tab.name…}`), so the missing
              tooltip here was the lone outlier across all three truncatable
              labels in the chrome (tab name, project title, file path).
              Mirror the filePath sibling's "value || placeholder" shape so
              empty manifest titles still surface the visible '未命名專案'
              fallback in the tooltip rather than nothing. */}
          <span
            className="text-sm font-medium truncate"
            title={title || '未命名專案'}
          >
            {title || '未命名專案'}
          </span>
        </div>
        <div
          className="text-[10px] text-muted-foreground truncate max-w-full"
          title={filePath || '尚未儲存到磁碟'}
          dir="rtl"
        >
          {shortPath}
        </div>
      </div>

      <Button
        size="icon"
        variant="ghost"
        onClick={onEnterFocus}
        title="進入專注模式 — 隱藏側欄與工具列以全幅檢視文件 (Ctrl+Shift+F)"
      >
        <FocusIcon className="h-4 w-4" />
      </Button>
      {/* Surface the Ctrl+, accelerator the same way every other button in
          this toolbar does (新專案 Ctrl+N, 開啟 Ctrl+O, 儲存 Ctrl+S, 匯出 Ctrl+E,
          復原 Ctrl+Z, 專注模式 Ctrl+Shift+F …). The shortcut is wired in
          menu.ts:180-183 → 'menu:openSettings' → setSettingsOpen(true) at
          App.tsx:597, so the keystroke really does open this dialog —
          previously the only hint was the AI submenu in the native app menu,
          which most users don't go hunting through. */}
      <Button size="icon" variant="ghost" onClick={onSettings} title="設定 (Ctrl+,)">
        <SettingsIcon className="h-4 w-4" />
      </Button>
    </div>
  );
}

/**
 * Right-edge resize handle for the file-explorer pane. Mirrors `AIPanel`'s
 * left-edge handle but flipped (drag right = wider). Bounded so the pane
 * can't shrink below "useful" or eat the whole window on a wide monitor.
 */
function ExplorerResizeHandle({
  onChange,
}: {
  width: number;
  onChange: (w: number) => void;
}): JSX.Element {
  return (
    <div
      className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-primary/40 active:bg-primary/60 z-10 transition-colors"
      onMouseDown={(e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = (e.currentTarget.parentElement as HTMLElement).offsetWidth;
        // See AIPanel ResizeHandle — same body-cursor / user-select lock
        // for the duration of the drag. Without it the cursor flickers
        // off the 1-px handle whenever the user drags faster than the
        // panel re-layouts, and text under the pointer becomes selected.
        const prevCursor = document.body.style.cursor;
        const prevUserSelect = document.body.style.userSelect;
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        const onMove = (ev: MouseEvent) => {
          const next = Math.max(180, Math.min(560, startWidth + (ev.clientX - startX)));
          onChange(next);
        };
        const onUp = () => {
          document.body.style.cursor = prevCursor;
          document.body.style.userSelect = prevUserSelect;
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      }}
    />
  );
}

/** Compress a long absolute path to "...\parent\file.gd". */
function shortenPath(path: string): string {
  const parts = path.split(/[\\/]/);
  if (parts.length <= 2) return path;
  return `…/${parts.slice(-2).join('/')}`;
}
