import { app, BrowserWindow, dialog, ipcMain, shell, screen } from 'electron';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { registerIpcHandlers, bootStorageHousekeeping, shutdownStorage } from './ipc';
import { buildAppMenu } from './menu';
import { loadConfig, patchConfig } from './storage/config';
import { cancelAll as cancelAllAiStreams } from './ai/anthropic';
import { IPC } from '../types/ipc';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

let mainWindow: BrowserWindow | null = null;

// Renderer pushes its dirty flag here on every change so the OS close path can
// prompt before destroying unsaved work. We also track `isQuitting` to break
// the close-handler loop after the user picks "Don't Save" or after a queued
// "Save then Quit" finishes.
let rendererDirty = false;
let isQuitting = false;

const isDev = !app.isPackaged;
const VITE_DEV_SERVER_URL = process.env['ELECTRON_RENDERER_URL'] ?? 'http://localhost:5173';

// Logo lives next to the app at packaging time and at the project root in dev.
// In dev __dirname = <project>/dist-electron/main; in prod = <app.asar>/dist-electron/main.
// `app.getAppPath()` resolves to the project root (dev) or app.asar root (prod).
const APP_ICON_PATH = join(app.getAppPath(), 'logo', 'logo.png');

async function createMainWindow(): Promise<void> {
  // Restore last-known window bounds, but only if they still land on a
  // currently-attached display — otherwise an unplugged external monitor
  // would strand the window off-screen. Validate by intersecting with each
  // display's work area; require at least 200×200 of overlap.
  const cfg = await loadConfig();
  const saved = cfg.windowBounds;
  let bounds: { x?: number; y?: number; width: number; height: number } = {
    width: 1440,
    height: 900,
  };
  if (saved) {
    const onScreen = screen.getAllDisplays().some((d) => {
      const a = d.workArea;
      const ix = Math.max(saved.x, a.x);
      const iy = Math.max(saved.y, a.y);
      const ax = Math.min(saved.x + saved.width, a.x + a.width);
      const ay = Math.min(saved.y + saved.height, a.y + a.height);
      return ax - ix >= 200 && ay - iy >= 200;
    });
    if (onScreen) bounds = saved;
  }

  mainWindow = new BrowserWindow({
    ...bounds,
    minWidth: 960,
    minHeight: 600,
    backgroundColor: '#ffffff',
    icon: APP_ICON_PATH,
    show: false,
    webPreferences: {
      preload: join(__dirname, '..', 'preload', 'preload.mjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  // Persist window bounds on resize/move. Debounce — these events fire many
  // times per second during a drag and we don't want to thrash the JSON file.
  let saveBoundsTimer: ReturnType<typeof setTimeout> | null = null;
  const saveBounds = () => {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    // R219 — also skip when the window is maximized. The doc-block below
    // claims maximized state isn't persisted ("restoring at the saved
    // bounds is friendlier than yanking the user back into a maximized
    // window they may have un-maximized for a reason"), but the previous
    // skip-list was minimized + fullscreen ONLY. Maximize fires a `resize`
    // event too, and `getBounds()` on a maximized window returns the
    // maximized rectangle (e.g. 1920×1080 on a 1080p monitor). So a user
    // who maximized their session ended up with windowBounds = full
    // monitor work-area persisted to disk; next launch restored that
    // size as a *non*-maximized window — visually filling the screen
    // but in restored state, so clicking the maximize button does
    // nothing visible (size already matches), drag-from-title-bar
    // moves the whole 1920×1080 frame around, and the user wonders
    // why the chrome behaves so oddly. Aligning the skip list with
    // the comment fixes both: maximize state isn't persisted (next
    // launch restores last un-maximized bounds), and the un-maximize
    // → resume-typical-bounds path works as before. Mirrors the
    // already-correct `isFullScreen()` skip exactly one line above —
    // both are "transient layout states the user can re-enter via
    // gesture, persisting them locks future sessions into a confusing
    // shape."
    if (
      mainWindow.isMinimized() ||
      mainWindow.isFullScreen() ||
      mainWindow.isMaximized()
    ) return;
    const b = mainWindow.getBounds();
    if (saveBoundsTimer) clearTimeout(saveBoundsTimer);
    saveBoundsTimer = setTimeout(() => {
      void patchConfig({ windowBounds: b }).catch(() => undefined);
    }, 400);
  };
  // Maximized state intentionally not persisted — restoring at the saved
  // bounds is friendlier than yanking the user back into a maximized window
  // they may have un-maximized for a reason. R219 brings the actual skip
  // list at the top of `saveBounds` in line with this claim.

  mainWindow.on('ready-to-show', () => mainWindow?.show());
  mainWindow.on('resize', saveBounds);
  mainWindow.on('move', saveBounds);

  // Save-before-quit prompt. If the renderer reports dirty state and the user
  // clicks the OS close button (or Cmd+Q on macOS routes through here via
  // `before-quit` below), block the close, ask Save / Don't Save / Cancel,
  // and resolve via an explicit ack from the renderer (`saveAndQuitResult`)
  // so a cancelled Save-As doesn't leave the window in a half-quit state.
  mainWindow.on('close', (event) => {
    if (isQuitting || !rendererDirty || !mainWindow) return;
    event.preventDefault();
    const choice = dialog.showMessageBoxSync(mainWindow, {
      type: 'warning',
      buttons: ['儲存', '不儲存', '取消'],
      defaultId: 0,
      cancelId: 2,
      title: 'Gen Doc',
      message: '目前的變更尚未儲存',
      detail: '是否要在離開前儲存？',
      noLink: true,
    });
    if (choice === 0) {
      // Save-and-quit. Renderer runs its save flow then signals success or
      // cancel back via `app:saveAndQuitResult` — see ipcMain handler below.
      isQuitting = true;
      mainWindow.webContents.send('menu:saveAndQuit');
    } else if (choice === 1) {
      isQuitting = true;
      mainWindow.destroy();
    }
    // 2 (cancel) = leave isQuitting=false; the next close attempt re-prompts.
  });

  // External links open in the OS default browser, not inside Electron.
  // `setWindowOpenHandler` only catches `window.open()` / `target="_blank"`
  // — plain `<a href="https://...">` clicks (which is what `marked` emits
  // for markdown preview links) fire `will-navigate` instead. Without the
  // second handler, clicking any link in the rendered preview replaces the
  // editor surface with the linked page; F5 reload is suppressed below so
  // the user has no way back short of quitting the app, losing unsaved
  // tab edits.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  // Capture `webContents` here so the closure doesn't depend on the module-
  // level `mainWindow` (which TS sees as possibly-null after the window's
  // `closed` handler nulls it).
  const wc = mainWindow.webContents;
  wc.on('will-navigate', (event, url) => {
    // Only divert cross-origin http(s) navigation. Same-origin (vite dev
    // server, internal app routes) and non-http schemes (file://, data:,
    // about:) must pass through — diverting the initial load or HMR
    // reload would break the app.
    if (!/^https?:\/\//i.test(url)) return;
    try {
      const target = new URL(url);
      const current = new URL(wc.getURL());
      if (target.origin === current.origin) return;
    } catch {
      // Unparseable URL — fall through to the safe path (prevent + external).
    }
    event.preventDefault();
    void shell.openExternal(url);
  });

  // Suppress Chromium's default F5 / Ctrl+R reload — F5 is bound to
  // PowerPoint presentation mode in the editor; an accidental reload would
  // wipe unsaved tab edits and look like a crash. The renderer still sees
  // the keydown so the editor can react to it.
  mainWindow.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown') return;
    const key = input.key;
    if (key === 'F5') {
      event.preventDefault();
      return;
    }
    if ((input.control || input.meta) && (key === 'r' || key === 'R')) {
      event.preventDefault();
    }
  });

  if (isDev) {
    await mainWindow.loadURL(VITE_DEV_SERVER_URL);
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    await mainWindow.loadFile(join(__dirname, '..', '..', 'dist', 'index.html'));
  }
}

// Renderer-side dirty flag mirror. Wire this once at module load — it has no
// per-window state, so registering inside `createMainWindow` would re-add
// listeners on `app.activate`-driven re-creates and leak.
ipcMain.on(IPC.app.setDirty, (_e, dirty: boolean) => {
  rendererDirty = !!dirty;
});

// Native Yes/Cancel dialog used as a drop-in replacement for renderer-side
// `window.confirm()`. We can't use Chromium's native confirm on Windows: it
// leaves the BrowserWindow in a state where keystrokes are no longer routed
// to webContents — the symptom was that pressing Ctrl+N then opening a fresh
// blank Word tab produced a contentEditable that accepted clicks
// (activeElement set, Selection range placed) but `document.hasFocus()` was
// false and no keydown ever reached the renderer. This handler is called
// by every renderer prompt (dirty-discard, tab close, slide / sheet /
// textbox delete, layout-apply); it routes OS focus correctly and we
// additionally re-focus the window + webContents after dismissal as a
// belt-and-braces guarantee.
ipcMain.handle(IPC.app.confirm, async (_e, message: string) => {
  if (!mainWindow || mainWindow.isDestroyed()) return false;
  const choice = dialog.showMessageBoxSync(mainWindow, {
    type: 'warning',
    buttons: ['確定', '取消'],
    defaultId: 0,
    cancelId: 1,
    title: 'Gen Doc',
    message,
    noLink: true,
  });
  // Restore OS-level keyboard routing regardless of the user's choice.
  // Without this the next contentEditable mounted in the renderer would not
  // receive typing on Windows.
  mainWindow.focus();
  mainWindow.webContents.focus();
  return choice === 0;
});

// Final ack of the save-and-quit handshake. `ok=true` means the renderer
// actually wrote to disk; `false` means the user cancelled (e.g. backed out
// of the Save-As dialog), in which case we must clear isQuitting so the
// window doesn't tear down with unsaved data on the next close attempt.
ipcMain.on(IPC.app.saveAndQuitResult, (_e, ok: boolean) => {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  if (ok) {
    mainWindow.destroy();
    // Windows/Linux exit via `window-all-closed → app.quit()` once the last
    // window dies, but on macOS that handler is intentionally a no-op so
    // closing the only window leaves the app idling in the dock. The user's
    // explicit "Save and Quit" intent should fully exit on every platform —
    // force it here. `isQuitting` is already true so before-quit's prompt
    // branch is bypassed and we fall straight through to storage cleanup.
    app.quit();
  } else {
    isQuitting = false;
  }
});

app.whenReady().then(async () => {
  await bootStorageHousekeeping();
  registerIpcHandlers();
  // Seed the menu with the persisted recent-files list so it's populated on
  // the very first paint, not only after the user opens/saves something.
  const cfg = await loadConfig();
  buildAppMenu(cfg.recentFiles);
  await createMainWindow();

  app.on('activate', async () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      await createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  // R222 — abort any in-flight Anthropic streams when the renderer goes
  // away. On Windows / Linux this is mostly redundant because the
  // immediate `app.quit()` below cascades through `before-quit` →
  // `app.exit(0)` and the process dies, killing the streams anyway. But
  // on macOS the window-closed-but-app-stays convention means a long
  // chat turn that was streaming when the user closed the only window
  // would otherwise keep iterating the Anthropic SDK's SSE response,
  // calling `send()` into `BrowserWindow.getAllWindows()[0]` (now
  // undefined → no-op), and billing the user for tokens nothing in the
  // app will ever read. cancelAll triggers each AbortController; the
  // SDK's stream iterator throws AbortError, the catch in `chat()` settles
  // back to idle, and the inflight Map empties. Safe to call even when
  // the inflight Map is empty (the for-loop body just doesn't run).
  cancelAllAiStreams();
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', async (event) => {
  // macOS Cmd+Q (and any programmatic app.quit) lands here before the window
  // close cascade fires. If we have unsaved changes and haven't yet shown the
  // prompt, route the quit through `mainWindow.close()` so the window's own
  // 'close' handler can run the Save / Don't Save / Cancel dialog. The user's
  // choice will set `isQuitting` (or destroy the window outright); we then
  // come back here on the next quit attempt and fall through to cleanup.
  if (
    !isQuitting &&
    rendererDirty &&
    mainWindow &&
    !mainWindow.isDestroyed()
  ) {
    event.preventDefault();
    mainWindow.close();
    return;
  }
  // Clean shutdown — best-effort, don't block exit if storage hiccups.
  event.preventDefault();
  await shutdownStorage().catch(() => undefined);
  app.exit(0);
});
