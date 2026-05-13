/**
 * Native application menu. Most actions are forwarded to the renderer over
 * dedicated channels so the React side owns the actual workspace state.
 *
 * R405 — bilingual menu. `buildAppMenu(recentFiles, locale)` takes an explicit
 * locale ('zh' | 'en') and translates every visible label. The renderer's
 * language selector + config.set({locale}) triggers rebuildAppMenu so the
 * native menu re-renders without an app restart. Inline `t(zh, en)` keeps
 * the original Traditional Chinese visible as living source of truth and
 * pairs it with English next to the call site — same idiom as lib/i18n.ts.
 */

import path from 'node:path';
import { app, Menu, type MenuItemConstructorOptions, BrowserWindow } from 'electron';

const isMac = process.platform === 'darwin';

function send(channel: string, ...args: unknown[]): void {
  const win = BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0];
  if (!win) return;
  win.webContents.send(channel, ...args);
}

/**
 * Compress a long absolute path for menu display: keep the parent directory
 * + filename so users can disambiguate "report.gd" entries living in
 * different folders. Anything deeper is collapsed.
 *
 * R398 — escape `&` in the returned string. Electron's MenuItem.label treats
 * a single `&` as a mnemonic accelerator hint: the character right after
 * the `&` gets underlined (Windows / Linux: also bound to Alt-<char>; Mac:
 * silently stripped) and the `&` itself is consumed. A user whose recent
 * file path contains `&` — folder「Q&A」, file「foo & bar.gd」, OneDrive's
 * `&` in the synced root name (`OneDrive - Company & Co`), legitimate
 * Chinese / English-mixed names — would see the menu entry rendered with
 * the `&` missing AND an accidental Alt-<next-char> binding pointing at
 * the recent file instead of whatever the user's app conventionally uses
 * that shortcut for. The tooltip (line 41 `toolTip: p`) carries the raw
 * path and is unaffected — only the menu label needs the escape because
 * that's the surface that goes through Electron's mnemonic parser.
 *
 * Doubling `&` → `&&` is the canonical escape per Electron's MenuItem
 * docs ("To display a literal &, use && instead"). The literal `&&`
 * renders as a single `&` cross-platform — same intent, just opted in
 * via the official escape.
 */
function displayPath(p: string): string {
  const parts = p.split(/[\\/]/);
  const trimmed = parts.length <= 2 ? p : `…${path.sep}${parts.slice(-2).join(path.sep)}`;
  return trimmed.replace(/&/g, '&&');
}

type Locale = 'zh' | 'en';

/** Bilingual label helper — same idiom as renderer's lib/i18n.ts:t(). */
function t(locale: Locale, zh: string, en: string): string {
  return locale === 'zh' ? zh : en;
}

/**
 * Build (or rebuild) the app menu. `recentFiles` is the most-recent-first
 * list of `.gd` paths that powers the File → 最近開啟 submenu. We rebuild
 * (not patch) because Electron menus are immutable once installed; the
 * caller invokes `rebuildAppMenu` after any open/save or locale change.
 */
export function buildAppMenu(recentFiles: string[] = [], locale: Locale = 'en'): void {
  const tt = (zh: string, en: string) => t(locale, zh, en);
  const recentSubmenu: MenuItemConstructorOptions[] =
    recentFiles.length === 0
      ? [{ label: tt('（無）', '(None)'), enabled: false }]
      : [
          ...recentFiles.slice(0, 10).map<MenuItemConstructorOptions>((p) => ({
            label: displayPath(p),
            toolTip: p,
            click: () => send('menu:openRecent', p),
          })),
          { type: 'separator' },
          {
            label: tt('清除最近開啟', 'Clear recent'),
            click: () => send('menu:clearRecent'),
          },
        ];

  const template: MenuItemConstructorOptions[] = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [
              { role: 'about' as const },
              { type: 'separator' as const },
              { role: 'services' as const },
              { type: 'separator' as const },
              { role: 'hide' as const },
              { role: 'hideOthers' as const },
              { role: 'unhide' as const },
              { type: 'separator' as const },
              { role: 'quit' as const },
            ],
          },
        ]
      : []),
    {
      label: tt('檔案', 'File'),
      submenu: [
        {
          label: tt('新專案', 'New Project'),
          accelerator: 'CmdOrCtrl+N',
          click: () => send('menu:newProject'),
        },
        {
          label: tt('開啟…', 'Open…'),
          accelerator: 'CmdOrCtrl+O',
          click: () => send('menu:open'),
        },
        {
          label: tt('開啟資料夾…', 'Open Folder…'),
          accelerator: 'CmdOrCtrl+K CmdOrCtrl+O',
          click: () => send('menu:openFolder'),
        },
        {
          label: tt('最近開啟', 'Open Recent'),
          submenu: recentSubmenu,
        },
        { type: 'separator' },
        {
          label: tt('儲存', 'Save'),
          accelerator: 'CmdOrCtrl+S',
          click: () => send('menu:save'),
        },
        {
          label: tt('另存新檔…', 'Save As…'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => send('menu:saveAs'),
        },
        { type: 'separator' },
        {
          label: tt('匯出目前頁籤…', 'Export Current Tab…'),
          accelerator: 'CmdOrCtrl+E',
          click: () => send('menu:exportTab'),
        },
        {
          label: tt('批次匯出多個頁籤…', 'Batch Export Tabs…'),
          accelerator: 'CmdOrCtrl+Shift+E',
          click: () => send('menu:batchExport'),
        },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: tt('編輯', 'Edit'),
      submenu: [
        // `registerAccelerator: false` shows the "Ctrl+Z" hint in the menu
        // label but does NOT bind the keystroke — without this, the menu
        // layer hijacks Ctrl+Z everywhere (including inside the AI prompt
        // textarea, rename inputs, and CodeMirror) and unconditionally fires
        // workspace-level AI changeset undo, surprising users who pressed
        // Ctrl+Z to undo their own typing. Keyboard ownership lives on the
        // focus-aware window handler in App.tsx, which correctly skips when
        // focus is inside any editor that owns its own undo stack. Clicking
        // the menu item still works (intentional invocation by the user).
        {
          label: tt('復原', 'Undo'),
          accelerator: 'CmdOrCtrl+Z',
          registerAccelerator: false,
          click: () => send('menu:undo'),
        },
        {
          label: tt('重做', 'Redo'),
          accelerator: 'CmdOrCtrl+Shift+Z',
          registerAccelerator: false,
          click: () => send('menu:redo'),
        },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
      ],
    },
    {
      label: tt('檢視', 'View'),
      submenu: [
        // Reload / forceReload / toggleDevTools are dev-only. In a packaged
        // build they're a user-facing footgun: clicking 重新整理 nukes
        // every unsaved tab without firing the window's close prompt
        // (reload doesn't close the window, it just navigates the
        // renderer). main.ts already blocks F5 / Ctrl+R via
        // before-input-event for the same reason — leaving the menu
        // entries in production silently undermined that guard. Devtools
        // hidden too so end users see a clean View menu (developers can
        // still reach it programmatically or via the dev build).
        ...(app.isPackaged
          ? []
          : ([
              { role: 'reload' as const },
              { role: 'forceReload' as const },
              { role: 'toggleDevTools' as const },
              { type: 'separator' as const },
            ] satisfies MenuItemConstructorOptions[])),
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
    {
      label: 'AI',
      submenu: [
        {
          label: tt('聚焦 AI 對話框', 'Focus AI Input'),
          accelerator: 'CmdOrCtrl+L',
          click: () => send('menu:focusAI'),
        },
        {
          label: tt('設定…', 'Settings…'),
          accelerator: 'CmdOrCtrl+,',
          click: () => send('menu:openSettings'),
        },
      ],
    },
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/** Convenience wrapper used by ipc.ts after a recent-files / locale mutation. */
export function rebuildAppMenu(recentFiles: string[], locale: Locale = 'en'): void {
  buildAppMenu(recentFiles, locale);
}
