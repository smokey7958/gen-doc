/**
 * Typed contract for the renderer ↔ main IPC bridge. The preload script
 * exposes this surface as `window.gendoc` (see {@link GenDocBridge}).
 *
 * All functions are async because they cross the IPC boundary; even simple
 * getters return Promises so that the bridge is uniform.
 */

import type { Manifest, TabDescriptor } from './manifest';

/** Raw bytes of a single file inside a `.gd` archive. */
export interface TabPayload {
  descriptor: TabDescriptor;
  bytes: Uint8Array;
}

/** What we hand back to the renderer when a `.gd` is opened. */
export interface OpenedWorkspace {
  /** Absolute path to the on-disk `.gd` file. */
  filePath: string;
  manifest: Manifest;
  tabs: TabPayload[];
}

/** Paired with `saveWorkspace`: full workspace state from renderer → main. */
export interface SaveWorkspaceRequest {
  /** Absolute path. If empty/undefined, main will prompt Save-As. */
  filePath?: string;
  manifest: Manifest;
  tabs: TabPayload[];
}

export interface SaveWorkspaceResult {
  filePath: string;
  modifiedAt: string;
}

/**
 * Export a single tab as its native format file (.md / .html / .docx /
 * .xlsx / .pptx). Distinct from `save` which writes the whole .gd archive
 * — this lets the user hand off one document for use outside Gen Doc.
 */
export interface ExportTabRequest {
  /** File extension without the dot, lower-case ("md", "html", "docx", "xlsx", "pptx"). */
  ext: 'md' | 'html' | 'docx' | 'xlsx' | 'pptx';
  /** Suggested filename (without extension). */
  suggestedName: string;
  // R361 — comment now lists every text-format that uses this path. Was「For
  // markdown tabs, UTF-8 encoded text」 which dated to before HTML support
  // (R336+/R337) and was inconsistent with the sibling ExportTabsRequest
  // doc-comment below (line 74) that already says「markdown / html」.
  // Caller sites at App.tsx::handleExportTab and TabBar::exportSingleTab
  // both branch on markdown OR html → TextEncoder.encode, matching this
  // updated description.
  /** Raw bytes to write. For markdown / html tabs, UTF-8 encoded text;
   *  for docx / xlsx / pptx, the raw OOXML zip bytes. */
  bytes: Uint8Array;
}

export interface ExportTabResult {
  filePath: string;
}

/**
 * Batch tab export — user prompted ONCE for a destination folder; main writes
 * every tab in `tabs` into that folder, deduping filename collisions by
 * appending " (2)" / " (3)" … before the extension (same convention Windows
 * Explorer uses for "Copy and paste" collisions). Distinct from
 * `exportTab` because (a) the UX is a folder picker (not a Save-As per
 * file — N stacked dialogs would be unusable for N > 2), (b) collisions
 * happen across N writes and must be handled in main where the existing
 * filesystem is known, and (c) partial failure reporting is meaningful
 * (e.g., "wrote 4 of 5, disk full on the last one") in a way the
 * single-tab API doesn't need.
 */
export interface ExportTabsRequest {
  tabs: Array<{
    /** File extension without the dot — same union as ExportTabRequest. */
    ext: 'md' | 'html' | 'docx' | 'xlsx' | 'pptx';
    /** Suggested filename (without extension). */
    suggestedName: string;
    /** Raw bytes to write. UTF-8 encoded text for markdown / html. */
    bytes: Uint8Array;
  }>;
}

export interface ExportTabsResult {
  /** Absolute path of the folder the user picked. */
  folderPath: string;
  /** Absolute file paths actually written, in input order. */
  filePaths: string[];
  /** Per-tab failure detail (index into the original `tabs` array + reason).
   *  An empty array means "every tab written cleanly". */
  failures: Array<{ index: number; error: string }>;
}

/**
 * Render the Markdown preview pane to PDF. Renderer hands across the already-
 * rendered HTML (the same string `marked.parse` produced for live preview),
 * plus a `<style>` block, and main spawns an offscreen BrowserWindow to drive
 * `webContents.printToPDF`. We pass HTML rather than the raw markdown so the
 * main process never needs to know about marked / the renderer's CSS — keeps
 * the surface tiny and avoids shipping a second markdown parser.
 */
export interface ExportMarkdownPdfRequest {
  /** Suggested filename without extension. */
  suggestedName: string;
  /** Document title — used for the PDF metadata + window title. */
  title: string;
  /** Already-rendered markdown HTML (output of `marked.parse`). */
  bodyHtml: string;
}

export interface ExportMarkdownPdfResult {
  filePath: string;
}

export interface PingResult {
  ok: boolean;
  /** When ok=false, a short reason for the UI. */
  error?: string;
  /** Echoed model name on success. */
  model?: string;
}

/**
 * One entry inside a directory listed by `fs:listDirectory`. Carries the
 * absolute path so the renderer doesn't have to re-join — different OSes
 * use different separators and we already know it on the main side.
 */
export interface FsEntry {
  name: string;
  /** Absolute path, OS-native separators. */
  path: string;
  isDirectory: boolean;
  /** Size in bytes for files; 0 for directories. */
  size: number;
  /** Last-modified epoch ms; null if stat failed. */
  mtime: number | null;
}

/** Result of `fs:readFile` — the renderer needs both bytes and a name to derive a tab. */
export interface FsFileContent {
  name: string;
  /** Lower-case extension without the dot, e.g. "md", "xlsx". Empty for files without an ext. */
  ext: string;
  bytes: Uint8Array;
}

export interface AppInfo {
  version: string;
  platform: NodeJS.Platform;
  isPackaged: boolean;
  /** Where the user-level config lives (~/.gendoc on *nix). */
  configDir: string;
}

/** Persisted user config (non-secret). API key is stored separately. */
export interface UserConfig {
  defaultModel: string;
  temperature: number;
  maxTokens: number;
  promptCache: boolean;
  embedChatHistoryDefault: boolean;
  theme: 'light' | 'dark' | 'system';
  /** ms; 0 disables. */
  autoSaveIntervalMs: number;
  recentFiles: string[];
  keymapOverrides: Record<string, string>;
  /**
   * Last-known main-window bounds. Restored on next launch so the user's
   * sized/positioned window persists. Null on first launch — main falls back
   * to the default 1440×900.
   */
  windowBounds: { x: number; y: number; width: number; height: number } | null;
  /** Whether to re-open the most recent .gd workspace on launch. */
  autoOpenLastWorkspace: boolean;
  /**
   * R405 — explicit UI-language preference. `null` = follow OS locale
   * (zh-* OS → 'zh', everything else → 'en'). Persisted via Settings dialog
   * or the toolbar language selector. Affects renderer-side React strings
   * via `useT()` hook and main-process app menu via rebuildAppMenu(locale).
   */
  locale: 'zh' | 'en' | null;
}

/** Stored chat-history row, returned to renderer for replay. */
export interface ChatRow {
  id: string;
  conversationId: string;
  role: 'user' | 'assistant' | 'tool_result';
  /** JSON string of {@link import('./ai').ContentBlock}[]. */
  content: string;
  toolUseId: string | null;
  createdAt: number;
  tokenInput: number | null;
  tokenOutput: number | null;
  cacheRead: number | null;
  cacheCreation: number | null;
}

export interface PersistChatRow extends Omit<ChatRow, 'id' | 'createdAt' | 'conversationId'> {
  id?: string;
  createdAt?: number;
  /** Optional — main also accepts the conversationId as the second argument. */
  conversationId?: string;
}

export interface UndoRow {
  id: string;
  changesetJson: string;
  appliedAt: number;
  workspaceId: string | null;
}

/**
 * The strongly-typed bridge object exposed on `window.gendoc`.
 * Implemented by `src/preload.ts` and consumed throughout the renderer.
 */
export interface GenDocBridge {
  app: {
    info(): Promise<AppInfo>;
    /**
     * Push the renderer's dirty flag to main so the OS-level window-close
     * handler can prompt the user before destroying unsaved work. Fire-and-
     * forget — we don't await an ack because the close path needs the latest
     * value synchronously when the user clicks the X.
     */
    setDirty(dirty: boolean): void;
    /**
     * Show a native modal Yes/Cancel confirmation dialog from the main
     * process. Used as a drop-in replacement for renderer-side
     * `window.confirm()`, which on Windows leaves the BrowserWindow in a
     * half-focused state: `document.activeElement` ends up correct but
     * `document.hasFocus()` is false, so subsequent keystrokes (typing into
     * a freshly-mounted contentEditable / new tab) are not routed to
     * webContents. The main-process dialog handles focus restoration
     * cleanly and the handler additionally calls
     * `mainWindow.focus() + webContents.focus()` after the user dismisses
     * the dialog. Returns true on confirm, false on cancel.
     *
     * Used for every dirty-prompt and destructive-action confirmation in
     * the renderer (open/save discard, tab close, slide / sheet / textbox
     * delete, layout-apply that wipes content). Default buttons are
     * 確定 / 取消 — sufficient for all current callers.
     */
    confirm(message: string): Promise<boolean>;
    /**
     * Final signal in the save-and-quit handshake. Main intercepts a window
     * close on a dirty workspace, prompts Save/Don't Save/Cancel, and on
     * "Save" sends `menu:saveAndQuit` to the renderer. The renderer runs its
     * save flow then calls this with `true` on success → main tears the
     * window down. `false` (e.g. user cancelled the Save-As dialog) lets
     * main re-arm the prompt for the next close attempt instead of leaving
     * the window in a half-quit state where data could later be lost.
     */
    saveAndQuitResult(ok: boolean): void;
    /**
     * R405 — return the OS locale string from Chromium (e.g. 'zh-TW',
     * 'en-US', 'ja-JP'). Used at boot to pick a default UI language when
     * UserConfig.locale is null (= "follow OS"). The renderer maps anything
     * starting with "zh" to its 'zh' locale, everything else to 'en'.
     */
    getOsLocale(): Promise<string>;
  };
  workspace: {
    open(): Promise<OpenedWorkspace | null>;
    openPath(filePath: string): Promise<OpenedWorkspace | null>;
    save(req: SaveWorkspaceRequest): Promise<SaveWorkspaceResult>;
    saveAs(req: SaveWorkspaceRequest): Promise<SaveWorkspaceResult>;
    /** Write a single tab to disk in its native format. Throws "export_cancelled" if user backs out. */
    exportTab(req: ExportTabRequest): Promise<ExportTabResult>;
    /**
     * Batch-export multiple tabs into a single user-picked folder. Resolves
     * to null if the user cancelled the folder picker. Per-tab write
     * failures appear in the `failures` array — overall promise still
     * resolves so partial successes are visible (renderer surfaces a
     * summary toast).
     */
    exportTabs(req: ExportTabsRequest): Promise<ExportTabsResult | null>;
  };
  markdown: {
    /**
     * Render the supplied HTML body to a PDF on disk. Resolves with the
     * chosen path on success, or `null` if the user cancels the save dialog.
     * Other failures (write error, printToPDF crash) reject — callers should
     * surface a toast.
     */
    exportPdf(req: ExportMarkdownPdfRequest): Promise<ExportMarkdownPdfResult | null>;
  };
  /**
   * Filesystem operations backing the IDE-style file explorer. We deliberately
   * keep this small — no write / delete / rename — because the explorer is
   * read-only browse + open in MVP; mutations stay scoped to the active tab
   * via the workspace API.
   */
  fs: {
    /** Show OS folder picker. Returns the chosen absolute path, or null on cancel. */
    pickDirectory(): Promise<string | null>;
    /** List one level of `dirPath`. Throws on permission / not-found. */
    listDirectory(dirPath: string): Promise<FsEntry[]>;
    /** Read a single file's bytes + derived name/ext. Used to open into a tab. */
    readFile(filePath: string): Promise<FsFileContent>;
  };
  config: {
    get(): Promise<UserConfig>;
    set(patch: Partial<UserConfig>): Promise<UserConfig>;
    /** Returns true if a key is currently stored (without revealing it). */
    hasApiKey(): Promise<boolean>;
    setApiKey(key: string): Promise<void>;
    clearApiKey(): Promise<void>;
  };
  ai: {
    /** Validates the stored key with a tiny ping to Claude. */
    ping(model: string): Promise<PingResult>;
    /** Streams a chat turn; chunks arrive via the `onChunk` channel. */
    chat(req: {
      requestId: string;
      model: string;
      system: string;
      messages: unknown[];
      tools: unknown[];
      maxTokens: number;
      temperature: number;
      cacheBreakpoints: number[];
      /** When false, suppress all `cache_control` markers (system + tools + per-
       * turn breakpoints) so Anthropic bills every token at non-cache rates.
       * Mirrors UserConfig.promptCache; previously orphaned (config persisted
       * but never reached this IPC payload), see store/ai.ts:46-53 for the
       * same-shape fix already in place for maxTokens / temperature. */
      promptCache: boolean;
    }): Promise<void>;
    /** Cancels an in-flight stream. */
    cancel(requestId: string): Promise<void>;
    onChunk(handler: (chunk: { requestId: string; chunk: unknown }) => void): () => void;
  };
  history: {
    listConversations(workspaceId: string | null): Promise<
      Array<{ id: string; title: string; createdAt: number; updatedAt: number }>
    >;
    listMessages(conversationId: string): Promise<ChatRow[]>;
    appendMessage(conversationId: string, row: PersistChatRow): Promise<ChatRow>;
    createConversation(opts: {
      id?: string;
      title: string;
      workspaceId: string | null;
    }): Promise<{ id: string; createdAt: number; updatedAt: number }>;
  };
  undo: {
    push(entry: { changesetJson: string; workspaceId: string | null }): Promise<UndoRow>;
    pop(workspaceId: string | null): Promise<UndoRow | null>;
    /** For redo: take the last popped entry and re-push it. */
    list(workspaceId: string | null, limit: number): Promise<UndoRow[]>;
    clear(workspaceId: string | null): Promise<void>;
    /**
     * R386 — re-tag undo_entries and conversations rows whose workspace_id is
     * `oldId` to `newId`. Called from performSave when a Save-As mints a new
     * filePath hash (workspaceIdFor changes), so pre-Save-As undo history and
     * chat conversations follow the workspace into its new identity instead of
     * orphaning under the old id.
     */
    relink(oldId: string | null, newId: string): Promise<void>;
  };
  /**
   * Slim wrapper around Electron's `webUtils` so the renderer can resolve
   * dropped-File objects to absolute paths. Electron 32+ removed `File.path`
   * — this is the only supported way. Synchronous because `webUtils.getPathForFile`
   * is itself sync.
   */
  webUtils: {
    getPathForFile(file: File): string;
  };
}

declare global {
  interface Window {
    gendoc: GenDocBridge;
  }
}

/** IPC channel names — single source of truth shared by main + preload. */
export const IPC = {
  app: {
    info: 'app:info',
    setDirty: 'app:setDirty',
    confirm: 'app:confirm',
    saveAndQuitResult: 'app:saveAndQuitResult',
    getOsLocale: 'app:getOsLocale',
  },
  workspace: {
    open: 'workspace:open',
    openPath: 'workspace:openPath',
    save: 'workspace:save',
    saveAs: 'workspace:saveAs',
    exportTab: 'workspace:exportTab',
    exportTabs: 'workspace:exportTabs',
  },
  markdown: {
    exportPdf: 'markdown:exportPdf',
  },
  fs: {
    pickDirectory: 'fs:pickDirectory',
    listDirectory: 'fs:listDirectory',
    readFile: 'fs:readFile',
  },
  config: {
    get: 'config:get',
    set: 'config:set',
    hasApiKey: 'config:hasApiKey',
    setApiKey: 'config:setApiKey',
    clearApiKey: 'config:clearApiKey',
  },
  ai: {
    ping: 'ai:ping',
    chat: 'ai:chat',
    cancel: 'ai:cancel',
    chunk: 'ai:chunk',
  },
  history: {
    listConversations: 'history:listConversations',
    listMessages: 'history:listMessages',
    appendMessage: 'history:appendMessage',
    createConversation: 'history:createConversation',
  },
  undo: {
    push: 'undo:push',
    pop: 'undo:pop',
    list: 'undo:list',
    clear: 'undo:clear',
    relink: 'undo:relink',
  },
} as const;
