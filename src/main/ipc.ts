/**
 * Main-process IPC handlers — wires every channel declared in `src/types/ipc.ts`.
 * Renderer never speaks to the filesystem, OS keystore, or Anthropic SDK
 * directly; it always crosses through here.
 */

import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import {
  IPC,
  type AppInfo,
  type ExportMarkdownPdfRequest,
  type ExportMarkdownPdfResult,
  type ExportTabRequest,
  type ExportTabsRequest,
  type ExportTabsResult,
  type FsEntry,
  type FsFileContent,
  type SaveWorkspaceRequest,
} from '../types/ipc';
import {
  readGdArchive,
  writeGdArchive,
  ensureSessionTempRoot,
  sweepStaleTempRoots,
  cleanupSessionTempRoot,
  atomicWrite,
} from './storage/gd';
import {
  loadConfig,
  patchConfig,
  hasApiKey,
  writeApiKey,
  clearApiKey,
  getConfigDir,
} from './storage/config';
import { rebuildAppMenu } from './menu';
import {
  appendMessage,
  closeDatabase,
  createConversation,
  listConversations,
  listMessages,
  pushUndo,
  popUndo,
  listUndo,
  clearUndo,
  relinkWorkspaceId,
} from './storage/sqlite';
import { chat, cancel, ping } from './ai/anthropic';

export function registerIpcHandlers(): void {
  // ── app ──────────────────────────────────────────────────────────────
  ipcMain.handle(IPC.app.info, async (): Promise<AppInfo> => {
    return {
      version: app.getVersion(),
      platform: process.platform,
      isPackaged: app.isPackaged,
      configDir: getConfigDir(),
    };
  });

  // ── workspace ────────────────────────────────────────────────────────
  ipcMain.handle(IPC.workspace.open, async () => {
    const result = await dialog.showOpenDialog({
      title: '開啟 .gd 檔案',
      filters: [{ name: 'Gen Doc workspace', extensions: ['gd'] }],
      properties: ['openFile'],
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    const opened = await readGdArchive(result.filePaths[0]);
    // R241 — pushRecent is post-open housekeeping; failure must not
    // surface as「開啟失敗」. The .gd is already in `opened` (read
    // succeeded), and the recent-files list is just for the menu's
    // 最近開啟 submenu next-time convenience. Same swallow as the
    // save siblings below — disk write priority over recent-list
    // bookkeeping.
    await pushRecent(result.filePaths[0]).catch(() => undefined);
    return opened;
  });

  ipcMain.handle(IPC.workspace.openPath, async (_e, filePath: string) => {
    const opened = await readGdArchive(filePath);
    // R241 — see workspace.open sibling above for rationale.
    await pushRecent(filePath).catch(() => undefined);
    return opened;
  });

  ipcMain.handle(IPC.workspace.save, async (_e, req: SaveWorkspaceRequest) => {
    if (!req.filePath) {
      const res = await invokeSaveAs(req);
      // R241 — same best-effort posture as R212's `rotateBackups` swallow:
      // pushRecent failures (config.json disk-full / EACCES from antivirus
      // pinning the config dir / patchConfig's chained promise hitting a
      // transient lock) MUST NOT surface to the renderer as「儲存失敗」.
      // The .gd is already on disk via invokeSaveAs / writeGdArchive
      // above; the recent-files list is post-save housekeeping. Without
      // this, the user retries Save, hits the same pushRecent failure
      // again, eventually gives up assuming nothing was saved — and may
      // overwrite the real successful save with re-typed content. Same
      // failure-vs-housekeeping priority order R212 already established
      // for rotateBackups.
      await pushRecent(res.filePath).catch(() => undefined);
      return res;
    }
    const res = await writeGdArchive(req);
    await pushRecent(res.filePath).catch(() => undefined);
    return res;
  });

  ipcMain.handle(IPC.workspace.saveAs, async (_e, req: SaveWorkspaceRequest) => {
    const res = await invokeSaveAs(req);
    // R241 — see workspace.save sibling above for rationale.
    await pushRecent(res.filePath).catch(() => undefined);
    return res;
  });

  ipcMain.handle(IPC.workspace.exportTab, async (_e, req: ExportTabRequest) => {
    return invokeExportTab(req);
  });

  ipcMain.handle(IPC.workspace.exportTabs, async (_e, req: ExportTabsRequest) => {
    return invokeExportTabs(req);
  });

  ipcMain.handle(
    IPC.markdown.exportPdf,
    async (_e, req: ExportMarkdownPdfRequest): Promise<ExportMarkdownPdfResult | null> => {
      return invokeExportMarkdownPdf(req);
    },
  );

  // ── fs (file explorer) ───────────────────────────────────────────────
  ipcMain.handle(IPC.fs.pickDirectory, async (): Promise<string | null> => {
    const result = await dialog.showOpenDialog({
      title: '選擇資料夾',
      properties: ['openDirectory'],
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });

  ipcMain.handle(IPC.fs.listDirectory, async (_e, dirPath: string): Promise<FsEntry[]> => {
    // withFileTypes avoids a per-entry `lstat` for the dir/file flag, but we
    // still need a stat for size + mtime — only attempt it if asked. Skip the
    // stat for entries that fail (broken symlinks etc.) so a single bad file
    // doesn't blow up the whole listing.
    const dirents = await fs.readdir(dirPath, { withFileTypes: true });
    const entries = await Promise.all(
      dirents.map(async (d) => {
        const full = path.join(dirPath, d.name);
        let size = 0;
        let mtime: number | null = null;
        try {
          const st = await fs.stat(full);
          size = d.isDirectory() ? 0 : st.size;
          mtime = st.mtimeMs;
        } catch {
          /* swallow — return entry with default size/mtime. */
        }
        return {
          name: d.name,
          path: full,
          isDirectory: d.isDirectory(),
          size,
          mtime,
        } satisfies FsEntry;
      }),
    );
    // Sort: directories first, then by case-insensitive name. Matches VS Code
    // / Finder defaults so users don't have to hunt for folders.
    entries.sort((a, b) => {
      if (a.isDirectory !== b.isDirectory) return a.isDirectory ? -1 : 1;
      return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
    });
    return entries;
  });

  ipcMain.handle(IPC.fs.readFile, async (_e, filePath: string): Promise<FsFileContent> => {
    const buf = await fs.readFile(filePath);
    const base = path.basename(filePath);
    const dot = base.lastIndexOf('.');
    const ext = dot > 0 ? base.slice(dot + 1).toLowerCase() : '';
    // Buffer is a Uint8Array under the hood — cross the IPC boundary as
    // Uint8Array (electron's structured clone preserves it).
    return { name: base, ext, bytes: new Uint8Array(buf) };
  });

  // ── config ───────────────────────────────────────────────────────────
  ipcMain.handle(IPC.config.get, async () => loadConfig());
  ipcMain.handle(IPC.config.set, async (_e, patch) => {
    const updated = await patchConfig(patch);
    // The menu's 最近開啟 submenu is baked at build time, so any change to
    // the recent-files list (including the renderer-driven "clear all") has
    // to trigger a rebuild here.
    // R242 — two robustness improvements paired with the same R212 / R241
    // best-effort housekeeping posture:
    //   (a) `patch &&` guard before the `in` operator. The `in` operator
    //       throws TypeError on null / undefined; while every renderer
    //       callsite passes an object literal today, the IPC boundary
    //       has no runtime type enforcement, so a malformed payload
    //       (renderer bug, future preload refactor) would crash the
    //       handler and surface as「儲存設定失敗」.
    //   (b) try/catch swallow around rebuildAppMenu. The disk write is
    //       already committed by `await patchConfig(patch)` on the line
    //       above; rebuildAppMenu is post-write housekeeping (just
    //       re-paints the native menu's 最近開啟 submenu). If
    //       Menu.setApplicationMenu throws — corrupted menu template
    //       cache, transient Electron internal state during a window
    //       teardown that races with this IPC, or any future menu
    //       extension that's brittle — the IPC would reject and the
    //       renderer would report the entire setting change as failed,
    //       even though config.json was successfully written. Same
    //       priority order R212 (rotateBackups) / R241 (pushRecent)
    //       established: primary operation success > housekeeping
    //       success. Renderer's caller already dispatched the
    //       gendoc:configChanged event in its own success path, and
    //       the next config change rebuilds the menu fresh — degraded
    //       state is just a stale 最近開啟 submenu until the next
    //       open / save.
    if (patch && 'recentFiles' in patch) {
      try {
        rebuildAppMenu(updated.recentFiles);
      } catch {
        /* swallow — see R242 doc-block above. */
      }
    }
    return updated;
  });
  ipcMain.handle(IPC.config.hasApiKey, async () => hasApiKey());
  ipcMain.handle(IPC.config.setApiKey, async (_e, key: string) => writeApiKey(key));
  ipcMain.handle(IPC.config.clearApiKey, async () => clearApiKey());

  // ── ai ───────────────────────────────────────────────────────────────
  ipcMain.handle(IPC.ai.ping, async (_e, model: string) => ping(model));
  ipcMain.handle(IPC.ai.chat, async (_e, req) => {
    // Fire-and-forget — chunks stream over IPC.ai.chunk. We resolve once
    // the stream is done so the renderer can await the full turn cleanly.
    await chat(req);
  });
  ipcMain.handle(IPC.ai.cancel, async (_e, requestId: string) => cancel(requestId));

  // ── history ──────────────────────────────────────────────────────────
  ipcMain.handle(IPC.history.listConversations, async (_e, workspaceId) =>
    listConversations(workspaceId),
  );
  ipcMain.handle(IPC.history.listMessages, async (_e, conversationId) =>
    listMessages(conversationId),
  );
  ipcMain.handle(IPC.history.appendMessage, async (_e, conversationId, row) =>
    appendMessage(conversationId, row),
  );
  ipcMain.handle(IPC.history.createConversation, async (_e, opts) => createConversation(opts));

  // ── undo ─────────────────────────────────────────────────────────────
  ipcMain.handle(IPC.undo.push, async (_e, entry) => pushUndo(entry));
  ipcMain.handle(IPC.undo.pop, async (_e, workspaceId) => popUndo(workspaceId));
  ipcMain.handle(IPC.undo.list, async (_e, workspaceId, limit) => listUndo(workspaceId, limit));
  ipcMain.handle(IPC.undo.clear, async (_e, workspaceId) => clearUndo(workspaceId));
  ipcMain.handle(IPC.undo.relink, async (_e, oldId, newId) => relinkWorkspaceId(oldId, newId));
}

async function invokeSaveAs(req: SaveWorkspaceRequest) {
  const result = await dialog.showSaveDialog({
    title: '另存新檔',
    defaultPath: req.filePath
      ? path.basename(req.filePath)
      : `${req.manifest.title || '未命名'}.gd`,
    filters: [{ name: 'Gen Doc workspace', extensions: ['gd'] }],
  });
  if (result.canceled || !result.filePath) {
    throw new Error('save_cancelled');
  }
  // R213 — case-insensitive extension check, matching the parity already
  // established by invokeExportTab (line 250 `filePath.toLowerCase()
  // .endsWith(dot)`) and invokeExportMarkdownPdf (line 293
  // `filePath.toLowerCase().endsWith('.pdf')`). Windows' file dialog
  // matches filter extensions case-insensitively but returns the user's
  // typed name verbatim. A user who types「Report.GD」 (or「report.Gd」,
  // or any non-lowercase variant) would slip past the bare `.endsWith
  // ('.gd')` check and get「Report.GD.gd」 written to disk — visibly
  // doubled and ugly in 最近開啟 / OS file manager. The two sibling
  // export paths already use `.toLowerCase().endsWith(...)`; this was
  // the lone holdout. The on-disk file would still open correctly
  // (Windows treats `.gd` and `.GD` as the same extension), but the
  // doubled name lies about what the user named their workspace.
  const filePath = result.filePath.toLowerCase().endsWith('.gd')
    ? result.filePath
    : `${result.filePath}.gd`;
  return writeGdArchive({ ...req, filePath });
}

const EXPORT_FILTERS: Record<ExportTabRequest['ext'], { name: string; extensions: string[] }> = {
  md: { name: 'Markdown', extensions: ['md'] },
  html: { name: 'HTML document', extensions: ['html', 'htm'] },
  docx: { name: 'Word document', extensions: ['docx'] },
  xlsx: { name: 'Excel spreadsheet', extensions: ['xlsx'] },
  pptx: { name: 'PowerPoint presentation', extensions: ['pptx'] },
};

/** Sanitize a tab name so it's safe as a filename on Windows / mac / Linux. */
function safeFileName(name: string, fallback: string): string {
  const trimmed = name.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

async function invokeExportTab(req: ExportTabRequest) {
  const filter = EXPORT_FILTERS[req.ext];
  const base = safeFileName(req.suggestedName, '未命名');
  // If the user-given name already ends with the extension, don't double it up.
  const dot = `.${req.ext}`;
  const defaultName = base.toLowerCase().endsWith(dot) ? base : `${base}${dot}`;
  const result = await dialog.showSaveDialog({
    title: `匯出為 .${req.ext}`,
    defaultPath: defaultName,
    filters: [filter],
  });
  if (result.canceled || !result.filePath) {
    throw new Error('export_cancelled');
  }
  let filePath = result.filePath;
  if (!filePath.toLowerCase().endsWith(dot)) filePath = `${filePath}${dot}`;
  // R251 — atomic write so a partial-write failure doesn't truncate the
  // user's pre-existing target file. Realistic when the user re-exports
  // over an existing .docx / .xlsx / .pptx (common workflow: edit in
  // Gen-Doc → export back to the same path) and the disk fills mid-
  // write, AV yanks the handle, or a network share drops. With raw
  // `fs.writeFile` they'd lose BOTH the prior version and the new
  // export. atomicWrite (write-tmp → fsync → rename → cleanup-on-fail)
  // is the same invariant R191 / R206 established for the .gd save
  // path, now reused here. See atomicWrite doc-block in storage/gd.ts.
  await atomicWrite(filePath, Buffer.from(req.bytes));
  return { filePath };
}

/**
 * Batch tab export — single folder picker, N atomic writes. Filename
 * collisions (two tabs with the same name, OR collision with a file
 * already in the picked folder) are resolved by appending " (2)" / " (3)"
 * etc. before the extension — same convention Windows Explorer uses for
 * "Copy and paste over existing file" → "filename (2).ext". We track
 * names we've already chosen IN this batch so two same-named tabs don't
 * both pick the same disambiguated name.
 *
 * Per-tab write failures are accumulated rather than thrown; the renderer
 * surfaces a per-batch summary toast ("匯出 5 個中 4 個成功，1 個失敗：…").
 * Cancelling the folder picker resolves to null (renderer no-op, no toast).
 */
async function invokeExportTabs(req: ExportTabsRequest): Promise<ExportTabsResult | null> {
  if (req.tabs.length === 0) return null;
  const pick = await dialog.showOpenDialog({
    title: `批次匯出 ${req.tabs.length} 個頁籤到資料夾`,
    properties: ['openDirectory', 'createDirectory'],
    buttonLabel: '匯出到此資料夾',
  });
  if (pick.canceled || pick.filePaths.length === 0) return null;
  const folderPath = pick.filePaths[0];

  // Track names we've assigned in this batch so duplicates inside the
  // request don't collide with each other. Lower-cased because Windows /
  // macOS default filesystems are case-insensitive — "Notes.md" and
  // "notes.md" are the same file on the disk.
  const usedLower = new Set<string>();
  const filePaths: string[] = [];
  const failures: Array<{ index: number; error: string }> = [];

  for (let i = 0; i < req.tabs.length; i++) {
    const tab = req.tabs[i];
    const dot = `.${tab.ext}`;
    const base = safeFileName(tab.suggestedName, '未命名');
    const stripped = base.toLowerCase().endsWith(dot)
      ? base.slice(0, -dot.length)
      : base;
    // Resolve final name via "stripped[ (N)].ext" iteration. Probe the
    // filesystem at each step so we don't clobber a file the user kept
    // in this folder from a previous export round.
    let candidate = `${stripped}${dot}`;
    let n = 2;
    // First, dedupe against names we've already picked in THIS batch.
    while (usedLower.has(candidate.toLowerCase())) {
      candidate = `${stripped} (${n})${dot}`;
      n += 1;
    }
    // Then, dedupe against existing files in the target folder. Loop
    // because `${stripped} (2).ext` itself might already exist.
    // eslint-disable-next-line no-await-in-loop
    while (await pathExists(path.join(folderPath, candidate))) {
      candidate = `${stripped} (${n})${dot}`;
      n += 1;
      if (usedLower.has(candidate.toLowerCase())) {
        // This batch already claimed this disambiguated name too; bump again.
        continue;
      }
    }
    usedLower.add(candidate.toLowerCase());
    const filePath = path.join(folderPath, candidate);
    try {
      // eslint-disable-next-line no-await-in-loop
      await atomicWrite(filePath, Buffer.from(tab.bytes));
      filePaths.push(filePath);
    } catch (err) {
      failures.push({
        index: i,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  return { folderPath, filePaths, failures };
}

async function pathExists(p: string): Promise<boolean> {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * Render the supplied markdown HTML to a PDF on disk.
 *
 * Strategy: spawn a hidden, isolated BrowserWindow, load a self-contained
 * `data:` URL holding the prose-styled markup, wait for it to settle, then
 * call `webContents.printToPDF`. We deliberately do NOT print the renderer's
 * own preview pane — it lives inside the application chrome (toolbar, file
 * tree, AI panel) and would either bleed those into the PDF or require a
 * pile of `@media print` rules that are easy to drift out of date. A clean
 * second window also lets us pick A4 + sensible margins without surprising
 * the user's on-screen view.
 *
 * Cancel semantics: the user backing out of the save dialog returns `null`
 * (the renderer shows nothing). Real failures (write error, printToPDF
 * crash) reject — the renderer surfaces a toast. We avoid the
 * `export_cancelled` Error convention used by `invokeExportTab` because
 * this surface returns a nullable result instead of throwing.
 *
 * The hidden window is `show: false` and built without a preload, sandboxed
 * with javascript disabled — there's no script to run, just a static HTML
 * blob, and shutting off JS removes any chance of user-supplied markdown
 * containing a `<script>` running in main's context. `data:` URLs avoid the
 * temp-file dance and the scheme is treated as same-origin-by-self by
 * Chromium so external resource loads (network images, etc.) still resolve
 * via standard CORS.
 */
async function invokeExportMarkdownPdf(
  req: ExportMarkdownPdfRequest,
): Promise<ExportMarkdownPdfResult | null> {
  const base = safeFileName(req.suggestedName, '未命名');
  const defaultName = base.toLowerCase().endsWith('.pdf') ? base : `${base}.pdf`;
  const result = await dialog.showSaveDialog({
    title: '輸出為 PDF',
    defaultPath: defaultName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (result.canceled || !result.filePath) return null;
  let filePath = result.filePath;
  if (!filePath.toLowerCase().endsWith('.pdf')) filePath = `${filePath}.pdf`;

  const fullHtml = buildPrintableHtml(req.title, req.bodyHtml);
  // Base64 to dodge any URL-encoding edge cases with `#`, `%`, raw newlines
  // etc. inside the markdown body. `data:` URL size is bounded by Chromium
  // at ~2MB; markdown documents that exceed this are vanishingly rare for
  // an editor-driven workflow.
  const dataUrl = 'data:text/html;charset=utf-8;base64,' + Buffer.from(fullHtml, 'utf8').toString('base64');

  const win = new BrowserWindow({
    show: false,
    webPreferences: {
      offscreen: false,
      sandbox: true,
      javascript: false,
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  try {
    // `did-finish-load` fires after the document and its synchronously-loaded
    // resources are done. With JS disabled there's no async layout work to
    // chase, so this is enough to start `printToPDF` against a stable layout.
    await new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Timed out loading print page')), 15000);
      win.webContents.once('did-finish-load', () => {
        clearTimeout(timer);
        resolve();
      });
      win.webContents.once('did-fail-load', (_e, code, desc) => {
        clearTimeout(timer);
        reject(new Error(`did-fail-load ${code}: ${desc}`));
      });
      win.loadURL(dataUrl).catch(reject);
    });
    const pdfBuffer = await win.webContents.printToPDF({
      printBackground: true,
      pageSize: 'A4',
      // Inches per Electron's API. ~1.5cm side, ~2cm vertical — comfortable
      // for body copy without wasting too much page on margins.
      margins: { top: 0.79, bottom: 0.79, left: 0.59, right: 0.59 },
      preferCSSPageSize: false,
      landscape: false,
    });
    // R251 — atomic write, see invokeExportTab sibling. PDF re-export over
    // an existing file is the same overwrite scenario; without atomicWrite
    // a partial-write failure leaves the user with neither the prior PDF
    // nor the new one.
    await atomicWrite(filePath, pdfBuffer);
  } finally {
    // `destroy` over `close` so we don't trip any beforeunload handlers (we
    // don't have any, but defensive is cheap here).
    if (!win.isDestroyed()) win.destroy();
  }

  // Reveal in OS file manager — same convenience as the single-tab export.
  // Wrap in try/catch so a failure here doesn't undo the successful write.
  try {
    shell.showItemInFolder(filePath);
  } catch {
    /* non-fatal */
  }
  return { filePath };
}

/**
 * Wrap the markdown body HTML in a print-ready document. The CSS here is a
 * self-contained mirror of `.markdown-preview` from `src/renderer/index.css`
 * — kept intentionally close so the PDF matches what the user just saw on
 * screen — plus print-specific tweaks (page-break hints, white background,
 * resolved colour values instead of the renderer's `hsl(var(--…))` tokens
 * which only resolve inside the React tree).
 *
 * The body is inserted via a verbatim string concatenation; `marked.parse`
 * already produces safe-ish HTML, and the source is the user's own buffer
 * (no remote/untrusted content) — same XSS posture as the live preview.
 */
function buildPrintableHtml(title: string, bodyHtml: string): string {
  const safeTitle = title.replace(/[<>&]/g, (c) => (c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&amp;'));
  return `<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8" />
<title>${safeTitle}</title>
<style>
  @page { size: A4; margin: 0; }
  html, body { background: #ffffff; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft JhengHei',
      'PingFang TC', 'Noto Sans CJK TC', 'Noto Sans TC', sans-serif;
    color: #1f2328;
    font-size: 13px;
    line-height: 1.7;
    margin: 0;
    padding: 0;
  }
  .markdown-preview {
    max-width: 720px;
    margin: 0 auto;
    padding: 0 4px;
  }
  .markdown-preview h1 { font-size: 1.875rem; font-weight: 700; margin: 1.5em 0 0.6em; line-height: 1.2; border-bottom: 1px solid #d0d7de; padding-bottom: 0.3em; }
  .markdown-preview h2 { font-size: 1.5rem;  font-weight: 700; margin: 1.3em 0 0.5em; line-height: 1.25; border-bottom: 1px solid #d0d7de; padding-bottom: 0.2em; }
  .markdown-preview h3 { font-size: 1.25rem; font-weight: 600; margin: 1.2em 0 0.5em; line-height: 1.3; }
  .markdown-preview h4 { font-size: 1.05rem; font-weight: 600; margin: 1.1em 0 0.4em; }
  .markdown-preview h5 { font-size: 1rem;    font-weight: 600; margin: 1em 0 0.3em; }
  .markdown-preview h6 { font-size: 0.95rem; font-weight: 600; margin: 1em 0 0.3em; color: #656d76; }
  .markdown-preview h1, .markdown-preview h2, .markdown-preview h3,
  .markdown-preview h4, .markdown-preview h5, .markdown-preview h6 {
    page-break-after: avoid;
  }
  .markdown-preview p { margin: 0.6em 0; }
  .markdown-preview a { color: #2563eb; text-decoration: underline; }
  .markdown-preview strong { font-weight: 700; }
  .markdown-preview em { font-style: italic; }
  .markdown-preview ul, .markdown-preview ol { padding-left: 1.5em; margin: 0.6em 0; }
  .markdown-preview ul { list-style: disc; }
  .markdown-preview ol { list-style: decimal; }
  .markdown-preview li { margin: 0.2em 0; }
  .markdown-preview li > input[type="checkbox"] { margin-right: 0.4em; }
  .markdown-preview blockquote { border-left: 3px solid #d0d7de; padding-left: 0.9em; color: #656d76; margin: 0.6em 0; }
  .markdown-preview code { background: #f1f3f5; padding: 0.1em 0.35em; border-radius: 4px; font-size: 0.9em; font-family: ui-monospace, 'SFMono-Regular', 'Menlo', 'Consolas', monospace; }
  .markdown-preview pre { background: #f6f8fa; padding: 0.8em 1em; border-radius: 6px; font-size: 0.85em; line-height: 1.55; margin: 0.8em 0; white-space: pre-wrap; word-break: break-word; page-break-inside: avoid; }
  .markdown-preview pre code { background: transparent; padding: 0; border-radius: 0; font-size: inherit; }
  .markdown-preview hr { border: none; border-top: 1px solid #d0d7de; margin: 1.5em 0; }
  .markdown-preview img { max-width: 100%; height: auto; border-radius: 4px; }
  .markdown-preview table { border-collapse: collapse; width: 100%; margin: 0.8em 0; font-size: 0.92em; page-break-inside: avoid; }
  .markdown-preview th, .markdown-preview td { border: 1px solid #d0d7de; padding: 0.4em 0.7em; }
  .markdown-preview th { background: #f6f8fa; font-weight: 600; }
  .markdown-preview del { text-decoration: line-through; color: #656d76; }
</style>
</head>
<body>
<div class="markdown-preview">
${bodyHtml}
</div>
</body>
</html>`;
}

/**
 * Add a `.gd` path to the recent-files list. De-dupes (if the path is already
 * in the list, it floats to the top) and caps at 10 entries — beyond that the
 * menu becomes unwieldy and old entries are unlikely to still exist on disk.
 * Triggers a menu rebuild so the new entry shows up immediately under
 * File → 最近開啟.
 */
async function pushRecent(filePath: string): Promise<void> {
  const cfg = await loadConfig();
  const next = [filePath, ...cfg.recentFiles.filter((p) => p !== filePath)].slice(0, 10);
  if (next.length === cfg.recentFiles.length && next.every((p, i) => p === cfg.recentFiles[i])) {
    return; // already at the top, nothing changed
  }
  const updated = await patchConfig({ recentFiles: next });
  rebuildAppMenu(updated.recentFiles);
}

/** Run on app startup: housekeeping for temp dirs from prior crashes. */
export async function bootStorageHousekeeping(): Promise<void> {
  await ensureSessionTempRoot();
  await sweepStaleTempRoots();
  await fs.mkdir(getConfigDir(), { recursive: true });
}

export async function shutdownStorage(): Promise<void> {
  // R194 — close the SQLite handle BEFORE removing the temp root.
  // On Windows the open handle blocks `fs.rm` with EBUSY; the
  // .catch in cleanupSessionTempRoot would swallow the error,
  // leaving chat.sqlite (+ -wal / -shm) plus the whole gendoc-XXX
  // temp dir behind for sweepStaleTempRoots's 7-day TTL to finally
  // claim. closeDatabase is sync + idempotent.
  closeDatabase();
  await cleanupSessionTempRoot();
}
