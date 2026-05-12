/**
 * R226 — module-level gate for `markdown.exportPdf` IPC re-entry.
 *
 * R196 introduced a component-local `exportPdfBusyRef` inside MarkdownEditor
 * to block rapid double-click on the「輸出為 PDF」 toolbar button. That
 * gate is correct for same-instance double-fire, but doesn't survive a
 * tab switch:
 *
 *   1. User clicks「輸出為 PDF」on a markdown tab. exportPdfBusyRef =
 *      true. main spawns a hidden BrowserWindow, awaits did-finish-load,
 *      printToPDF, fs.writeFile — the full chain takes 1-3 seconds.
 *   2. While the IPC is still in flight, user switches to a non-markdown
 *      tab (.docx / .xlsx / .pptx / unsaved blank). MarkdownEditor unmounts.
 *      The ref dies with the React component instance.
 *   3. User switches BACK to the same markdown tab. MarkdownEditor remounts
 *      with a fresh `useRef(false)`. The toolbar's `exportPdfBusy` React
 *      state is also reset (button enabled).
 *   4. User clicks「輸出為 PDF」 again. exportPdfBusyRef is false → gate
 *      passes → second IPC fires. main spawns ANOTHER hidden window, the
 *      OS now has two save dialogs stacked on the BrowserWindow.
 *
 * Module-level state survives mount/unmount, so the second click in step
 * 4 sees `busy === true` (set in step 1, never released because the IPC
 * is still running) and short-circuits. Released in the finally of
 * `handleExportPdf` — same try/finally as the existing component-local
 * ref already uses, so the release reliably covers cancel / error /
 * success paths.
 *
 * Sibling to:
 *   - lib/workspace-open-busy.ts (R199) — workspace replacement entry points
 *   - lib/tab-close-busy.ts (R221) — per-tab close confirms
 *   - lib/export-tab-busy.ts (R225) — workspace.exportTab IPC
 *
 * Same shape as R225 (single global boolean — only one PDF can be
 * dialog-stacked at a time anyway, the OS save dialog is app-wide
 * modal-ish).
 */

let busy = false;

export function tryEnterExportPdf(): boolean {
  if (busy) return false;
  busy = true;
  return true;
}

export function exitExportPdf(): void {
  busy = false;
}
