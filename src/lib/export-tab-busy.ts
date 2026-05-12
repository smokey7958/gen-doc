/**
 * R225 — module-level gate shared between the two paths that fire
 * `window.gendoc.workspace.exportTab`:
 *
 *   1. `App.tsx::handleExportTab` — toolbar Download button + Ctrl+E
 *      menu accelerator (was R197's `exportInFlightRef`).
 *   2. `TabBar::exportSingleTab` — per-tab download icon next to the ×
 *      (was the module-local `singleTabExportBusy`).
 *
 * Both ultimately route through main's `invokeExportTab` →
 * `dialog.showSaveDialog`. Two separate gates meant rapid cross-fire
 * (click the toolbar icon AND the per-tab icon, or Ctrl+E + per-tab,
 * within the same hover frame) queued two save IPCs and main stacked
 * two OS save dialogs on the BrowserWindow. The user dismisses one,
 * the next pops on top with the same「匯出為 .md」 / .docx / etc. —
 * exactly the dialog stack-up R207 / R209 / R220 / R221 closed for
 * other entry classes, here at the export granularity.
 *
 * The OS save dialog is an app-wide modal-ish surface (one save can
 * land at a time anyway), so a single global gate is correct rather
 * than per-tab — gating per-tab would still allow stacking two
 * dialogs from different tabs in rapid succession.
 */

let busy = false;

export function tryEnterExportTab(): boolean {
  if (busy) return false;
  busy = true;
  return true;
}

export function exitExportTab(): void {
  busy = false;
}

/**
 * R320 — read-only query so non-IPC entry points (e.g., the BatchExportDialog
 * open-trigger from the toolbar button / Ctrl+Shift+E menu accelerator) can
 * check whether an export is in flight WITHOUT claiming the gate. Without
 * this, the dialog would happily re-open on top of an already-running export
 * — the OS save / folder picker would sit underneath, and the user's
 * subsequent click on the freshly-reopened dialog's confirm button would
 * silently bounce off `tryEnterExportTab()` with no UI feedback.
 */
export function isExportTabBusy(): boolean {
  return busy;
}
