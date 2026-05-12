/**
 * Module-level sync gate for "workspace open" entry points. R199 — every
 * code path that opens a confirm dialog AND/OR an OS file picker on the
 * way to `loadFromOpened()` must claim this gate before its first await,
 * so rapid double-fire (auto-repeat Ctrl+O, accidental double-clicks on
 * a recent menu entry, FileExplorer entry mash) doesn't queue two
 * `dialog.showMessageBoxSync` calls in main and stack two OS-level dialogs
 * on the user's screen.
 *
 * R168 / R187 introduced a sibling counter (`workspace-load-gen.ts`) that
 * orders the *post-IPC* load results — latest wins, others drop. R199 is
 * complementary: it gates the *pre-IPC* OS dialog UI so the user only ever
 * sees one set of confirm + open dialogs at a time. Without this, the
 * race between the renderer's IPC queue and main's `showMessageBoxSync`
 * blocking behavior produces visibly stacked dialogs that are hard to
 * trace back to which Ctrl+O / recent-click triggered which.
 *
 * Usage at every entry point:
 *
 * ```ts
 * if (!tryEnterOpen()) return;
 * try {
 *   if (dirty && !(await confirm(...))) return;
 *   const myGen = bumpLoadGen();
 *   const opened = await window.gendoc.workspace.openPath(filePath);
 *   if (myGen !== currentLoadGen()) return;
 *   if (opened) useWorkspace.getState().loadFromOpened(opened);
 * } finally {
 *   exitOpen();
 * }
 * ```
 *
 * Module-level state is correct (renderer-wide invariant: one user, one
 * dialog stack), and the boolean has no React-reactive consumers — UI
 * disabled-state for menu items / buttons isn't needed because the gate
 * windows are short (confirm dismissal + IPC roundtrip = sub-second).
 */

let busy = false;

export function tryEnterOpen(): boolean {
  if (busy) return false;
  busy = true;
  return true;
}

export function exitOpen(): void {
  busy = false;
}
