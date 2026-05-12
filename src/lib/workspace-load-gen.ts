/**
 * Module-level monotonic counter that orders all workspace-replacement
 * operations across the renderer. Every code path that resolves to
 * `loadFromOpened()` / `newWorkspace()` must capture its own gen at the
 * sync entry point and re-check before mutating the store; the latest
 * call wins, older ones drop on the floor.
 *
 * R168 / R169 introduced this counter as a `useRef` inside `App.tsx` and
 * threaded it through 5 entry points (auto-open, handleOpen,
 * handleOpenRecent, drag-drop .gd, handleNew). R187 moves it to a shared
 * module so FileExplorer's own .gd-open path (`useOpenFile` at
 * FileExplorer.tsx:670) — which is a 6th entry point that's outside
 * App.tsx's component scope — can join the same exclusion class without
 * prop drilling or context plumbing. Module-level state is correct here:
 * the counter has app-wide semantics (only one workspace is active at a
 * time, regardless of which component triggered the load), and there's no
 * App.tsx unmount to worry about (it's the root component).
 *
 * Usage at every entry point:
 *
 * ```ts
 * const myGen = bumpLoadGen();
 * const opened = await window.gendoc.workspace.openPath(target);
 * if (myGen !== currentLoadGen()) return;          // newer load won
 * if (opened) useWorkspace.getState().loadFromOpened(opened);
 * ```
 *
 * For sync entries (`handleNew`'s `newWorkspace()` call), only `bumpLoadGen`
 * is needed — it invalidates any in-flight async load whose myGen is now
 * stale.
 */

let counter = 0;

export function bumpLoadGen(): number {
  return ++counter;
}

export function currentLoadGen(): number {
  return counter;
}
