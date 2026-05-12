/**
 * R221 — module-level per-tab close-confirm gate, shared between TabBar's
 * three close entry points (middle-click, ×, context menu) and App.tsx's
 * Ctrl+W keymap.
 *
 * R220 gated the three TabBar sites with a TabBar-component-local Set, but
 * the Ctrl+W keymap lives in App.tsx and didn't see that Set — so a user
 * holding Ctrl+W (auto-repeat) or mashing the keystroke on a dirty tab
 * still queued multiple `app.confirm` IPCs and watched stacked OS dialogs
 * pop one after another, exactly the bug R220 closed for the mouse paths.
 *
 * Promoting the Set to a module makes the invariant renderer-wide: at any
 * moment, each tab id has at most one close-confirm in flight, no matter
 * which entry point started it. Different tabs can still confirm in
 * parallel — the gate is scoped per-tab, not global, because closing tab
 * A and tab B are independent user intents that shouldn't block each
 * other.
 *
 * Usage at every entry point:
 *
 * ```ts
 * if (!tryEnterClose(tabId)) return;
 * try {
 *   if (dirty && !(await app.confirm(...))) return;
 *   await flushEditors();
 *   removeTab(tabId);
 * } finally {
 *   exitClose(tabId);
 * }
 * ```
 *
 * Same shape as workspace-open-busy.ts (R199's renderer-wide gate for
 * workspace replacement) but per-tab granularity instead of global. The
 * two gates are intentionally orthogonal — opening a workspace and
 * closing a tab can happen concurrently without conflict.
 */

const closing = new Set<string>();

export function tryEnterClose(tabId: string): boolean {
  if (closing.has(tabId)) return false;
  closing.add(tabId);
  return true;
}

export function exitClose(tabId: string): void {
  closing.delete(tabId);
}
