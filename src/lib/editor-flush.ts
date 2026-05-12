/**
 * Editor flush registry — bridges debounced editors (PPTX, DOCX) to the
 * save flow.
 *
 * The pptx/docx editors batch re-serialize work behind a 400ms debounce so
 * a burst of keystrokes doesn't re-zip the entire archive on each one.
 * Without this registry, a Ctrl+S issued within that window would call
 * `serializeForSave()` on stale `tab.data` bytes — silently dropping the
 * last burst of edits.
 *
 * Each editor mounts a flush callback that, when invoked, cancels its
 * pending debounce timer and synchronously serialises the latest model
 * into the workspace store. The save flow awaits `flushEditors()` before
 * snapshotting tab bytes; the close-tab flow awaits it before checking
 * `tab.dirty` so an unflushed edit still triggers the unsaved-changes
 * prompt instead of being discarded.
 *
 * Editors register on mount and unregister on unmount, so a tab that
 * isn't currently mounted (inactive in the tab strip) contributes
 * nothing — its bytes already live in the store from when it was last
 * active.
 */

type FlushFn = () => Promise<void>;

const flushers = new Set<FlushFn>();

export function registerEditorFlush(fn: FlushFn): () => void {
  flushers.add(fn);
  return () => {
    flushers.delete(fn);
  };
}

export async function flushEditors(): Promise<void> {
  if (flushers.size === 0) return;
  // Snapshot first — a flush may itself trigger an unmount (e.g. during a
  // close-tab race) which mutates the set mid-iteration.
  const fns = Array.from(flushers);
  await Promise.all(fns.map((f) => f().catch(() => undefined)));
}
