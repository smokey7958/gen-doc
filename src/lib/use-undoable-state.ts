/**
 * Editor-local undo/redo hook.
 *
 * Why editor-local rather than App-level:
 *   The Docx / Xlsx / Pptx editors each maintain their own model state
 *   (`useState<DocxModel>` etc.) in parallel with the workspace's bytes
 *   buffer. App-level undo would require either re-parsing bytes back into
 *   each editor's local shape on every undo, or re-architecting state
 *   ownership. A ring buffer of model snapshots inside the editor sidesteps
 *   that — undo just swaps the model and the existing debounced
 *   serialize-to-bytes effect picks up the new value automatically.
 *
 * Coalescing strategy:
 *   Pushing one snapshot per setState call would make every keystroke a
 *   separate undo step ("hello" → 5 undos to clear). Instead we coalesce
 *   consecutive mutations within a 500 ms window — the first edit in a
 *   window pushes the *previous* state to the past stack, and subsequent
 *   edits within the same window only update `present`. After 500 ms of
 *   inactivity the next edit opens a new window and pushes again. That
 *   matches PowerPoint / Word's "one word ≈ one undo step" behaviour
 *   without any per-call commit boundary work at the call site.
 *
 *   Structural ops (insert / delete block, reorder slide, etc.) usually
 *   pass through the same setState path and benefit from the same window —
 *   if the user pastes ten blocks in 200 ms they collapse into one undo
 *   step, which matches user expectation.
 *
 * The hook also exposes `undo()` / `redo()` callbacks that callers wire to
 * Ctrl/Cmd+Z and Ctrl/Cmd+Y / Ctrl/Cmd+Shift+Z. We don't bind keys here so
 * each editor can scope the shortcut to its own DOM root (same pattern as
 * `use-format-shortcuts.ts`).
 */

import { useCallback, useEffect, useRef, useState } from 'react';

interface UndoableApi<T> {
  undo: () => void;
  redo: () => void;
  /** True when there is at least one snapshot to roll back to. */
  canUndo: boolean;
  /** True when an undo has happened and a redo is available. */
  canRedo: boolean;
  /** Replace the model without touching the undo history (e.g., after a
   *  fresh re-parse from disk). Resets past / future stacks. */
  resetHistory: (next: T) => void;
}

const COALESCE_MS = 500;
const MAX_HISTORY = 100;

/**
 * Drop-in replacement for `useState<T>` that records changes for undo/redo.
 *
 * Returns the same `[state, setState]` shape as `useState`, plus an `api`
 * object with undo / redo / canUndo / canRedo / resetHistory.
 */
export function useUndoableState<T>(
  initial: T | (() => T),
): [T, (updater: T | ((prev: T) => T)) => void, UndoableApi<T>] {
  const [state, setState] = useState<T>(initial);
  const past = useRef<T[]>([]);
  const future = useRef<T[]>([]);
  const lastPushTime = useRef<number>(0);
  // Tick state so canUndo / canRedo (which depend on ref-backed stacks)
  // still trigger a re-render when the stacks change.
  const [, setTick] = useState(0);
  const bump = useCallback(() => setTick((v) => v + 1), []);

  const set = useCallback(
    (updater: T | ((prev: T) => T)) => {
      setState((prev) => {
        const next =
          typeof updater === 'function' ? (updater as (p: T) => T)(prev) : updater;
        if (Object.is(next, prev)) return prev;
        const now = Date.now();
        // Open a new "edit window" on first change after the coalesce gap;
        // within the window we keep updating `present` without growing past.
        if (now - lastPushTime.current > COALESCE_MS) {
          past.current.push(prev);
          if (past.current.length > MAX_HISTORY) past.current.shift();
          // Any fresh edit invalidates the redo branch.
          if (future.current.length > 0) future.current = [];
          bump();
        }
        lastPushTime.current = now;
        return next;
      });
    },
    [bump],
  );

  const undo = useCallback(() => {
    setState((prev) => {
      const last = past.current.pop();
      if (last === undefined) return prev;
      future.current.push(prev);
      // Reset window so the next user edit starts a fresh push.
      lastPushTime.current = 0;
      bump();
      return last;
    });
  }, [bump]);

  const redo = useCallback(() => {
    setState((prev) => {
      const next = future.current.pop();
      if (next === undefined) return prev;
      past.current.push(prev);
      lastPushTime.current = 0;
      bump();
      return next;
    });
  }, [bump]);

  const resetHistory = useCallback(
    (next: T) => {
      past.current = [];
      future.current = [];
      lastPushTime.current = 0;
      setState(next);
      bump();
    },
    [bump],
  );

  const api: UndoableApi<T> = {
    undo,
    redo,
    canUndo: past.current.length > 0,
    canRedo: future.current.length > 0,
    resetHistory,
  };

  return [state, set, api];
}

/**
 * Bind Ctrl/Cmd+Z (undo) and Ctrl/Cmd+Y / Ctrl/Cmd+Shift+Z (redo) scoped to
 * a particular editor root. Mirrors `useFormatShortcuts` — only fires when
 * focus is inside the matched root, so it doesn't hijack the Markdown
 * editor's CM6-internal undo or the global window.
 */
interface UseUndoShortcutsOpts {
  rootSelector: string;
  undo: () => void;
  redo: () => void;
}

export function useUndoShortcuts(opts: UseUndoShortcutsOpts): void {
  const ref = useRef(opts);
  ref.current = opts;
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const k = e.key.toLowerCase();
      const isUndo = k === 'z' && !e.shiftKey;
      const isRedo = (k === 'z' && e.shiftKey) || k === 'y';
      if (!isUndo && !isRedo) return;
      const ae = document.activeElement as HTMLElement | null;
      if (!ae || typeof ae.closest !== 'function') return;
      if (!ae.closest(ref.current.rootSelector)) return;
      e.preventDefault();
      e.stopPropagation();
      if (isUndo) ref.current.undo();
      else ref.current.redo();
    };
    document.addEventListener('keydown', handler, true);
    return () => document.removeEventListener('keydown', handler, true);
  }, [opts.rootSelector]);
}
