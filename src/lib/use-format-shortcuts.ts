/**
 * Reusable Ctrl/Cmd+B / +I / +U keyboard shortcut hook for the Word / Excel /
 * PowerPoint editors.
 *
 * Why a custom hook instead of CodeMirror's keymap or per-editor handlers:
 *   - These editors render plain `<textarea>` / `<input>` elements (not CM6),
 *     so the browser's default `Ctrl+B` (no-op in textareas) is what we
 *     intercept.
 *   - All three editors already have their own `toggleStyle` / `updateStyle`
 *     functions wired into the toolbar — the hook just calls those, so the
 *     visual outcome matches clicking the toolbar button exactly.
 *
 * Usage:
 *   1. Add `data-{kind}-editor-root` attribute to the editor's outermost div.
 *   2. Call `useFormatShortcuts({ rootSelector, isActive, toggle })` inside
 *      the editor component. `isActive` gates the shortcut on whether a
 *      target block / shape / cell is focused.
 *   3. The hook only fires when document.activeElement is inside that root,
 *      so it does NOT hijack Ctrl+B in unrelated tabs / panels.
 */

import { useEffect, useRef } from 'react';

export type FormatShortcutKey = 'bold' | 'italic' | 'underline';

interface UseFormatShortcutsOpts {
  /** CSS selector that scopes the shortcut to one editor instance. */
  rootSelector: string;
  /** Return true when there is a focused target the toggle can act on. */
  isActive: () => boolean;
  /** Apply the format toggle (delegates to existing toolbar logic). */
  toggle: (key: FormatShortcutKey) => void;
  /** Subset of keys to enable. Default: all three. */
  keys?: ReadonlyArray<FormatShortcutKey>;
}

const KEY_MAP: Record<string, FormatShortcutKey> = {
  b: 'bold',
  i: 'italic',
  u: 'underline',
};

export function useFormatShortcuts(opts: UseFormatShortcutsOpts): void {
  // Stash latest opts in a ref so the document listener does not need to be
  // re-bound on every render.
  const ref = useRef(opts);
  ref.current = opts;

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      // Require Ctrl (Win/Linux) or Cmd (mac); reject Shift/Alt to avoid
      // colliding with Ctrl+Shift+B etc.
      const mod = e.ctrlKey || e.metaKey;
      if (!mod || e.shiftKey || e.altKey) return;
      const k = e.key.toLowerCase();
      const formatKey = KEY_MAP[k];
      if (!formatKey) return;
      const allowed = ref.current.keys;
      if (allowed && !allowed.includes(formatKey)) return;

      // Scope: only fire when focus is somewhere inside this editor's root.
      const ae = document.activeElement as HTMLElement | null;
      if (!ae || typeof ae.closest !== 'function') return;
      if (!ae.closest(ref.current.rootSelector)) return;
      if (!ref.current.isActive()) return;

      e.preventDefault();
      ref.current.toggle(formatKey);
    };
    document.addEventListener('keydown', handler);
    return () => document.removeEventListener('keydown', handler);
  }, [opts.rootSelector]);
}
