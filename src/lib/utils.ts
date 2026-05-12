import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}

/**
 * Clamp a floating element's `(x, y)` anchor (typically a right-click
 * point) so the element stays inside the viewport. Without this, a
 * right-click near the right or bottom edge of the window renders the
 * context menu with its later items clipped off-screen — Esc still
 * dismisses, but discoverability of the cut-off options is zero.
 *
 * We use caller-provided estimates for width/height rather than
 * measuring the rendered menu post-mount. Measuring would mean a
 * one-frame flicker (initial render at the unclamped position, then
 * a snap into bounds via `useLayoutEffect`) — visually worse than a
 * static estimate that's slightly conservative for our small fixed-
 * content menus. Generous estimates lose at worst one or two items'
 * worth of margin near the edge; an under-estimate clips the same
 * way the unclamped version did. Pass values that bound the actual
 * rendered size.
 *
 * `margin` keeps a small visual gap between the menu and the viewport
 * edge so the shadow has room to breathe.
 */
/**
 * R382 — Unicode-aware string truncation. JS's `.slice(0, n)` is UTF-16
 * code-unit indexed and can split a surrogate pair in half — for
 * emoji-bearing input where code unit n-1 is a high surrogate (the first
 * half of a 2-unit emoji like「🚀」), the result is an ORPHAN HIGH
 * SURROGATE that renders as「�」 / fails strict JSON serialization.
 * `Array.from(text)` uses the code-point iterator so each emoji is one
 * array element; slicing and joining keeps boundaries intact.
 *
 * Two flavors:
 *   • `sliceCodePoints` — raw slice, no suffix. Use when the truncated
 *     output is itself consumed (search query, search highlight, etc.)
 *     and any extra suffix character would corrupt downstream logic.
 *   • `slicePreview` — slice + `…` suffix when truncated. Use for
 *     human-readable display where the ellipsis tells the reader「more
 *     content was cut」 (badge text, summary strings, undo descriptions).
 *
 * Both return the input unchanged when no truncation is needed and
 * share a two-step optimisation: skip `Array.from` allocation when
 * `text.length` (an upper bound on code-point count) is already within
 * the cap.
 */
export function sliceCodePoints(text: string, maxCodePoints: number): string {
  if (text.length <= maxCodePoints) return text;
  const codePoints = Array.from(text);
  if (codePoints.length <= maxCodePoints) return text;
  return codePoints.slice(0, maxCodePoints).join('');
}

export function slicePreview(text: string, maxCodePoints: number): string {
  if (text.length <= maxCodePoints) return text;
  const codePoints = Array.from(text);
  if (codePoints.length <= maxCodePoints) return text;
  return codePoints.slice(0, maxCodePoints).join('') + '…';
}

/**
 * R392 / R395 — true ONLY for "the file is gone from disk". Used by the
 * three open-failure catch arms (App.tsx auto-open + handleOpenRecent,
 * EditorSurface.onOpenRecent) to decide whether to silently prune the
 * dead entry from `recentFiles`.
 *
 * Other failure classes — corrupted manifest (R376 bad type / R377
 * duplicate id / R390 duplicate file path), EACCES from antivirus lock,
 * EISDIR if the path was a directory, IPC bridge throws, app.confirm
 * rejection — all leave the .gd intact on disk; the user may legitimately
 * want a second attempt after fixing the situation externally. Pruning
 * for those failures permanently erases the menu entry, surprising the
 * user with「我剛才看到的最近檔案怎麼不見了」 next session.
 *
 * Node's `fs.readFile` produces error messages with the OS error code
 * embedded as a prefix —「ENOENT: no such file or directory, open
 * '/path'」 — and that prefix survives the Electron IPC structured-clone
 * round-trip (Error.message is preserved verbatim, even though
 * `err.code` typically isn't). Substring-match on the canonical prefix
 * is the simplest discriminator that survives the IPC boundary without
 * needing to change the main-side error shape.
 *
 * We intentionally DO NOT match other ENOENT-adjacent codes (ENOTDIR
 * etc.) — those are rare on a workflow that just resolved a saved
 * absolute path, and the "in doubt, keep the entry" posture is the
 * right default. False negatives (failing to prune a genuinely-gone
 * file) re-prompt the user next session; false positives (pruning a
 * recoverable failure) permanently erase user state — asymmetric, so
 * we err toward false negatives.
 *
 * R395 — hoisted out of App.tsx into lib/utils.ts so the EditorSurface
 * welcome-screen「最近開啟」 list's onOpenRecent (which has the same
 * unconditional-prune-on-any-error bug R392 closed for App.tsx) can
 * share the same gate. Three-way consistency across every recent-list
 * eviction site, same predicate, same false-positive avoidance.
 */
export function isFileMissingError(err: unknown): boolean {
  const msg = err instanceof Error ? err.message : String(err);
  return msg.startsWith('ENOENT') || msg.includes('ENOENT:');
}

export function clampToViewport(
  x: number,
  y: number,
  estimatedWidth: number,
  estimatedHeight: number,
  margin = 8,
): { left: number; top: number } {
  const maxX = Math.max(margin, window.innerWidth - estimatedWidth - margin);
  const maxY = Math.max(margin, window.innerHeight - estimatedHeight - margin);
  return {
    left: Math.min(Math.max(margin, x), maxX),
    top: Math.min(Math.max(margin, y), maxY),
  };
}
