/**
 * Lightweight toast notifications. Replaces renderer-side `window.alert()`
 * calls used for fire-and-forget error / info notices ("開啟失敗", "匯出
 * 失敗", "沒有頁籤可以匯出"…).
 *
 * Why a toast and not `alert()`:
 *   - `alert()` is synchronous and blocking — it freezes the renderer
 *     event loop, breaks any in-flight async (auto-save, streaming AI
 *     response) until dismissed.
 *   - On Windows / Electron, native `alert()` has the same OS-focus quirk
 *     as `confirm()` (see comments around App.tsx::handleNew): keystrokes
 *     stop being routed to webContents until the window is re-focused.
 *   - Toast doesn't pull the user out of context — they can keep typing,
 *     keep clicking; the message fades on its own (or they dismiss it).
 *
 * Single-store design (Zustand, like the rest of the app) so the trigger
 * site doesn't need any context plumbing — just `notify('msg', 'error')`.
 * The `<Toaster />` component renders the visible stack from this store.
 */

import { create } from 'zustand';

export type ToastVariant = 'info' | 'success' | 'warning' | 'error';

export interface Toast {
  id: string;
  message: string;
  variant: ToastVariant;
  /** Wall-clock ms when this toast should auto-dismiss; undefined = sticky. */
  expiresAt: number | undefined;
}

interface ToastState {
  toasts: Toast[];
  /**
   * Push a toast. Returns the id so callers can dismiss early — useful for
   * "in progress" → "done" transitions where the caller wants to swap one
   * toast for another. `duration` is ms; pass 0 to make it sticky (user
   * must click ×). Defaults: error 8s, warning 6s, info / success 4s — error
   * lingers because users frequently miss a 4s flash that looked important.
   */
  notify: (
    message: string,
    variant?: ToastVariant,
    options?: { duration?: number },
  ) => string;
  dismiss: (id: string) => void;
  clear: () => void;
  /**
   * Push a toast's `expiresAt` forward by `deltaMs`. Used by the Toaster's
   * pause-on-hover: the viewport tracks how long the cursor sat over a
   * toast, then on mouse-leave it credits that time back to the timer so
   * a user reading a 4 s info toast for 3 s still gets the full remaining
   * 4 s after they look away. Sticky toasts (`expiresAt === undefined`)
   * are no-ops — there's nothing to extend.
   */
  extend: (id: string, deltaMs: number) => void;
}

const DEFAULT_DURATION_MS: Record<ToastVariant, number> = {
  info: 4000,
  success: 4000,
  warning: 6000,
  error: 8000,
};

let toastCounter = 0;

export const useToasts = create<ToastState>((set, get) => ({
  toasts: [],
  notify: (message, variant = 'info', options) => {
    const duration = options?.duration ?? DEFAULT_DURATION_MS[variant];
    const expiresAt = duration > 0 ? Date.now() + duration : undefined;
    // Coalesce an identical (message+variant) toast that's still visible:
    // common when the same warning fires repeatedly (Ctrl+E mashed with no
    // open tab → "沒有開啟的頁籤…", a flapping disk → repeated "儲存失敗…",
    // drag-drop where the same path errors twice). Without dedupe the stack
    // grows by one per call and the user has to dismiss each duplicate by
    // hand. Refreshing expiresAt also gives them a fresh read window if
    // they were about to lose the original to its timeout. Stable id so any
    // caller holding it can still dismiss; position in the stack stays put
    // (the message is identical, so re-ordering would just visually flicker).
    const existing = get().toasts.find(
      (t) => t.message === message && t.variant === variant,
    );
    if (existing) {
      set((s) => ({
        toasts: s.toasts.map((t) =>
          t.id === existing.id ? { ...t, expiresAt } : t,
        ),
      }));
      return existing.id;
    }
    toastCounter += 1;
    const id = `t${toastCounter}`;
    set((s) => ({ toasts: [...s.toasts, { id, message, variant, expiresAt }] }));
    return id;
  },
  dismiss: (id) =>
    set((s) => ({ toasts: s.toasts.filter((t) => t.id !== id) })),
  clear: () => set({ toasts: [] }),
  extend: (id, deltaMs) =>
    set((s) => ({
      toasts: s.toasts.map((t) =>
        t.id === id && t.expiresAt !== undefined
          ? { ...t, expiresAt: t.expiresAt + deltaMs }
          : t,
      ),
    })),
}));

/**
 * Imperative API for non-React call sites. Equivalent to
 * `useToasts.getState().notify(...)` but reads better at the callsite.
 */
export function notify(
  message: string,
  variant: ToastVariant = 'info',
  options?: { duration?: number },
): string {
  return useToasts.getState().notify(message, variant, options);
}
