/**
 * Toast viewport — renders the live stack of toasts from `useToasts`.
 *
 * Mount once near the App root. Listens to the store, auto-dismisses by
 * comparing each toast's `expiresAt` against a 250ms tick (cheap; running
 * only while at least one toast is visible). The only animation is the
 * R410 entrance keyframe (`animate-toast-in`, a 150ms transform/opacity
 * ease — GPU-composited, so fast keyboard / mouse users can still fire
 * several toasts back-to-back without frame jank).
 *
 * Position is fixed bottom-right so it never overlaps the editor's
 * top-bar / find dialog (top-right) or the AI panel (right edge above the
 * status bar). Each toast is `pointer-events-auto` inside a
 * `pointer-events-none` container so clicks anywhere else fall through to
 * the editor instead of being eaten by the empty viewport.
 */

import { useEffect, useRef } from 'react';
import { AlertCircle, AlertTriangle, CheckCircle2, Info, X } from 'lucide-react';
import { useToasts, type ToastVariant } from '../store/toast';
import { tImp } from '../lib/i18n';
import { cn } from '../lib/utils';

const VARIANT_STYLES: Record<ToastVariant, string> = {
  info: 'bg-background border-border text-foreground',
  success: 'bg-background border-emerald-500/40 text-foreground',
  warning: 'bg-background border-amber-500/50 text-foreground',
  error: 'bg-background border-destructive/50 text-foreground',
};

const VARIANT_ICON: Record<ToastVariant, React.ComponentType<{ className?: string }>> = {
  info: Info,
  success: CheckCircle2,
  warning: AlertTriangle,
  error: AlertCircle,
};

const VARIANT_ICON_TONE: Record<ToastVariant, string> = {
  info: 'text-muted-foreground',
  success: 'text-emerald-600',
  warning: 'text-amber-600',
  error: 'text-destructive',
};

export function Toaster(): JSX.Element {
  const toasts = useToasts((s) => s.toasts);
  const dismiss = useToasts((s) => s.dismiss);
  const extend = useToasts((s) => s.extend);

  /**
   * Per-toast pause bookkeeping. Maps toast id → { since, sources } where
   * `sources` tracks every active pause reason (mouse hover, keyboard focus,
   * future ones). The toast is considered paused while `sources.size > 0`;
   * we only credit elapsed time back to `expiresAt` (via `extend`) when the
   * last source releases, so a user who Tab-focuses a toast they were
   * hovering doesn't accidentally resume the countdown when the cursor
   * drifts off (and vice versa).
   *
   * Without this multi-source model, R157's keyboard parity collided with
   * the existing hover handlers: mouseLeave fired while focus was still
   * held would credit the elapsed time and unpause, leaving the focused
   * (still-being-read) toast on a live timer. Same race in the other
   * direction. Tracking sources independently makes both signals composable.
   *
   * Ref (not state) on purpose: the dismissal tick reads it, and we don't
   * want pause transitions to re-run the auto-dismiss effect.
   */
  const pausedRef = useRef<Map<string, { since: number; sources: Set<'hover' | 'focus'> }>>(
    new Map(),
  );
  const pause = (id: string, source: 'hover' | 'focus') => {
    const cur = pausedRef.current.get(id);
    if (cur) {
      cur.sources.add(source);
    } else {
      pausedRef.current.set(id, { since: Date.now(), sources: new Set([source]) });
    }
  };
  const resume = (id: string, source: 'hover' | 'focus') => {
    const cur = pausedRef.current.get(id);
    if (!cur) return;
    cur.sources.delete(source);
    if (cur.sources.size === 0) {
      extend(id, Date.now() - cur.since);
      pausedRef.current.delete(id);
    }
  };

  // Auto-dismiss tick. Only run while there's something to track — when
  // the array is empty we don't even mount the interval, which keeps the
  // idle app cost at zero. 250ms is the same cadence the FocusDiagnostic
  // overlay used; perceptually instant without burning CPU.
  useEffect(() => {
    if (toasts.length === 0) return;
    const id = window.setInterval(() => {
      const now = Date.now();
      for (const t of toasts) {
        if (pausedRef.current.has(t.id)) continue;
        if (t.expiresAt !== undefined && t.expiresAt <= now) {
          dismiss(t.id);
        }
      }
    }, 250);
    return () => window.clearInterval(id);
  }, [toasts, dismiss]);

  // Drop pause bookkeeping for toasts that no longer exist (× clicked
  // mid-hover, or `clear()` from somewhere else). Otherwise the map
  // would slowly accumulate dead ids across the session.
  useEffect(() => {
    const live = new Set(toasts.map((t) => t.id));
    for (const id of pausedRef.current.keys()) {
      if (!live.has(id)) pausedRef.current.delete(id);
    }
  }, [toasts]);

  if (toasts.length === 0) return <></>;

  return (
    <div
      // pointer-events-none so the empty viewport doesn't intercept
      // clicks on whatever is underneath. Each toast re-enables pointer
      // events for its own area.
      className="pointer-events-none fixed bottom-3 right-3 z-50 flex flex-col gap-2 max-w-[min(420px,90vw)]"
      role="region"
      aria-label={tImp('通知', 'Notifications')}
    >
      {toasts.map((t) => {
        const Icon = VARIANT_ICON[t.variant];
        return (
          <div
            key={t.id}
            role="status"
            aria-live={t.variant === 'error' ? 'assertive' : 'polite'}
            onMouseEnter={() => pause(t.id, 'hover')}
            onMouseLeave={() => resume(t.id, 'hover')}
            // R157 — keyboard / SR parity for hover-pause. The mouseEnter/
            // mouseLeave pair above freezes the dismissal countdown while a
            // sighted user reads the toast, but a Tab user landing on the
            // close button — or a SR user navigating through this region —
            // would see / hear the toast vanish mid-read because the timer
            // kept ticking. React synthesizes onFocus / onBlur to bubble
            // (unlike native focus/blur which don't), so binding here on the
            // toast `<div>` catches focus on any descendant — currently just
            // the close button below, but any future interactive (action
            // button, link in the message) is covered automatically. Hover
            // and focus track as independent pause sources via `pause` /
            // `resume` so a user who hovers, then Tab-focuses, then mouse-
            // leaves does NOT prematurely unpause the countdown — the toast
            // stays paused until BOTH sources release.
            onFocus={() => pause(t.id, 'focus')}
            onBlur={() => resume(t.id, 'focus')}
            className={cn(
              // R410 — animate-toast-in: entrance keyframe (index.css).
              'pointer-events-auto flex items-start gap-2 px-3 py-2 rounded-md border shadow-md text-sm animate-toast-in',
              VARIANT_STYLES[t.variant],
            )}
          >
            <Icon className={cn('h-4 w-4 mt-0.5 shrink-0', VARIANT_ICON_TONE[t.variant])} />
            {/* whitespace-pre-line preserves "\n" (used for
                multi-failure-list alerts); break-words guards against
                un-spaced URLs / paths overflowing the toast width. */}
            <div className="flex-1 whitespace-pre-line break-words leading-snug">
              {t.message}
            </div>
            {/* R121 — cross-component scope-disclosure parity. Bare 「關閉」
                was the lone outlier among the three dismissal-X buttons in
                the codebase:
                  AIPanel.tsx:588        `關閉錯誤訊息`  ← scope-disclosed
                  FileExplorer.tsx:318   `關閉`          ← bare (also fixed)
                  Toaster.tsx:132        `關閉`          ← was here
                The dissonance was sharpest because AIPanel's own button
                comment (line 583) explicitly says "Mirrors the Toaster ×
                pattern (Toaster.tsx:130-136)" — the cross-reference
                established a parity claim that the actual strings broke,
                with the mirror disclosing more than the source. R119 just
                established that bare 關閉 is the outlier vocabulary on
                close-X buttons (TabBar.tsx:467 → 關閉此頁籤). 通知 is the
                generic term because the toast viewport hosts all four
                variants (info / success / warning / error — see
                VARIANT_STYLES at line 23) so 關閉錯誤訊息 would lie on
                non-error toasts; FileExplorer's banner is variant-fixed
                error so it can use the AIPanel string verbatim. */}
            <button
              type="button"
              title={tImp('關閉通知', 'Dismiss notification')}
              // R152 — icon-only-button accessible-name parity. Pairs with
              // AIPanel.tsx:609-610 / TabBar.tsx (close-X) / FileExplorer
              // .tsx (banner X), all updated this round. The toast's
              // surrounding container at line 100-101 already has
              // `role="region" aria-label="通知"` for the viewport itself;
              // the close button needs its own action-scoped label so SR
              // users reach the dismiss control without falling back to
              // the unannounced `<X>` icon.
              aria-label={tImp('關閉通知', 'Dismiss notification')}
              onClick={() => dismiss(t.id)}
              className="text-muted-foreground hover:text-foreground transition-colors rounded p-0.5 -mr-1 -mt-0.5"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        );
      })}
    </div>
  );
}
