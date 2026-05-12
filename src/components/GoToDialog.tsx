import { useEffect, useRef, useState } from 'react';
import { ArrowRight } from 'lucide-react';
import { notify } from '../store/toast';

/**
 * Tiny "Go to N" dialog. Modeled on FindReplaceDialog but stripped down:
 * single number input + jump button. The host (PptxEditor / DocxEditor)
 * binds Ctrl+G and toggles `open`; `focusNonce` re-fires the input focus
 * each Ctrl+G press so users who hit Ctrl+G while the dialog is already
 * open re-focus the input (matches VS Code's "Go to Line" behaviour).
 *
 * Generic across formats: PowerPoint passes slide count + slide label
 * ("跳到第幾張投影片？"); Word passes block count + paragraph label
 * ("跳到第幾段？"). The dialog itself doesn't know what N means — it just
 * clamps to [1, max] and hands the integer back via onJump.
 *
 * Why a separate component (rather than inlined in each editor):
 * - Two callers already; inlining would duplicate ~70 lines.
 * - The Ctrl+G keymap belongs to the host (it knows the editor's scope
 *   selector); the dialog only owns the input UI.
 */
export interface GoToDialogProps {
  open: boolean;
  onClose: () => void;
  focusNonce?: number;
  /** Total item count (slides, blocks, …) — input is clamped to [1, max]. */
  max: number;
  /** 1-based jump target. Receives a clamped, validated integer. */
  onJump: (oneBased: number) => void;
  /** Label shown above the input. Caller picks domain-appropriate wording. */
  label?: string;
}

export function GoToDialog({
  open,
  onClose,
  focusNonce = 0,
  max,
  onJump,
  label = '跳到第幾項？',
}: GoToDialogProps): JSX.Element | null {
  const inputRef = useRef<HTMLInputElement>(null);
  const [value, setValue] = useState('');
  // R93 — empty-submit feedback. Without this, clicking 跳轉 (or pressing
  // Enter) with the input still empty silently returned at submit's
  // `!Number.isFinite(n)` guard (line ~131): button looked responsive but
  // produced zero signal. Mirrors LinkInsertDialog's `urlError` pattern at
  // MarkdownEditor.tsx:1017-1026 where the same dialog-shape (input +
  // primary action button) flips a red border + inline message on empty
  // submit instead of silently no-op'ing. The author already closed the
  // out-of-range silent-jump gap (see comment at lines ~104-113) — empty
  // input was the parallel case that slipped through.
  const [emptyError, setEmptyError] = useState(false);
  // Same focus-restore contract as FindReplaceDialog: capture the element
  // that had focus when the dialog opened (typically the editor surface or
  // the toolbar button), restore on close so Esc/× doesn't strand focus on
  // <body> with no editor accepting keystrokes. Only restore when nothing
  // else has claimed focus post-close — onJump may legitimately have
  // scrolled the editor and focused a slide/paragraph element.
  const restoreFocusToRef = useRef<HTMLElement | null>(null);

  // Clear + focus on open / Ctrl+G re-press. queueMicrotask delays the focus
  // until after React has committed the JSX so the input element exists.
  useEffect(() => {
    if (!open) return;
    setValue('');
    setEmptyError(false);
    queueMicrotask(() => {
      const el = inputRef.current;
      if (!el) return;
      el.focus();
      el.select();
    });
  }, [open, focusNonce]);

  useEffect(() => {
    if (open) {
      restoreFocusToRef.current = document.activeElement as HTMLElement | null;
      return;
    }
    const target = restoreFocusToRef.current;
    restoreFocusToRef.current = null;
    if (!target) return;
    if (!document.body.contains(target)) return;
    const ae = document.activeElement;
    if (ae !== null && ae !== document.body) return;
    target.focus();
  }, [open]);

  // Document-level Esc — closes even when focus has wandered. Skipped when
  // the user is actively editing another input, mirroring FindReplaceDialog.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return;
      const ae = document.activeElement as HTMLElement | null;
      if (ae) {
        const tag = ae.tagName;
        // We do want Esc to close even when our own input has focus, so
        // only skip if focus is in a *different* editable element. Crude
        // check: if focus is our input, allow close; otherwise, if it's
        // any other input/textarea/contentEditable, defer.
        if (ae !== inputRef.current) {
          if (tag === 'INPUT' || tag === 'TEXTAREA' || ae.isContentEditable) return;
        }
      }
      onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  if (!open) return null;

  // Visual feedback for out-of-range / non-numeric input. Without this,
  // typing "100" in a 50-slide deck and hitting Enter silently jumped
  // to slide 50 — the user got no signal that their input was clamped
  // or that the deck was shorter than they thought. The HTML `max`
  // attribute only constrains the spinner buttons, not typed values.
  // Mirrors the FindReplaceDialog's red-ring pattern (Round 33). We
  // intentionally keep submit's behaviour permissive (clamp and jump)
  // — refusing the keystroke would be a worse experience than letting
  // it through with a visible "we adjusted this" cue. The hint also
  // names the actual valid range in red so the user can re-aim.
  const trimmed = value.trim();
  const parsed = trimmed === '' ? null : Number.parseInt(trimmed, 10);
  const isOutOfRange =
    parsed !== null &&
    (Number.isNaN(parsed) || !Number.isFinite(parsed) || parsed < 1 || parsed > max);

  const submit = () => {
    const trimmedValue = value.trim();
    if (!trimmedValue) {
      // R93 — flag the empty input + refocus instead of silently returning.
      // Mirrors LinkInsertDialog at MarkdownEditor.tsx:1019-1026.
      setEmptyError(true);
      inputRef.current?.focus();
      return;
    }
    const n = Number.parseInt(trimmedValue, 10);
    if (!Number.isFinite(n) || max <= 0) return;
    const clamped = Math.max(1, Math.min(max, n));
    onJump(clamped);
    // Deliver the "we adjusted this" cue the comment above claims exists.
    // The red ring only shows *while typing* — it vanishes the instant the
    // dialog unmounts on close, so a user who types "100" in a 50-slide
    // deck and hits Enter currently lands on slide 50 with zero signal that
    // their typed number was beyond the deck. Without the toast, the
    // permissive-clamp behaviour silently lies about where the user ended
    // up. Toast fires only when actual clamping occurred (n !== clamped),
    // so the in-range fast path stays quiet.
    if (n !== clamped) {
      notify(`已調整為第 ${clamped} 項（輸入 ${n} 超出 1 – ${max} 範圍）`, 'info');
    }
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-label={label}
      className="absolute top-3 right-3 z-30 w-[280px] rounded-md border bg-background shadow-lg p-3 text-sm"
      // Stop the host editor's keymap (Ctrl+B/I/U etc.) from firing while
      // typing in this dialog — focus is on the input but the bubbling
      // keydown still reaches the document-level shortcut listeners.
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation();
          onClose();
        } else if (e.key === 'Enter') {
          e.stopPropagation();
          submit();
        }
      }}
    >
      <label className="block text-xs text-muted-foreground mb-1.5">{label}</label>
      <div className="flex items-center gap-1.5">
        <input
          ref={inputRef}
          type="number"
          min={1}
          max={Math.max(1, max)}
          value={value}
          onChange={(e) => {
            setValue(e.target.value);
            // Calm the empty-submit red the moment the user starts typing —
            // same in-place clearing pattern as LinkInsertDialog's `urlError`
            // (MarkdownEditor.tsx:1089-1096) and FindReplaceDialog's
            // regex-error reset.
            if (emptyError) setEmptyError(false);
          }}
          placeholder={`1 – ${max}`}
          aria-invalid={isOutOfRange || emptyError ? true : undefined}
          className={
            isOutOfRange || emptyError
              ? 'flex-1 h-8 rounded border border-destructive/60 bg-background px-2 text-sm focus:outline-none focus:ring-1 focus:ring-destructive'
              : 'flex-1 h-8 rounded border border-border bg-background px-2 text-sm focus:outline-none focus:ring-1 focus:ring-primary'
          }
        />
        {/* Sibling-shortcut-in-tooltip parity: every other primary action
            button in the app advertises its keystroke (AIPanel.tsx:1162
            "送出 (Enter)", SettingsDialog.tsx:258 "儲存 API key (Enter)",
            DiffPreview.tsx:89 "套用變更 (Ctrl+Enter)", FindReplaceDialog
            .tsx:620/632 prev/next, etc.). The button's own keystroke is
            wired at line 151-154 (Enter on the dialog calls submit), and
            the footer hint at line 195 already mentions "Enter 跳轉" —
            but that footer is *conditional*: when the user types an
            out-of-range value the footer flips to the red "請輸入 1 – N
            之間的數字" message (line 190-192), and the Enter hint
            disappears at exactly the moment they're most likely looking
            for it. The button tooltip is the only stable surface, so
            advertising "(Enter)" here closes the gap and brings the
            dialog in line with every other primary-action button. */}
        <button
          type="button"
          onClick={submit}
          title="跳轉 (Enter)"
          className="h-8 px-2 inline-flex items-center gap-1 rounded text-xs bg-primary text-primary-foreground hover:bg-primary/90 transition-colors"
        >
          <ArrowRight className="h-3.5 w-3.5" />
          跳轉
        </button>
      </div>
      {/* Inline range error: only shown when the user has typed something
          out of range, so the dialog stays quiet on first open. The hint
          text repeats the placeholder's range in destructive colour to
          make the violation concrete (telling the user what valid input
          looks like beats the abstract "invalid"). */}
      {/* R154 — both validation branches carry `role="alert"` (implicit
          `aria-live="assertive"`) so SR users typing an out-of-range or
          empty value hear the constraint immediately, matching the
          FindReplaceDialog regex-error pattern this same round added at
          FindReplaceDialog.tsx:849. The default informational branch
          (Esc / Enter help text) stays a plain div — it's static help that
          appears alongside the dialog itself; SR users get the dialog's
          aria-label on open and don't need this re-announced. The Submit
          path's clamp-toast (line 153) was the only post-action signal SR
          users got before this fix; the pre-submit visual red-border was
          sighted-only. */}
      {emptyError ? (
        <div role="alert" className="mt-1.5 text-[10px] text-destructive">
          請輸入要跳轉的項次
        </div>
      ) : isOutOfRange ? (
        <div role="alert" className="mt-1.5 text-[10px] text-destructive">
          請輸入 1 – {max} 之間的數字
        </div>
      ) : (
        <div className="mt-1.5 text-[10px] text-muted-foreground">
          Esc 關閉 · Enter 跳轉
        </div>
      )}
    </div>
  );
}
