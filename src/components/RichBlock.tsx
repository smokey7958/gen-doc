/**
 * Phase F-2: contentEditable surface for one Word block.
 *
 * Replaces the textarea-based BlockRow body when we want selection-aware
 * inline B/I/U editing. The component owns:
 *
 *   - Reading the latest `runs` from props and writing them to the DOM via
 *     `innerHTML` only when the model genuinely diverges from what's there.
 *     We avoid rewriting on every keystroke because that would either
 *     stomp the caret or force us to re-resolve it constantly. Instead we
 *     compare a serialized hash of the runs we *last wrote* with the new
 *     incoming runs.
 *   - IME composition guard. While the user is composing CJK input the
 *     browser inserts intermediate characters into the DOM that don't yet
 *     belong in the model. We hold off `onChange` and any innerHTML rewrite
 *     until `compositionend` fires.
 *   - Caret preservation across rewrites. When we *do* rewrite (e.g., after
 *     a programmatic style toggle from the toolbar), we save the caret
 *     position in character space, replace innerHTML, and restore it.
 *   - Imperative API for the parent to call `applyRange()` after toggling
 *     B/I/U on the toolbar — the parent computes new runs from the current
 *     selection and hands them back.
 */

import { forwardRef, useCallback, useEffect, useImperativeHandle, useLayoutEffect, useRef } from 'react';
import type { DocxRun } from '../lib/docx-adapter';
import { domToRuns, getCharRange, runsToHtml, runsToText, setCharRange } from '../lib/rich-text';
import { cn } from '../lib/utils';

/** Imperative handle exposed to the parent so the toolbar can read the
 *  selection and write fresh runs without re-rendering noise. */
export interface RichBlockHandle {
  /** True if the user has a non-empty selection inside this block. */
  hasSelection: () => boolean;
  /** Get the current selection in character offsets, or null. */
  getRange: () => { start: number; end: number } | null;
  /** Force a re-read from the DOM (e.g., before a toolbar action that
   *  needs the latest text the user typed). Returns the runs that the
   *  block currently shows. */
  readRuns: () => DocxRun[];
  /** Re-focus and place the caret at the end of the block's text. */
  focusEnd: () => void;
}

interface Props {
  blockId: string;
  runs: DocxRun[];
  textClass: string;
  inlineStyle: React.CSSProperties;
  placeholder?: string;
  active: boolean;
  /** Print-preview gate. When true, the block becomes non-editable and
   *  drops every editor-only visual (active ring, focus tint) so what the
   *  user sees matches what would print. */
  previewMode?: boolean;
  /** Called whenever the user types / pastes — receives the new runs. */
  onChange: (runs: DocxRun[]) => void;
  onFocus: () => void;
  /** Enter (no shift) splits the block — parent inserts a new block after. */
  onEnter: () => void;
}

export const RichBlock = forwardRef<RichBlockHandle, Props>(function RichBlock(
  { blockId, runs, textClass, inlineStyle, placeholder, active, previewMode, onChange, onFocus, onEnter },
  ref,
) {
  const elRef = useRef<HTMLDivElement>(null);
  const composingRef = useRef(false);
  // Snapshot of what runs we last wrote to the DOM. We compare via a quick
  // canonical string — runs are small so the cost is negligible.
  const lastWrittenKey = useRef<string>('');

  const runsKey = useCallback((rs: DocxRun[]): string => {
    // Bold/italic/underline collapse to a 3-bit prefix; text follows. Cheap
    // and deterministic.
    let out = '';
    for (const r of rs) {
      const s = r.style;
      out +=
        (s?.bold ? 'B' : '_') +
        (s?.italic ? 'I' : '_') +
        (s?.underline ? 'U' : '_') +
        '|' +
        r.text +
        '\u0001';
    }
    return out;
  }, []);

  // Sync model → DOM. We do this in useLayoutEffect so the rewrite happens
  // before paint (no visible flash). Skip when the runs match what we last
  // wrote — that case means the change came from this block's own oninput
  // handler and the DOM is already correct.
  useLayoutEffect(() => {
    const el = elRef.current;
    if (!el) return;
    if (composingRef.current) return; // Don't disturb IME.
    const key = runsKey(runs);
    if (key === lastWrittenKey.current) return;
    // Save caret in char space so we can restore it after innerHTML wipe.
    const saved = el.contains(document.activeElement) ? getCharRange(el) : null;
    // Empty-paragraph IME anchor.
    //
    // When `runsToHtml` would emit just a zero-width-space (the "no text"
    // shape of `[{ text: '' }]`), Chromium's IME compositor can't bind a
    // composition context to that single zero-width text node — so the very
    // first keystroke into a fresh empty paragraph is silently dropped, and
    // only switching IME (which forces a context rebuild) recovers it. Users
    // saw "open new project → blank Word → can't type until I switch IME".
    // The browser's own fix for empty contentEditable lines is `<br>`, so we
    // mirror that: when the model says "no characters", we write a `<br>`
    // instead of a ZWSP. `domToRuns` already ignores `<br>`, so this stays
    // round-trip-clean — the model never sees the placeholder element.
    const isEmpty = runsToText(runs).length === 0;
    if (isEmpty) {
      // Empty-paragraph representation: an explicit (empty) text node
      // alongside the visual `<br>`. Why both:
      //   - The `<br>` gives the line a visible height and a stable
      //     point for Chromium's IME compositor to bind to (a ZWSP-only
      //     element can't host an IME composition context — that's the
      //     "first keystroke is lost until I switch IME" bug from a
      //     previous round).
      //   - The empty text node gives `Selection.collapse` something to
      //     ANCHOR to. We discovered via the FocusDiagnostic overlay
      //     that on the Ctrl+N → blank Word path, focus lands on the
      //     contentEditable correctly (`document.activeElement` IS the
      //     RichBlock div) but no caret renders and keypresses are
      //     dropped. Reason: a Range collapsed at offset 0 of an
      //     element whose only child is `<br>` is technically valid but
      //     Chromium fails to render a caret for it on this re-mount
      //     path. Anchoring to a real text node (even one of length 0)
      //     fixes it deterministically.
      // `domToRuns` walks SHOW_TEXT and skips zero-length text nodes,
      // so this stays round-trip-clean — the model never sees the anchor.
      while (el.firstChild) el.removeChild(el.firstChild);
      el.appendChild(document.createTextNode(''));
      el.appendChild(document.createElement('br'));
    } else {
      el.innerHTML = runsToHtml(runs);
    }
    lastWrittenKey.current = key;
    if (saved && el.contains(document.activeElement)) {
      // Clamp to text length (the toolbar may have shortened the run).
      const total = runsToText(runs).length;
      const start = Math.min(saved.start, total);
      const end = Math.min(saved.end, total);
      setCharRange(el, start, end);
    }
  }, [runs, runsKey]);

  const fireChange = useCallback(() => {
    const el = elRef.current;
    if (!el) return;
    const next = domToRuns(el);
    // Update lastWrittenKey so the layout effect above won't fight the
    // user's input: the model is about to receive `next`, and the DOM
    // already reflects it.
    lastWrittenKey.current = runsKey(next);
    onChange(next);
  }, [onChange, runsKey]);

  useImperativeHandle(
    ref,
    () => ({
      hasSelection: () => {
        const el = elRef.current;
        if (!el) return false;
        const r = getCharRange(el);
        return !!r && r.end > r.start;
      },
      getRange: () => {
        const el = elRef.current;
        if (!el) return null;
        return getCharRange(el);
      },
      readRuns: () => {
        const el = elRef.current;
        if (!el) return runs;
        return domToRuns(el);
      },
      focusEnd: () => {
        const el = elRef.current;
        if (!el) return;
        const total = runsToText(domToRuns(el)).length;
        el.focus();
        setCharRange(el, total, total);
      },
    }),
    [runs],
  );

  // Caret-placement fallback for freshly-mounted empty blocks.
  //
  // Repro: open the app → blank Word, type fine. New Project → blank Word,
  // click the paragraph → no caret appears, no typing possible. The tab is
  // freshly mounted (EditorSurface keys the ErrorBoundary on `active.id`,
  // and `newWorkspace` issues a new uuid), so React just constructed a
  // brand-new contentEditable element. Chromium fails to bind a usable
  // selection to a contentEditable whose only child is `<br>` until the
  // element has been focused at least once *and* a selection has been
  // explicitly placed in it. Click → focus alone isn't enough.
  //
  // Strategy: place the caret programmatically whenever the empty block
  // gets focus (covers click + Tab) AND auto-focus on mount when the whole
  // document is empty (covers "user opens blank Word and never clicks —
  // just starts typing"). We don't preventDefault on mousedown — that
  // turned out to break Electron's IME path on Windows.
  const placeCaretIfEmpty = useCallback(() => {
    const el = elRef.current;
    if (!el) return;
    if (runsToText(domToRuns(el)).length !== 0) return;
    const sel = window.getSelection();
    if (!sel) return;
    const range = document.createRange();
    // Prefer the (empty) text node we wrote alongside `<br>` — anchoring
    // to a real text node is what actually makes the caret render on the
    // post-Ctrl+N re-mount path. See the layout effect above for why.
    // Fall back to selectNodeContents in case some path produced an
    // element with no text node yet (defensive — shouldn't fire in
    // normal flows but cheap insurance).
    const firstText = Array.from(el.childNodes).find(
      (n): n is Text => n.nodeType === Node.TEXT_NODE,
    );
    if (firstText) {
      range.setStart(firstText, 0);
      range.setEnd(firstText, 0);
    } else {
      range.selectNodeContents(el);
      range.collapse(true);
    }
    sel.removeAllRanges();
    sel.addRange(range);
  }, []);

  const handleFocus = useCallback(() => {
    onFocus();
    // rAF so the browser's own post-focus selection logic settles first;
    // we then overwrite with a known-good range. Without rAF, Chromium has
    // been seen to wipe the selection we set in the same tick.
    requestAnimationFrame(placeCaretIfEmpty);
  }, [onFocus, placeCaretIfEmpty]);

  // Auto-focus + caret-place on mount when this block is the (single)
  // empty paragraph of a fresh blank doc — DocxEditor pre-sets
  // `activeBlockId` for that case so we get `active=true` here.
  //
  // FocusDiagnostic confirmed focus itself lands reliably on this element
  // after Ctrl+N; the previous retry loop with `window.focus()` was treating
  // the wrong symptom. The actual missing piece was caret placement against
  // a real text-node anchor, handled now by `placeCaretIfEmpty` above and
  // by the layout effect's `<text node> + <br>` empty representation.
  useEffect(() => {
    if (!active) return;
    const el = elRef.current;
    if (!el) return;
    if (runsToText(domToRuns(el)).length !== 0) return;
    el.focus();
    placeCaretIfEmpty();
  }, [active, placeCaretIfEmpty]);

  // Strip pasted formatting we can't represent. We accept text only and let
  // the next render coalesce; the user can re-apply B/I/U with the toolbar.
  // (Native paste of <strong>/<em>/<u> spans would mostly survive domToRuns,
  // but Word/Web sources also drop in spans with inline color/font, which
  // we don't model and don't want lingering as zombie nodes.)
  const onPaste = useCallback((e: React.ClipboardEvent<HTMLDivElement>) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    if (!text) return;
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    range.insertNode(document.createTextNode(text));
    range.collapse(false);
    sel.removeAllRanges();
    sel.addRange(range);
    fireChange();
  }, [fireChange]);

  // We don't preventDefault on plain typing — let the browser do its thing
  // and fire `input` for us to read the resulting DOM. Enter is the one
  // exception: we want to *replace* the default <div>/<br> insertion with
  // a fresh new block in the parent's model.
  const onKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      if (e.key === 'Enter' && !e.shiftKey && !composingRef.current) {
        e.preventDefault();
        // Flush any pending typed text before splitting; otherwise the
        // last keystroke before Enter would be lost.
        fireChange();
        onEnter();
        return;
      }
      // Tab inside a paragraph would otherwise focus the next focusable
      // sibling (the insert / delete buttons in the parent's hover gutter),
      // which yanks the user out of the document. Word inserts a literal
      // tab character; we mirror that — soft tab, not block reordering.
      if (e.key === 'Tab' && !composingRef.current) {
        e.preventDefault();
        const sel = window.getSelection();
        if (!sel || sel.rangeCount === 0) return;
        const range = sel.getRangeAt(0);
        range.deleteContents();
        range.insertNode(document.createTextNode('\t'));
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
        fireChange();
      }
    },
    [fireChange, onEnter],
  );

  // Showing/hiding the placeholder via CSS attribute since it has to live
  // outside the contentEditable subtree (or contentEditable would let the
  // user type into it). We toggle by data attribute — see CSS in index.css.
  useEffect(() => {
    const el = elRef.current;
    if (!el) return;
    const empty = runsToText(runs).length === 0;
    el.dataset.empty = empty ? '1' : '0';
  }, [runs]);

  return (
    <div
      ref={elRef}
      data-block-id={blockId}
      data-rich-block="1"
      data-empty={runsToText(runs).length === 0 ? '1' : '0'}
      data-placeholder={placeholder ?? ''}
      // Preview mode only hides editing chrome; the block stays editable so
      // the user can keep typing while inspecting the print appearance.
      contentEditable
      suppressContentEditableWarning
      onFocus={handleFocus}
      onInput={() => {
        if (composingRef.current) return;
        fireChange();
      }}
      onCompositionStart={() => {
        composingRef.current = true;
      }}
      onCompositionEnd={() => {
        composingRef.current = false;
        fireChange();
      }}
      onKeyDown={onKeyDown}
      onPaste={onPaste}
      className={cn(
        'w-full bg-transparent outline-none rounded px-2 py-1 whitespace-pre-wrap break-words',
        textClass,
        // Editor-only visuals — in preview mode we want the page to read as
        // pure document content, no focus tint and no active ring.
        !previewMode && 'focus:bg-secondary/30',
        active && !previewMode && 'ring-1 ring-primary/20',
      )}
      style={inlineStyle}
      // No children prop — the layout effect above is the sole writer of
      // innerHTML. React doesn't reconcile descendants when the element
      // has no children, so the user's typed DOM survives between renders.
    />
  );
});
