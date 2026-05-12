/**
 * Rich-text helpers for the Word editor's contentEditable surface.
 *
 * The editor's data model is `DocxRun[]` — an ordered list of `{ text, style }`
 * fragments where `style` is a subset of `{ bold, italic, underline }`. This
 * file bridges that model to a contentEditable DOM and back:
 *
 *   - `runsToHtml` / `domToRuns`        — model ↔ DOM serialization
 *   - `getCharRange` / `setCharRange`   — caret position in *characters*
 *                                          (not DOM offsets), so we can save
 *                                          and restore selection across an
 *                                          innerHTML rewrite without caring
 *                                          how the browser split text nodes
 *   - `applyStyleToRange`               — toggle B/I/U over `[start, end)`
 *   - `coalesceRuns` / `runsToText`     — bookkeeping shared with the adapter
 *
 * Why characters, not DOM offsets:
 *   When the user types into a B/I/U span the browser may split or merge
 *   text nodes; restoring a saved (Node, offset) pair after rebuilding the
 *   DOM from `runs` is fragile. Counting plain-text characters from the
 *   editor root is stable as long as the visible text didn't change — and
 *   when it did, we know exactly how to advance the caret.
 *
 * HTML shape we emit:
 *   `<strong>`, `<em>`, `<u>` for bold/italic/underline (semantic, matches
 *   what mammoth produces and what `domToRuns` reads back). We escape any
 *   user-typed HTML metacharacters so a paragraph that literally contains
 *   "<script>" round-trips as text. Empty runs render as a single `\u200B`
 *   zero-width space inside their span so contentEditable can place a caret
 *   inside an otherwise-empty bold region — but we strip those again on
 *   read so the model never carries phantom characters.
 */

import type { DocxRun, DocxRunStyle } from './docx-adapter';

const ZWSP = '\u200B';

/** Concatenate all run text into a single string. */
export function runsToText(runs: DocxRun[]): string {
  let out = '';
  for (const r of runs) out += r.text;
  return out;
}

/** True if two run styles produce the same B/I/U triple. */
export function runStylesEqual(
  a: DocxRunStyle | undefined,
  b: DocxRunStyle | undefined,
): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return !!a.bold === !!b.bold && !!a.italic === !!b.italic && !!a.underline === !!b.underline;
}

/** Drop falsy fields; return undefined when no flag is set. */
export function normalizeRunStyle(s: DocxRunStyle | undefined): DocxRunStyle | undefined {
  if (!s) return undefined;
  const out: DocxRunStyle = {};
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  return Object.keys(out).length === 0 ? undefined : out;
}

/** Merge consecutive runs that share the same style. Empty-text runs are
 *  dropped (they carry no information and waste DOM nodes). */
export function coalesceRuns(runs: DocxRun[]): DocxRun[] {
  const out: DocxRun[] = [];
  for (const r of runs) {
    if (!r.text) continue;
    const last = out[out.length - 1];
    if (last && runStylesEqual(last.style, r.style)) {
      last.text += r.text;
    } else {
      out.push({ text: r.text, style: normalizeRunStyle(r.style) });
    }
  }
  return out;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Serialize runs to an HTML string suitable for `innerHTML`. We always wrap
 * each fragment in a `<span>` carrying optional B/I/U tags around it; even
 * unstyled runs get a span so the boundary between adjacent runs survives a
 * subsequent `domToRuns` read.
 *
 * Empty runs aren't expected at the call site (callers should `coalesceRuns`
 * first to drop them), but if one slips through we substitute a ZWSP to
 * keep the contentEditable layout stable.
 */
export function runsToHtml(runs: DocxRun[]): string {
  if (runs.length === 0) return '';
  let out = '';
  for (const r of runs) {
    const text = r.text === '' ? ZWSP : escapeHtml(r.text);
    let inner = text;
    const s = r.style;
    if (s?.underline) inner = `<u>${inner}</u>`;
    if (s?.italic) inner = `<em>${inner}</em>`;
    if (s?.bold) inner = `<strong>${inner}</strong>`;
    out += inner;
  }
  return out;
}

/**
 * Extract `DocxRun[]` from a contentEditable subtree. We walk every text
 * node, collect its B/I/U ancestors, and emit one run per node, then
 * coalesce. ZWSPs we ourselves injected to support empty bold regions are
 * stripped on the way out so the model stays clean.
 *
 * `<br>` elements (Chrome inserts these for empty contentEditable lines)
 * are treated as no-op — a single block doesn't carry hard breaks in our
 * model; if the user wants a new paragraph they press Enter and we split.
 */
export function domToRuns(root: Element): DocxRun[] {
  const out: DocxRun[] = [];
  const doc = root.ownerDocument ?? document;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n: Node | null = walker.nextNode();
  while (n) {
    let text = n.textContent ?? '';
    // Strip ZWSPs we inserted to host carets in empty styled spans. If the
    // node was *only* a ZWSP it becomes empty and we skip it.
    text = text.replace(/\u200B/g, '');
    if (text.length > 0) {
      const style: DocxRunStyle = {};
      let p: Element | null = n.parentElement;
      while (p && p !== root.parentElement && p !== root) {
        const tag = p.tagName.toUpperCase();
        if (tag === 'STRONG' || tag === 'B') style.bold = true;
        if (tag === 'EM' || tag === 'I') style.italic = true;
        if (tag === 'U') style.underline = true;
        p = p.parentElement;
      }
      out.push({ text, style: normalizeRunStyle(style) });
    }
    n = walker.nextNode();
  }
  return coalesceRuns(out);
}

// ── caret bookkeeping ───────────────────────────────────────────────────

/**
 * Resolve a (node, offset) pair to the cumulative character index from the
 * editor root. We walk every text node via TreeWalker and accumulate
 * lengths, stripping ZWSPs the same way `domToRuns` does so the index
 * matches `runsToText(domToRuns(root))`.
 *
 * Returns -1 if the node isn't inside the root.
 */
export function getCharOffset(root: Element, node: Node, offset: number): number {
  // If `node` is an element (e.g., a span), `offset` counts child indices.
  // Convert to a leaf text node and an in-text offset.
  if (node.nodeType === Node.ELEMENT_NODE) {
    const el = node as Element;
    if (offset === 0) {
      // Caret before any child — char position equals chars before this elt.
      return charsBefore(root, el, 0);
    }
    const child = el.childNodes[offset - 1];
    if (!child) return charsBefore(root, el, 0);
    // Caret after `child` — chars before el's start + chars inside `child`.
    return charsBefore(root, el, 0) + textLengthOf(child);
  }
  return charsBefore(root, node, offset);
}

/**
 * Sum text-node character counts (after ZWSP stripping) that appear strictly
 * before `(target, offset)` in document order, where `target` is itself a
 * text node included partially up to `offset`.
 */
function charsBefore(root: Element, target: Node, offset: number): number {
  const doc = root.ownerDocument ?? document;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let total = 0;
  let n: Node | null = walker.nextNode();
  while (n) {
    if (n === target) {
      const piece = (n.textContent ?? '').slice(0, offset).replace(/\u200B/g, '');
      return total + piece.length;
    }
    total += ((n.textContent ?? '').replace(/\u200B/g, '')).length;
    n = walker.nextNode();
  }
  // Target wasn't a text node descendant — fall back: walk again counting
  // until we either pass `target` (an element) or exhaust the tree.
  return charsBeforeElement(root, target);
}

function charsBeforeElement(root: Element, target: Node): number {
  // Use a generic walker that visits both elements and text in DFS order.
  const doc = root.ownerDocument ?? document;
  let total = 0;
  let found = false;
  function visit(n: Node): void {
    if (found) return;
    if (n === target) {
      found = true;
      return;
    }
    if (n.nodeType === Node.TEXT_NODE) {
      total += ((n.textContent ?? '').replace(/\u200B/g, '')).length;
      return;
    }
    for (const c of Array.from(n.childNodes)) visit(c);
  }
  for (const c of Array.from(root.childNodes)) visit(c);
  // Reference doc to avoid unused-var warning under strict TS settings.
  void doc;
  return total;
}

function textLengthOf(node: Node): number {
  return ((node.textContent ?? '').replace(/\u200B/g, '')).length;
}

/** Get the caret/selection as a {start, end} character range relative to
 *  `root`. Returns null when there is no selection or the selection lies
 *  outside `root`. `start <= end` always. */
export function getCharRange(root: Element): { start: number; end: number } | null {
  const sel = (root.ownerDocument ?? document).getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) {
    return null;
  }
  const a = getCharOffset(root, range.startContainer, range.startOffset);
  const b = getCharOffset(root, range.endContainer, range.endOffset);
  return a <= b ? { start: a, end: b } : { start: b, end: a };
}

/**
 * Place the selection at character positions `[start, end]` within `root`.
 * Walks text nodes in order and binds the range to the first text node
 * whose cumulative length covers each endpoint. ZWSPs are skipped during
 * the count but counted when picking the in-node offset (so the caret
 * still has a real position to land on).
 */
export function setCharRange(root: Element, start: number, end: number): void {
  const sel = (root.ownerDocument ?? document).getSelection();
  if (!sel) return;
  const startPos = locateChar(root, start);
  const endPos = locateChar(root, end);
  if (!startPos || !endPos) return;
  const range = (root.ownerDocument ?? document).createRange();
  range.setStart(startPos.node, startPos.offset);
  range.setEnd(endPos.node, endPos.offset);
  sel.removeAllRanges();
  sel.addRange(range);
}

function locateChar(root: Element, target: number): { node: Node; offset: number } | null {
  if (target < 0) return null;
  const doc = root.ownerDocument ?? document;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let acc = 0;
  let lastNode: Node | null = null;
  let lastInNodeOffset = 0;
  let n: Node | null = walker.nextNode();
  while (n) {
    const raw = n.textContent ?? '';
    // Build a map of "visible char index" → "raw offset in node" so we can
    // place the caret on the correct side of any ZWSP padding.
    let visibleLen = 0;
    let inNodeOffset = raw.length;
    for (let i = 0; i < raw.length; i++) {
      if (raw.charCodeAt(i) === 0x200b) continue;
      if (acc + visibleLen === target) {
        inNodeOffset = i;
        break;
      }
      visibleLen += 1;
    }
    if (acc + visibleLen >= target) {
      return { node: n, offset: inNodeOffset };
    }
    acc += visibleLen;
    lastNode = n;
    lastInNodeOffset = raw.length;
    n = walker.nextNode();
  }
  if (lastNode) return { node: lastNode, offset: lastInNodeOffset };
  // Empty editor — bind to root with offset 0.
  return { node: root, offset: 0 };
}

// ── style mutation ──────────────────────────────────────────────────────

/**
 * Toggle (or set) one B/I/U flag over `[start, end)` in character space.
 * If `value` is omitted the toggle is decided by the *first* run that
 * intersects the range — if it already has the flag we clear it across the
 * whole range; otherwise we set it. That matches Word and Google Docs:
 * pressing Ctrl+B over partially-bold text turns the whole selection bold.
 *
 * Empty selection (start === end) is a no-op; the caller should fall back
 * to a block-level toggle.
 *
 * Returns a fresh runs array; the input is not mutated.
 */
export function applyStyleToRange(
  runs: DocxRun[],
  start: number,
  end: number,
  key: keyof DocxRunStyle,
  value?: boolean,
): DocxRun[] {
  if (end <= start) return runs;
  const split = splitRunsAt(splitRunsAt(runs, start), end);
  // Decide value if not given.
  let target = value;
  if (target === undefined) {
    let pos = 0;
    for (const r of split) {
      const len = r.text.length;
      if (pos + len > start && pos < end) {
        target = !r.style?.[key];
        break;
      }
      pos += len;
    }
    if (target === undefined) target = true;
  }
  // Apply to runs whose [pos, pos+len) overlaps [start, end).
  let pos = 0;
  const out: DocxRun[] = split.map((r) => {
    const len = r.text.length;
    const overlaps = pos + len > start && pos < end;
    pos += len;
    if (!overlaps) return r;
    const nextStyle: DocxRunStyle = { ...(r.style ?? {}) };
    if (target) nextStyle[key] = true;
    else delete nextStyle[key];
    return { text: r.text, style: normalizeRunStyle(nextStyle) };
  });
  return coalesceRuns(out);
}

/** Split any run that straddles char position `at` into two runs at that
 *  boundary. Existing run order and styles are preserved. */
function splitRunsAt(runs: DocxRun[], at: number): DocxRun[] {
  if (at <= 0) return runs.slice();
  const out: DocxRun[] = [];
  let pos = 0;
  for (const r of runs) {
    const len = r.text.length;
    if (pos < at && at < pos + len) {
      const cut = at - pos;
      out.push({ text: r.text.slice(0, cut), style: r.style });
      out.push({ text: r.text.slice(cut), style: r.style });
    } else {
      out.push(r);
    }
    pos += len;
  }
  return out;
}
