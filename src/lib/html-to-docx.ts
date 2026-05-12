/**
 * HTML source → `DocxBlock[]` converter.
 *
 * Sibling of `markdown-to-docx.ts`. Used by DocxEditor's「插入 HTML」command
 * so a user can paste an HTML snippet (or hand-craft one in the HTML tab and
 * cross-paste it) directly into a Word document and have it become real
 * Word blocks — headings, lists, tables, and styled inline runs.
 *
 * Coverage (HTML → DocxBlock):
 *   - `<h1>`..`<h6>`                  → kind: `heading1`..`heading6`
 *   - `<p>`                            → kind: `paragraph`, runs[] with bold/italic/underline
 *   - `<ul><li>` / `<ol><li>`         → kind: `bullet` / `numbered`, one block per item
 *   - `<table><tr><td|th>`            → kind: `table`, rows: string[][]
 *   - `<pre>`                          → one paragraph per source line (preserves whitespace)
 *   - `<blockquote>`                   → italic paragraph(s)
 *   - `<hr>`                           → skipped (no Word equivalent in our model)
 *   - `<img>`                          → paragraph carrying the alt text (no remote
 *                                         fetch — privacy + offline-safety, mirrors
 *                                         markdown-to-docx's `![alt](url)` rule)
 *   - `<script>` / `<style>` /
 *     `<noscript>` / `<template>` /
 *     `<iframe>` / `<object>` /
 *     `<embed>` / `<svg>` / `<canvas>` → skipped entirely (R317 — see SKIPPED_TAGS
 *                                         doc-block: their text children would
 *                                         otherwise leak CSS / JS source as
 *                                         paragraphs into the Word document)
 *   - free `#text` nodes at body root  → paragraph
 *
 * Inline coverage:
 *   - `<strong>` / `<b>`              → run.style.bold
 *   - `<em>` / `<i>`                  → run.style.italic
 *   - `<u>`                           → run.style.underline
 *   - `<a href>`                      → run with label text only (no per-run hyperlink
 *                                         in DocxRunStyle yet — same as markdown-to-docx)
 *   - `<br>`                          → split into adjacent runs separated by a space
 *                                         (Word's TextRun doesn't honour `\n` inside)
 *   - `<code>`                        → run with no special style (no monospace in
 *                                         DocxRunStyle yet — falls back to plain text
 *                                         rather than silently dropping content)
 *
 * Parsing strategy: DOMParser is available in the renderer (no extra dependency),
 * round-trips well-formed HTML reliably, and gracefully recovers from malformed
 * input by emitting `<parsererror>` which we detect and surface to the caller.
 * For really malformed input the parser still produces a tree we can walk —
 * worst case we get extra text nodes, never a crash.
 *
 * Each emitted block gets a fresh id via the `genId` callback the caller
 * supplies, matching markdown-to-docx's contract.
 */

import type { DocxBlock, DocxBlockKind, DocxRun, DocxRunStyle } from './docx-adapter';

interface InlineCtx {
  bold: boolean;
  italic: boolean;
  underline: boolean;
}

/**
 * R317 — markup-only / non-content tags that must be SKIPPED entirely rather
 * than walked-as-unknown-wrapper. DOMParser keeps `<script>` and `<style>`
 * bodies as a single text-node child; the default `walkBlock` / `nodeToInline
 * Runs` fall-through path would then emit that CSS / JS source as a docx
 * paragraph. `<noscript>` / `<template>` carry user-invisible markup too.
 * `<iframe>` / `<object>` / `<embed>` are media-by-reference that we can't
 * faithfully render as docx blocks anyway. Skipping all of them mirrors how
 * a browser's text-content extraction conventionally treats these tags
 * (DOMParser doesn't strip them automatically because they're legitimate
 * DOM nodes for live document rendering — we strip them at the
 * docx-conversion boundary specifically).
 */
const SKIPPED_TAGS = new Set([
  'script',
  'style',
  'noscript',
  'template',
  'iframe',
  'object',
  'embed',
  'svg',
  'canvas',
]);

function emptyStyle(ctx: InlineCtx): DocxRunStyle | undefined {
  const s: DocxRunStyle = {};
  if (ctx.bold) s.bold = true;
  if (ctx.italic) s.italic = true;
  if (ctx.underline) s.underline = true;
  return Object.keys(s).length > 0 ? s : undefined;
}

function pushRun(out: DocxRun[], text: string, ctx: InlineCtx): void {
  if (!text) return;
  const style = emptyStyle(ctx);
  // Coalesce adjacent runs that share the same style — saves bytes in the
  // serializer and keeps the runs[] list short for the DocxEditor renderer.
  const last = out[out.length - 1];
  if (last && JSON.stringify(last.style ?? null) === JSON.stringify(style ?? null)) {
    last.text += text;
    return;
  }
  out.push(style ? { text, style } : { text });
}

/**
 * Walk an inline subtree and emit DocxRun[]. Block-level descendants
 * (paragraph, heading, list, table) interrupt the inline walk — we don't
 * try to flatten them into the surrounding run stream; the caller's
 * block-walk will pick them up at the next level up.
 */
function nodeToInlineRuns(node: Node, ctx: InlineCtx, out: DocxRun[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    pushRun(out, node.textContent ?? '', ctx);
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;
  const el = node as Element;
  const tag = el.tagName.toLowerCase();
  // R317 — skip markup-only tags (script/style/noscript/template/iframe/…).
  // See SKIPPED_TAGS doc-block. Returning here means children are NOT walked,
  // which is what we want — `<script>foo()</script>`'s body is a single text
  // node we don't want pushed into the run stream.
  if (SKIPPED_TAGS.has(tag)) return;
  switch (tag) {
    case 'br':
      // Word's TextRun can't carry `\n`; insert a space as the soft-break
      // approximation. Mirrors markdown-to-docx's hard-break handling.
      pushRun(out, ' ', ctx);
      return;
    case 'strong':
    case 'b':
      el.childNodes.forEach((c) => nodeToInlineRuns(c, { ...ctx, bold: true }, out));
      return;
    case 'em':
    case 'i':
      el.childNodes.forEach((c) => nodeToInlineRuns(c, { ...ctx, italic: true }, out));
      return;
    case 'u':
      el.childNodes.forEach((c) => nodeToInlineRuns(c, { ...ctx, underline: true }, out));
      return;
    case 'a':
    case 'span':
    case 'code':
      // Strip the wrapping element, walk children with current style. Code
      // and links degrade to plain text — DocxRunStyle doesn't carry monospace
      // or per-run hyperlinks yet (Phase H+).
      el.childNodes.forEach((c) => nodeToInlineRuns(c, ctx, out));
      return;
    case 'img': {
      const alt = el.getAttribute('alt') ?? '';
      if (alt) pushRun(out, alt, ctx);
      return;
    }
    default:
      // Unknown inline-ish element — strip wrapper, keep children.
      el.childNodes.forEach((c) => nodeToInlineRuns(c, ctx, out));
      return;
  }
}

/** Plain text inside a block, with whitespace normalized (single spaces). */
function collectInlineText(el: Element): string {
  const runs: DocxRun[] = [];
  el.childNodes.forEach((c) => nodeToInlineRuns(c, { bold: false, italic: false, underline: false }, runs));
  return runs.map((r) => r.text).join('').replace(/\s+/g, ' ').trim();
}

/** Inline-walk wrapper that returns runs, used for paragraph/heading bodies. */
function collectInlineRuns(el: Element): DocxRun[] {
  const out: DocxRun[] = [];
  el.childNodes.forEach((c) => nodeToInlineRuns(c, { bold: false, italic: false, underline: false }, out));
  // Trim leading/trailing whitespace-only runs and collapse runs of whitespace
  // inside individual runs — HTML's freeform whitespace would otherwise
  // produce visible double-spaces inside the Word paragraph.
  const trimmed: DocxRun[] = [];
  for (const r of out) {
    const text = r.text.replace(/\s+/g, ' ');
    if (!text) continue;
    trimmed.push({ ...r, text });
  }
  if (trimmed.length === 0) return [];
  trimmed[0] = { ...trimmed[0], text: trimmed[0].text.replace(/^\s+/, '') };
  trimmed[trimmed.length - 1] = {
    ...trimmed[trimmed.length - 1],
    text: trimmed[trimmed.length - 1].text.replace(/\s+$/, ''),
  };
  return trimmed.filter((r) => r.text.length > 0);
}

const HEADING_KINDS: Record<string, DocxBlockKind> = {
  h1: 'heading1',
  h2: 'heading2',
  h3: 'heading3',
  h4: 'heading4',
  h5: 'heading5',
  h6: 'heading6',
};

/**
 * R332 — reverse-lookup set of the DocxBlockKind values that HEADING_KINDS
 * maps to. Used by the `<blockquote>` walker to detect whether a child
 * block is a heading (vs. plain paragraph / list / table) before
 * down-grading it to an italic paragraph. The previous code checked
 * `HEADING_KINDS[tag] === b.kind` where `tag` is the OUTER element tag
 * (`'blockquote'`), so the lookup was always `undefined` and the heading
 * branch never fired — `<blockquote><h2>X</h2><p>Y</p></blockquote>` came
 * out as「normal heading2」+「italic paragraph」, mixing styles inside
 * what the user wrote as a single quoted block. Reverse-lookup is the
 * cleanest fix — `Object.values(HEADING_KINDS)` is the authoritative list
 * (kept in lockstep with the forward map by construction). Same idiom as
 * dispatcher's `TABLE_BLOCK_KINDS` / `LIST_BLOCK_KINDS` patterns nearby.
 */
const HEADING_BLOCK_KINDS = new Set<DocxBlockKind>(Object.values(HEADING_KINDS));

/** Walk a block-level node and append zero or more DocxBlocks to `out`. */
function walkBlock(node: Node, genId: () => string, out: DocxBlock[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const txt = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
    if (txt) {
      out.push({ id: genId(), kind: 'paragraph', text: txt });
    }
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;
  const el = node as Element;
  const tag = el.tagName.toLowerCase();
  // R317 — block-level skip mirrors the inline branch above. Without this,
  // a top-level `<style>` (legal in HTML5 inside `<body>` since the
  // `scoped` proposal lapsed but browsers still parse it) lands its CSS
  // text as a Word paragraph. Same for `<script>`, `<noscript>`,
  // `<template>` and the media-by-reference family. Tags listed in
  // SKIPPED_TAGS are dropped at this layer — no children walked.
  if (SKIPPED_TAGS.has(tag)) return;

  if (HEADING_KINDS[tag]) {
    const runs = collectInlineRuns(el);
    const text = runs.map((r) => r.text).join('');
    if (!text) return;
    out.push({
      id: genId(),
      kind: HEADING_KINDS[tag],
      text,
      runs: runs.length > 1 || runs.some((r) => r.style) ? runs : undefined,
    });
    return;
  }
  if (tag === 'p') {
    const runs = collectInlineRuns(el);
    const text = runs.map((r) => r.text).join('');
    if (!text) return;
    out.push({
      id: genId(),
      kind: 'paragraph',
      text,
      runs: runs.length > 1 || runs.some((r) => r.style) ? runs : undefined,
    });
    return;
  }
  if (tag === 'ul' || tag === 'ol') {
    const itemKind: DocxBlockKind = tag === 'ul' ? 'bullet' : 'numbered';
    el.childNodes.forEach((child) => {
      if (
        child.nodeType === Node.ELEMENT_NODE &&
        (child as Element).tagName.toLowerCase() === 'li'
      ) {
        const runs = collectInlineRuns(child as Element);
        const text = runs.map((r) => r.text).join('');
        if (!text) return;
        out.push({
          id: genId(),
          kind: itemKind,
          text,
          runs: runs.length > 1 || runs.some((r) => r.style) ? runs : undefined,
        });
      }
    });
    return;
  }
  if (tag === 'table') {
    const rows: string[][] = [];
    el.querySelectorAll('tr').forEach((tr) => {
      const cells: string[] = [];
      tr.querySelectorAll('th, td').forEach((cell) => {
        cells.push(collectInlineText(cell));
      });
      if (cells.length > 0) rows.push(cells);
    });
    if (rows.length === 0) return;
    out.push({ id: genId(), kind: 'table', text: '', rows });
    return;
  }
  if (tag === 'pre') {
    // Each line in <pre> becomes its own paragraph block; preserves the
    // whitespace structure roughly (Word's TextRun can't render literal '\n').
    //
    // R393 — strip the LEADING newline per HTML spec § 4.4.3:
    //
    //   > If the contents of a pre element have a leading newline character,
    //   > that newline character is dropped when rendering the element.
    //
    // The canonical output of every markdown → HTML converter (marked,
    // remark, Pandoc, etc.) for a fenced code block ```foo``` is
    //   `<pre><code>foo\n</code></pre>`
    // — both the leading slot (`<pre>\n...`, depending on the serializer)
    // and the trailing `\n` are common. Browsers consume the leading
    // newline so the rendered code starts on its own line WITHOUT a
    // blank line above it. Pre-R393 our walker faithfully split the raw
    // text on `\n`, so a leading `\n` produced an empty leading element
    // in `lines` → first paragraph emitted was `text: ''` → user pastes
    // HTML from a markdown source and gets a blank paragraph above
    // every code block in their Word doc. The trailing `\n` ends up as
    // a trailing empty paragraph too — that's actually closer to spec
    // (browsers render `<pre>foo\n</pre>` as「foo + blank line below」)
    // so we leave it alone; only the leading-newline-stripping is what
    // HTML spec mandates.
    //
    // CRLF (`\r\n`) survives the renderer's contentEditable / clipboard
    // path on Windows occasionally; the `\r?` covers both LF and CRLF
    // leading forms. Single `\r` (old-Mac) is too rare to bother with
    // and would arrive only via deliberately crafted clipboard data.
    let raw = el.textContent ?? '';
    raw = raw.replace(/^\r?\n/, '');
    const lines = raw.split('\n');
    for (const line of lines) {
      out.push({ id: genId(), kind: 'paragraph', text: line });
    }
    return;
  }
  if (tag === 'blockquote') {
    el.childNodes.forEach((child) => {
      if (child.nodeType === Node.ELEMENT_NODE) {
        // Recursive walk for nested blocks; we re-emit them as italic
        // paragraphs by post-processing. Simplification: treat the whole
        // quote as italic paragraphs derived from the quoted block's text.
        const sub: DocxBlock[] = [];
        walkBlock(child, genId, sub);
        for (const b of sub) {
          // R332 — was `HEADING_KINDS[tag] === b.kind`, where `tag` is the
          // outer `'blockquote'` literal at this site — the map only has
          // `h1`..`h6` keys, so the comparison was always `undefined ===
          // b.kind` → false for every real block. Heading children of a
          // blockquote therefore fell through to the `else` branch and
          // were emitted as their original heading kind WITHOUT italic
          // styling, breaking the "the whole quote is italicized" invariant
          // the surrounding doc-block promises. Reverse-lookup
          // HEADING_BLOCK_KINDS (defined above the function) correctly
          // recognises every heading flavor.
          if (b.kind === 'paragraph' || HEADING_BLOCK_KINDS.has(b.kind)) {
            const runs = (b.runs ?? [{ text: b.text }]).map((r) => ({
              ...r,
              style: { ...(r.style ?? {}), italic: true } as DocxRunStyle,
            }));
            out.push({ ...b, runs, kind: 'paragraph' });
          } else {
            out.push(b);
          }
        }
      } else if (child.nodeType === Node.TEXT_NODE) {
        const txt = (child.textContent ?? '').replace(/\s+/g, ' ').trim();
        if (!txt) return;
        out.push({
          id: genId(),
          kind: 'paragraph',
          text: txt,
          runs: [{ text: txt, style: { italic: true } }],
        });
      }
    });
    return;
  }
  if (tag === 'hr') {
    // No Word equivalent in our model — skip (matches markdown-to-docx).
    return;
  }
  if (tag === 'img') {
    const alt = el.getAttribute('alt') ?? '';
    if (alt) {
      out.push({ id: genId(), kind: 'paragraph', text: alt });
    }
    return;
  }
  if (tag === 'br') {
    // Standalone <br> at block level — emit empty paragraph as visual gap.
    out.push({ id: genId(), kind: 'paragraph', text: '' });
    return;
  }
  // Unknown wrapper — descend into children (handles <div>, <article>,
  // <section>, <main>, etc. without losing their block-level descendants).
  el.childNodes.forEach((c) => walkBlock(c, genId, out));
}

export interface HtmlToDocxResult {
  blocks: DocxBlock[];
  /** Set when the HTML couldn't be parsed (DOMParser produced a `<parsererror>`).
   *  Caller surfaces this to the user instead of silently inserting nothing. */
  parseError?: string;
}

/**
 * Convert an HTML source string to a sequence of DocxBlocks ready to splice
 * into a docx model. Returns an empty array if the input is whitespace or
 * malformed beyond recovery.
 */
export function htmlToDocxBlocks(source: string, genId: () => string): HtmlToDocxResult {
  const trimmed = source.trim();
  if (!trimmed) return { blocks: [] };
  // Wrap fragments in a body so DOMParser produces a stable root. Full
  // documents (with their own <html><body>) parse just as cleanly because
  // we read from `body.childNodes` regardless.
  const wrapped = /<body\b/i.test(trimmed)
    ? trimmed
    : `<html><body>${trimmed}</body></html>`;
  const doc = new DOMParser().parseFromString(wrapped, 'text/html');
  const errEl = doc.querySelector('parsererror');
  if (errEl) {
    return {
      blocks: [],
      parseError: errEl.textContent ?? 'HTML parse error',
    };
  }
  const body = doc.body;
  if (!body) return { blocks: [] };
  const out: DocxBlock[] = [];
  body.childNodes.forEach((n) => walkBlock(n, genId, out));
  return { blocks: out };
}
