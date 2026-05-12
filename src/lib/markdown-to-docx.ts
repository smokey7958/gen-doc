/**
 * Markdown source → `DocxBlock[]` converter.
 *
 * Used by the DocxEditor's "插入 Markdown" command so a user can paste a
 * markdown snippet (or whole AI-drafted section) directly into a Word
 * document and have it become real Word blocks — headings, lists, tables,
 * and styled inline runs — instead of a raw text dump.
 *
 * Why an AST walk over `marked.lexer` rather than HTML round-trip:
 *   - We already have `marked` as a dep (markdown preview), so no new bytes.
 *   - The token tree is a much better fit for our paragraph-flat block
 *     model than `htmlToBlocks` (which has to undo `marked`'s nesting just
 *     to produce the same block list, with edge-case escapes that bite).
 *   - We need to attach inline B/I via `DocxRun.style`, which means walking
 *     `strong`/`em`/`codespan` etc. anyway — doing it during the source
 *     walk avoids materialising the intermediate HTML.
 *
 * Coverage (what becomes what in the docx model):
 *   - `# Heading 1..6` → kind: `heading1`..`heading6`
 *   - paragraph        → kind: `paragraph`, runs[] with bold/italic from `**`/`*`/`_`
 *   - `- ` / `* `      → kind: `bullet`, one block per list item
 *   - `1. `            → kind: `numbered`, one block per list item
 *   - fenced ``` code  → one paragraph per source line, monospace block style
 *   - `> ` blockquote  → italic paragraph(s); our model has no native quote
 *   - `| a | b |` GFM  → kind: `table`, rows: string[][] (cell inline marks
 *                        flattened to plain text — table cells don't carry
 *                        run-level styling in the DocxBlock model yet)
 *   - `---` hr         → skipped (no Word equivalent in our model)
 *   - `![alt](url)`    → paragraph carrying the alt text; we don't fetch
 *                        remote images at insert time (privacy + offline-
 *                        safety) and data-URI inlining is rarer than worth
 *                        the bytes for an MVP
 *
 * Inline coverage:
 *   - `**bold**`, `__bold__`            → run.style.bold
 *   - `*em*`, `_em_`                    → run.style.italic
 *   - `` `code` ``                      → run with no special style (inline
 *                                         font swap requires per-run fontFamily,
 *                                         which our DocxRunStyle doesn't carry —
 *                                         falls back to plain text rather than
 *                                         silently dropping the content)
 *   - `[label](url)`                    → run with `label` text only (per-run
 *                                         hyperlinks are a Phase H+ feature in
 *                                         the docx adapter)
 *   - `~~del~~`                         → run with no special style (no
 *                                         strikethrough in DocxRunStyle yet)
 *   - hard break (`  \n` or `\\n`)      → split into adjacent runs with a
 *                                         space; Word's TextRun won't honour
 *                                         a `\n` inside the text
 *
 * Each emitted block gets a fresh id — caller is responsible for splicing
 * into the model alongside their own id space.
 */

import { marked, type Tokens } from 'marked';
import type { DocxBlock, DocxBlockKind, DocxRun, DocxRunStyle } from './docx-adapter';

interface InlineCtx {
  bold: boolean;
  italic: boolean;
}

/** Recursively flatten an inline token tree into `DocxRun[]`. */
function tokensToRuns(tokens: Tokens.Generic[] | undefined, ctx: InlineCtx): DocxRun[] {
  if (!tokens || tokens.length === 0) return [];
  const out: DocxRun[] = [];
  for (const tok of tokens) {
    switch (tok.type) {
      case 'text': {
        // marked sometimes nests further `text` tokens (with .tokens) and
        // sometimes emits a leaf (.text only). The leaf carries entity-decoded
        // text already.
        const child = (tok as Tokens.Text).tokens;
        if (child && child.length > 0) {
          out.push(...tokensToRuns(child, ctx));
        } else {
          pushRun(out, (tok as Tokens.Text).text, ctx);
        }
        break;
      }
      case 'strong': {
        out.push(...tokensToRuns((tok as Tokens.Strong).tokens, { ...ctx, bold: true }));
        break;
      }
      case 'em': {
        out.push(...tokensToRuns((tok as Tokens.Em).tokens, { ...ctx, italic: true }));
        break;
      }
      case 'codespan':
        // No per-run fontFamily in DocxRunStyle — drop to plain text rather
        // than losing the content.
        pushRun(out, (tok as Tokens.Codespan).text, ctx);
        break;
      case 'del':
        // No strikethrough on DocxRunStyle; preserve the text without the mark.
        out.push(...tokensToRuns((tok as Tokens.Del).tokens, ctx));
        break;
      case 'link':
        // Per-run links not in MVP docx model — preserve the visible label.
        out.push(...tokensToRuns((tok as Tokens.Link).tokens, ctx));
        break;
      case 'br':
        // Word's TextRun won't render a literal \n. Adjacent runs already
        // separate visually; collapse to a space so words don't run together.
        pushRun(out, ' ', ctx);
        break;
      case 'image': {
        // Same MVP rationale as link — preserve the alt text inline.
        const t = tok as Tokens.Image;
        if (t.text) pushRun(out, t.text, ctx);
        break;
      }
      case 'escape':
        pushRun(out, (tok as Tokens.Escape).text, ctx);
        break;
      case 'html': {
        // Unrecognised inline HTML — keep as plain text. The user almost
        // certainly didn't mean to nest raw HTML inside a Word paste.
        const raw = (tok as Tokens.HTML).text ?? '';
        pushRun(out, raw, ctx);
        break;
      }
      default: {
        // Fallback — unknown inline token type. Try to recover any nested
        // text so we don't drop content silently.
        const generic = tok as { text?: string; tokens?: Tokens.Generic[] };
        if (generic.tokens) out.push(...tokensToRuns(generic.tokens, ctx));
        else if (typeof generic.text === 'string') pushRun(out, generic.text, ctx);
      }
    }
  }
  return out;
}

/**
 * Append a `DocxRun` to `out`, coalescing with the trailing run when the
 * style matches — produces tighter run lists which the docx serializer can
 * write as fewer `<w:r>` elements. The simple style equality check is
 * sufficient because `DocxRunStyle` is a flat boolean record.
 */
function pushRun(out: DocxRun[], text: string, ctx: InlineCtx): void {
  if (text.length === 0) return;
  const style: DocxRunStyle | undefined =
    ctx.bold || ctx.italic ? { ...(ctx.bold ? { bold: true } : {}), ...(ctx.italic ? { italic: true } : {}) } : undefined;
  const last = out[out.length - 1];
  if (last && runStylesEqual(last.style, style)) {
    last.text += text;
    return;
  }
  out.push(style ? { text, style } : { text });
}

function runStylesEqual(a: DocxRunStyle | undefined, b: DocxRunStyle | undefined): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return !!a.bold === !!b.bold && !!a.italic === !!b.italic && !!a.underline === !!b.underline;
}

/** Concatenate all run.text fields. Caller usually keeps `runs` as the
 *  source-of-truth and uses this only to populate `block.text` for callers
 *  that read plain text. */
function runsToText(runs: DocxRun[]): string {
  return runs.map((r) => r.text).join('');
}

/** Build a paragraph/heading/list block from a token's inline children. */
function makeRichBlock(
  id: string,
  kind: DocxBlockKind,
  inlineTokens: Tokens.Generic[] | undefined,
): DocxBlock {
  const runs = tokensToRuns(inlineTokens, { bold: false, italic: false });
  return {
    id,
    kind,
    text: runsToText(runs),
    runs: runs.length > 0 ? runs : undefined,
  };
}

/**
 * Walk a list token, emitting one block per item. Nested lists are
 * flattened (the docx model has no per-block indentation), but inline
 * styling on each item survives. Task list checkboxes (`- [x]`) are
 * preserved as a `[x] ` / `[ ] ` text prefix because DocxRunStyle has no
 * checkbox affordance.
 */
function emitListBlocks(
  list: Tokens.List,
  out: DocxBlock[],
  newId: () => string,
): void {
  const kind: DocxBlockKind = list.ordered ? 'numbered' : 'bullet';
  for (const item of list.items) {
    const block = makeRichBlock(newId(), kind, item.tokens);
    if (item.task) {
      const prefix = item.checked ? '[x] ' : '[ ] ';
      // Prepend to the first run so styling on the rest of the line stays put.
      if (block.runs && block.runs.length > 0) {
        block.runs = [{ text: prefix }, ...block.runs];
        block.text = prefix + block.text;
      } else {
        block.runs = [{ text: prefix }];
        block.text = prefix;
      }
    }
    out.push(block);
    // Recursively expand nested lists / sub-tokens that produced their own
    // block-level children. marked stuffs nested lists inside `item.tokens`.
    if (item.tokens) {
      for (const sub of item.tokens) {
        if (sub.type === 'list') {
          emitListBlocks(sub as Tokens.List, out, newId);
        }
      }
    }
  }
}

/**
 * Convert markdown source text to a flat list of `DocxBlock`. Caller
 * supplies `newId()` so the new blocks share id-space conventions with the
 * surrounding model (avoids collisions with the editor's own counter).
 *
 * Empty / whitespace-only input returns an empty array — caller decides
 * whether that's a no-op or a "nothing to insert" toast.
 */
export function markdownToDocxBlocks(source: string, newId: () => string): DocxBlock[] {
  const trimmed = source.trim();
  if (trimmed.length === 0) return [];
  const tokens = marked.lexer(trimmed);
  const out: DocxBlock[] = [];
  for (const tok of tokens) {
    switch (tok.type) {
      case 'heading': {
        const t = tok as Tokens.Heading;
        const depth = Math.min(Math.max(t.depth, 1), 6) as 1 | 2 | 3 | 4 | 5 | 6;
        out.push(makeRichBlock(newId(), `heading${depth}` as DocxBlockKind, t.tokens));
        break;
      }
      case 'paragraph': {
        out.push(makeRichBlock(newId(), 'paragraph', (tok as Tokens.Paragraph).tokens));
        break;
      }
      case 'list':
        emitListBlocks(tok as Tokens.List, out, newId);
        break;
      case 'code': {
        // Each source line becomes its own paragraph so long blocks wrap
        // cleanly inside the page. Style.fontFamily marks them monospace —
        // there's no run-level fontFamily in our model, so block-level is
        // the only signal we can give.
        const t = tok as Tokens.Code;
        const lines = t.text.split('\n');
        for (const line of lines) {
          const block: DocxBlock = {
            id: newId(),
            kind: 'paragraph',
            text: line,
            runs: line.length > 0 ? [{ text: line }] : undefined,
            style: { fontFamily: 'Consolas' },
          };
          out.push(block);
        }
        break;
      }
      case 'blockquote': {
        // Recurse: the contained tokens are themselves block-level. We
        // re-walk and apply italic at the block level (model has no quote
        // affordance).
        const inner = markdownToDocxBlocks(
          (tok as Tokens.Blockquote).text,
          newId,
        );
        for (const b of inner) {
          b.style = { ...(b.style ?? {}), italic: true };
          out.push(b);
        }
        break;
      }
      case 'table': {
        const t = tok as Tokens.Table;
        const rows: string[][] = [];
        // Header row first — same convention as the toolbar's "插入表格"
        // template. Inline marks inside cells flatten to plain text because
        // DocxBlock.rows is a string[][].
        rows.push(t.header.map((cell) => runsToText(tokensToRuns(cell.tokens, { bold: false, italic: false }))));
        for (const row of t.rows) {
          rows.push(row.map((cell) => runsToText(tokensToRuns(cell.tokens, { bold: false, italic: false }))));
        }
        out.push({ id: newId(), kind: 'table', text: '', rows });
        break;
      }
      case 'hr':
      case 'space':
        // hr has no equivalent in the block model; space is just blank lines
        // between blocks (already handled by the outer loop).
        break;
      case 'html': {
        // Raw HTML at the block level — best-effort: emit as a paragraph so
        // the user can clean it up post-paste. Stripping it would lose data.
        const raw = (tok as Tokens.HTML).text ?? '';
        if (raw.trim().length > 0) {
          out.push({ id: newId(), kind: 'paragraph', text: raw, runs: [{ text: raw }] });
        }
        break;
      }
      default: {
        // Unknown block-level token — try to recover its `.text`.
        const generic = tok as { text?: string };
        if (typeof generic.text === 'string' && generic.text.trim().length > 0) {
          out.push({
            id: newId(),
            kind: 'paragraph',
            text: generic.text,
            runs: [{ text: generic.text }],
          });
        }
      }
    }
  }
  return out;
}
