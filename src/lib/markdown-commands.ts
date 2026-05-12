/**
 * Markdown editing commands shared by `MarkdownToolbar` (button clicks) and
 * `MarkdownEditor` (keyboard shortcuts). Each command operates on a CodeMirror
 * 6 EditorView and dispatches a transaction.
 *
 * Command shape: `(view: EditorView) => boolean`. Returning `true` swallows
 * the keystroke (so the default Enter / typed-char behaviour doesn't also
 * fire); `false` means we didn't handle this case — let CM6 fall through.
 */

import { EditorSelection, type ChangeSpec } from '@codemirror/state';
import type { EditorView } from '@codemirror/view';

/** Wrap each non-empty selection with `before`/`after`; cursor only → insert + place caret between. */
export function wrapSelection(view: EditorView, before: string, after: string): boolean {
  const { state } = view;
  const changes: ChangeSpec[] = [];
  const ranges: ReturnType<typeof EditorSelection.range>[] = [];
  for (const range of state.selection.ranges) {
    if (range.empty) {
      changes.push({ from: range.from, insert: before + after });
      ranges.push(EditorSelection.cursor(range.from + before.length));
    } else {
      const text = state.sliceDoc(range.from, range.to);
      changes.push({ from: range.from, to: range.to, insert: before + text + after });
      ranges.push(
        EditorSelection.range(
          range.from + before.length,
          range.from + before.length + text.length,
        ),
      );
    }
  }
  view.dispatch({
    changes,
    selection: EditorSelection.create(ranges, 0),
  });
  return true;
}

/**
 * Replace any existing markdown line prefix (heading / list / quote) on each
 * selected line with `prefix`, or strip it if already that prefix.
 *
 * Toggle detection compares the *matched* prefix to the requested one — a
 * naïve `startsWith(prefix)` would mis-fire on deeper heading levels (e.g.
 * `'### foo'.startsWith('## ')` is true, so converting H3 → H2 would strip
 * the heading entirely instead of demoting it).
 *
 * Recognised line prefixes: ATX heading `# `..`###### `, GFM task list
 * `- [ ] ` / `- [x] ` / `- [X] `, plain bullets `- ` / `* ` / `+ `,
 * ordered list `N. `, blockquote `> `. The checkbox alternative MUST
 * appear before plain `- ` in the alternation — otherwise `- [ ] task`
 * matches the bare `- ` first, leaving `[ ] task` as the "stripped" body
 * and producing `# [ ] task` when promoted to H1 (the user expected
 * `# task`). Same logic for `* ` and `+ ` bullets which `smartListContinue`
 * already accepts: keeping the prefix list in sync means converting
 * imported markdown (Word→md, Notion exports) to a heading no longer
 * leaves orphan bullet characters.
 */
export function setLinePrefix(view: EditorView, prefix: string): boolean {
  const { state } = view;
  const changes: ChangeSpec[] = [];
  const seenLines = new Set<number>();
  const prefixRe = /^(#{1,6} |- \[[ xX]\] |- |\* |\+ |\d+\. |> )/;
  for (const range of state.selection.ranges) {
    const fromLine = state.doc.lineAt(range.from).number;
    const toLine = state.doc.lineAt(range.to).number;
    for (let n = fromLine; n <= toLine; n++) {
      if (seenLines.has(n)) continue;
      seenLines.add(n);
      const line = state.doc.line(n);
      const m = line.text.match(prefixRe);
      const currentPrefix = m ? m[0] : '';
      const stripped = currentPrefix ? line.text.slice(currentPrefix.length) : line.text;
      const final = currentPrefix === prefix ? stripped : prefix + stripped;
      changes.push({ from: line.from, to: line.to, insert: final });
    }
  }
  view.dispatch({ changes });
  return true;
}

/**
 * Insert a markdown link `[label](url)` at the current selection. Pure — no
 * UI side effects. The host (MarkdownEditor) owns the dialog that collects
 * `url` / `label`; replacing this command's previous `window.prompt` flow
 * was the whole point of the refactor (modal browser dialogs are jarring,
 * don't honour the app theme, and offer no way to also edit the label in
 * the same step). Selection text becomes the label by default; an empty
 * selection inserts the placeholder so users can immediately overtype.
 */
export function applyLink(view: EditorView, url: string, label?: string): void {
  const { state } = view;
  const range = state.selection.main;
  const text = label && label.length > 0 ? label : '連結文字';
  view.dispatch({
    changes: { from: range.from, to: range.to, insert: `[${text}](${url})` },
    selection: { anchor: range.from + 1, head: range.from + 1 + text.length },
  });
}

/** Same as applyLink but produces a markdown image `![alt](url)`. */
export function applyImage(view: EditorView, url: string, alt?: string): void {
  const range = view.state.selection.main;
  const a = alt && alt.length > 0 ? alt : '圖片描述';
  view.dispatch({ changes: { from: range.from, to: range.to, insert: `![${a}](${url})` } });
}

export function insertCodeBlock(view: EditorView): boolean {
  const { state } = view;
  const range = state.selection.main;
  const sel = state.sliceDoc(range.from, range.to);
  const body = sel || 'code';
  const block = `\n\`\`\`\n${body}\n\`\`\`\n`;
  view.dispatch({
    changes: { from: range.from, to: range.to, insert: block },
    selection: { anchor: range.from + 4, head: range.from + 4 },
  });
  return true;
}

export function insertTable(view: EditorView): boolean {
  const range = view.state.selection.main;
  const tpl = [
    '| 欄位 1 | 欄位 2 | 欄位 3 |',
    '| --- | --- | --- |',
    '| a | b | c |',
    '| d | e | f |',
    '',
  ].join('\n');
  view.dispatch({ changes: { from: range.from, to: range.to, insert: tpl } });
  return true;
}

export function insertHr(view: EditorView): boolean {
  const range = view.state.selection.main;
  view.dispatch({ changes: { from: range.from, to: range.to, insert: '\n---\n' } });
  return true;
}

/**
 * Smart list / quote continuation. Bound to Enter.
 *
 * Rules:
 *  - On a non-empty list item (`- foo` / `1. foo` / `- [ ] foo` / `> foo`),
 *    inserting Enter adds a fresh prefix on the new line so the user can
 *    keep typing without re-typing `- `.
 *  - On an *empty* list item (just the prefix, no content), Enter strips
 *    the prefix and exits the list — matches the universal Markdown editor
 *    convention (Obsidian / Typora / Notion).
 *  - Ordered lists auto-increment: pressing Enter after `3. foo` produces
 *    `4. `.
 *  - Returning false lets the default newline insertion proceed (so plain
 *    paragraphs Enter normally).
 */
export function smartListContinue(view: EditorView): boolean {
  const { state } = view;
  const sel = state.selection.main;
  if (!sel.empty) return false;
  const line = state.doc.lineAt(sel.from);
  const text = line.text;

  // Match: optional indent, then bullet/numbered/checkbox/quote prefix
  const m = text.match(/^(\s*)(- \[[ xX]\] |- |\* |\+ |(\d+)\. |> )(.*)$/);
  if (!m) return false;
  const [, indent, prefix, num, body] = m;

  // Empty item → strip the prefix (exit list)
  if (body.trim() === '') {
    view.dispatch({
      changes: { from: line.from, to: line.to, insert: '' },
      selection: { anchor: line.from },
    });
    return true;
  }

  // Non-empty item → insert newline + prefix
  let nextPrefix = prefix;
  if (num) {
    nextPrefix = `${Number(num) + 1}. `;
  } else if (prefix.startsWith('- [')) {
    // Reset checkbox state
    nextPrefix = '- [ ] ';
  }
  const insert = `\n${indent}${nextPrefix}`;
  view.dispatch({
    changes: { from: sel.from, insert },
    selection: { anchor: sel.from + insert.length },
  });
  return true;
}
