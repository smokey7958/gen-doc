/**
 * ChangeSet apply / inverse / undo helpers — works on the in-memory tab list
 * provided by the workspace store. Each ChangeOp carries `before` and `after`
 * so undo is a simple swap (spec §7.2).
 */

import type { ChangeSet, ChangeOp } from '../types/changeset';
import { type Tab, tabFromDescriptor } from '../types/tab';

export interface ApplyResult {
  tabs: Tab[];
  /** Description used for the undo toast. */
  description: string;
}

/** Apply forward — produce a new tabs array. */
export function applyChangeset(tabs: Tab[], cs: ChangeSet): ApplyResult {
  let next = tabs;
  for (const op of cs.ops) {
    next = applyOp(next, op, /* forward */ true);
  }
  return { tabs: next, description: cs.description };
}

/** Apply inverse (undo) — flip before/after on every op. */
export function undoChangeset(tabs: Tab[], cs: ChangeSet): ApplyResult {
  let next = tabs;
  for (let i = cs.ops.length - 1; i >= 0; i--) {
    next = applyOp(next, cs.ops[i], /* forward */ false);
  }
  return { tabs: next, description: `Undo: ${cs.description}` };
}

function applyOp(tabs: Tab[], op: ChangeOp, forward: boolean): Tab[] {
  switch (op.type) {
    case 'md_text': {
      const value = forward ? op.after : op.before;
      return tabs.map((t) =>
        t.id === op.tabId && t.type === 'markdown' ? { ...t, content: value, dirty: true } : t,
      );
    }
    case 'tab_create': {
      if (forward) {
        // Add the tab if not present.
        if (tabs.some((t) => t.id === op.tabId)) return tabs;
        // R347 — was a hand-rolled `op.tab.type === 'markdown' ? text : data
        // as Tab` branch that silently fell through to the binary `data`
        // shape for ANY non-markdown type. The cast worked for docx/xlsx/
        // pptx (all carry `data: Uint8Array`) but produced a structurally
        // invalid HtmlTab when `op.tab.type === 'html'` — HtmlTab declares
        // `content: string`, not `data`, so the resulting tab would carry
        // a phantom `data` field that HtmlEditor doesn't read and a
        // missing `content` field that it does. CM6's
        // `EditorState.create({doc: undefined})` survives via JS's
        // implicit-empty coercion, so the bug surfaces as「I created an
        // HTML tab via this future op, and it opens blank」 rather than a
        // crash — exactly the kind of dormant silent-corruption that
        // tab_create's lack of constructors today doesn't surface but
        // would fire the moment any new dispatcher tool emits this op.
        // tabFromDescriptor (types/tab.ts:64) already has the exhaustive
        // type-switch + `as never` exhaustiveness check; reusing it keeps
        // the construction shape in lockstep with `loadFromOpened`'s
        // sibling path (both build a Tab from a TabDescriptor + bytes).
        // `dirty: true` after the helper because tab_create implies "new
        // content the user hasn't saved yet", overriding the helper's
        // default `dirty: false` (which is correct for `loadFromOpened`
        // — the user opened a file that's still in sync with disk).
        const fresh = tabFromDescriptor(op.tab, op.data);
        return [...tabs, { ...fresh, order: tabs.length, dirty: true }];
      }
      // inverse: remove
      return tabs.filter((t) => t.id !== op.tabId).map((t, i) => ({ ...t, order: i }));
    }
    case 'tab_delete': {
      if (forward) {
        return tabs.filter((t) => t.id !== op.tabId).map((t, i) => ({ ...t, order: i }));
      }
      // inverse: re-add. R347 — same tabFromDescriptor fix as tab_create
      // sibling; sibling reasoning applies symmetrically.
      const fresh = tabFromDescriptor(op.tab, op.data);
      return [...tabs, { ...fresh, order: tabs.length, dirty: true }];
    }
    case 'binary_replace': {
      const value = forward ? op.after : op.before;
      return tabs.map((t) => {
        if (t.id !== op.tabId) return t;
        // R347 — also extend the "text-shaped tabs don't carry `data`" skip
        // to html. Previous `t.type === 'markdown'` check was technically a
        // defensive no-op (the dispatcher only emits binary_replace for
        // docx/xlsx/pptx so this branch was a "shouldn't happen" guard),
        // but with HTML added as a second text-shaped type the guard's
        // condition is incomplete: a future tool emitting binary_replace
        // on an html tab would fall to the `data: value` line and add a
        // phantom data field on an HtmlTab whose content went unchanged.
        // Skip the same way for both text-formats — they both have
        // `content: string` and no `data` slot to assign.
        if (t.type === 'markdown' || t.type === 'html') return t; // shouldn't happen
        return { ...t, data: value, dirty: true };
      });
    }
    case 'word_paragraph':
    case 'excel_cell':
    case 'pptx_text':
      // These typed ops are kept in the schema so the spec stays sound, but
      // we never actually emit them — the dispatcher always lowers them to a
      // binary_replace once it has the new bytes in hand. Treat as no-op.
      return tabs;
    default: {
      const _exhaustive: never = op;
      throw new Error(`Unknown ChangeOp type: ${(_exhaustive as ChangeOp).type}`);
    }
  }
}
