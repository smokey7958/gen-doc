/**
 * ChangeSet — see docs/MVP-SPEC.md §7.2.
 * The atomic unit AI tools emit; doc engine apply / undo work at this level.
 */

import type { TabDescriptor } from './manifest';

export type ChangeOrigin = 'ai' | 'manual';

export interface ParagraphData {
  text: string;
  style?: Record<string, unknown>;
}

export type CellValue = string | number | boolean | null;

export interface MdTextOp {
  tabId: string;
  type: 'md_text';
  before: string;
  after: string;
  /** Optional [start, end] character range; omit to replace whole doc. */
  range?: [number, number];
}

export interface WordParagraphOp {
  tabId: string;
  type: 'word_paragraph';
  paraIndex: number;
  before: ParagraphData;
  after: ParagraphData;
}

export interface ExcelCellOp {
  tabId: string;
  type: 'excel_cell';
  sheet: string;
  /** A1 notation, e.g. "B7". */
  address: string;
  before: CellValue;
  after: CellValue;
}

export interface PptxTextOp {
  tabId: string;
  type: 'pptx_text';
  slideIndex: number;
  shapeId: string;
  before: string;
  after: string;
}

/**
 * Atomic byte-level replacement of a binary-format tab's `data` field
 * (docx / xlsx / pptx — the three tab kinds whose runtime representation
 * is `Uint8Array`, in contrast to markdown / html which store `content:
 * string`). Built by the dispatcher after parse → mutate → re-serialize.
 * apply / undo are byte swaps. Carries a human-readable `description` so
 * DiffPreview can show what changed without re-parsing.
 *
 * R362 — was「non-markdown tab's data」 which was correct pre-HTML but
 * mislead after html was added as a second text-format tab kind. HTML
 * tabs have `content: string` (same shape as markdown), not `data:
 * Uint8Array`; the changeset-apply path at changeset-apply.ts:88-89
 * already excludes BOTH text-format kinds from binary_replace via
 * `if (t.type === 'markdown' || t.type === 'html') return t` (R347).
 * Updating this doc-comment closes the doc/code consistency gap —
 * future readers see the same「binary format = docx/xlsx/pptx」 mental
 * model the apply code already enforces.
 */
export interface BinaryReplaceOp {
  tabId: string;
  type: 'binary_replace';
  before: Uint8Array;
  after: Uint8Array;
  /** Free-form summary of what changed (used in DiffPreview & undo). */
  description: string;
}

export interface TabCreateOp {
  tabId: string;
  type: 'tab_create';
  tab: TabDescriptor;
  data: Uint8Array;
}

export interface TabDeleteOp {
  tabId: string;
  type: 'tab_delete';
  tab: TabDescriptor;
  data: Uint8Array;
}

export type ChangeOp =
  | MdTextOp
  | WordParagraphOp
  | ExcelCellOp
  | PptxTextOp
  | BinaryReplaceOp
  | TabCreateOp
  | TabDeleteOp;

export interface ChangeSet {
  id: string;
  origin: ChangeOrigin;
  /** ISO-8601 timestamp. */
  createdAt: string;
  /** Human-readable summary, used in undo toasts. */
  description: string;
  ops: ChangeOp[];
}
