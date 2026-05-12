/**
 * Excel-side clipboard helpers — TSV in/out plus a private rich payload
 * (cells with style + formula) so in-app round-trip preserves everything
 * SheetJS would otherwise drop.
 *
 * Design:
 *   - `serializeRangeToTsv` → text/plain, lets the user paste into Excel /
 *      Google Sheets / a code editor.
 *   - `serializeRangeToJson` → application/x-gendoc-xlsx, used when both
 *      ends are us so styles + formulas (with relative-ref shifting)
 *      survive.
 *   - `parseTsv` is the fallback when only text/plain is present.
 *   - `applyPaste` is the workhorse: it takes the active sheet, the anchor
 *      coordinate, and a payload, and returns a new sheet with the paste
 *      applied. The caller picks between rich JSON / TSV beforehand.
 */

import type { XlsxCell, XlsxSheet } from './xlsx-adapter';
import { shiftFormula } from './xlsx-formula';

/** Mime type for our private payload. Custom types survive ClipboardEvent. */
export const GENDOC_XLSX_MIME = 'application/x-gendoc-xlsx';

export interface RichClipboardPayload {
  /** Origin (top-left) of the copied rectangle, used for relative-ref shifting. */
  origin: { r: number; c: number };
  /** Cells indexed [row][col] starting at the origin. */
  cells: XlsxCell[][];
}

/** Extract a cell rectangle from the sheet (inclusive corners). */
export function extractRange(sheet: XlsxSheet, r1: number, c1: number, r2: number, c2: number): XlsxCell[][] {
  const out: XlsxCell[][] = [];
  for (let r = r1; r <= r2; r += 1) {
    const row: XlsxCell[] = [];
    for (let c = c1; c <= c2; c += 1) {
      const src = sheet.cells[r]?.[c];
      // Snapshot the cell so the clipboard payload doesn't alias the model.
      row.push(src ? { ...src, style: src.style ? { ...src.style } : undefined } : { text: '' });
    }
    out.push(row);
  }
  return out;
}

/**
 * TSV serialization. Excel's interchange format: rows on `\n`, cells on
 * `\t`. Cells containing tab / newline / quote get wrapped in double
 * quotes with `""` escape — same as CSV-in-TSV-clothes that Excel emits.
 *
 * For formula cells we emit the *formula source* (not the computed text)
 * so pasting back into Excel reconstructs the formula. This loses the
 * cached value for non-Gen-Doc destinations, but Excel recomputes anyway.
 */
export function serializeRangeToTsv(cells: XlsxCell[][]): string {
  const rows: string[] = [];
  for (const row of cells) {
    const out: string[] = [];
    for (const cell of row) {
      // Prefer the formula source so round-trip into Excel works.
      const raw = cell.formula ?? cell.text ?? '';
      out.push(escapeTsvCell(raw));
    }
    rows.push(out.join('\t'));
  }
  return rows.join('\n');
}

function escapeTsvCell(s: string): string {
  if (s.includes('\t') || s.includes('\n') || s.includes('"')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

/** Rich payload as JSON — paired with TSV on the system clipboard. */
export function serializeRangeToJson(payload: RichClipboardPayload): string {
  return JSON.stringify(payload);
}

/**
 * Parse Excel-style TSV / CSV-ish text. Recognises the same quoting rules
 * we emit. Always returns at least a 1×1 grid.
 */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let inQuotes = false;
  let i = 0;
  const len = text.length;
  while (i < len) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      cell += ch;
      i += 1;
      continue;
    }
    if (ch === '"') {
      // Open quote — only at start of cell. If we're already mid-cell,
      // treat as literal (Excel does the same).
      if (cell === '') {
        inQuotes = true;
        i += 1;
        continue;
      }
      cell += ch;
      i += 1;
      continue;
    }
    if (ch === '\t') {
      row.push(cell);
      cell = '';
      i += 1;
      continue;
    }
    if (ch === '\r') {
      // Swallow — handled with following \n or treated as line end.
      i += 1;
      continue;
    }
    if (ch === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      i += 1;
      continue;
    }
    cell += ch;
    i += 1;
  }
  // Final cell / row.
  if (cell !== '' || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  // Normalise to a rectangular shape — pad short rows with empty strings.
  const cols = rows.reduce((m, r) => Math.max(m, r.length), 0);
  for (const r of rows) while (r.length < cols) r.push('');
  return rows.length > 0 ? rows : [['']];
}

/**
 * Apply a paste payload starting at (anchorR, anchorC). Returns a new
 * XlsxSheet — caller is responsible for plugging it back into the model.
 *
 * If the payload extends past the current sheet bounds we grow rowCount
 * / colCount accordingly so the user doesn't lose data; growing also
 * matches what Excel does (pastes silently extend the used range).
 *
 * `payload` may carry a different `origin` than the destination anchor —
 * that delta is what drives `shiftFormula` for any formula cells in the
 * payload, so relative refs follow the move.
 */
export function applyPaste(
  sheet: XlsxSheet,
  anchorR: number,
  anchorC: number,
  payload: RichClipboardPayload,
): XlsxSheet {
  const rows = payload.cells.length;
  const cols = payload.cells[0]?.length ?? 0;
  if (rows === 0 || cols === 0) return sheet;

  const dr = anchorR - payload.origin.r;
  const dc = anchorC - payload.origin.c;

  const newRowCount = Math.max(sheet.rowCount, anchorR + rows);
  const newColCount = Math.max(sheet.colCount, anchorC + cols);

  // Deep-copy existing cells, growing the grid where needed so we never
  // alias rows/cells with the previous model.
  const cells: XlsxCell[][] = [];
  for (let r = 0; r < newRowCount; r += 1) {
    const row: XlsxCell[] = [];
    for (let c = 0; c < newColCount; c += 1) {
      const src = sheet.cells[r]?.[c];
      row.push(src ? { ...src } : { text: '' });
    }
    cells.push(row);
  }

  // Stamp the payload, shifting formula refs.
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const srcCell = payload.cells[r][c];
      const tr = anchorR + r;
      const tc = anchorC + c;
      if (tr >= newRowCount || tc >= newColCount) continue;
      const next: XlsxCell = {
        text: srcCell.text ?? '',
        ...(srcCell.style ? { style: { ...srcCell.style } } : {}),
        ...(srcCell.rawType ? { rawType: srcCell.rawType } : {}),
      };
      if (srcCell.formula) {
        next.formula = shiftFormula(srcCell.formula, dr, dc);
        // Echo the source into text so the recompute pass that follows
        // the paste has a sensible placeholder; recompute will overwrite.
        next.text = next.formula;
      }
      cells[tr][tc] = next;
    }
  }

  return { ...sheet, cells, rowCount: newRowCount, colCount: newColCount };
}

/**
 * Blank cells in [r1..r2, c1..c2] that don't fall inside the protected
 * rectangle [pr1..pr2, pc1..pc2]. Used by cut+paste so a paste-in-place
 * doesn't immediately erase the freshly-pasted cells. Style is preserved.
 */
export function clearRangeExcept(
  sheet: XlsxSheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  pr1: number,
  pc1: number,
  pr2: number,
  pc2: number,
): XlsxSheet {
  const cells = sheet.cells.map((row, r) => {
    if (r < r1 || r > r2) return row;
    return row.map((cell, c) => {
      if (c < c1 || c > c2) return cell;
      // Inside protected rect? Leave as-is.
      if (r >= pr1 && r <= pr2 && c >= pc1 && c <= pc2) return cell;
      const { formula: _f, ...rest } = cell;
      void _f;
      return { ...rest, text: '' };
    });
  });
  return { ...sheet, cells };
}

/** Clear (blank-out) cells in a rectangle. Style is preserved. */
export function clearRange(sheet: XlsxSheet, r1: number, c1: number, r2: number, c2: number): XlsxSheet {
  const cells = sheet.cells.map((row, r) =>
    r < r1 || r > r2
      ? row
      : row.map((cell, c) => {
          if (c < c1 || c > c2) return cell;
          // Drop text + formula but keep formatting — same as Excel's Delete.
          const { formula: _f, ...rest } = cell;
          void _f;
          return { ...rest, text: '' };
        }),
  );
  return { ...sheet, cells };
}

/**
 * Build a payload from raw TSV text — used when the system clipboard
 * holds text that didn't come from us (e.g., copied from Excel itself).
 * Single-line single-cell text comes back as a 1×1 payload; the caller
 * decides whether to delegate to native paste instead.
 */
export function tsvToPayload(text: string, anchorR: number, anchorC: number): RichClipboardPayload {
  const rows = parseTsv(text);
  const cells: XlsxCell[][] = rows.map((row) =>
    row.map((s) => (s.startsWith('=') ? { text: s, formula: s } : { text: s })),
  );
  return { origin: { r: anchorR, c: anchorC }, cells };
}

/**
 * Heuristic — does this clipboard text look like a multi-cell range?
 * Used to decide whether to intercept paste vs. let native paste fill an
 * input. Single-line single-cell text falls through to native handling.
 */
export function isMultiCellTsv(text: string): boolean {
  if (text.includes('\t')) return true;
  // Strip a single trailing newline (Excel ends with one) before counting.
  const trimmed = text.replace(/\r?\n$/, '');
  return trimmed.includes('\n');
}
