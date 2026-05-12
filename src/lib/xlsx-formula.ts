/**
 * Excel formula evaluator. Lives in the renderer; formula sources are read /
 * written by the xlsx adapter as `cell.formula`, and `recomputeAllFormulas`
 * is invoked after every model mutation in the Xlsx editor.
 *
 * Scope (matches MVP-SPEC §A.1's 50 functions):
 *   - Cell refs: A1, $A$1, $A1, A$1, Sheet1!A1
 *   - Ranges: A1:B5, Sheet1!A1:B5
 *   - Operators: + - * / ^ & = <> < > <= >= and unary minus + percent suffix
 *   - Literals: numbers, "strings" with "" escape, TRUE / FALSE
 *   - Errors: #DIV/0!, #VALUE!, #REF!, #NAME?, #N/A, #NUM!, #CYCLE!
 *   - Function calls dispatched to formulajs (50 listed in A.1 plus more)
 *
 * Out of scope: external workbook refs, structured table refs, array
 * formulas with `{...}` syntax, defined names. Dynamic arrays spill from
 * functions like UNIQUE / SORT collapse to first cell (Phase H+).
 *
 * Recompute strategy: iterate-to-fixpoint with a max-pass cap (8). Tracks
 * a per-cell `evaluating` flag during a single pass to detect cycles and
 * mark them `#CYCLE!`. Avoids the complexity of a full dependency DAG —
 * MVP sheets are small enough that 8 sweeps over O(N) formulas is cheap.
 */
/* eslint-disable @typescript-eslint/no-explicit-any */

import * as fj from '@formulajs/formulajs';

// ── Types ────────────────────────────────────────────────────────────────

export type FormulaErrorCode =
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#N/A'
  | '#NUM!'
  | '#CYCLE!'
  | '#ERROR!';

export const ERROR_CODES = new Set<FormulaErrorCode>([
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#N/A',
  '#NUM!',
  '#CYCLE!',
  '#ERROR!',
]);

export class FormulaError extends Error {
  readonly code: FormulaErrorCode;
  constructor(code: FormulaErrorCode, msg?: string) {
    super(msg ?? code);
    this.code = code;
  }
}

export type CellValue = number | string | boolean | FormulaError | null;

/**
 * Sheet contract the evaluator works against. We deliberately keep this
 * minimal — XlsxSheet from the adapter satisfies it structurally without
 * an explicit conversion step.
 */
export interface FormulaCell {
  text: string;
  formula?: string;
}

export interface FormulaSheet {
  name: string;
  cells: FormulaCell[][];
  rowCount: number;
  colCount: number;
}

// ── Tokenizer ────────────────────────────────────────────────────────────

type Tok =
  | { kind: 'num'; v: number }
  | { kind: 'str'; v: string }
  | { kind: 'bool'; v: boolean }
  | { kind: 'err'; v: FormulaErrorCode }
  | { kind: 'ref'; v: string } // raw text "A1" or "Sheet1!A1"
  | { kind: 'name'; v: string } // function name or unrecognized identifier
  | { kind: 'op'; v: string }
  | { kind: 'lparen' }
  | { kind: 'rparen' }
  | { kind: 'comma' }
  | { kind: 'colon' };

const REF_RE = /^(?:'([^']+)'|([A-Za-z_][\w]*))!\$?([A-Z]+)\$?(\d+)|^\$?([A-Z]+)\$?(\d+)/;

function tokenize(src: string): Tok[] {
  const tokens: Tok[] = [];
  let i = 0;
  const len = src.length;

  while (i < len) {
    const c = src[i];

    // Whitespace
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
      i += 1;
      continue;
    }

    // Strings
    if (c === '"') {
      let j = i + 1;
      let out = '';
      while (j < len) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            out += '"';
            j += 2;
            continue;
          }
          break;
        }
        out += src[j];
        j += 1;
      }
      if (j >= len) throw new FormulaError('#VALUE!', 'Unterminated string');
      tokens.push({ kind: 'str', v: out });
      i = j + 1;
      continue;
    }

    // Numbers (must come before bare letters)
    if ((c >= '0' && c <= '9') || (c === '.' && src[i + 1] >= '0' && src[i + 1] <= '9')) {
      let j = i;
      while (j < len && ((src[j] >= '0' && src[j] <= '9') || src[j] === '.')) j += 1;
      // Scientific notation (1e5, 1.2E-3)
      if (j < len && (src[j] === 'e' || src[j] === 'E')) {
        j += 1;
        if (j < len && (src[j] === '+' || src[j] === '-')) j += 1;
        while (j < len && src[j] >= '0' && src[j] <= '9') j += 1;
      }
      const num = Number(src.slice(i, j));
      if (!Number.isFinite(num)) throw new FormulaError('#NUM!');
      tokens.push({ kind: 'num', v: num });
      i = j;
      continue;
    }

    // Errors (#NAME?, #DIV/0!, etc.)
    if (c === '#') {
      let j = i + 1;
      while (j < len && (/[A-Z0-9/]/.test(src[j]) || src[j] === '!' || src[j] === '?')) j += 1;
      const code = src.slice(i, j) as FormulaErrorCode;
      if (ERROR_CODES.has(code)) {
        tokens.push({ kind: 'err', v: code });
        i = j;
        continue;
      }
      throw new FormulaError('#NAME?', `Unknown error literal ${code}`);
    }

    // Identifiers, refs, function names, booleans
    if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '$' || c === "'") {
      // Try cell ref first (handles $A$1, Sheet1!A1, 'My Sheet'!A1)
      const refMatch = REF_RE.exec(src.slice(i));
      if (refMatch) {
        // Bare-ref shape (col-letters + digits with no sheet qualifier) is
        // also satisfied by function names ending in digits — LOG10, ATAN2,
        // LOG2 — and by mixed-token identifiers like BIN2DEC. Without this
        // check, `=LOG10(100)` tokenises as ref:LOG10 + lparen and parses
        // as `Trailing tokens`, returning #VALUE! for a perfectly valid
        // formula. Heuristic: if what follows the matched ref is `(`
        // (function call) or another identifier-continuation char (more
        // letters/digits/_/.), the user wrote a name, not a cell. Sheet-
        // qualified refs (`Sheet1!A1`) keep ref priority — `Sheet1!A1(...)`
        // isn't valid syntax anyway.
        const sheetQualified = !!(refMatch[1] || refMatch[2]);
        const after = src[i + refMatch[0].length];
        const looksLikeIdentifierContinuation =
          after === '(' || (after !== undefined && /[A-Za-z0-9_.]/.test(after));
        if (sheetQualified || !looksLikeIdentifierContinuation) {
          tokens.push({ kind: 'ref', v: refMatch[0] });
          i += refMatch[0].length;
          continue;
        }
      }
      // Fall back to identifier (function name)
      let j = i;
      while (j < len && /[A-Za-z0-9_.]/.test(src[j])) j += 1;
      const ident = src.slice(i, j);
      const upper = ident.toUpperCase();
      if (upper === 'TRUE') tokens.push({ kind: 'bool', v: true });
      else if (upper === 'FALSE') tokens.push({ kind: 'bool', v: false });
      else tokens.push({ kind: 'name', v: upper });
      i = j;
      continue;
    }

    // Two-char operators first
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      tokens.push({ kind: 'op', v: two });
      i += 2;
      continue;
    }

    // Single-char operators / punctuation
    if ('+-*/^&=<>%'.includes(c)) {
      tokens.push({ kind: 'op', v: c });
      i += 1;
      continue;
    }
    if (c === '(') {
      tokens.push({ kind: 'lparen' });
      i += 1;
      continue;
    }
    if (c === ')') {
      tokens.push({ kind: 'rparen' });
      i += 1;
      continue;
    }
    if (c === ',') {
      tokens.push({ kind: 'comma' });
      i += 1;
      continue;
    }
    if (c === ':') {
      tokens.push({ kind: 'colon' });
      i += 1;
      continue;
    }

    throw new FormulaError('#VALUE!', `Unexpected character '${c}'`);
  }

  return tokens;
}

// ── AST + Parser ─────────────────────────────────────────────────────────

type Node =
  | { kind: 'num'; v: number }
  | { kind: 'str'; v: string }
  | { kind: 'bool'; v: boolean }
  | { kind: 'err'; v: FormulaErrorCode }
  | { kind: 'ref'; sheet?: string; r: number; c: number }
  | { kind: 'range'; sheet?: string; r1: number; c1: number; r2: number; c2: number }
  | { kind: 'unary'; op: '-' | '+' | '%'; arg: Node }
  | { kind: 'bin'; op: string; l: Node; r: Node }
  | { kind: 'call'; name: string; args: Node[] };

class Parser {
  private toks: Tok[];
  private pos = 0;
  constructor(toks: Tok[]) {
    this.toks = toks;
  }
  private peek(): Tok | undefined {
    return this.toks[this.pos];
  }
  private eat(): Tok | undefined {
    return this.toks[this.pos++];
  }
  private expect(kind: Tok['kind']): Tok {
    const t = this.eat();
    if (!t || t.kind !== kind) throw new FormulaError('#VALUE!', `Expected ${kind}`);
    return t;
  }

  parse(): Node {
    const node = this.parseComparison();
    if (this.pos < this.toks.length) {
      throw new FormulaError('#VALUE!', 'Trailing tokens');
    }
    return node;
  }

  // Precedence (low → high): comparison, &, +/-, *//, ^, unary, primary
  private parseComparison(): Node {
    let left = this.parseConcat();
    while (true) {
      const t = this.peek();
      if (!t || t.kind !== 'op') break;
      if (!['=', '<>', '<', '>', '<=', '>='].includes(t.v as string)) break;
      this.eat();
      const right = this.parseConcat();
      left = { kind: 'bin', op: (t as { v: string }).v, l: left, r: right };
    }
    return left;
  }
  private parseConcat(): Node {
    let left = this.parseAdditive();
    while (true) {
      const t = this.peek();
      if (!t || t.kind !== 'op' || (t as { v: string }).v !== '&') break;
      this.eat();
      const right = this.parseAdditive();
      left = { kind: 'bin', op: '&', l: left, r: right };
    }
    return left;
  }
  private parseAdditive(): Node {
    let left = this.parseMultiplicative();
    while (true) {
      const t = this.peek();
      if (!t || t.kind !== 'op' || ((t as { v: string }).v !== '+' && (t as { v: string }).v !== '-')) break;
      this.eat();
      const right = this.parseMultiplicative();
      left = { kind: 'bin', op: (t as { v: string }).v, l: left, r: right };
    }
    return left;
  }
  private parseMultiplicative(): Node {
    let left = this.parseExponent();
    while (true) {
      const t = this.peek();
      if (!t || t.kind !== 'op' || ((t as { v: string }).v !== '*' && (t as { v: string }).v !== '/')) break;
      this.eat();
      const right = this.parseExponent();
      left = { kind: 'bin', op: (t as { v: string }).v, l: left, r: right };
    }
    return left;
  }
  private parseExponent(): Node {
    const left = this.parseUnary();
    const t = this.peek();
    if (t && t.kind === 'op' && (t as { v: string }).v === '^') {
      this.eat();
      const right = this.parseExponent(); // right-assoc
      return { kind: 'bin', op: '^', l: left, r: right };
    }
    return left;
  }
  private parseUnary(): Node {
    const t = this.peek();
    if (t && t.kind === 'op' && ((t as { v: string }).v === '-' || (t as { v: string }).v === '+')) {
      this.eat();
      const arg = this.parseUnary();
      return { kind: 'unary', op: (t as { v: string }).v as '-' | '+', arg };
    }
    return this.parsePostfix();
  }
  private parsePostfix(): Node {
    let node = this.parsePrimary();
    // Percent suffix: 50% → 0.5
    while (true) {
      const t = this.peek();
      if (!t || t.kind !== 'op' || (t as { v: string }).v !== '%') break;
      this.eat();
      node = { kind: 'unary', op: '%', arg: node };
    }
    return node;
  }
  private parsePrimary(): Node {
    const t = this.eat();
    if (!t) throw new FormulaError('#VALUE!', 'Unexpected end of input');
    switch (t.kind) {
      case 'num':
        return { kind: 'num', v: (t as { v: number }).v };
      case 'str':
        return { kind: 'str', v: (t as { v: string }).v };
      case 'bool':
        return { kind: 'bool', v: (t as { v: boolean }).v };
      case 'err':
        return { kind: 'err', v: (t as { v: FormulaErrorCode }).v };
      case 'lparen': {
        const inner = this.parseComparison();
        this.expect('rparen');
        return inner;
      }
      case 'ref': {
        const refNode = parseRefToken((t as { v: string }).v);
        // Range? Look for `:` followed by another ref.
        const next = this.peek();
        if (next?.kind === 'colon') {
          this.eat();
          const tt = this.eat();
          if (!tt || tt.kind !== 'ref') throw new FormulaError('#REF!', 'Expected ref after :');
          const r2Node = parseRefToken((tt as { v: string }).v);
          return {
            kind: 'range',
            sheet: refNode.sheet ?? r2Node.sheet,
            r1: Math.min(refNode.r, r2Node.r),
            c1: Math.min(refNode.c, r2Node.c),
            r2: Math.max(refNode.r, r2Node.r),
            c2: Math.max(refNode.c, r2Node.c),
          };
        }
        return refNode;
      }
      case 'name': {
        const name = (t as { v: string }).v;
        // Function call?
        const next = this.peek();
        if (next?.kind === 'lparen') {
          this.eat();
          const args: Node[] = [];
          if (this.peek()?.kind !== 'rparen') {
            args.push(this.parseComparison());
            while (this.peek()?.kind === 'comma') {
              this.eat();
              args.push(this.parseComparison());
            }
          }
          this.expect('rparen');
          return { kind: 'call', name, args };
        }
        // Bare identifier with no parens — treat as #NAME? at eval time.
        throw new FormulaError('#NAME?', `Unknown name ${name}`);
      }
      default:
        throw new FormulaError('#VALUE!', `Unexpected token ${t.kind}`);
    }
  }
}

/** "$A$1" / "Sheet1!B12" → { sheet?, r, c }. r/c are 0-based. */
function parseRefToken(raw: string): { kind: 'ref'; sheet?: string; r: number; c: number } {
  const m = /^(?:'([^']+)'|([A-Za-z_][\w]*))!(.+)|^(.+)/.exec(raw);
  if (!m) throw new FormulaError('#REF!');
  const sheet = m[1] ?? m[2];
  const cellPart = m[3] ?? m[4];
  const cm = /^\$?([A-Z]+)\$?(\d+)$/.exec(cellPart);
  if (!cm) throw new FormulaError('#REF!');
  let c = 0;
  for (const ch of cm[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { kind: 'ref', sheet, r: parseInt(cm[2], 10) - 1, c: c - 1 };
}

// ── Evaluator ────────────────────────────────────────────────────────────

interface EvalContext {
  sheets: FormulaSheet[];
  /** Active sheet name — used to resolve unqualified refs. */
  activeSheet: string;
  /** Map of formula keys ("sheetName!r:c") currently being evaluated — cycle guard. */
  inProgress: Set<string>;
  /** Cache of computed values keyed by "sheetName!r:c" so a single pass doesn't
   *  re-evaluate the same formula cell multiple times when many cells reference it. */
  cache: Map<string, CellValue>;
}

/** Cell-value coerce helpers (number, string, boolean). */
function toNumber(v: CellValue): number {
  if (v instanceof FormulaError) throw v;
  if (v === null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(String(v).replace(/,/g, ''));
  if (!Number.isFinite(n)) throw new FormulaError('#VALUE!');
  return n;
}
function toStr(v: CellValue): string {
  if (v instanceof FormulaError) throw v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function findSheet(ctx: EvalContext, name: string | undefined): FormulaSheet | null {
  const target = name ?? ctx.activeSheet;
  return ctx.sheets.find((s) => s.name === target) ?? null;
}

function cellKey(sheet: string, r: number, c: number): string {
  return `${sheet}!${r}:${c}`;
}

/** Read the value of a single cell, recursively evaluating its formula if needed. */
function readCell(ctx: EvalContext, sheetName: string, r: number, c: number): CellValue {
  const sheet = findSheet(ctx, sheetName);
  if (!sheet) return new FormulaError('#REF!');
  if (r < 0 || c < 0 || r >= sheet.rowCount || c >= sheet.colCount) {
    return null;
  }
  const cell = sheet.cells[r]?.[c];
  if (!cell) return null;
  if (cell.formula) {
    const key = cellKey(sheet.name, r, c);
    if (ctx.cache.has(key)) return ctx.cache.get(key) ?? null;
    if (ctx.inProgress.has(key)) {
      // Cycle — return error and let the parent see it.
      return new FormulaError('#CYCLE!');
    }
    ctx.inProgress.add(key);
    try {
      const v = evaluateFormula(ctx, cell.formula, sheet.name);
      ctx.cache.set(key, v);
      return v;
    } finally {
      ctx.inProgress.delete(key);
    }
  }
  return parseLiteral(cell.text);
}

/** Best-effort literal parse — same heuristic Excel uses for typed values. */
function parseLiteral(text: string): CellValue {
  if (text === '') return null;
  if (text === 'TRUE') return true;
  if (text === 'FALSE') return false;
  if (ERROR_CODES.has(text as FormulaErrorCode)) {
    return new FormulaError(text as FormulaErrorCode);
  }
  // Number with optional thousands separator and trailing %.
  const cleaned = text.replace(/,/g, '');
  const pct = cleaned.endsWith('%');
  const num = Number(pct ? cleaned.slice(0, -1) : cleaned);
  if (Number.isFinite(num) && cleaned !== '') {
    return pct ? num / 100 : num;
  }
  return text;
}

function readRange(
  ctx: EvalContext,
  sheetName: string,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
): CellValue[][] {
  const out: CellValue[][] = [];
  for (let r = r1; r <= r2; r += 1) {
    const row: CellValue[] = [];
    for (let c = c1; c <= c2; c += 1) {
      row.push(readCell(ctx, sheetName, r, c));
    }
    out.push(row);
  }
  return out;
}

function evalNode(ctx: EvalContext, node: Node, scope: string): CellValue {
  switch (node.kind) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return new FormulaError(node.v);
    case 'ref':
      return readCell(ctx, node.sheet ?? scope, node.r, node.c);
    case 'range': {
      // Bare range outside of a function is rare — return an array. Functions
      // handle this; arithmetic on it is #VALUE!.
      const arr = readRange(ctx, node.sheet ?? scope, node.r1, node.c1, node.r2, node.c2);
      // Single cell -> scalar
      if (arr.length === 1 && arr[0].length === 1) return arr[0][0];
      // Otherwise return a sentinel-ish 2D array; called-function path handles 2D.
      return arr as unknown as CellValue;
    }
    case 'unary': {
      if (node.op === '%') {
        const v = evalNode(ctx, node.arg, scope);
        return toNumber(v) / 100;
      }
      const v = evalNode(ctx, node.arg, scope);
      const n = toNumber(v);
      return node.op === '-' ? -n : +n;
    }
    case 'bin': {
      const lv = evalNode(ctx, node.l, scope);
      const rv = evalNode(ctx, node.r, scope);
      if (lv instanceof FormulaError) return lv;
      if (rv instanceof FormulaError) return rv;
      switch (node.op) {
        case '+':
          return toNumber(lv) + toNumber(rv);
        case '-':
          return toNumber(lv) - toNumber(rv);
        case '*':
          return toNumber(lv) * toNumber(rv);
        case '/': {
          const denom = toNumber(rv);
          if (denom === 0) return new FormulaError('#DIV/0!');
          return toNumber(lv) / denom;
        }
        case '^':
          return Math.pow(toNumber(lv), toNumber(rv));
        case '&':
          return toStr(lv) + toStr(rv);
        case '=':
          return looseEq(lv, rv);
        case '<>':
          return !looseEq(lv, rv);
        case '<':
          return cmp(lv, rv) < 0;
        case '>':
          return cmp(lv, rv) > 0;
        case '<=':
          return cmp(lv, rv) <= 0;
        case '>=':
          return cmp(lv, rv) >= 0;
      }
      return new FormulaError('#VALUE!');
    }
    case 'call':
      return callFunction(ctx, node, scope);
  }
}

/** Excel-ish equality: numbers compare numerically, booleans coerced to numbers,
 *  strings case-insensitive equal, types differ → false. */
function looseEq(a: CellValue, b: CellValue): boolean {
  if (a === null) a = 0;
  if (b === null) b = 0;
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
  if (typeof a === 'number' && typeof b === 'boolean') return a === (b ? 1 : 0);
  if (typeof b === 'number' && typeof a === 'boolean') return b === (a ? 1 : 0);
  return false;
}
function cmp(a: CellValue, b: CellValue): number {
  if (a instanceof FormulaError || b instanceof FormulaError) {
    throw a instanceof FormulaError ? a : (b as FormulaError);
  }
  if (a === null) a = 0;
  if (b === null) b = 0;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return a < b ? -1 : a > b ? 1 : 0;
  // Mixed: coerce to number; if either is non-numeric, coerce to string.
  const an = Number(a);
  const bn = Number(b);
  if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
  return String(a) < String(b) ? -1 : String(a) > String(b) ? 1 : 0;
}

// ── Function dispatch ────────────────────────────────────────────────────

/**
 * Convert a 2D CellValue range to a flat array of plain JS values that
 * formulajs accepts. formulajs's signature for SUM / AVERAGE etc. is
 * variadic: `SUM(1, [2, 3], 4)`. We always pass a flattened array so any
 * mix of refs / ranges / scalars normalises.
 */
function rangeToFlat(node: Node, ctx: EvalContext, scope: string): any[] {
  if (node.kind === 'range') {
    const arr = readRange(ctx, node.sheet ?? scope, node.r1, node.c1, node.r2, node.c2);
    const out: any[] = [];
    for (const row of arr) {
      for (const v of row) {
        if (v === null) continue;
        if (v instanceof FormulaError) throw v;
        out.push(v);
      }
    }
    return out;
  }
  const v = evalNode(ctx, node, scope);
  if (v instanceof FormulaError) throw v;
  if (v === null) return [];
  return [v];
}

/** For lookup-style functions that need the 2D shape preserved. */
function rangeTo2D(node: Node, ctx: EvalContext, scope: string): any[][] {
  if (node.kind === 'range') {
    const arr = readRange(ctx, node.sheet ?? scope, node.r1, node.c1, node.r2, node.c2);
    return arr.map((row) =>
      row.map((v) => {
        if (v instanceof FormulaError) throw v;
        return v === null ? '' : v;
      }),
    );
  }
  const v = evalNode(ctx, node, scope);
  if (v instanceof FormulaError) throw v;
  return [[v === null ? '' : v]];
}

function callFunction(ctx: EvalContext, node: { kind: 'call'; name: string; args: Node[] }, scope: string): CellValue {
  const fn = (fj as Record<string, any>)[node.name];
  if (typeof fn !== 'function') {
    return new FormulaError('#NAME?', `Unknown function ${node.name}`);
  }
  // A few functions need raw range arrays preserved — rather than scalar-evalling
  // each arg, we hand them flattened arrays / 2D arrays as appropriate.
  try {
    const args = mapArgsForFunction(node.name, node.args, ctx, scope);
    const result = fn.apply(null, args);
    return normalizeResult(result);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    return new FormulaError('#VALUE!', (e as Error).message);
  }
}

const RANGE_FLATTEN_FNS = new Set([
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA',
  'SUMIF', 'COUNTIF', 'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS',
  'PRODUCT', 'STDEV', 'STDEVP', 'VAR', 'VARP', 'MEDIAN', 'MODE',
]);
const RANGE_2D_FNS = new Set([
  'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP',
]);

function mapArgsForFunction(name: string, args: Node[], ctx: EvalContext, scope: string): any[] {
  // Default: scalar-evaluate each arg.
  const out: any[] = [];
  // SUMIF / COUNTIF / etc. take (range, criterion, [sumRange]) — first / third
  // args are ranges, criterion is scalar. For simplicity we flatten ranges
  // for ALL args except known criterion positions.
  if (name === 'SUMIF' || name === 'AVERAGEIF') {
    out.push(rangeToFlat(args[0], ctx, scope));
    out.push(scalar(evalNode(ctx, args[1], scope)));
    if (args[2]) out.push(rangeToFlat(args[2], ctx, scope));
    return out;
  }
  if (name === 'COUNTIF') {
    out.push(rangeToFlat(args[0], ctx, scope));
    out.push(scalar(evalNode(ctx, args[1], scope)));
    return out;
  }
  if (name === 'SUMIFS' || name === 'AVERAGEIFS' || name === 'COUNTIFS') {
    // SUMIFS(sumRange, range1, criteria1, range2, criteria2, ...)
    // COUNTIFS(range1, criteria1, range2, criteria2, ...)
    const start = name === 'COUNTIFS' ? 0 : 1;
    if (start === 1) out.push(rangeToFlat(args[0], ctx, scope));
    for (let i = start; i < args.length; i += 2) {
      out.push(rangeToFlat(args[i], ctx, scope));
      out.push(scalar(evalNode(ctx, args[i + 1], scope)));
    }
    return out;
  }
  if (RANGE_2D_FNS.has(name)) {
    // VLOOKUP(lookup, table, col, [exact]) → first arg scalar, second 2D, third scalar
    if (args.length >= 1) out.push(scalar(evalNode(ctx, args[0], scope)));
    if (args.length >= 2) out.push(rangeTo2D(args[1], ctx, scope));
    for (let i = 2; i < args.length; i += 1) out.push(scalar(evalNode(ctx, args[i], scope)));
    return out;
  }
  if (RANGE_FLATTEN_FNS.has(name)) {
    for (const a of args) out.push(rangeToFlat(a, ctx, scope));
    return out;
  }
  // Default — every arg is a scalar.
  for (const a of args) {
    if (a.kind === 'range') {
      // Auto-flatten ranges to scalar of first cell, like Excel's implicit
      // intersection. Better than #VALUE! for casual use.
      const arr = readRange(ctx, a.sheet ?? scope, a.r1, a.c1, a.r2, a.c2);
      out.push(scalar(arr[0]?.[0] ?? null));
    } else {
      out.push(scalar(evalNode(ctx, a, scope)));
    }
  }
  return out;
}

function scalar(v: CellValue): any {
  if (v instanceof FormulaError) throw v;
  return v;
}

function normalizeResult(v: any): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) {
      if (Number.isNaN(v)) return new FormulaError('#NUM!');
      return new FormulaError('#NUM!');
    }
    return v;
  }
  if (typeof v === 'string') {
    if (ERROR_CODES.has(v as FormulaErrorCode)) return new FormulaError(v as FormulaErrorCode);
    return v;
  }
  if (typeof v === 'boolean') return v;
  if (v instanceof Error) {
    // formulajs surfaces some errors as Error subclasses.
    const msg = v.message;
    if (ERROR_CODES.has(msg as FormulaErrorCode)) return new FormulaError(msg as FormulaErrorCode);
    return new FormulaError('#VALUE!', msg);
  }
  if (Array.isArray(v)) {
    // Spill array → take first cell.
    const first = (v[0] as any) ?? null;
    return normalizeResult(Array.isArray(first) ? first[0] : first);
  }
  return String(v);
}

// ── Public API ───────────────────────────────────────────────────────────

/**
 * Evaluate one formula source ("=A1+B1") in the context of the given
 * workbook + active sheet. Returns the resulting CellValue. Throws never —
 * any error short-circuits to a FormulaError CellValue.
 */
export function evaluateFormula(ctx: EvalContext, source: string, sheetName: string): CellValue {
  const trimmed = source.trim();
  if (!trimmed.startsWith('=')) return parseLiteral(trimmed);
  const body = trimmed.slice(1);
  if (body === '') return new FormulaError('#VALUE!');
  try {
    const toks = tokenize(body);
    const ast = new Parser(toks).parse();
    const v = evalNode(ctx, ast, sheetName);
    if (v instanceof FormulaError) return v;
    return v;
  } catch (e) {
    if (e instanceof FormulaError) return e;
    return new FormulaError('#ERROR!', (e as Error).message);
  }
}

/**
 * Render a CellValue as the user-visible string. Mirrors Excel's auto-
 * formatting: numbers as plain decimals (no trailing zeros), booleans as
 * TRUE / FALSE, errors as their code, null as empty.
 */
export function formatValue(v: CellValue): string {
  if (v === null) return '';
  if (v instanceof FormulaError) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return '#NUM!';
    // Trim trailing zeros while keeping precision reasonable.
    return Number(v.toPrecision(15)).toString();
  }
  return v;
}

/**
 * Re-evaluate every cell with a `formula` field across all sheets, mutating
 * a deep copy of the sheets array and writing computed display strings into
 * `cell.text`. Returns the new sheets array (referentially fresh; safe to
 * pass straight back to setModel).
 *
 * Iterate-to-fixpoint with a max-pass cap. The cache is rebuilt per pass so
 * dependency depth N converges in N passes; max 8 covers the sheets we
 * realistically expect, while still bounding runtime.
 */
export function recomputeAllFormulas<T extends FormulaSheet>(sheets: T[]): T[] {
  // Shallow-copy sheet structure so we can mutate cells without aliasing.
  // Generic over the input element so extra adapter fields (merges, raw cell
  // types, styles) survive the round-trip — we only touch `text` here.
  const next = sheets.map((s) => ({
    ...s,
    cells: s.cells.map((row) => row.map((c) => ({ ...c }))),
  })) as T[];

  const ctx: EvalContext = {
    sheets: next,
    activeSheet: next[0]?.name ?? '',
    inProgress: new Set(),
    cache: new Map(),
  };

  // Precompute the formulas to evaluate so we don't waste passes on plain cells.
  const formulaCells: Array<{ sheet: FormulaSheet; r: number; c: number }> = [];
  for (const sheet of next) {
    for (let r = 0; r < sheet.cells.length; r += 1) {
      const row = sheet.cells[r];
      for (let c = 0; c < row.length; c += 1) {
        if (row[c]?.formula) formulaCells.push({ sheet, r, c });
      }
    }
  }
  if (formulaCells.length === 0) return next;

  // R375 — raise pass cap from 8 to 32. The fixed-point iteration propagates
  // one dependency-chain level per pass: A=1, B=A+1, C=B+1, ..., I=H+1 (9
  // cells) needs 9 passes for I to reach its final 9 — pass 8 would leave I
  // showing an off-by-one stale value because pass 8 sees H still stale.
  // Excel docs / spec /「N 條深 chain」 reports of cells displaying「one off」
  // values trace back to this exact off-by-N pattern. 8 is enough for
  // shallow formulas (most workbooks: 2-4 deep) but **silently truncates
  // computation for 9+ deep chains** — the most common offender is summary
  // sheets where each row references the row above's accumulator. The
  // change-detection break (`if (!changed) break`) keeps shallow cases
  // cheap: short chains converge in 2-3 passes and exit early, paying no
  // perf cost for the higher cap. Circular references stay bounded by
  // `ctx.inProgress` (set at line 871) — that's the real safety net; the
  // pass cap is a secondary guard for non-circular but deep chains.
  // 32 covers virtually any realistic spreadsheet depth (Excel's own
  // calculation engine uses higher iterative caps with the same fixed-
  // point shape).
  const MAX_PASSES = 32;
  for (let pass = 0; pass < MAX_PASSES; pass += 1) {
    ctx.cache = new Map();
    let changed = false;
    for (const { sheet, r, c } of formulaCells) {
      ctx.activeSheet = sheet.name;
      ctx.inProgress = new Set();
      const formula = sheet.cells[r][c].formula!;
      const v = evaluateFormula(ctx, formula, sheet.name);
      const display = formatValue(v);
      if (sheet.cells[r][c].text !== display) {
        sheet.cells[r][c] = { ...sheet.cells[r][c], text: display };
        changed = true;
      }
    }
    if (!changed) break;
  }
  return next;
}

/**
 * Utility — true if `text` looks like a formula source (starts with `=` after
 * trim). Used by the editor to decide whether to set the `formula` field.
 */
export function isFormulaSource(text: string): boolean {
  return text.trimStart().startsWith('=');
}

// ── Reference shifting (used by copy / paste) ────────────────────────────

/** Match a single cell ref at the *start* of input. Captures sheet (quoted
 *  or bare), absolute markers ($), col letters, row digits. Mirrors the
 *  shape used by the tokenizer's REF_RE but exposes the absolute markers so
 *  we can decide which axis to shift. */
const REF_AT_START =
  /^(?:'([^']+)'!|([A-Za-z_][\w]*)!)?(\$)?([A-Z]+)(\$)?(\d+)/;

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function colIndexToLetters(idx: number): string {
  let s = '';
  let n = idx;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/**
 * Rewrite every cell reference in `src` by `(dr, dc)`. Absolute parts
 * (the `$` prefix) stay locked. Sheet-qualified refs keep their sheet —
 * we only shift the cell coordinate.
 *
 * Out-of-bounds (negative row/col) collapse to `#REF!` so Excel-style
 * "you cut this off the edge" feedback shows up at evaluation time.
 *
 * Returns the original string unchanged if it isn't a formula or if the
 * shift is zero.
 */
export function shiftFormula(src: string, dr: number, dc: number): string {
  if (!src.startsWith('=')) return src;
  if (dr === 0 && dc === 0) return src;

  const len = src.length;
  let out = '=';
  let i = 1;

  while (i < len) {
    const ch = src[i];

    // String literal — copy through verbatim, honoring "" escape.
    if (ch === '"') {
      let j = i + 1;
      while (j < len) {
        if (src[j] === '"' && src[j + 1] === '"') {
          j += 2;
          continue;
        }
        if (src[j] === '"') {
          j += 1;
          break;
        }
        j += 1;
      }
      out += src.slice(i, j);
      i = j;
      continue;
    }

    // Identifier-or-ref start: try to match a ref greedily; otherwise
    // consume an identifier token (function name / TRUE / FALSE) without
    // mangling it. This also avoids matching `SUM(` as `SUM` col + `(`.
    if (/[A-Za-z_$']/.test(ch)) {
      const m = REF_AT_START.exec(src.slice(i));
      // Same gate as the tokenizer: don't shift function names that happen
      // to look like refs (LOG10, ATAN2) or identifier prefixes (BIN2DEC).
      // Without this, copy-pasting a cell with `=LOG10(A1)` would mangle
      // the function name on shift.
      const sheetQualified = !!(m && (m[1] || m[2]));
      const after = m ? src[i + m[0].length] : undefined;
      const looksLikeIdentifierContinuation = m
        ? after === '(' || (after !== undefined && /[A-Za-z0-9_.]/.test(after))
        : false;
      if (m && (sheetQualified || !looksLikeIdentifierContinuation)) {
        const sheetQuoted = m[1];
        const sheetBare = m[2];
        const colAbs = !!m[3];
        const colLetters = m[4];
        const rowAbs = !!m[5];
        const rowDigits = m[6];

        let newCol = colLettersToIndex(colLetters);
        let newRow = parseInt(rowDigits, 10) - 1;
        if (!colAbs) newCol += dc;
        if (!rowAbs) newRow += dr;

        if (newCol < 0 || newRow < 0) {
          out += '#REF!';
        } else {
          let s = '';
          if (sheetQuoted) s += `'${sheetQuoted}'!`;
          else if (sheetBare) s += `${sheetBare}!`;
          if (colAbs) s += '$';
          s += colIndexToLetters(newCol);
          if (rowAbs) s += '$';
          s += String(newRow + 1);
          out += s;
        }
        i += m[0].length;
        continue;
      }

      // Plain identifier — consume as a unit.
      let j = i;
      while (j < len && /[A-Za-z0-9_.]/.test(src[j])) j += 1;
      out += src.slice(i, j);
      i = j;
      continue;
    }

    // Anything else — punctuation, numbers, operators — copy as-is.
    out += ch;
    i += 1;
  }
  return out;
}
