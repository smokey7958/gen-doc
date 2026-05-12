/**
 * xlsx ↔ in-memory model adapter.
 *
 * Reading uses SheetJS (`xlsx`) which already supports cellStyles.
 * Writing uses `xlsx-js-style` so we can persist font / alignment / fill on
 * cells the user has formatted — vanilla SheetJS strips style on write.
 *
 * Round-trip caveat: SheetJS only surfaces fill colours via `.s` on read,
 * so our editor's toolbar state can only reflect fill exactly. Bold/italic/
 * align/font-color *are* written into styles.xml (Excel renders them) but
 * we won't see them in our toolbar after a save→reload cycle. The user can
 * re-apply via the toolbar if they need to mutate further. This is flagged
 * as an MVP limitation in the editor banner.
 */

import * as XLSX from 'xlsx';
import * as XLSXStyle from 'xlsx-js-style';
import JSZip from 'jszip';

export interface XlsxCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** "left" | "center" | "right" */
  align?: 'left' | 'center' | 'right';
  /** ARGB-ish hex (no #), e.g. "FF0000". Falsy = inherit. */
  fontColor?: string;
  /** Background fill colour, hex (no #). */
  bgColor?: string;
  /** Font size in points. */
  fontSize?: number;
  /** Font family name (maps to xlsx `font.name`). */
  fontFamily?: string;
  /**
   * Excel number format string (e.g. "#,##0.00", "0%", "yyyy-mm-dd"). Empty
   * / undefined = "General". Round-trips via SheetJS's `s.numFmt`.
   */
  numberFormat?: string;
}

export interface XlsxCell {
  /** Display string the editor shows + writes. Empty string = blank cell.
   *  For formula cells this is the *computed* value — `formula` holds the
   *  source ("=A1+B1"). */
  text: string;
  /** Original SheetJS cell type, if known — drives serialization. */
  rawType?: 'n' | 's' | 'b' | 'd' | 'e';
  /** Optional formatting; absent = inherit cell default. */
  style?: XlsxCellStyle;
  /** Excel formula source including the leading "=" (e.g. "=SUM(A1:A5)").
   *  Absent for plain literal cells. The xlsx-formula module writes computed
   *  results back into `text`; we round-trip the source via SheetJS's `cell.f`. */
  formula?: string;
}

/**
 * Merged-cell rectangle. Inclusive on both corners — a 2×2 merge of
 * A1:B2 is `{ r1: 0, c1: 0, r2: 1, c2: 1 }`. SheetJS represents the same
 * rectangle as `{ s: {r,c}, e: {r,c} }`; we translate at the adapter
 * boundary so the editor doesn't have to know about that shape.
 */
export interface MergeRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * Floating image anchored over a sheet. We store everything we need to emit
 * the OOXML drawing parts on serialize: the raw bytes, the extension/MIME,
 * the top-left anchor cell, and the display size in EMU.
 *
 * Round-trip caveat: SheetJS / xlsx-js-style do *not* model drawings, so
 * pre-existing images in an opened xlsx are lost on parse and never come
 * back. We only emit images that the user inserted *in this session* (or
 * round-trip ones that survive in the model). Same lossy profile already
 * disclosed in the editor's banner.
 */
export interface XlsxImage {
  /** Stable id for React keys / panel UI. Not used in OOXML. */
  id: string;
  /** Raw image bytes — stamped into `xl/media/imageN.<ext>` on save. */
  data: Uint8Array;
  /** File extension as it should appear in the media path. */
  ext: 'png' | 'jpg' | 'gif' | 'bmp';
  /** MIME type — used for `<Default ContentType=…>` and the preview data URL. */
  mime: string;
  /** Top-left anchor cell (0-based). The image floats over the grid. */
  anchorRow: number;
  anchorCol: number;
  /**
   * Display size in EMU. 914400 EMU = 1 inch; 96 DPI ⇒ 1 px = 9525 EMU.
   * Stored in EMU because that's what oneCellAnchor's `<xdr:ext>` consumes —
   * pre-multiplying once at insert time keeps the serializer trivial.
   */
  widthEmu: number;
  heightEmu: number;
  /** Optional preview data URL for the side-panel thumbnail. */
  dataUrl?: string;
}

export interface XlsxSheet {
  name: string;
  /** rows × cols, dense — cell at [r][c]. */
  cells: XlsxCell[][];
  rowCount: number;
  colCount: number;
  /** Optional list of merged rectangles. Empty / undefined = no merges. */
  merges?: MergeRange[];
  /** Optional list of floating images anchored over the grid. */
  images?: XlsxImage[];
}

export interface XlsxModel {
  sheets: XlsxSheet[];
  /** Active sheet name as recorded in the workbook. */
  activeSheet: string;
}

// Column R = index 17 (0-based), so MIN_COLS=18 ensures A..R are visible.
const MIN_ROWS = 200;
const MIN_COLS = 18;

/**
 * Parse xlsx bytes → editable model. For empty (zero-byte) input — i.e. a
 * freshly created Excel tab — we synthesize a single empty sheet.
 */
export function parseXlsx(bytes: Uint8Array): XlsxModel {
  if (bytes.byteLength === 0) {
    return {
      sheets: [
        {
          name: 'Sheet1',
          cells: emptyGrid(MIN_ROWS, MIN_COLS),
          rowCount: MIN_ROWS,
          colCount: MIN_COLS,
        },
      ],
      activeSheet: 'Sheet1',
    };
  }
  const wb = XLSX.read(bytes, { type: 'array', cellDates: true, cellStyles: true });
  const sheets: XlsxSheet[] = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    return readSheet(name, ws);
  });
  return {
    sheets,
    activeSheet: wb.SheetNames[0] ?? 'Sheet1',
  };
}

/**
 * Serialize a model back to xlsx bytes.
 *
 * IMPORTANT: To preserve formatting/charts/etc on round-trip we re-parse the
 * original bytes (when given) and only overwrite cell values for sheets we
 * touched. Caller passes `originalBytes` when available; otherwise we build
 * a brand-new workbook (formatting will be plain).
 */
export function serializeXlsx(model: XlsxModel, originalBytes?: Uint8Array): Uint8Array {
  // Use xlsx-js-style for both read (preserves .s on its own writes) and
  // write (it's the one that actually emits font/alignment/fill).
  const wb =
    originalBytes && originalBytes.byteLength > 0
      ? XLSXStyle.read(originalBytes, { type: 'array', cellDates: true, cellStyles: true })
      : XLSXStyle.utils.book_new();

  // Track sheet names we have in the model. Add new sheets if missing.
  const modelNames = new Set(model.sheets.map((s) => s.name));
  for (const sheet of model.sheets) {
    const existing = wb.Sheets[sheet.name];
    if (existing) {
      writeSheetInto(existing, sheet);
    } else {
      const ws = XLSXStyle.utils.aoa_to_sheet(toAoA(sheet));
      // Apply styles after aoa_to_sheet since helper drops them.
      applyStylesToSheet(ws, sheet);
      XLSXStyle.utils.book_append_sheet(wb, ws, sheet.name);
    }
  }
  // Drop any sheets removed in the model.
  for (const name of [...wb.SheetNames]) {
    if (!modelNames.has(name)) {
      delete wb.Sheets[name];
      wb.SheetNames = wb.SheetNames.filter((n) => n !== name);
    }
  }

  const out = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer | Uint8Array;
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}

/**
 * Inject the model's floating images into freshly-serialized xlsx bytes.
 *
 * SheetJS / xlsx-js-style don't carry drawings on round-trip, so we run
 * this *after* `serializeXlsx` returns. We open the bytes as a zip, add
 * `xl/media/imageN.<ext>` for each image, write `xl/drawings/drawingN.xml`
 * (one per sheet that has images) plus its `_rels`, register the drawing in
 * the sheet's `_rels` and inject `<drawing r:id="…"/>` into the worksheet
 * XML, and add Default + Override entries in `[Content_Types].xml`.
 *
 * Returns the original bytes unchanged when no sheet has images.
 *
 * Caller is expected to invoke this from a writeBack pipeline that's allowed
 * to be async (the operation has to be — JSZip's load/generate are async).
 */
export async function injectXlsxImages(bytes: Uint8Array, model: XlsxModel): Promise<Uint8Array> {
  const sheetsWithImages = model.sheets.filter((s) => (s.images?.length ?? 0) > 0);
  if (sheetsWithImages.length === 0) return bytes;

  const zip = await JSZip.loadAsync(bytes);

  const ctEntry = zip.file('[Content_Types].xml');
  if (!ctEntry) return bytes;
  let contentTypes = await ctEntry.async('string');

  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const wbRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wbXml || !wbRelsXml) return bytes;

  // Map sheet name → sheet xml path via the workbook's <sheet r:id="…"/>
  // entries and the workbook rels' Target attributes.
  const wbRels: Record<string, string> = {};
  for (const m of wbRelsXml.matchAll(/<Relationship\s+([^/>]*)\/>/g)) {
    const attrs = m[1];
    const id = /Id\s*=\s*"([^"]*)"/.exec(attrs)?.[1];
    const target = /Target\s*=\s*"([^"]*)"/.exec(attrs)?.[1];
    if (id && target) wbRels[id] = target;
  }
  const sheetPathByName: Record<string, string> = {};
  for (const m of wbXml.matchAll(/<sheet\s+([^/>]*)\/?>/g)) {
    const attrs = m[1];
    const name = /name\s*=\s*"([^"]*)"/.exec(attrs)?.[1];
    const rid = /r:id\s*=\s*"([^"]*)"/.exec(attrs)?.[1];
    if (name && rid && wbRels[rid]) {
      sheetPathByName[name] = `xl/${wbRels[rid]}`;
    }
  }

  // Pick free drawing/media file numbers so we don't collide with anything
  // the underlying writer left behind (defensive — typically nothing).
  const filePaths = Object.keys(zip.files);
  let nextDrawing =
    filePaths.reduce((mx, p) => {
      const m = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(p);
      return m ? Math.max(mx, Number(m[1])) : mx;
    }, 0) + 1;
  let nextMedia =
    filePaths.reduce((mx, p) => {
      const m = /^xl\/media\/image(\d+)\./.exec(p);
      return m ? Math.max(mx, Number(m[1])) : mx;
    }, 0) + 1;

  const extsNeeded = new Set<string>();

  for (const sheet of sheetsWithImages) {
    const sheetPath = sheetPathByName[sheet.name];
    if (!sheetPath) continue;
    const sheetEntry = zip.file(sheetPath);
    if (!sheetEntry) continue;
    let sheetXml = await sheetEntry.async('string');

    const drawingNum = nextDrawing++;
    const drawingPath = `xl/drawings/drawing${drawingNum}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${drawingNum}.xml.rels`;

    const drawingRels: string[] = [];
    const anchors: string[] = [];
    let imgRid = 1;
    for (const img of sheet.images ?? []) {
      const mediaNum = nextMedia++;
      const mediaName = `image${mediaNum}.${img.ext}`;
      zip.file(`xl/media/${mediaName}`, img.data);
      extsNeeded.add(img.ext);

      const ridStr = `rId${imgRid}`;
      drawingRels.push(
        `<Relationship Id="${ridStr}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${mediaName}"/>`,
      );
      anchors.push(buildOneCellAnchor(img, ridStr, imgRid));
      imgRid += 1;
    }

    const drawingXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      anchors.join('') +
      `</xdr:wsDr>`;
    zip.file(drawingPath, drawingXml);

    const drawingRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      drawingRels.join('') +
      `</Relationships>`;
    zip.file(drawingRelsPath, drawingRelsXml);

    // Register the drawing in the sheet's rels file, creating it if absent.
    const sheetFileName = sheetPath.split('/').pop()!; // "sheet1.xml"
    const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;
    let sheetRelsXml =
      (await zip.file(sheetRelsPath)?.async('string')) ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const ridNums = [...sheetRelsXml.matchAll(/Id\s*=\s*"rId(\d+)"/g)].map((m) => Number(m[1]));
    // R371 — Math.max via reduce, not spread. ridNums grows with the number
    // of relationships in the sheet's .rels file; for a sheet that
    // accumulated thousands of drawing / hyperlink / external-link
    // relationships (uncommon but realistic for long-lived workbooks
    // with embedded report images), the spread form would explode V8's
    // ~65K argument limit. Same R328 / R370 / R371-DocxEditor fix
    // shape — defensive even when current data sizes are well under
    // the limit, since this is a write path that runs every image
    // injection.
    const maxRid = ridNums.reduce<number>((m, n) => Math.max(m, n), 0);
    const drawingRid = `rId${maxRid + 1}`;
    const newRel = `<Relationship Id="${drawingRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/>`;
    sheetRelsXml = sheetRelsXml.replace(/<\/Relationships>/, `${newRel}</Relationships>`);
    zip.file(sheetRelsPath, sheetRelsXml);

    // Inject `<drawing r:id="…"/>` into the worksheet XML. If a stale tag
    // already exists (e.g. left over from a prior write that we partially
    // touched), replace it — we own the drawing for this sheet now.
    if (/<drawing\s+[^/>]*\/>/.test(sheetXml)) {
      sheetXml = sheetXml.replace(/<drawing\s+[^/>]*\/>/, `<drawing r:id="${drawingRid}"/>`);
    } else {
      sheetXml = sheetXml.replace(/<\/worksheet>/, `<drawing r:id="${drawingRid}"/></worksheet>`);
    }
    if (!/xmlns:r=/.test(sheetXml)) {
      sheetXml = sheetXml.replace(
        /<worksheet([^>]*)>/,
        `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"$1>`,
      );
    }
    zip.file(sheetPath, sheetXml);

    const drawingOverride = `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
    if (!contentTypes.includes(drawingOverride)) {
      contentTypes = contentTypes.replace(/<\/Types>/, `${drawingOverride}</Types>`);
    }
  }

  for (const ext of extsNeeded) {
    if (!new RegExp(`<Default[^/>]*Extension="${ext}"[^/>]*/>`).test(contentTypes)) {
      const ct = mimeForImageExt(ext);
      contentTypes = contentTypes.replace(/<\/Types>/, `<Default Extension="${ext}" ContentType="${ct}"/></Types>`);
    }
  }
  zip.file('[Content_Types].xml', contentTypes);

  const out = await zip.generateAsync({ type: 'uint8array' });
  return out;
}

function mimeForImageExt(ext: string): string {
  switch (ext) {
    case 'png':
      return 'image/png';
    case 'jpg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    default:
      return 'application/octet-stream';
  }
}

function buildOneCellAnchor(img: XlsxImage, rid: string, idx: number): string {
  return (
    `<xdr:oneCellAnchor>` +
    `<xdr:from>` +
    `<xdr:col>${img.anchorCol}</xdr:col>` +
    `<xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${img.anchorRow}</xdr:row>` +
    `<xdr:rowOff>0</xdr:rowOff>` +
    `</xdr:from>` +
    `<xdr:ext cx="${img.widthEmu}" cy="${img.heightEmu}"/>` +
    `<xdr:pic>` +
    `<xdr:nvPicPr>` +
    `<xdr:cNvPr id="${idx + 1}" name="Picture ${idx}"/>` +
    `<xdr:cNvPicPr/>` +
    `</xdr:nvPicPr>` +
    `<xdr:blipFill>` +
    `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/>` +
    `<a:stretch><a:fillRect/></a:stretch>` +
    `</xdr:blipFill>` +
    `<xdr:spPr>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</xdr:spPr>` +
    `</xdr:pic>` +
    `<xdr:clientData/>` +
    `</xdr:oneCellAnchor>`
  );
}

// ── helpers ──────────────────────────────────────────────────────────────

function readSheet(name: string, ws: XLSX.WorkSheet | undefined): XlsxSheet {
  // Hidden / very-hidden sheets may show up in `SheetNames` without a matching
  // entry in `Sheets`. Render them as empty rather than crashing — the user
  // can still see the tab and edit if they want.
  if (!ws) {
    return {
      name,
      cells: emptyGrid(MIN_ROWS, MIN_COLS),
      rowCount: MIN_ROWS,
      colCount: MIN_COLS,
    };
  }
  const ref = ws['!ref'];
  if (!ref) {
    return {
      name,
      cells: emptyGrid(MIN_ROWS, MIN_COLS),
      rowCount: MIN_ROWS,
      colCount: MIN_COLS,
    };
  }
  const range = XLSX.utils.decode_range(ref);
  const rowCount = Math.max(range.e.r + 1, MIN_ROWS);
  const colCount = Math.max(range.e.c + 1, MIN_COLS);
  const cells: XlsxCell[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const row: XlsxCell[] = [];
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSX.CellObject | undefined;
      if (!cell) {
        row.push({ text: '' });
      } else {
        const rawType =
          cell.t === 'n' || cell.t === 's' || cell.t === 'b' || cell.t === 'd' || cell.t === 'e'
            ? cell.t
            : undefined;
        // SheetJS exposes the formula source via `cell.f` (without a leading `=`).
        // We canonicalize to always carry the "=" so callers can round-trip
        // user-typed input directly.
        const fSrc = (cell as { f?: string }).f;
        const formula = typeof fSrc === 'string' && fSrc.length > 0 ? `=${fSrc}` : undefined;
        row.push({
          text: formatCellDisplay(cell),
          rawType,
          style: extractCellStyle(cell),
          ...(formula ? { formula } : {}),
        });
      }
    }
    cells.push(row);
  }
  // SheetJS exposes merges as `!merges`: Array<{ s: {r,c}, e: {r,c} }>. Keep
  // the field undefined (not [] ) when there are none so the model stays
  // structurally identical to a fresh sheet.
  const wsMerges = (ws as { '!merges'?: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> })['!merges'];
  const merges =
    Array.isArray(wsMerges) && wsMerges.length > 0
      ? wsMerges.map((m) => ({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }))
      : undefined;
  return { name, cells, rowCount, colCount, merges };
}

function formatCellDisplay(cell: XLSX.CellObject): string {
  if (cell.w !== undefined) return String(cell.w);
  if (cell.v === null || cell.v === undefined) return '';
  if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
  return String(cell.v);
}

/** Pull whatever subset SheetJS surfaces in cell.s into our schema. */
function extractCellStyle(cell: XLSX.CellObject): XlsxCellStyle | undefined {
  const s = (cell as { s?: Record<string, unknown> }).s;
  if (!s || typeof s !== 'object') return undefined;
  const out: XlsxCellStyle = {};
  const font = s.font as { bold?: boolean; italic?: boolean; underline?: boolean; sz?: number; name?: string; color?: { rgb?: string } } | undefined;
  if (font) {
    if (font.bold) out.bold = true;
    if (font.italic) out.italic = true;
    if (font.underline) out.underline = true;
    if (typeof font.sz === 'number') out.fontSize = font.sz;
    if (typeof font.name === 'string' && font.name) out.fontFamily = font.name;
    if (font.color?.rgb) out.fontColor = stripAlpha(font.color.rgb);
  }
  const alignment = s.alignment as { horizontal?: string } | undefined;
  if (alignment?.horizontal === 'left' || alignment?.horizontal === 'center' || alignment?.horizontal === 'right') {
    out.align = alignment.horizontal;
  }
  // SheetJS may store bg under `fgColor` (foreground of the patternFill).
  const fill = s.fgColor as { rgb?: string } | undefined;
  if (fill?.rgb) out.bgColor = stripAlpha(fill.rgb);
  // Some files put it under `patternFill.fgColor`.
  const pattern = (s as { patternFill?: { fgColor?: { rgb?: string } } }).patternFill;
  if (!out.bgColor && pattern?.fgColor?.rgb) out.bgColor = stripAlpha(pattern.fgColor.rgb);
  const numFmt = (s as { numFmt?: string | number }).numFmt;
  if (typeof numFmt === 'string' && numFmt && numFmt !== 'General') out.numberFormat = numFmt;
  return Object.keys(out).length === 0 ? undefined : out;
}

/** "FFFF0000" → "FF0000". OOXML often prefixes alpha. */
function stripAlpha(rgb: string): string {
  return rgb.length === 8 ? rgb.slice(2) : rgb;
}

/** Build a SheetJS-style style object from our schema. */
function toSheetStyle(style: XlsxCellStyle): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const font: Record<string, unknown> = {};
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.underline) font.underline = true;
  if (style.fontSize) font.sz = style.fontSize;
  if (style.fontFamily) font.name = style.fontFamily;
  if (style.fontColor) font.color = { rgb: style.fontColor };
  if (Object.keys(font).length > 0) out.font = font;
  if (style.align) out.alignment = { horizontal: style.align };
  if (style.bgColor) out.fill = { patternType: 'solid', fgColor: { rgb: style.bgColor } };
  if (style.numberFormat) out.numFmt = style.numberFormat;
  return out;
}

function writeSheetInto(ws: XLSX.WorkSheet, sheet: XlsxSheet): void {
  // Drop any cells outside the new model bounds so row/col deletions actually
  // shrink the saved sheet (otherwise the trailing original cells stay).
  for (const key of Object.keys(ws)) {
    if (key.startsWith('!')) continue;
    const addr = XLSX.utils.decode_cell(key);
    if (addr.r >= sheet.rowCount || addr.c >= sheet.colCount) {
      delete ws[key];
    }
  }

  for (let r = 0; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const modelCell = sheet.cells[r]?.[c];
      const text = modelCell?.text ?? '';
      const prev = ws[addr] as XLSX.CellObject | undefined;
      if (text === '' && !modelCell?.style) {
        if (prev) delete ws[addr];
        continue;
      }
      // Build style: prefer our model's style; otherwise inherit from prev.
      const prevStyle = prev ? (prev as { s?: Record<string, unknown> }).s : undefined;
      const newStyle = modelCell?.style ? toSheetStyle(modelCell.style) : prevStyle;

      if (text === '' && newStyle) {
        // Empty cell carrying just a style — write a blank string cell.
        const cell: XLSX.CellObject = { t: 's', v: '' };
        if (newStyle) (cell as { s?: Record<string, unknown> }).s = newStyle;
        delete (cell as { w?: string }).w;
        ws[addr] = cell;
        continue;
      }
      const inferred = inferCell(text);
      const cell: XLSX.CellObject = {
        ...(prev ?? { t: inferred.t, v: inferred.v }),
        t: inferred.t,
        v: inferred.v,
      };
      if (newStyle) (cell as { s?: Record<string, unknown> }).s = newStyle;
      delete (cell as { w?: string }).w;
      // Round-trip the formula source. SheetJS expects `cell.f` *without* the
      // leading "=" — strip it. We also drop a stale `f` if the cell is no
      // longer a formula (user replaced the formula with a literal).
      if (modelCell?.formula) {
        (cell as { f?: string }).f = modelCell.formula.replace(/^=/, '');
      } else {
        delete (cell as { f?: string }).f;
      }
      ws[addr] = cell;
    }
  }
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(0, sheet.rowCount - 1), c: Math.max(0, sheet.colCount - 1) },
  });
  // Persist merges in SheetJS shape. We always overwrite (vs. merging into
  // any pre-existing `!merges`) so explicit unmerges propagate.
  if (sheet.merges && sheet.merges.length > 0) {
    (ws as { '!merges'?: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> })['!merges'] =
      sheet.merges.map((m) => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
  } else {
    delete (ws as { '!merges'?: unknown })['!merges'];
  }
}

function applyStylesToSheet(ws: XLSX.WorkSheet, sheet: XlsxSheet): void {
  for (let r = 0; r < sheet.rowCount; r++) {
    for (let c = 0; c < sheet.colCount; c++) {
      const modelCell = sheet.cells[r]?.[c];
      if (!modelCell?.style) continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSX.CellObject | undefined;
      if (!cell) continue;
      (cell as { s?: Record<string, unknown> }).s = toSheetStyle(modelCell.style);
    }
  }
}

function inferCell(text: string): { t: 'n' | 's' | 'b'; v: string | number | boolean } {
  if (text === 'TRUE' || text === 'true') return { t: 'b', v: true };
  if (text === 'FALSE' || text === 'false') return { t: 'b', v: false };
  if (/^-?\d+(\.\d+)?$/.test(text)) {
    const n = Number(text);
    if (Number.isFinite(n)) return { t: 'n', v: n };
  }
  return { t: 's', v: text };
}

function emptyGrid(rows: number, cols: number): XlsxCell[][] {
  const grid: XlsxCell[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: XlsxCell[] = [];
    for (let c = 0; c < cols; c++) row.push({ text: '' });
    grid.push(row);
  }
  return grid;
}

function toAoA(sheet: XlsxSheet): Array<Array<string | number | boolean>> {
  const out: Array<Array<string | number | boolean>> = [];
  for (let r = 0; r < sheet.rowCount; r++) {
    const row: Array<string | number | boolean> = [];
    for (let c = 0; c < sheet.colCount; c++) {
      const text = sheet.cells[r]?.[c]?.text ?? '';
      if (text === '') {
        row.push('');
        continue;
      }
      const inferred = inferCell(text);
      row.push(inferred.v);
    }
    out.push(row);
  }
  return out;
}

// ── row / col mutators ──────────────────────────────────────────────────
// These edit the in-memory model's cells grid. The serializer rebuilds the
// sheet from the model, so structural changes propagate on next save.

/**
 * Apply a row-axis transform to every merge in the sheet, dropping any merge
 * that collapses into a < 2-cell range. `op` returns the new (r1, r2) or null
 * to drop the merge entirely.
 */
function shiftMergesRow(
  merges: MergeRange[] | undefined,
  op: (m: MergeRange) => { r1: number; r2: number } | null,
): MergeRange[] | undefined {
  if (!merges || merges.length === 0) return merges;
  const out: MergeRange[] = [];
  for (const m of merges) {
    const next = op(m);
    if (!next) continue;
    if (next.r1 === next.r2 && m.c1 === m.c2) continue; // collapsed to 1 cell
    out.push({ ...m, r1: next.r1, r2: next.r2 });
  }
  return out.length > 0 ? out : undefined;
}

function shiftMergesCol(
  merges: MergeRange[] | undefined,
  op: (m: MergeRange) => { c1: number; c2: number } | null,
): MergeRange[] | undefined {
  if (!merges || merges.length === 0) return merges;
  const out: MergeRange[] = [];
  for (const m of merges) {
    const next = op(m);
    if (!next) continue;
    if (next.c1 === next.c2 && m.r1 === m.r2) continue;
    out.push({ ...m, c1: next.c1, c2: next.c2 });
  }
  return out.length > 0 ? out : undefined;
}

/** Insert a blank row at index `r` (existing rows shift down). */
export function insertRowAt(sheet: XlsxSheet, r: number): XlsxSheet {
  const cells = [...sheet.cells];
  const blank: XlsxCell[] = Array.from({ length: sheet.colCount }, () => ({ text: '' }));
  cells.splice(Math.max(0, Math.min(r, cells.length)), 0, blank);
  // Merge bookkeeping: if the inserted index lands at-or-before the merge's
  // top, both edges shift down; if it lands strictly inside, only the
  // bottom expands (the merge grows by one row).
  const merges = shiftMergesRow(sheet.merges, (m) => {
    if (r <= m.r1) return { r1: m.r1 + 1, r2: m.r2 + 1 };
    if (r <= m.r2) return { r1: m.r1, r2: m.r2 + 1 };
    return { r1: m.r1, r2: m.r2 };
  });
  return { ...sheet, cells, rowCount: cells.length, merges };
}

/** Remove the row at index `r`. Refuses to drop below 1 row. */
export function deleteRowAt(sheet: XlsxSheet, r: number): XlsxSheet {
  if (sheet.cells.length <= 1) return sheet;
  const cells = sheet.cells.filter((_, i) => i !== r);
  // Merges that become invalid (entirely on the deleted row, or shrink to a
  // single cell) drop out via the helper's collapse check.
  const merges = shiftMergesRow(sheet.merges, (m) => {
    if (r < m.r1) return { r1: m.r1 - 1, r2: m.r2 - 1 };
    if (r > m.r2) return { r1: m.r1, r2: m.r2 };
    // r is within [r1, r2] — shrink the bottom.
    if (m.r1 === m.r2) return null;
    return { r1: m.r1, r2: m.r2 - 1 };
  });
  return { ...sheet, cells, rowCount: cells.length, merges };
}

/** Insert a blank column at index `c` (existing cols shift right). */
export function insertColAt(sheet: XlsxSheet, c: number): XlsxSheet {
  const cells = sheet.cells.map((row) => {
    const next = [...row];
    next.splice(Math.max(0, Math.min(c, next.length)), 0, { text: '' });
    return next;
  });
  const merges = shiftMergesCol(sheet.merges, (m) => {
    if (c <= m.c1) return { c1: m.c1 + 1, c2: m.c2 + 1 };
    if (c <= m.c2) return { c1: m.c1, c2: m.c2 + 1 };
    return { c1: m.c1, c2: m.c2 };
  });
  return { ...sheet, cells, colCount: cells[0]?.length ?? sheet.colCount + 1, merges };
}

/** Remove the column at index `c`. Refuses to drop below 1 col. */
export function deleteColAt(sheet: XlsxSheet, c: number): XlsxSheet {
  if (sheet.colCount <= 1) return sheet;
  const cells = sheet.cells.map((row) => row.filter((_, i) => i !== c));
  const merges = shiftMergesCol(sheet.merges, (m) => {
    if (c < m.c1) return { c1: m.c1 - 1, c2: m.c2 - 1 };
    if (c > m.c2) return { c1: m.c1, c2: m.c2 };
    if (m.c1 === m.c2) return null;
    return { c1: m.c1, c2: m.c2 - 1 };
  });
  return {
    ...sheet,
    cells,
    colCount: cells[0]?.length ?? Math.max(1, sheet.colCount - 1),
    merges,
  };
}

/**
 * Add a merge over the rectangle [r1..r2, c1..c2]. Drops any pre-existing
 * merge that overlaps the rectangle (so a "merge over a merge" does what
 * Excel does — replaces, doesn't try to compose). Refuses single-cell merges.
 */
export function mergeRange(sheet: XlsxSheet, r1: number, c1: number, r2: number, c2: number): XlsxSheet {
  if (r1 === r2 && c1 === c2) return sheet;
  const next: MergeRange = {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
  const filtered = (sheet.merges ?? []).filter((m) => !rectsOverlap(m, next));
  return { ...sheet, merges: [...filtered, next] };
}

/**
 * Remove every merge that contains (r, c). `unmergeAt(sheet, anchor.r, anchor.c)`
 * is enough for the toolbar — the user clicks any cell inside the merge and
 * we drop it.
 */
export function unmergeAt(sheet: XlsxSheet, r: number, c: number): XlsxSheet {
  if (!sheet.merges || sheet.merges.length === 0) return sheet;
  const next = sheet.merges.filter((m) => !(r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2));
  if (next.length === sheet.merges.length) return sheet;
  return { ...sheet, merges: next.length > 0 ? next : undefined };
}

/** True iff (r, c) is covered by some merge but isn't its top-left anchor. */
export function isMergeCovered(sheet: XlsxSheet, r: number, c: number): boolean {
  if (!sheet.merges) return false;
  for (const m of sheet.merges) {
    if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2 && !(r === m.r1 && c === m.c1)) {
      return true;
    }
  }
  return false;
}

/** Returns the merge whose anchor is exactly (r, c), or null. */
export function mergeAtAnchor(sheet: XlsxSheet, r: number, c: number): MergeRange | null {
  if (!sheet.merges) return null;
  for (const m of sheet.merges) {
    if (m.r1 === r && m.c1 === c) return m;
  }
  return null;
}

function rectsOverlap(a: MergeRange, b: MergeRange): boolean {
  return !(a.r2 < b.r1 || a.r1 > b.r2 || a.c2 < b.c1 || a.c1 > b.c2);
}

// ── sheet (workbook-level) mutators ─────────────────────────────────────
// These operate on the XlsxModel rather than a single XlsxSheet. They keep
// `activeSheet` pointing at a real sheet name so the editor's selected-tab
// index stays in sync after add/delete.

/** Pick a sheet name not already in use, starting from `Sheet{n}`. */
function uniqueSheetName(sheets: XlsxSheet[], base = 'Sheet'): string {
  const taken = new Set(sheets.map((s) => s.name));
  let i = sheets.length + 1;
  while (taken.has(`${base}${i}`)) i++;
  return `${base}${i}`;
}

/** Append a new blank sheet. Returns the updated model and the new sheet's index. */
export function addSheet(model: XlsxModel, name?: string): { model: XlsxModel; index: number } {
  const finalName = (name ?? '').trim() || uniqueSheetName(model.sheets);
  // Disambiguate against existing names.
  const taken = new Set(model.sheets.map((s) => s.name));
  let resolved = finalName;
  let n = 2;
  while (taken.has(resolved)) resolved = `${finalName} (${n++})`;
  const sheet: XlsxSheet = {
    name: resolved,
    cells: emptyGrid(MIN_ROWS, MIN_COLS),
    rowCount: MIN_ROWS,
    colCount: MIN_COLS,
  };
  return {
    model: { ...model, sheets: [...model.sheets, sheet], activeSheet: resolved },
    index: model.sheets.length,
  };
}

/** Rename the sheet at `idx`. Refuses blank names and collisions (returns model unchanged). */
export function renameSheet(model: XlsxModel, idx: number, newName: string): XlsxModel {
  const trimmed = newName.trim();
  if (!trimmed) return model;
  if (idx < 0 || idx >= model.sheets.length) return model;
  if (model.sheets.some((s, i) => i !== idx && s.name === trimmed)) return model;
  const oldName = model.sheets[idx].name;
  return {
    ...model,
    sheets: model.sheets.map((s, i) => (i === idx ? { ...s, name: trimmed } : s)),
    activeSheet: model.activeSheet === oldName ? trimmed : model.activeSheet,
  };
}

/** Delete the sheet at `idx`. Refuses to drop the last remaining sheet. */
export function deleteSheet(model: XlsxModel, idx: number): XlsxModel {
  if (model.sheets.length <= 1) return model;
  if (idx < 0 || idx >= model.sheets.length) return model;
  const removed = model.sheets[idx];
  const sheets = model.sheets.filter((_, i) => i !== idx);
  return {
    ...model,
    sheets,
    activeSheet:
      model.activeSheet === removed.name ? sheets[Math.max(0, idx - 1)].name : model.activeSheet,
  };
}

/** Move the sheet at `from` to position `to` (both 0-based, in pre-move
 *  indexing — `to` is the destination index in the resulting array). Returns
 *  the model unchanged when the move would be a no-op or either index is
 *  out of range. `activeSheet` is preserved by name (not by index), so
 *  whatever sheet was active stays active even if its position shifted. */
export function moveSheet(model: XlsxModel, from: number, to: number): XlsxModel {
  if (from === to) return model;
  if (from < 0 || from >= model.sheets.length) return model;
  if (to < 0 || to >= model.sheets.length) return model;
  const sheets = [...model.sheets];
  const [moved] = sheets.splice(from, 1);
  sheets.splice(to, 0, moved);
  return { ...model, sheets };
}

/** Duplicate the sheet at `idx`, inserting the copy right after it. */
export function duplicateSheet(model: XlsxModel, idx: number): { model: XlsxModel; index: number } {
  if (idx < 0 || idx >= model.sheets.length) return { model, index: idx };
  const src = model.sheets[idx];
  const taken = new Set(model.sheets.map((s) => s.name));
  let copyName = `${src.name} (copy)`;
  let n = 2;
  while (taken.has(copyName)) copyName = `${src.name} (copy ${n++})`;
  // Deep-copy the cells grid so future edits don't alias the source.
  const cloned: XlsxSheet = {
    ...src,
    name: copyName,
    cells: src.cells.map((row) => row.map((c) => ({ ...c, style: c.style ? { ...c.style } : undefined }))),
    merges: src.merges ? src.merges.map((m) => ({ ...m })) : undefined,
    // Image bytes are immutable from the model's POV; we copy the array but
    // share the underlying Uint8Array so duplication stays cheap.
    images: src.images ? src.images.map((img) => ({ ...img })) : undefined,
  };
  const sheets = [...model.sheets];
  sheets.splice(idx + 1, 0, cloned);
  return { model: { ...model, sheets, activeSheet: copyName }, index: idx + 1 };
}

/** Convert column index (0-based) → letter ("A", "B", ..., "AA"). */
export function colIndexToLetter(c: number): string {
  let s = '';
  let n = c;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/** "B7" → { r: 6, c: 1 }. */
export function parseA1(addr: string): { r: number; c: number } | null {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}
