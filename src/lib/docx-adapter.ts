/**
 * docx ↔ in-memory model adapter.
 *
 * Round-trip story (path B, MVP):
 *   - Read: mammoth.convertToHtml gives us a clean structural HTML view.
 *     We walk it into a flat list of "blocks" (paragraph / heading-N / bullet).
 *     We sniff at the dominant inline style of the block text — if the entire
 *     paragraph is wrapped in <strong>/<em>/<u>, we lift those to block-level
 *     style flags. Mixed-style runs collapse to plain text (Phase 2 work).
 *   - Write: rebuild a fresh .docx via the `docx` library. Each block becomes
 *     a Paragraph with one TextRun carrying the block-level style + paragraph
 *     alignment.
 *
 * The user-facing banner in DocxEditor still says "lossy" — Phase 2 will
 * upgrade to per-run formatting via a contentEditable surface.
 */

import mammoth from 'mammoth';
import JSZip from 'jszip';
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LineRuleType,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  type IParagraphOptions,
  type ParagraphChild,
  type IRunOptions,
  type FileChild,
} from 'docx';

// R360 — single shared numbering reference for `kind: 'numbered'` blocks.
// Previously buildParagraph mapped both 'bullet' and 'numbered' to
// `bullet: { level: 0 }` (a bullet point) — losing the numbered-vs-bullet
// distinction every round-trip:
//   parseDocx (<ol> from mammoth) → kind: 'numbered'
//   editor displays as bullet styling (DocxEditor.tsx:2999)
//   serializeDocx → bullet output
//   open in Word → numbered list became bullets
// Same data-loss shape R268 (run styling) / R359 (table/image text) family.
// docx package wants a numbering config on the Document; declare it once and
// reference it from every numbered block via `numbering: { reference, level }`.
// `LevelFormat.DECIMAL` + `text: '%1.'` produces standard `1. 2. 3.` markers.
// Single shared reference means continuous numbering across the whole doc —
// matches markdown's mental model where ordered lists implicitly continue
// when adjacent. Per-list-restart (e.g., separate `<ol>` groups) would need
// per-list references; that's a Phase H+ refinement, the current single-
// reference is strictly better than the bullet fallback for any document
// with a single ordered-list pass.
const NUMBERED_LIST_REF = 'gd-numbered-default';

export type DocxBlockKind =
  | 'paragraph'
  | 'heading1'
  | 'heading2'
  | 'heading3'
  | 'heading4'
  | 'heading5'
  | 'heading6'
  | 'bullet'
  | 'numbered'
  | 'table'
  | 'image';

export type DocxAlign = 'left' | 'center' | 'right' | 'justify';

/** docx named highlight colors the editor's palette offers (subset of OOXML's w:highlight). */
export type DocxHighlightColor =
  | 'yellow'
  | 'green'
  | 'cyan'
  | 'magenta'
  | 'red'
  | 'darkYellow';

export interface DocxBlockStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  /** Font size in half-points × 2 — i.e. 24 = 12pt. We accept points; convert on write. */
  fontSize?: number;
  /** Hex without leading '#', e.g. "FF0000". */
  color?: string;
  /** Font family name, e.g. "Calibri" or "Microsoft JhengHei". */
  fontFamily?: string;
  /** Named highlight applied to every run of the block. */
  highlight?: DocxHighlightColor;
  /** Line-spacing multiplier (1 = single, 1.5, 2 …). Serialized as 240ths. */
  lineSpacing?: number;
}

/**
 * Per-character / per-run inline style. Only the toggle attributes (B / I /
 * U) live here for now — color / size / family are still tracked at the
 * block level because mammoth doesn't surface them per-run by default.
 */
export interface DocxRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}

/**
 * One text fragment with its inline style. `runs[]` survives parse →
 * serialize so a docx with mixed bold/italic spans round-trips losslessly
 * as long as the user doesn't edit the block. The textarea-based editor in
 * this phase still collapses `runs[]` to plain text on edit (same loss
 * profile as before) — selection-aware inline editing arrives in Phase F-2
 * once we move to a contentEditable surface.
 */
export interface DocxRun {
  text: string;
  style?: DocxRunStyle;
}

/**
 * Inline image block payload. The bytes survive parse → serialize so an
 * image inserted by the user round-trips losslessly. SVG isn't included
 * because docx's ImageRun requires a raster fallback for it (Phase 2).
 *
 * Dimensions are in pixels at the natural resolution of the image; the
 * docx serializer feeds them to ImageRun's `transformation` which
 * interprets them as on-page pixels (~96 dpi). Editing the dimensions
 * (via drag-resize) updates `widthPx`/`heightPx` in place.
 */
export interface DocxImage {
  data: Uint8Array;
  mediaType: 'jpg' | 'png' | 'gif' | 'bmp';
  widthPx: number;
  heightPx: number;
  /** Cached `data:` URL for display. Optional — re-derivable from `data` + `mediaType`. */
  dataUrl?: string;
}

export interface DocxBlock {
  /** Stable id so React keys survive edits. */
  id: string;
  kind: DocxBlockKind;
  /** Plain-text body for paragraph/heading/list blocks. Unused for tables. */
  text: string;
  style?: DocxBlockStyle;
  align?: DocxAlign;
  /** Populated only when kind === 'table'. rows[r][c] is plain cell text. */
  rows?: string[][];
  /**
   * Block-level external hyperlink. If set, the entire block.text becomes a
   * single clickable hyperlink in the rendered docx. We intentionally do
   * NOT support per-character link fragments at MVP — that would require
   * the run-level rich-text model planned for a later phase.
   */
  link?: string;
  /**
   * Optional run-level breakdown. When present, runs[] is the source of
   * truth for the block's text + inline B/I/U styling, and `text` is a
   * cached concatenation kept in sync. The serializer emits one TextRun
   * per `runs[]` entry; callers that only care about plain text can keep
   * reading `block.text`. Editing collapses runs back to a single entry
   * for now (Phase F-1).
   */
  runs?: DocxRun[];
  /** Populated only when kind === 'image'. */
  image?: DocxImage;
  /** Emit a page break before this block on serialize. */
  pageBreakBefore?: boolean;
}

/**
 * Page size in twips (= twentieths of a point, the unit used by `<w:pgSz>`).
 * 1440 twips = 1 inch. Default below is A4 (210×297 mm).
 */
export interface DocxPageSize {
  /** Width in twips. */
  w: number;
  /** Height in twips. */
  h: number;
}

/** Page margins in twips, mirroring `<w:pgMar>`. Default = 1 inch on all sides. */
export interface DocxPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
}

export const DEFAULT_PAGE_SIZE: DocxPageSize = { w: 11906, h: 16838 }; // A4
export const DEFAULT_PAGE_MARGINS: DocxPageMargins = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
};

export interface DocxModel {
  blocks: DocxBlock[];
  /** Parsed from `<w:pgSz>` in the first `<w:sectPr>`, falling back to A4. */
  pageSize: DocxPageSize;
  /** Parsed from `<w:pgMar>`, falling back to 1-inch margins. */
  pageMargins: DocxPageMargins;
}

let blockIdCounter = 0;
function nextBlockId(): string {
  blockIdCounter += 1;
  return `b${blockIdCounter}`;
}

/**
 * Parse docx bytes → editable block model. Empty input synthesizes a single
 * empty paragraph so the editor can render.
 *
 * mammoth gives us the structural HTML view but discards section properties
 * (page size / margins). We open the zip a second time to pluck `<w:pgSz>`
 * and `<w:pgMar>` from `word/document.xml` so the editor can render a
 * realistic page boundary.
 */
export async function parseDocx(bytes: Uint8Array): Promise<DocxModel> {
  if (bytes.byteLength === 0) {
    return {
      blocks: [{ id: nextBlockId(), kind: 'paragraph', text: '' }],
      pageSize: DEFAULT_PAGE_SIZE,
      pageMargins: DEFAULT_PAGE_MARGINS,
    };
  }
  const ab = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(ab).set(bytes);
  const [htmlResult, pageProps] = await Promise.all([
    mammoth.convertToHtml({ arrayBuffer: ab }),
    readPageProps(bytes),
  ]);
  return {
    blocks: htmlToBlocks(htmlResult.value),
    pageSize: pageProps.pageSize,
    pageMargins: pageProps.pageMargins,
  };
}

/**
 * Best-effort page-size/margin extraction. Failures fall back to defaults so
 * the editor is never blocked by an unusual sectPr layout.
 */
async function readPageProps(
  bytes: Uint8Array,
): Promise<{ pageSize: DocxPageSize; pageMargins: DocxPageMargins }> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const file = zip.file('word/document.xml');
    if (!file) return { pageSize: DEFAULT_PAGE_SIZE, pageMargins: DEFAULT_PAGE_MARGINS };
    const xml = await file.async('string');
    const pgSzMatch = xml.match(/<w:pgSz\b[^/>]*\/?>/);
    const pgMarMatch = xml.match(/<w:pgMar\b[^/>]*\/?>/);
    const pageSize: DocxPageSize = pgSzMatch
      ? {
          w: numAttr(pgSzMatch[0], 'w:w') ?? DEFAULT_PAGE_SIZE.w,
          h: numAttr(pgSzMatch[0], 'w:h') ?? DEFAULT_PAGE_SIZE.h,
        }
      : DEFAULT_PAGE_SIZE;
    const pageMargins: DocxPageMargins = pgMarMatch
      ? {
          top: numAttr(pgMarMatch[0], 'w:top') ?? DEFAULT_PAGE_MARGINS.top,
          right: numAttr(pgMarMatch[0], 'w:right') ?? DEFAULT_PAGE_MARGINS.right,
          bottom: numAttr(pgMarMatch[0], 'w:bottom') ?? DEFAULT_PAGE_MARGINS.bottom,
          left: numAttr(pgMarMatch[0], 'w:left') ?? DEFAULT_PAGE_MARGINS.left,
          header: numAttr(pgMarMatch[0], 'w:header') ?? DEFAULT_PAGE_MARGINS.header,
          footer: numAttr(pgMarMatch[0], 'w:footer') ?? DEFAULT_PAGE_MARGINS.footer,
        }
      : DEFAULT_PAGE_MARGINS;
    return { pageSize, pageMargins };
  } catch {
    return { pageSize: DEFAULT_PAGE_SIZE, pageMargins: DEFAULT_PAGE_MARGINS };
  }
}

function numAttr(tag: string, name: string): number | null {
  // Match attr="123"; allow either single or double quotes.
  const re = new RegExp(`\\b${name.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')}\\s*=\\s*["']([-0-9]+)["']`);
  const m = tag.match(re);
  return m ? Number(m[1]) : null;
}

/** Build a fresh docx Uint8Array from the block model. */
export async function serializeDocx(model: DocxModel): Promise<Uint8Array> {
  const children: FileChild[] = model.blocks.map((b) => {
    if (b.kind === 'table') return buildTable(b);
    if (b.kind === 'image' && b.image) return buildImageParagraph(b);
    return buildParagraph(b);
  });
  const doc = new Document({
    // R360 — register a single decimal numbering reference so blocks with
    // `kind: 'numbered'` can produce real numbered lists instead of falling
    // back to bullet styling. Defined once at Document level; every numbered
    // paragraph (via buildParagraph's switch) cites it. AlignmentType.START
    // is import-equivalent to docx's "start" — picked over LEFT so RTL
    // documents render markers on the leading edge automatically.
    numbering: {
      config: [
        {
          reference: NUMBERED_LIST_REF,
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.START,
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: model.pageSize.w, height: model.pageSize.h },
            margin: {
              top: model.pageMargins.top,
              right: model.pageMargins.right,
              bottom: model.pageMargins.bottom,
              left: model.pageMargins.left,
              header: model.pageMargins.header,
              footer: model.pageMargins.footer,
            },
          },
        },
        children,
      },
    ],
  });
  const blob = await Packer.toBlob(doc);
  const ab = await blob.arrayBuffer();
  return new Uint8Array(ab);
}

function buildTable(b: DocxBlock): Table {
  const rows = b.rows ?? [['']];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(
      (row) =>
        new TableRow({
          children: row.map(
            (cellText) =>
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: cellText })] })],
              }),
          ),
        }),
    ),
  });
}

// ── HTML → blocks ────────────────────────────────────────────────────────

function htmlToBlocks(html: string): DocxBlock[] {
  if (!html.trim()) {
    return [{ id: nextBlockId(), kind: 'paragraph', text: '' }];
  }
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div id="root">${html}</div>`, 'text/html');
  const root = doc.getElementById('root');
  if (!root) return [{ id: nextBlockId(), kind: 'paragraph', text: '' }];

  const blocks: DocxBlock[] = [];
  walk(root, blocks);
  if (blocks.length === 0) {
    blocks.push({ id: nextBlockId(), kind: 'paragraph', text: '' });
  }
  return blocks;
}

function walk(node: Element, out: DocxBlock[]): void {
  for (const child of Array.from(node.children)) {
    const tag = child.tagName.toLowerCase();
    switch (tag) {
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6': {
        const lvl = Number(tag.slice(1)) as 1 | 2 | 3 | 4 | 5 | 6;
        out.push(makeBlockFromElement(`heading${lvl}` as DocxBlockKind, child));
        break;
      }
      case 'p': {
        out.push(makeBlockFromElement('paragraph', child));
        break;
      }
      case 'ul':
        for (const li of Array.from(child.children)) {
          if (li.tagName.toLowerCase() !== 'li') continue;
          out.push(makeBlockFromElement('bullet', li));
        }
        break;
      case 'ol':
        for (const li of Array.from(child.children)) {
          if (li.tagName.toLowerCase() !== 'li') continue;
          out.push(makeBlockFromElement('numbered', li));
        }
        break;
      case 'table': {
        const trs = Array.from(child.querySelectorAll('tr'));
        const rows: string[][] = trs.map((tr) =>
          Array.from(tr.querySelectorAll('td, th')).map((td) => (td.textContent ?? '').trim()),
        );
        if (rows.length > 0) {
          // Pad ragged rows to a uniform column count.
          // R370 — compute cols via reduce (not `Math.max(...spread)`). For very
          // large tables — e.g., a docx exported from a spreadsheet with 65K+
          // rows — the spread form blows V8's argument-count limit (~65K
          // function args) and throws `RangeError: Maximum call stack size
          // exceeded`. mammoth's convertToHtml preserves every `<tr>` so
          // tall data dumps hit this every time. Same R328 fix shape for
          // dispatcher.ts:585 (`doExcelSetRange`'s maxRowLen), now
          // mirrored on the parse side. reduce has no spread, handles any
          // array size. Fallback to 0 when rows is empty (already filtered
          // above but defensive).
          const cols = rows.reduce<number>((m, r) => Math.max(m, r.length), 0);
          const padded = rows.map((r) => {
            const pad = [...r];
            while (pad.length < cols) pad.push('');
            return pad;
          });
          out.push({ id: nextBlockId(), kind: 'table', text: '', rows: padded });
        }
        break;
      }
      default:
        walk(child, out);
    }
  }
}

/**
 * Build a block from an element, sniffing style. We treat the block as
 * "bold" when *all* of its non-whitespace text content sits inside a
 * <strong>/<b>; same for italic (<em>/<i>) and underline (<u>). Mixed
 * inline styling collapses to plain text.
 */
function makeBlockFromElement(kind: DocxBlockKind, el: Element): DocxBlock {
  const text = (el.textContent ?? '').trim();
  if (!text) {
    return { id: nextBlockId(), kind, text: '' };
  }
  const style: DocxBlockStyle = {};
  if (allTextWrappedIn(el, ['strong', 'b'])) style.bold = true;
  if (allTextWrappedIn(el, ['em', 'i'])) style.italic = true;
  if (allTextWrappedIn(el, ['u'])) style.underline = true;
  if (allTextWrappedIn(el, ['s', 'del', 'strike'])) style.strikethrough = true;
  // Best-effort highlight recovery: mammoth doesn't surface w:highlight by
  // default, but if any span carries a background-color matching one of the
  // six palette colors, lift it to block level.
  const hl = sniffHighlight(el);
  if (hl) style.highlight = hl;
  // Hyperlink: lift the first descendant <a href> to a block-level link.
  // We only set it if the entire block's text sits inside that anchor —
  // otherwise the model can't represent a partial-line link without
  // dropping back to runs (planned for Phase F).
  let link: string | undefined;
  const firstA = el.querySelector('a[href]');
  if (firstA) {
    const href = firstA.getAttribute('href') ?? '';
    if (href && allTextWrappedIn(el, ['a'])) {
      link = href;
    }
  }
  // Run-level extraction. Walks every text node and collects the union of
  // its B/I/U ancestors. We only attach `runs` when at least one fragment
  // diverges from the block-level style — otherwise the existing
  // block.style path is sufficient and skipping `runs` keeps the model
  // tidy for plain blocks.
  const runs = extractRuns(el);
  const interesting = runs.length > 1 || runs.some((r) => r.style && Object.keys(r.style).length > 0);
  // mammoth doesn't preserve color/size by default — Phase 2.
  return {
    id: nextBlockId(),
    kind,
    text,
    style: Object.keys(style).length === 0 ? undefined : style,
    ...(link ? { link } : {}),
    ...(interesting ? { runs } : {}),
  };
}

/**
 * Walk every text node descendant and emit one DocxRun per leaf, capturing
 * which inline tags (b/strong/em/i/u) wrap it. Adjacent runs that share
 * the same style get coalesced so we don't emit five identical
 * `{bold:true}` runs for one bold word.
 *
 * NOTE: The walker preserves leading/trailing whitespace inside the
 * paragraph (no .trim()) so spaces between formatted spans round-trip.
 */
function extractRuns(el: Element): DocxRun[] {
  const out: DocxRun[] = [];
  const walker = (el.ownerDocument ?? document).createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let n: Node | null = walker.nextNode();
  while (n) {
    const text = n.textContent ?? '';
    if (text.length > 0) {
      const style: DocxRunStyle = {};
      let p: Element | null = n.parentElement;
      while (p && p !== el.parentElement) {
        const tag = p.tagName.toUpperCase();
        if (tag === 'STRONG' || tag === 'B') style.bold = true;
        if (tag === 'EM' || tag === 'I') style.italic = true;
        if (tag === 'U') style.underline = true;
        if (tag === 'S' || tag === 'DEL' || tag === 'STRIKE') style.strikethrough = true;
        p = p.parentElement;
      }
      const finalStyle = Object.keys(style).length > 0 ? style : undefined;
      const last = out[out.length - 1];
      if (last && runStylesEqual(last.style, finalStyle)) {
        last.text += text;
      } else {
        out.push({ text, style: finalStyle });
      }
    }
    n = walker.nextNode();
  }
  return out;
}

function runStylesEqual(a: DocxRunStyle | undefined, b: DocxRunStyle | undefined): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return (
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    !!a.underline === !!b.underline &&
    !!a.strikethrough === !!b.strikethrough
  );
}

/**
 * Map an inline `background-color` (on the element or any descendant) to one
 * of the six named docx highlight colors. Browsers normalize inline styles
 * to `rgb(...)`, so we match both hex and rgb forms plus CSS keywords.
 */
const HIGHLIGHT_BG_MAP: Record<string, DocxHighlightColor> = {
  '#ffff00': 'yellow', 'rgb(255,255,0)': 'yellow', yellow: 'yellow',
  '#00ff00': 'green', 'rgb(0,255,0)': 'green', lime: 'green',
  '#00ffff': 'cyan', 'rgb(0,255,255)': 'cyan', cyan: 'cyan', aqua: 'cyan',
  '#ff00ff': 'magenta', 'rgb(255,0,255)': 'magenta', magenta: 'magenta', fuchsia: 'magenta',
  '#ff0000': 'red', 'rgb(255,0,0)': 'red', red: 'red',
  '#808000': 'darkYellow', 'rgb(128,128,0)': 'darkYellow', olive: 'darkYellow',
};

function sniffHighlight(el: Element): DocxHighlightColor | undefined {
  const candidates = [el, ...Array.from(el.querySelectorAll('[style]'))];
  for (const c of candidates) {
    const bg = (c as HTMLElement).style?.backgroundColor ?? '';
    if (!bg) continue;
    const key = bg.toLowerCase().replace(/\s+/g, '');
    const named = HIGHLIGHT_BG_MAP[key];
    if (named) return named;
  }
  return undefined;
}

/** True if every text node descendant is inside one of `tags`. */
function allTextWrappedIn(el: Element, tags: string[]): boolean {
  const upperTags = tags.map((t) => t.toUpperCase());
  let allWrapped = true;
  let sawText = false;
  const walker = (el.ownerDocument ?? document).createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let n: Node | null = walker.nextNode();
  while (n) {
    if ((n.textContent ?? '').trim() !== '') {
      sawText = true;
      let p: Element | null = n.parentElement;
      let inside = false;
      while (p && p !== el.parentElement) {
        if (upperTags.includes(p.tagName)) {
          inside = true;
          break;
        }
        p = p.parentElement;
      }
      if (!inside) {
        allWrapped = false;
        break;
      }
    }
    n = walker.nextNode();
  }
  return sawText && allWrapped;
}

// ── blocks → docx ────────────────────────────────────────────────────────

function buildParagraph(b: DocxBlock): Paragraph {
  const s = b.style;
  // When the block carries a hyperlink we force blue + underline on the run
  // so it visually matches the convention. We override (not merge) any
  // user-set color so the link signal is unambiguous; plain block style
  // returns when the link is removed.
  const isLink = !!b.link;
  // Common styling that lives at the block level (font, size, color) and
  // applies to every run. Inline B/I/U comes from the run when runs[]
  // exists; otherwise from block.style.
  const blockShared: Partial<IRunOptions> = {
    ...(s?.color && !isLink ? { color: s.color } : {}),
    ...(isLink ? { color: '0563C1' } : {}),
    ...(s?.fontSize ? { size: s.fontSize * 2 } : {}), // pt → half-pt
    ...(s?.fontFamily ? { font: s.fontFamily } : {}),
    ...(s?.highlight ? { highlight: s.highlight } : {}),
  };
  // Build text runs. If runs[] is present we emit one TextRun per fragment,
  // each with its own B/I/U flags merged with the block defaults; the
  // hyperlink underline still wins via the OR. Otherwise we fall back to a
  // single run carrying block-level B/I/U.
  const textRuns: TextRun[] = b.runs && b.runs.length > 0
    ? b.runs.map((r) => {
        const rs = r.style ?? {};
        return new TextRun({
          ...blockShared,
          text: r.text,
          ...(rs.bold || s?.bold ? { bold: true } : {}),
          ...(rs.italic || s?.italic ? { italics: true } : {}),
          ...(rs.underline || s?.underline || isLink ? { underline: {} } : {}),
          ...(rs.strikethrough || s?.strikethrough ? { strike: true } : {}),
        });
      })
    : [
        new TextRun({
          ...blockShared,
          text: b.text,
          ...(s?.bold ? { bold: true } : {}),
          ...(s?.italic ? { italics: true } : {}),
          ...(s?.underline || isLink ? { underline: {} } : {}),
          ...(s?.strikethrough ? { strike: true } : {}),
        }),
      ];
  const children: ParagraphChild[] = isLink
    ? [new ExternalHyperlink({ link: b.link!, children: textRuns })]
    : textRuns;

  const baseOpts: IParagraphOptions = { children };
  let opts: IParagraphOptions = baseOpts;
  switch (b.kind) {
    case 'heading1':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_1 };
      break;
    case 'heading2':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_2 };
      break;
    case 'heading3':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_3 };
      break;
    case 'heading4':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_4 };
      break;
    case 'heading5':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_5 };
      break;
    case 'heading6':
      opts = { ...baseOpts, heading: HeadingLevel.HEADING_6 };
      break;
    case 'bullet':
      opts = { ...baseOpts, bullet: { level: 0 } };
      break;
    case 'numbered':
      // R360 — split from the shared `bullet: { level: 0 }` fallback.
      // Numbered blocks now cite the document-level numbering config
      // declared above (NUMBERED_LIST_REF), producing `1. 2. 3.` markers
      // in the output docx instead of the previous bullet impostors.
      opts = {
        ...baseOpts,
        numbering: { reference: NUMBERED_LIST_REF, level: 0 },
      };
      break;
    case 'paragraph':
    default:
      break;
  }
  if (b.align) {
    opts = { ...opts, alignment: alignToDocx(b.align) };
  }
  if (s?.lineSpacing) {
    // 240 = single spacing in 240ths of a line.
    opts = {
      ...opts,
      spacing: { line: Math.round(240 * s.lineSpacing), lineRule: LineRuleType.AUTO },
    };
  }
  if (b.pageBreakBefore) {
    opts = { ...opts, pageBreakBefore: true };
  }
  return new Paragraph(opts);
}

/**
 * Wrap a `DocxImage` block in a Paragraph carrying a single `ImageRun`. The
 * paragraph alignment is honoured (so users can center an image with the
 * align toolbar) and the image's pixel dimensions are passed through to
 * docx's `transformation` field — docx maps those to EMU on its own.
 *
 * `data` is copied into a fresh ArrayBuffer because the `docx` library
 * reads the bytes lazily during `Packer.toBlob`; if the caller mutates the
 * source Uint8Array between adding the run and packing, the wrong pixels
 * land in the document. The copy makes the call site fire-and-forget safe.
 */
function buildImageParagraph(b: DocxBlock): Paragraph {
  const img = b.image!;
  // Clone the bytes — see comment above. Subarray would alias.
  const buf = new Uint8Array(img.data.byteLength);
  buf.set(img.data);
  const run = new ImageRun({
    data: buf,
    type: img.mediaType,
    transformation: { width: img.widthPx, height: img.heightPx },
  });
  const opts: IParagraphOptions = { children: [run] };
  return new Paragraph(b.align ? { ...opts, alignment: alignToDocx(b.align) } : opts);
}

function alignToDocx(a: DocxAlign): typeof AlignmentType[keyof typeof AlignmentType] {
  switch (a) {
    case 'center':
      return AlignmentType.CENTER;
    case 'right':
      return AlignmentType.RIGHT;
    case 'justify':
      return AlignmentType.JUSTIFIED;
    case 'left':
    default:
      return AlignmentType.LEFT;
  }
}
