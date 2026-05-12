/**
 * pptx ↔ in-memory model adapter.
 *
 * Path B MVP+: pptx is OOXML in a zip. We don't have a Univer slides
 * package, so we go direct: parse `ppt/slides/slide*.xml` for `<a:r>` text
 * runs, expose each run's text + style (bold / italic / color / size) as
 * editable, and on write we rebuild each `<a:r>` with the model's text and
 * a merged-in `<a:rPr>` reflecting the toolbar choices.
 *
 * What this preserves: layout, images, theme, positioning, animations,
 * non-edited slides — anything not a text run is byte-preserved by JSZip.
 *
 * Limitations (called out in UI):
 *   - Orphan `<a:t>` not wrapped in `<a:r>` (rare; legacy fields) is skipped.
 *   - Add / remove slides + new text boxes are Phase 2.
 *
 * Brand-new pptx tabs are bootstrapped via `createBlankPptx()` — the editor
 * synthesizes a minimal valid one-slide deck so the user can edit immediately.
 */

import JSZip from 'jszip';

export interface PptxRunStyle {
  bold?: boolean;
  italic?: boolean;
  /** Underline. OOXML supports several underline styles (`sng`, `dbl`,
   * `wavy`, ...) on the `<a:rPr u="...">` attribute; we collapse them all
   * to a boolean for round-trip purposes and write back `u="sng"` when
   * true. Reading any non-`none` value preserves the underline state. */
  underline?: boolean;
  /** Hex without leading '#', e.g. "FF0000". */
  color?: string;
  /** Font size in points (OOXML stores hundredths-of-points, we expose pt). */
  size?: number;
  /** Font family (mapped to `<a:latin typeface="...">` inside `<a:rPr>`). */
  fontFamily?: string;
}

/** Position + size of a `<p:sp>` shape, in OOXML EMU (914400 EMU = 1 inch). */
export interface PptxFrame {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

export interface PptxTextRun {
  /** Stable id so React keys survive edits & re-parses. */
  id: string;
  /** Index of the slide this run belongs to (0-based, in slide-XML order). */
  slideIndex: number;
  /** 0-based position among the run's siblings within its slide, document-order. */
  runIndex: number;
  /** 0-based index of the parent `<p:sp>` shape, or -1 if the run lives in a
   *  `<p:graphicFrame>` table or other non-`<p:sp>` container. Runs sharing
   *  the same shapeIndex live in the same on-slide frame. */
  shapeIndex: number;
  /** Bounding box of the parent `<p:sp>`, copied so the canvas can position
   *  it without the renderer re-walking the XML. EMU units. For runs that
   *  inherit position from a layout (no `<a:xfrm>`), or live in a table, we
   *  synthesize a sensible fallback frame. */
  frame: PptxFrame;
  text: string;
  style?: PptxRunStyle;
}

export interface PptxSlide {
  index: number;
  zipPath: string;
  runs: PptxTextRun[];
  /**
   * Speaker notes plain text (concatenated across runs). Empty string =
   * either no notes part, or notes part exists but is empty. Track
   * `notesPath` separately to know whether to write back into an existing
   * notesSlide or synthesize a fresh one.
   */
  notesText: string;
  /** Zip path of the slide's notesSlide{N}.xml, or null if none exists. */
  notesPath: string | null;
  /**
   * Map of shapeIndex → image data URL for shapes whose `<p:spPr>` contains
   * a `<a:blipFill>` (i.e. picture-shapes inserted via `addPictureToSlide`,
   * or any shape with an image fill). The canvas overlays the data URL on
   * the shape's frame so users see the actual picture instead of an empty
   * box. Absent shapes are not images. Generated lazily during parsePptx.
   */
  pictures?: Record<number, string>;
}

export interface PptxModel {
  empty: boolean;
  slides: PptxSlide[];
  /** Slide width / height in EMU, from `<p:sldSz>` in `presentation.xml`.
   *  Defaults to the OOXML legacy 4:3 (10"×7.5") if missing. The canvas
   *  uses this to compute aspect ratio and to scale frame positions. */
  slideSize: { cx: number; cy: number };
}

/** OOXML defaults: 10"×7.5" (4:3). PowerPoint widescreen is 13.33"×7.5". */
export const DEFAULT_SLIDE_SIZE = { cx: 9144000, cy: 6858000 };

let runIdCounter = 0;
function nextRunId(): string {
  runIdCounter += 1;
  return `r${runIdCounter}`;
}

/** Parse pptx bytes → editable model. Empty input returns `empty: true`. */
export async function parsePptx(bytes: Uint8Array): Promise<PptxModel> {
  if (bytes.byteLength === 0) {
    return { empty: true, slides: [], slideSize: DEFAULT_SLIDE_SIZE };
  }
  const zip = await JSZip.loadAsync(bytes);
  const slidePaths = await resolveSlideOrder(zip);
  const slideSize = await readSlideSize(zip);
  if (slidePaths.length === 0) {
    return { empty: true, slides: [], slideSize };
  }
  const slides: PptxSlide[] = [];
  for (let i = 0; i < slidePaths.length; i++) {
    const path = slidePaths[i];
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async('string');
    const runs = extractRuns(xml, i, slideSize);
    // Notes are best-effort: a malformed _rels or notesSlide must not stop
    // the whole parse — the slide is still editable without speaker notes.
    let notesText = '';
    let notesPath: string | null = null;
    try {
      const notes = await readSlideNotes(zip, path);
      notesText = notes.text;
      notesPath = notes.path;
    } catch {
      // swallow — slide opens with empty notes
    }
    // Picture extraction is also best-effort — a missing media file or a
    // malformed rels entry should not block text editing.
    let pictures: Record<number, string> | undefined;
    try {
      const blipMap = extractShapeBlipRIds(xml);
      if (blipMap.size > 0) {
        pictures = await loadSlidePictures(zip, path, blipMap);
      }
    } catch {
      // swallow — slide opens without picture overlays
    }
    slides.push({
      index: i,
      zipPath: path,
      runs,
      notesText,
      notesPath,
      pictures,
    });
  }
  return { empty: slides.length === 0, slides, slideSize };
}

/** Read the deck's slide dimensions from `presentation.xml`. */
async function readSlideSize(zip: JSZip): Promise<{ cx: number; cy: number }> {
  const xml = await readStringOrNull(zip, 'ppt/presentation.xml');
  if (!xml) return DEFAULT_SLIDE_SIZE;
  const m = xml.match(/<p:sldSz\b[^>]*?\bcx="(\d+)"[^>]*?\bcy="(\d+)"/);
  if (!m) return DEFAULT_SLIDE_SIZE;
  return { cx: Number(m[1]), cy: Number(m[2]) };
}

/**
 * Build a minimal valid one-slide pptx from scratch — used when a brand-new
 * pptx tab is opened. Includes everything PowerPoint / LibreOffice need to
 * load the file: presentation root, one slide master + layout, a theme, and
 * a single blank slide carrying a title placeholder so the user has a text
 * frame to edit immediately.
 */
export async function createBlankPptx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', BLANK_CONTENT_TYPES);
  zip.file('_rels/.rels', BLANK_ROOT_RELS);
  zip.file('ppt/presentation.xml', BLANK_PRESENTATION);
  zip.file('ppt/_rels/presentation.xml.rels', BLANK_PRESENTATION_RELS);
  zip.file('ppt/slides/slide1.xml', BLANK_SLIDE);
  zip.file('ppt/slides/_rels/slide1.xml.rels', BLANK_SLIDE_RELS);
  zip.file('ppt/slideLayouts/slideLayout1.xml', BLANK_SLIDE_LAYOUT);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', BLANK_SLIDE_LAYOUT_RELS);
  zip.file('ppt/slideMasters/slideMaster1.xml', BLANK_SLIDE_MASTER);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', BLANK_SLIDE_MASTER_RELS);
  zip.file('ppt/theme/theme1.xml', BLANK_THEME);
  return await zip.generateAsync({ type: 'uint8array' });
}

// ── structural slide ops ────────────────────────────────────────────────

/**
 * Duplicate the slide at `slideIndex` (insert the copy right after it). The
 * copy carries the same text/layout/relationships; users edit it like any
 * other slide. Returns new pptx bytes.
 */
export async function duplicateSlide(originalBytes: Uint8Array, slideIndex: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const src = order[slideIndex];
  const srcXml = await mustReadString(zip, src.path);
  const srcRelsPath = `ppt/slides/_rels/${basename(src.path)}.rels`;
  const srcRelsXml = await readStringOrNull(zip, srcRelsPath);

  // Pick a free filename slide{N}.xml.
  const usedNums = new Set<number>();
  zip.forEach((p) => {
    const m = p.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (m) usedNums.add(Number(m[1]));
  });
  let nextNum = 1;
  while (usedNums.has(nextNum)) nextNum += 1;
  const newPath = `ppt/slides/slide${nextNum}.xml`;
  const newRelsPath = `ppt/slides/_rels/slide${nextNum}.xml.rels`;

  zip.file(newPath, srcXml);
  if (srcRelsXml) zip.file(newRelsPath, srcRelsXml);

  // Update [Content_Types].xml — add Override for new slide.
  await mutateContentTypes(zip, (xml) => {
    if (xml.includes(`PartName="/${newPath}"`)) return xml;
    const insert = `<Override PartName="/${newPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    return xml.replace(/<\/Types>\s*$/, insert + '</Types>');
  });

  // Update presentation.xml.rels — add Relationship for new slide.
  const newRId = await addPresentationRel(zip, `slides/slide${nextNum}.xml`);

  // Update presentation.xml — insert <p:sldId> after the source.
  await mutatePresentation(zip, (xml) => insertSldIdAfter(xml, src.rId, newRId));

  return await zip.generateAsync({ type: 'uint8array' });
}

/** Remove the slide at `slideIndex`. Returns new pptx bytes. */
export async function deleteSlide(originalBytes: Uint8Array, slideIndex: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  if (order.length <= 1) {
    throw new Error('pptx_min_one_slide: a presentation must keep at least one slide');
  }
  const target = order[slideIndex];

  // Drop from presentation.xml
  await mutatePresentation(zip, (xml) => removeSldId(xml, target.rId));
  // Drop the Relationship
  await mutatePresentationRels(zip, (xml) =>
    xml.replace(new RegExp(`<Relationship[^/]*?Id="${target.rId}"[^/]*?/>`, 'g'), ''),
  );
  // Drop the Override
  await mutateContentTypes(zip, (xml) =>
    xml.replace(new RegExp(`<Override[^/]*?PartName="/${escapeRegex(target.path)}"[^/]*?/>`, 'g'), ''),
  );
  // Remove slide file + its rels (best-effort).
  zip.remove(target.path);
  zip.remove(`ppt/slides/_rels/${basename(target.path)}.rels`);

  return await zip.generateAsync({ type: 'uint8array' });
}

/** Move the slide at `from` to position `to` (0-based, post-removal index). */
export async function reorderSlides(originalBytes: Uint8Array, from: number, to: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (from < 0 || from >= order.length || to < 0 || to >= order.length) {
    throw new Error(`slide_out_of_range: from=${from} to=${to}`);
  }
  if (from === to) return originalBytes;

  await mutatePresentation(zip, (xml) => moveSldId(xml, from, to));
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Re-load original bytes and patch each `<a:r>` block per the model. Non-text
 * content is preserved verbatim. Also writes back speaker notes per slide.
 */
export async function serializePptx(model: PptxModel, originalBytes: Uint8Array): Promise<Uint8Array> {
  if (originalBytes.byteLength === 0) {
    throw new Error('pptx_serialize_empty: brand-new pptx not yet supported');
  }
  const zip = await JSZip.loadAsync(originalBytes);
  for (const slide of model.slides) {
    const file = zip.file(slide.zipPath);
    if (!file) continue;
    const xml = await file.async('string');
    const updated = applyRuns(xml, slide.runs);
    zip.file(slide.zipPath, updated);
    await writeSlideNotes(zip, slide);
  }
  const out = await zip.generateAsync({ type: 'uint8array' });
  return out;
}

/**
 * Append a new text-box `<p:sp>` to the slide's `<p:spTree>`, carrying a
 * single placeholder text run. Returns new pptx bytes.
 */
export async function addTextBoxToSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  text = '新文字框',
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);

  // Find max cNvPr id within the slide so the new sp doesn't collide.
  let maxId = 1;
  const idRe = /<p:cNvPr\b[^>]*?\bid="(\d+)"/g;
  for (let m = idRe.exec(xml); m !== null; m = idRe.exec(xml)) {
    const v = Number(m[1]);
    if (v > maxId) maxId = v;
  }
  const newId = maxId + 1;
  const sp =
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${newId}" name="TextBox ${newId}"/>` +
    `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="1524000" y="1524000"/><a:ext cx="3048000" cy="685800"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:r><a:rPr lang="zh-TW" sz="1800"/><a:t>${encodeXmlEntities(text)}</a:t></a:r></a:p>` +
    `</p:txBody>` +
    `</p:sp>`;
  const updated = xml.replace(/<\/p:spTree>/, sp + '</p:spTree>');
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Built-in shape kinds we expose in the editor's shape picker. The string
 * values map directly to OOXML `prst` preset names so we can drop them into
 * `<a:prstGeom prst="...">` without translation.
 */
export type PptxShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightArrow';

/**
 * Append a new shape `<p:sp>` to the slide's `<p:spTree>`. Differs from
 * `addTextBoxToSlide` in three ways:
 *  1. No `txBox="1"` flag — this is a real shape, not a chrome-less text box.
 *  2. Default solid fill (`4472C4` = PowerPoint accent 1) so the shape is
 *     visible against the slide background.
 *  3. White, center-aligned label run so the shape participates in the
 *     parser's run inventory (the parser only surfaces shapes that contain
 *     `<a:r>`). Without a run, the shape would render but be invisible to
 *     edit / delete / F&R flows.
 */
export async function addShapeToSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  kind: PptxShapeKind,
  text = '',
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);

  let maxId = 1;
  const idRe = /<p:cNvPr\b[^>]*?\bid="(\d+)"/g;
  for (let m = idRe.exec(xml); m !== null; m = idRe.exec(xml)) {
    const v = Number(m[1]);
    if (v > maxId) maxId = v;
  }
  const newId = maxId + 1;
  // Right-arrow looks better wider than tall; everything else stays square-ish.
  const cx = kind === 'rightArrow' ? 3048000 : 2286000;
  const cy = kind === 'rightArrow' ? 1143000 : 1524000;
  const sp =
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${newId}" name="Shape ${newId}"/>` +
    `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="1524000" y="1524000"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${kind}"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"/>` +
    `<a:r><a:rPr lang="zh-TW" sz="1800"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>` +
    `<a:t>${encodeXmlEntities(text)}</a:t></a:r></a:p>` +
    `</p:txBody>` +
    `</p:sp>`;
  const updated = xml.replace(/<\/p:spTree>/, sp + '</p:spTree>');
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Delete the `<p:sp>` containing the run at runIndex on the given slide.
 * Returns new pptx bytes. Refuses to delete if removing the sp would empty
 * the slide of all editable runs.
 */
export async function deleteTextBoxFromSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  runIndex: number,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);

  // Walk <p:sp> blocks and figure out which one contains the runIndex'th <a:r>.
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let runCounter = 0;
  let target: { start: number; end: number } | null = null;
  for (let m = spRe.exec(xml); m !== null; m = spRe.exec(xml)) {
    const block = m[0];
    const runsInBlock = block.match(/<a:r>[\s\S]*?<\/a:r>/g) ?? [];
    if (runIndex >= runCounter && runIndex < runCounter + runsInBlock.length) {
      target = { start: m.index, end: m.index + block.length };
      break;
    }
    runCounter += runsInBlock.length;
  }
  if (!target) throw new Error(`run_out_of_range: slide ${slideIndex} run ${runIndex}`);

  const updated = xml.slice(0, target.start) + xml.slice(target.end);
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Duplicate the Nth `<p:sp>` on a slide and land the copy at `(x, y)` (EMU).
 * The original stays in place — this is the Adobe PS/AI/ID/Figma Alt-drag
 * convention (and PowerPoint's Ctrl-drag). The new shape inherits the
 * source's geometry/fill/text and is inserted immediately after the source
 * in `<p:spTree>` order so it stacks above the original (duplicate-above-
 * source matches every Adobe app's z-order behaviour).
 *
 * Why we rewrite cNvPr id rather than copy it: PPTX requires unique
 * `<p:cNvPr id>` within `<p:cSld>`. PowerPoint tolerates duplicates on read
 * but other consumers (Keynote, LibreOffice, our own parser's run inventory)
 * key on it, so we always assign max(existing) + 1.
 */
export async function duplicateShapeOnSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  shapeIndex: number,
  patch: { x: number; y: number },
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);

  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let i = 0;
  let target: { start: number; end: number; block: string } | null = null;
  for (let m = spRe.exec(xml); m !== null; m = spRe.exec(xml)) {
    if (i === shapeIndex) {
      target = { start: m.index, end: m.index + m[0].length, block: m[0] };
      break;
    }
    i += 1;
  }
  if (!target) throw new Error(`shape_out_of_range: slide ${slideIndex} shape ${shapeIndex}`);

  // Allocate a fresh cNvPr id (slide-wide max + 1). We only rewrite the
  // *first* cNvPr in the cloned block — that's the shape's own id; nested
  // cNvPr's (e.g. on grpSpPr children) don't apply here since we're cloning
  // a flat <p:sp>.
  let maxId = 1;
  const idRe = /<p:cNvPr\b[^>]*?\bid="(\d+)"/g;
  for (let m = idRe.exec(xml); m !== null; m = idRe.exec(xml)) {
    const v = Number(m[1]);
    if (v > maxId) maxId = v;
  }
  const newId = maxId + 1;
  const reidBlock = target.block.replace(/<p:cNvPr\b([^>]*?)\bid="\d+"/, `<p:cNvPr$1id="${newId}"`);
  const movedBlock = patchShapeXfrm(reidBlock, { x: patch.x, y: patch.y });

  const updated =
    xml.slice(0, target.end) + movedBlock + xml.slice(target.end);
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Insert an image as a new picture-shape on the slide. We model the picture
 * as a `<p:sp>` (rather than `<p:pic>`) so the existing drag/resize/delete
 * pipeline — which is keyed on shapeIndex — works without specialization.
 * The shape carries:
 *   - `<a:prstGeom prst="rect">` so the image displays in a plain rectangle.
 *   - `<a:blipFill r:embed="rIdN">` referencing a media part we add below.
 *   - A synthetic empty `<a:r>` so `extractRuns` surfaces the shape (the
 *     parser only inventories shapes that contain at least one run; without
 *     this stub the picture would render but be invisible to the editor).
 *
 * Default placement: 4-inch-wide preserving aspect, centered on the slide.
 * Caller passes `imageSize` (px) so we can compute the aspect; pass {1,1}
 * if the size is unknown — the picture will load square at first and the
 * user can resize manually.
 */
export async function addPictureToSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  imageBytes: Uint8Array,
  ext: string,
  mime: string,
  imageSize: { width: number; height: number },
  // Optional drop-point anchor in EMU (top-left of the picture). When
  // provided, the picture is placed at this position instead of centered;
  // the caller (e.g. PptxEditor's drop handler) is responsible for converting
  // pixel coords to EMU. The picture is still clamped to the slide bounds so
  // a careless drop near the edge can't push it off-screen.
  anchor?: { x: number; y: number },
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const slidePath = order[slideIndex].path;
  const xml = await mustReadString(zip, slidePath);

  // 1. Pick a free media filename. Walk every `ppt/media/imageN.*` so we
  // never collide with images already in the deck (a slide we're inserting
  // into might be slide 5, but image10 could already exist).
  const usedMedia = new Set<number>();
  zip.forEach((p) => {
    const m = p.match(/^ppt\/media\/image(\d+)\./);
    if (m) usedMedia.add(Number(m[1]));
  });
  let n = 1;
  while (usedMedia.has(n)) n += 1;
  const mediaPath = `ppt/media/image${n}.${ext}`;
  zip.file(mediaPath, imageBytes);

  // 2. Make sure [Content_Types].xml advertises this extension. If a
  // `<Default Extension="png" ...>` already exists we leave it; otherwise
  // append one so PowerPoint/Keynote know how to read the new media part.
  await mutateContentTypes(zip, (ctXml) => {
    if (new RegExp(`<Default\\b[^>]*?\\bExtension="${escapeRegex(ext)}"`).test(ctXml)) {
      return ctXml;
    }
    const insert = `<Default Extension="${ext}" ContentType="${mime}"/>`;
    return ctXml.replace(/<Types\b[^>]*>/, (m) => m + insert);
  });

  // 3. Add a Relationship in the slide's _rels file pointing to the media
  // part. The blipFill in the shape will reference this rId.
  const slideRelsPath = `ppt/slides/_rels/${basename(slidePath)}.rels`;
  const slideRelsXml =
    (await readStringOrNull(zip, slideRelsPath)) ??
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const usedRIds = new Set<number>();
  const rIdRe = /\bId="rId(\d+)"/g;
  for (let m = rIdRe.exec(slideRelsXml); m !== null; m = rIdRe.exec(slideRelsXml)) {
    usedRIds.add(Number(m[1]));
  }
  let rIdN = 1;
  while (usedRIds.has(rIdN)) rIdN += 1;
  const newRId = `rId${rIdN}`;
  const insertRel =
    `<Relationship Id="${newRId}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
    `Target="../media/image${n}.${ext}"/>`;
  const updatedSlideRels = slideRelsXml.replace(
    /<\/Relationships>\s*$/,
    insertRel + '</Relationships>',
  );
  zip.file(slideRelsPath, updatedSlideRels);

  // 4. Allocate a fresh slide-wide cNvPr id (PPTX requires uniqueness inside
  // <p:cSld>). Same pattern as addShapeToSlide / duplicateShapeOnSlide.
  let maxId = 1;
  const idRe = /<p:cNvPr\b[^>]*?\bid="(\d+)"/g;
  for (let m = idRe.exec(xml); m !== null; m = idRe.exec(xml)) {
    const v = Number(m[1]);
    if (v > maxId) maxId = v;
  }
  const newId = maxId + 1;

  // 5. Compute geometry: 4-inch-wide cap (or smaller if slide is narrow),
  // height by aspect ratio. EMU: 914400 = 1 inch.
  const slideSize = await readSlideSize(zip);
  const targetCx = Math.min(3657600, slideSize.cx - 914400);
  const aspect = imageSize.height / Math.max(1, imageSize.width);
  const cx = Math.max(914400, targetCx);
  const cy = Math.max(914400, Math.round(cx * aspect));
  // Position: drop anchor when provided (Round 75), otherwise centered.
  // When anchored, treat (anchor.x, anchor.y) as the desired top-left and
  // clamp into [0, slideSize − cx/cy] so a corner drop can't push the shape
  // off-stage. A negative coordinate (drop above/left of canvas) clamps to 0.
  let x: number;
  let y: number;
  if (anchor) {
    x = Math.max(0, Math.min(slideSize.cx - cx, Math.round(anchor.x)));
    y = Math.max(0, Math.min(slideSize.cy - cy, Math.round(anchor.y)));
  } else {
    x = Math.max(0, Math.round((slideSize.cx - cx) / 2));
    y = Math.max(0, Math.round((slideSize.cy - cy) / 2));
  }

  // 6. Append the picture-shape to <p:spTree>. The blipFill replaces the
  // solidFill that addShapeToSlide uses — same family of OOXML fill tags.
  const sp =
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${newId}" name="Picture ${newId}"/>` +
    `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:blipFill><a:blip r:embed="${newRId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` +
    `</p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>` +
    `<a:p><a:r><a:rPr lang="zh-TW" sz="1800"/><a:t></a:t></a:r></a:p>` +
    `</p:txBody>` +
    `</p:sp>`;
  const updated = xml.replace(/<\/p:spTree>/, sp + '</p:spTree>');
  zip.file(slidePath, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Walk the slide's `<p:sp>` blocks and collect the embed rId of any shape
 * carrying a `<a:blipFill>`. Returned map is keyed by the shape's index in
 * the same order as `extractRuns` assigns shapeIndex, so callers can join
 * by shapeIndex without ambiguity.
 */
function extractShapeBlipRIds(xml: string): Map<number, string> {
  const out = new Map<number, string>();
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let i = 0;
  for (let m = spRe.exec(xml); m !== null; m = spRe.exec(xml)) {
    const block = m[0];
    const blip = block.match(/<a:blipFill\b[\s\S]*?<a:blip\b[^>]*?r:embed="([^"]+)"/);
    if (blip) out.set(i, blip[1]);
    i += 1;
  }
  return out;
}

/**
 * Resolve each blipFill rId in `shapeBlipRIds` to a `data:` URL by looking
 * up the slide's _rels file and reading the underlying media part as
 * base64. Best-effort — missing rels or media files just drop that entry,
 * never throw, so the slide remains editable.
 */
async function loadSlidePictures(
  zip: JSZip,
  slidePath: string,
  shapeBlipRIds: Map<number, string>,
): Promise<Record<number, string>> {
  const out: Record<number, string> = {};
  const relsPath = `ppt/slides/_rels/${basename(slidePath)}.rels`;
  const relsXml = await readStringOrNull(zip, relsPath);
  if (!relsXml) return out;
  const relMap = new Map<string, string>();
  const re = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/g;
  for (let m = re.exec(relsXml); m !== null; m = re.exec(relsXml)) {
    relMap.set(m[1], m[2]);
  }
  for (const [shapeIdx, rId] of shapeBlipRIds) {
    const target = relMap.get(rId);
    if (!target) continue;
    const mediaPath = resolveRelTarget('ppt/slides/', target);
    const file = zip.file(mediaPath);
    if (!file) continue;
    const ext = (mediaPath.match(/\.([a-z0-9]+)$/i)?.[1] ?? 'png').toLowerCase();
    const mime =
      ext === 'jpg' || ext === 'jpeg'
        ? 'image/jpeg'
        : ext === 'gif'
          ? 'image/gif'
          : ext === 'svg'
            ? 'image/svg+xml'
            : ext === 'bmp'
              ? 'image/bmp'
              : ext === 'webp'
                ? 'image/webp'
                : 'image/png';
    const base64 = await file.async('base64');
    out[shapeIdx] = `data:${mime};base64,${base64}`;
  }
  return out;
}

/**
 * Reposition (and optionally resize) the Nth `<p:sp>` on a slide. Mutates
 * the shape's `<a:xfrm>` so PowerPoint and round-trips both see the new
 * geometry. If the shape has no explicit `<a:xfrm>` (because it inherits
 * from the slide layout), one is injected before `</p:spPr>`. Coordinates
 * are EMU; null fields keep the current value.
 */
export async function moveShapeOnSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  shapeIndex: number,
  patch: { x?: number; y?: number; cx?: number; cy?: number },
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);

  // Find the Nth <p:sp> block by walking the same regex used elsewhere so
  // index semantics line up with the parsed model's shapeIndex.
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  let i = 0;
  let target: { start: number; end: number; block: string } | null = null;
  for (let m = spRe.exec(xml); m !== null; m = spRe.exec(xml)) {
    if (i === shapeIndex) {
      target = { start: m.index, end: m.index + m[0].length, block: m[0] };
      break;
    }
    i += 1;
  }
  if (!target) throw new Error(`shape_out_of_range: slide ${slideIndex} shape ${shapeIndex}`);

  const updatedBlock = patchShapeXfrm(target.block, patch);
  const updated = xml.slice(0, target.start) + updatedBlock + xml.slice(target.end);
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

/**
 * Inside a single `<p:sp>` block, rewrite (or inject) `<a:off>` / `<a:ext>`
 * inside the shape's `<a:xfrm>`. Falls back to a default 1in × 1in box at
 * the requested position if no current geometry can be found, so callers
 * never silently no-op.
 */
function patchShapeXfrm(
  block: string,
  patch: { x?: number; y?: number; cx?: number; cy?: number },
): string {
  // Try to locate an existing <a:xfrm>...</a:xfrm> inside <p:spPr>.
  const xfrmRe = /<a:xfrm\b[^>]*>([\s\S]*?)<\/a:xfrm>/;
  const xfrmMatch = block.match(xfrmRe);
  if (xfrmMatch) {
    const inner = xfrmMatch[1];
    const off = inner.match(/<a:off\b[^/>]*\/?>/);
    const ext = inner.match(/<a:ext\b[^/>]*\/?>/);
    const curX = off ? Number(numAttrOf(off[0], 'x') ?? 0) : 0;
    const curY = off ? Number(numAttrOf(off[0], 'y') ?? 0) : 0;
    const curCx = ext ? Number(numAttrOf(ext[0], 'cx') ?? 0) : 914400;
    const curCy = ext ? Number(numAttrOf(ext[0], 'cy') ?? 0) : 914400;
    const nx = patch.x ?? curX;
    const ny = patch.y ?? curY;
    const ncx = patch.cx ?? curCx;
    const ncy = patch.cy ?? curCy;
    const newXfrm =
      `<a:xfrm><a:off x="${nx}" y="${ny}"/><a:ext cx="${ncx}" cy="${ncy}"/></a:xfrm>`;
    return block.slice(0, xfrmMatch.index!) + newXfrm + block.slice(xfrmMatch.index! + xfrmMatch[0].length);
  }
  // No xfrm — inject one before the </p:spPr> close.
  const nx = patch.x ?? 1524000;
  const ny = patch.y ?? 1524000;
  const ncx = patch.cx ?? 3048000;
  const ncy = patch.cy ?? 685800;
  const xfrm = `<a:xfrm><a:off x="${nx}" y="${ny}"/><a:ext cx="${ncx}" cy="${ncy}"/></a:xfrm>`;
  return block.replace(/<\/p:spPr>/, xfrm + '</p:spPr>');
}

function numAttrOf(tag: string, name: string): string | null {
  const re = new RegExp(`\\b${name}="([-0-9]+)"`);
  const m = tag.match(re);
  return m ? m[1] : null;
}

// ── layout presets ───────────────────────────────────────────────────────

export type PptxLayoutId =
  | 'titleSlide'
  | 'titleContent'
  | 'twoContent'
  | 'sectionHeader'
  | 'titleOnly'
  | 'blank';

export interface PptxLayoutDef {
  id: PptxLayoutId;
  /** Localized label shown in the picker. */
  label: string;
  /** Short description for the menu hint. */
  hint: string;
}

export const PPTX_LAYOUTS: PptxLayoutDef[] = [
  { id: 'titleSlide', label: '標題投影片', hint: '置中的主標題 + 副標題' },
  { id: 'titleContent', label: '標題與內容', hint: '上方標題 + 一個內容區' },
  { id: 'twoContent', label: '兩欄內容', hint: '上方標題 + 左右兩個內容區' },
  { id: 'sectionHeader', label: '章節標題', hint: '大標題 + 小說明' },
  { id: 'titleOnly', label: '僅標題', hint: '只有頂部一個標題列' },
  { id: 'blank', label: '空白', hint: '清空所有文字框' },
];

interface LayoutBoxSpec {
  /** All in EMU. */
  x: number;
  y: number;
  cx: number;
  cy: number;
  text: string;
  /** Hundredths-of-pt for OOXML <a:rPr sz="..."/>. e.g. 4400 = 44pt. */
  size: number;
  bold?: boolean;
  /** Paragraph horizontal alignment for <a:pPr algn>. */
  align?: 'l' | 'ctr' | 'r';
  /** Vertical anchor for <a:bodyPr anchor>. */
  anchor?: 't' | 'ctr' | 'b';
}

/** Build the placeholder list for a given layout, sized to the deck's slide. */
function buildLayoutBoxes(
  layout: PptxLayoutId,
  slideSize: { cx: number; cy: number },
): LayoutBoxSpec[] {
  const W = slideSize.cx;
  const H = slideSize.cy;
  const m = 457200; // 0.5 inch margin
  switch (layout) {
    case 'titleSlide':
      return [
        { x: m, y: Math.round(H * 0.35), cx: W - 2 * m, cy: Math.round(H * 0.18), text: '標題', size: 4400, align: 'ctr', anchor: 'ctr' },
        { x: m, y: Math.round(H * 0.55), cx: W - 2 * m, cy: Math.round(H * 0.10), text: '副標題', size: 2400, align: 'ctr', anchor: 'ctr' },
      ];
    case 'titleContent':
      return [
        { x: m, y: m, cx: W - 2 * m, cy: Math.round(H * 0.13), text: '標題', size: 3600, bold: true, align: 'l', anchor: 'ctr' },
        { x: m, y: Math.round(H * 0.18), cx: W - 2 * m, cy: Math.round(H * 0.78), text: '在這裡輸入內容…', size: 2000, align: 'l', anchor: 't' },
      ];
    case 'twoContent': {
      const colW = Math.round((W - 3 * m) / 2);
      return [
        { x: m, y: m, cx: W - 2 * m, cy: Math.round(H * 0.13), text: '標題', size: 3600, bold: true, align: 'l', anchor: 'ctr' },
        { x: m, y: Math.round(H * 0.18), cx: colW, cy: Math.round(H * 0.78), text: '左側內容', size: 2000, align: 'l', anchor: 't' },
        { x: m * 2 + colW, y: Math.round(H * 0.18), cx: colW, cy: Math.round(H * 0.78), text: '右側內容', size: 2000, align: 'l', anchor: 't' },
      ];
    }
    case 'sectionHeader':
      return [
        { x: m, y: Math.round(H * 0.30), cx: W - 2 * m, cy: Math.round(H * 0.20), text: '章節標題', size: 5400, bold: true, align: 'l', anchor: 'b' },
        { x: m, y: Math.round(H * 0.55), cx: W - 2 * m, cy: Math.round(H * 0.10), text: '說明文字', size: 2400, align: 'l', anchor: 't' },
      ];
    case 'titleOnly':
      return [
        { x: m, y: m, cx: W - 2 * m, cy: Math.round(H * 0.13), text: '標題', size: 3600, bold: true, align: 'l', anchor: 'ctr' },
      ];
    case 'blank':
      return [];
  }
}

function buildLayoutShapeXml(spec: LayoutBoxSpec, id: number): string {
  const algn = spec.align ? ` algn="${spec.align}"` : '';
  const anchor = spec.anchor ? ` anchor="${spec.anchor}"` : '';
  const bold = spec.bold ? ' b="1"' : '';
  return (
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/>` +
    `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${spec.x}" y="${spec.y}"/><a:ext cx="${spec.cx}" cy="${spec.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"${anchor}/><a:lstStyle/>` +
    `<a:p><a:pPr${algn}/><a:r><a:rPr lang="zh-TW" sz="${spec.size}"${bold}/>` +
    `<a:t>${encodeXmlEntities(spec.text)}</a:t></a:r></a:p>` +
    `</p:txBody>` +
    `</p:sp>`
  );
}

/**
 * Replace the slide's `<p:spTree>` contents with placeholders matching the
 * named layout. The existing text/shapes on the slide are dropped — callers
 * should confirm with the user first. Returns new pptx bytes.
 */
export async function applyLayoutToSlide(
  originalBytes: Uint8Array,
  slideIndex: number,
  layoutId: PptxLayoutId,
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(originalBytes);
  const order = await resolvePresentationOrder(zip);
  if (slideIndex < 0 || slideIndex >= order.length) {
    throw new Error(`slide_out_of_range: ${slideIndex}`);
  }
  const path = order[slideIndex].path;
  const xml = await mustReadString(zip, path);
  const slideSize = await readSlideSize(zip);
  const specs = buildLayoutBoxes(layoutId, slideSize);
  // Keep the standard <p:nvGrpSpPr>/<p:grpSpPr> prefix so PowerPoint accepts
  // the slide; replace only the shape children with our placeholders.
  const newSpTree =
    `<p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr/>` +
    specs.map((s, i) => buildLayoutShapeXml(s, i + 2)).join('') +
    `</p:spTree>`;
  const updated = xml.replace(/<p:spTree>[\s\S]*?<\/p:spTree>/, newSpTree);
  zip.file(path, updated);
  return await zip.generateAsync({ type: 'uint8array' });
}

// ── helpers ──────────────────────────────────────────────────────────────

interface SlideRef {
  /** Position in <p:sldIdLst> (0-based). */
  index: number;
  /** Relationship id, e.g. "rId2". */
  rId: string;
  /** Resolved zip path, e.g. "ppt/slides/slide1.xml". */
  path: string;
  /** sldId attribute (numeric, ≥256). */
  sldId: string;
}

/**
 * Get slides in *presentation* order — the truth is `<p:sldIdLst>` in
 * `presentation.xml`, not the filename suffix. Falls back to filename sort
 * if presentation.xml is missing/malformed.
 */
async function resolvePresentationOrder(zip: JSZip): Promise<SlideRef[]> {
  const presXml = await readStringOrNull(zip, 'ppt/presentation.xml');
  const relsXml = await readStringOrNull(zip, 'ppt/_rels/presentation.xml.rels');
  if (!presXml || !relsXml) {
    return collectSlidePathsFallback(zip);
  }

  // Parse rels → { rId: target }.
  const relMap = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/g;
  for (let m = relRe.exec(relsXml); m !== null; m = relRe.exec(relsXml)) {
    relMap.set(m[1], m[2]);
  }

  // Pull <p:sldId id="X" r:id="rIdY"/> (or rs:id depending on namespace alias).
  const sldIdRe = /<p:sldId\b[^>]*?\bid="(\d+)"[^>]*?\b(?:r:id|rs:id)="([^"]+)"/g;
  const refs: SlideRef[] = [];
  let i = 0;
  for (let m = sldIdRe.exec(presXml); m !== null; m = sldIdRe.exec(presXml)) {
    const sldId = m[1];
    const rId = m[2];
    const target = relMap.get(rId);
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `ppt/${target}`;
    refs.push({ index: i, rId, path, sldId });
    i += 1;
  }
  if (refs.length === 0) return collectSlidePathsFallback(zip);
  return refs;
}

/** Just the paths (kept for parsePptx where rId/sldId aren't needed). */
async function resolveSlideOrder(zip: JSZip): Promise<string[]> {
  const refs = await resolvePresentationOrder(zip);
  return refs.map((r) => r.path);
}

function collectSlidePathsFallback(zip: JSZip): SlideRef[] {
  const paths: string[] = [];
  zip.forEach((relPath) => {
    if (/^ppt\/slides\/slide\d+\.xml$/.test(relPath)) paths.push(relPath);
  });
  paths.sort((a, b) => {
    const na = Number(a.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
    const nb = Number(b.match(/slide(\d+)\.xml$/)?.[1] ?? 0);
    return na - nb;
  });
  return paths.map((p, i) => ({ index: i, rId: `rId${i + 1}`, path: p, sldId: String(256 + i) }));
}

async function mustReadString(zip: JSZip, path: string): Promise<string> {
  const f = zip.file(path);
  if (!f) throw new Error(`pptx_file_missing: ${path}`);
  return await f.async('string');
}

async function readStringOrNull(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path);
  if (!f) return null;
  return await f.async('string');
}

function basename(p: string): string {
  const i = p.lastIndexOf('/');
  return i < 0 ? p : p.slice(i + 1);
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function mutateContentTypes(zip: JSZip, fn: (xml: string) => string): Promise<void> {
  const path = '[Content_Types].xml';
  const xml = await mustReadString(zip, path);
  zip.file(path, fn(xml));
}

async function mutatePresentation(zip: JSZip, fn: (xml: string) => string): Promise<void> {
  const path = 'ppt/presentation.xml';
  const xml = await mustReadString(zip, path);
  zip.file(path, fn(xml));
}

async function mutatePresentationRels(zip: JSZip, fn: (xml: string) => string): Promise<void> {
  const path = 'ppt/_rels/presentation.xml.rels';
  const xml = await mustReadString(zip, path);
  zip.file(path, fn(xml));
}

/** Add a Relationship to presentation.xml.rels for a slide; returns the new rId. */
async function addPresentationRel(zip: JSZip, target: string): Promise<string> {
  const path = 'ppt/_rels/presentation.xml.rels';
  const xml = await mustReadString(zip, path);
  // Find max rId.
  const existing = new Set<number>();
  const re = /\bId="rId(\d+)"/g;
  for (let m = re.exec(xml); m !== null; m = re.exec(xml)) existing.add(Number(m[1]));
  let n = 1;
  while (existing.has(n)) n += 1;
  const rId = `rId${n}`;
  const insert = `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${target}"/>`;
  const updated = xml.replace(/<\/Relationships>\s*$/, insert + '</Relationships>');
  zip.file(path, updated);
  return rId;
}

/** Insert <p:sldId id="X" r:id="newRId"/> right after the entry referencing `afterRId`. */
function insertSldIdAfter(xml: string, afterRId: string, newRId: string): string {
  // Find the existing sldId element to anchor on.
  const anchorRe = new RegExp(`<p:sldId\\b[^/>]*?(?:r:id|rs:id)="${afterRId}"[^/>]*?/>`);
  const anchor = xml.match(anchorRe);
  if (!anchor) {
    // Fall back: append at end of <p:sldIdLst>.
    return xml.replace(/<\/p:sldIdLst>/, `${buildSldId(xml, newRId)}</p:sldIdLst>`);
  }
  const newEl = buildSldId(xml, newRId);
  return xml.replace(anchorRe, anchor[0] + newEl);
}

function buildSldId(xml: string, rId: string): string {
  // Find max id and pick max+1, but at least 256.
  const idRe = /<p:sldId\b[^>]*?\bid="(\d+)"/g;
  let max = 255;
  for (let m = idRe.exec(xml); m !== null; m = idRe.exec(xml)) {
    const v = Number(m[1]);
    if (v > max) max = v;
  }
  return `<p:sldId id="${max + 1}" r:id="${rId}"/>`;
}

function removeSldId(xml: string, rId: string): string {
  const re = new RegExp(`<p:sldId\\b[^/>]*?(?:r:id|rs:id)="${rId}"[^/>]*?/>`, 'g');
  return xml.replace(re, '');
}

function moveSldId(xml: string, from: number, to: number): string {
  const lstStart = xml.indexOf('<p:sldIdLst');
  const lstEnd = xml.indexOf('</p:sldIdLst>');
  if (lstStart < 0 || lstEnd < 0) return xml;
  const open = xml.indexOf('>', lstStart) + 1;
  const inner = xml.slice(open, lstEnd);
  const items = [...inner.matchAll(/<p:sldId\b[^/>]*?\/>/g)].map((m) => m[0]);
  if (from < 0 || from >= items.length || to < 0 || to >= items.length) return xml;
  const [it] = items.splice(from, 1);
  items.splice(to, 0, it);
  return xml.slice(0, open) + items.join('') + xml.slice(lstEnd);
}

const RUN_BLOCK_RE = /<a:r>([\s\S]*?)<\/a:r>/g;
const RPR_RE = /<a:rPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:rPr>)/;
const TEXT_RE = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/;
const SOLID_FILL_RE = /<a:solidFill>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6,8})"/;
const LATIN_FONT_RE = /<a:latin\b[^>]*?\btypeface="([^"]+)"[^>]*?\/>/;

/**
 * Extract editable runs from a slide XML. We match `<a:r>...</a:r>` blocks
 * globally (preserving the document-order runIndex contract used by
 * applyRuns / deleteTextBoxFromSlide), then look up each run's enclosing
 * `<p:sp>` to attach a position. Runs that don't sit inside a `<p:sp>`
 * (e.g. table cells inside `<p:graphicFrame>`) get a fallback frame and
 * shapeIndex=-1 so the canvas can still surface them.
 */
function extractRuns(
  xml: string,
  slideIndex: number,
  slideSize: { cx: number; cy: number },
): PptxTextRun[] {
  // Step 1: collect <p:sp> ranges + per-shape frame.
  const shapes: { start: number; end: number; frame: PptxFrame }[] = [];
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  for (let m = spRe.exec(xml); m !== null; m = spRe.exec(xml)) {
    const block = m[0];
    const frame = parseXfrm(block) ?? defaultFrame(shapes.length, slideSize);
    shapes.push({ start: m.index, end: m.index + block.length, frame });
  }
  // Step 2: walk every <a:r> and assign to its enclosing <p:sp> (if any).
  const runs: PptxTextRun[] = [];
  let i = 0;
  RUN_BLOCK_RE.lastIndex = 0;
  for (let m = RUN_BLOCK_RE.exec(xml); m !== null; m = RUN_BLOCK_RE.exec(xml)) {
    const inner = m[1];
    const tMatch = inner.match(TEXT_RE);
    if (!tMatch) continue;
    const text = decodeXmlEntities(tMatch[1]);
    const style = parseRPr(inner);
    const idx = shapes.findIndex((s) => s.start <= m.index && m.index < s.end);
    const frame = idx >= 0 ? shapes[idx].frame : defaultFrame(shapes.length + i, slideSize);
    runs.push({
      id: nextRunId(),
      slideIndex,
      runIndex: i,
      shapeIndex: idx,
      frame,
      text,
      style,
    });
    i += 1;
  }
  return runs;
}

/** Pull `<a:xfrm><a:off/><a:ext/></a:xfrm>` out of a shape block. The off/ext
 *  attribute order isn't guaranteed in OOXML, so we extract them as
 *  self-closing tags and pluck attributes individually. */
function parseXfrm(spBlock: string): PptxFrame | null {
  const xfrm = spBlock.match(/<a:xfrm\b[^>]*>([\s\S]*?)<\/a:xfrm>/);
  if (!xfrm) return null;
  const off = xfrm[1].match(/<a:off\b[^>]*?\/>/);
  const ext = xfrm[1].match(/<a:ext\b[^>]*?\/>/);
  if (!off || !ext) return null;
  const x = off[0].match(/\bx="(-?\d+)"/);
  const y = off[0].match(/\by="(-?\d+)"/);
  const cx = ext[0].match(/\bcx="(\d+)"/);
  const cy = ext[0].match(/\bcy="(\d+)"/);
  if (!x || !y || !cx || !cy) return null;
  return { x: Number(x[1]), y: Number(y[1]), cx: Number(cx[1]), cy: Number(cy[1]) };
}

/** Stack-fallback frame for shapes that inherit position from a layout (no
 *  explicit `<a:xfrm>`), or for runs that sit outside any `<p:sp>`. */
function defaultFrame(stackIndex: number, slideSize: { cx: number; cy: number }): PptxFrame {
  const margin = 457200; // 0.5 inch
  const h = 914400; // 1 inch tall placeholder
  return {
    x: margin,
    y: margin + stackIndex * (h + margin / 2),
    cx: slideSize.cx - 2 * margin,
    cy: h,
  };
}

function parseRPr(innerR: string): PptxRunStyle | undefined {
  const rprMatch = innerR.match(RPR_RE);
  if (!rprMatch) return undefined;
  const attrs = rprMatch[1] ?? '';
  const inner = rprMatch[3] ?? '';
  const out: PptxRunStyle = {};
  if (/\bb="1"/.test(attrs)) out.bold = true;
  if (/\bi="1"/.test(attrs)) out.italic = true;
  // Any `u` value other than "none" counts as underlined; downgrades to
  // boolean per the schema comment on PptxRunStyle.underline.
  const u = attrs.match(/\bu="([^"]+)"/);
  if (u && u[1] !== 'none') out.underline = true;
  const sz = attrs.match(/\bsz="(\d+)"/);
  if (sz) out.size = Number(sz[1]) / 100; // OOXML uses hundredths-of-pt
  const fill = inner.match(SOLID_FILL_RE);
  if (fill) {
    const hex = fill[1].length === 8 ? fill[1].slice(2) : fill[1];
    out.color = hex.toUpperCase();
  }
  const latin = inner.match(LATIN_FONT_RE);
  if (latin) out.fontFamily = decodeXmlEntities(latin[1]);
  return Object.keys(out).length === 0 ? undefined : out;
}

/**
 * Walk `<a:r>` blocks in document order. For each, replace its `<a:t>` body
 * with the model text, and merge the model style into its `<a:rPr>`.
 */
function applyRuns(xml: string, runs: PptxTextRun[]): string {
  let out = '';
  let cursor = 0;
  let i = 0;
  RUN_BLOCK_RE.lastIndex = 0;
  for (let m = RUN_BLOCK_RE.exec(xml); m !== null; m = RUN_BLOCK_RE.exec(xml)) {
    const fullBlock = m[0];
    const modelRun = runs[i];
    out += xml.slice(cursor, m.index);
    if (modelRun !== undefined) {
      out += rewriteRunBlock(fullBlock, modelRun);
    } else {
      out += fullBlock;
    }
    cursor = m.index + fullBlock.length;
    i += 1;
  }
  out += xml.slice(cursor);
  return out;
}

function rewriteRunBlock(block: string, run: PptxTextRun): string {
  // Replace inner `<a:t>` first (always present per extractRuns).
  let next = block.replace(TEXT_RE, () => `<a:t>${encodeXmlEntities(run.text)}</a:t>`);

  // Then merge style into `<a:rPr>`.
  const rprMatch = next.match(RPR_RE);
  if (rprMatch) {
    const newRpr = mergeRPr(rprMatch[0], run.style);
    next = next.replace(RPR_RE, () => newRpr);
  } else if (run.style && Object.keys(run.style).length > 0) {
    // No existing rPr — inject one right after `<a:r>`.
    next = next.replace(/^<a:r>/, `<a:r>${buildRPr(run.style)}`);
  }
  return next;
}

function mergeRPr(existing: string, style: PptxRunStyle | undefined): string {
  // Split into attrs + inner.
  const m = existing.match(/^<a:rPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:rPr>)/);
  if (!m) return existing;
  let attrs = m[1] ?? '';
  let inner = m[3] ?? '';

  // Toggle b / i / sz on the attribute string.
  attrs = setOrRemoveAttr(attrs, 'b', style?.bold ? '1' : null);
  attrs = setOrRemoveAttr(attrs, 'i', style?.italic ? '1' : null);
  attrs = setOrRemoveAttr(attrs, 'u', style?.underline ? 'sng' : null);
  attrs = setOrRemoveAttr(attrs, 'sz', style?.size ? String(Math.round(style.size * 100)) : null);

  // Update or remove `<a:solidFill>` for the colour.
  // We always strip any existing solidFill we set previously and re-emit if needed.
  inner = inner.replace(/<a:solidFill>\s*<a:srgbClr\s+val="[0-9A-Fa-f]{6,8}"\s*\/>\s*<\/a:solidFill>/g, '');
  if (style?.color) {
    inner = `<a:solidFill><a:srgbClr val="${style.color.toUpperCase()}"/></a:solidFill>` + inner;
  }

  // Update or remove `<a:latin>` typeface. Strip any existing one then re-emit.
  inner = inner.replace(/<a:latin\b[^/]*?\/>/g, '');
  if (style?.fontFamily) {
    inner += `<a:latin typeface="${encodeXmlEntities(style.fontFamily)}"/>`;
  }

  if (inner.trim() === '') {
    return `<a:rPr${attrs}/>`;
  }
  return `<a:rPr${attrs}>${inner}</a:rPr>`;
}

function buildRPr(style: PptxRunStyle): string {
  let attrs = '';
  if (style.bold) attrs += ' b="1"';
  if (style.italic) attrs += ' i="1"';
  if (style.underline) attrs += ' u="sng"';
  if (style.size) attrs += ` sz="${Math.round(style.size * 100)}"`;
  let inner = '';
  if (style.color) {
    inner += `<a:solidFill><a:srgbClr val="${style.color.toUpperCase()}"/></a:solidFill>`;
  }
  if (style.fontFamily) {
    inner += `<a:latin typeface="${encodeXmlEntities(style.fontFamily)}"/>`;
  }
  if (inner) {
    return `<a:rPr${attrs}>${inner}</a:rPr>`;
  }
  return `<a:rPr${attrs}/>`;
}

/** Add or remove an attribute on the rPr's attribute string ("\sb=\"1\"" etc). */
function setOrRemoveAttr(attrs: string, name: string, value: string | null): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  const stripped = attrs.replace(re, '');
  if (value === null) return stripped;
  return stripped + ` ${name}="${value}"`;
}

// ── speaker notes ──────────────────────────────────────────────────────
// pptx notes live in `ppt/notesSlides/notesSlideN.xml`, linked from each
// slide's `_rels` file. We read all <a:r> text in the body placeholder and
// expose it as a single editable string. Writes patch existing notes in
// place; if a slide has no notesSlide yet, the first non-empty save
// synthesizes the file + relationship + content type.

const NOTES_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';

async function readSlideNotes(zip: JSZip, slidePath: string): Promise<{ text: string; path: string | null }> {
  const relsPath = `ppt/slides/_rels/${basename(slidePath)}.rels`;
  const relsXml = await readStringOrNull(zip, relsPath);
  if (!relsXml) return { text: '', path: null };
  const m = relsXml.match(
    new RegExp(`<Relationship\\b[^>]*?Type="${escapeRegex(NOTES_REL_TYPE)}"[^>]*?Target="([^"]+)"`),
  );
  if (!m) return { text: '', path: null };
  const target = m[1];
  // Resolve relative path. Notes target is typically "../notesSlides/notesSlideN.xml".
  const resolved = resolveRelTarget('ppt/slides/', target);
  const notesXml = await readStringOrNull(zip, resolved);
  if (!notesXml) return { text: '', path: resolved };
  const text = extractNotesBodyText(notesXml);
  return { text, path: resolved };
}

async function writeSlideNotes(zip: JSZip, slide: PptxSlide): Promise<void> {
  const text = slide.notesText ?? '';
  if (slide.notesPath) {
    const existing = await readStringOrNull(zip, slide.notesPath);
    if (existing) {
      zip.file(slide.notesPath, replaceNotesBodyText(existing, text));
      return;
    }
  }
  if (text.trim() === '') return; // nothing to do
  // No existing notesSlide — synthesize one.
  await createNotesSlideForSlide(zip, slide.zipPath, text);
}

function extractNotesBodyText(notesXml: string): string {
  // Find the <p:sp> that's a body placeholder, then concatenate <a:t> text.
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  for (let m = spRe.exec(notesXml); m !== null; m = spRe.exec(notesXml)) {
    const block = m[0];
    if (!/<p:ph\b[^>]*?\btype="body"/.test(block)) continue;
    const parts: string[] = [];
    const tRe = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g;
    for (let tm = tRe.exec(block); tm !== null; tm = tRe.exec(block)) {
      parts.push(decodeXmlEntities(tm[1]));
    }
    return parts.join('\n');
  }
  return '';
}

function replaceNotesBodyText(notesXml: string, newText: string): string {
  const spRe = /<p:sp\b[\s\S]*?<\/p:sp>/g;
  return notesXml.replace(spRe, (block) => {
    if (!/<p:ph\b[^>]*?\btype="body"/.test(block)) return block;
    const newTxBody = buildNotesTxBody(newText);
    return block.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
  });
}

function buildNotesTxBody(text: string): string {
  const lines = text.split(/\r?\n/);
  const paragraphs = lines
    .map((line) =>
      line === ''
        ? '<a:p/>'
        : `<a:p><a:r><a:rPr lang="zh-TW"/><a:t>${encodeXmlEntities(line)}</a:t></a:r></a:p>`,
    )
    .join('');
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>`;
}

async function createNotesSlideForSlide(zip: JSZip, slidePath: string, text: string): Promise<void> {
  // Pick a free notesSlideN.xml path.
  const used = new Set<number>();
  zip.forEach((p) => {
    const m = p.match(/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/);
    if (m) used.add(Number(m[1]));
  });
  let n = 1;
  while (used.has(n)) n += 1;
  const notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
  const notesRelsPath = `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`;
  const notesXml =
    XML_DECL +
    '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
    '<p:cSld><p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '<p:sp>' +
    '<p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr/>' +
    buildNotesTxBody(text) +
    '</p:sp>' +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:notes>';
  zip.file(notesPath, notesXml);

  // Notes -> slide back-reference (optional but proper).
  const slideRelTarget = `../slides/${basename(slidePath)}`;
  const notesRelsXml =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${slideRelTarget}"/>` +
    '</Relationships>';
  zip.file(notesRelsPath, notesRelsXml);

  // Slide -> notes relationship.
  const slideRelsPath = `ppt/slides/_rels/${basename(slidePath)}.rels`;
  const slideRelsXml = (await readStringOrNull(zip, slideRelsPath)) ?? `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const existing = new Set<number>();
  const idRe = /\bId="rId(\d+)"/g;
  for (let m = idRe.exec(slideRelsXml); m !== null; m = idRe.exec(slideRelsXml)) existing.add(Number(m[1]));
  let rIdN = 1;
  while (existing.has(rIdN)) rIdN += 1;
  const insertRel = `<Relationship Id="rId${rIdN}" Type="${NOTES_REL_TYPE}" Target="../notesSlides/notesSlide${n}.xml"/>`;
  const updatedSlideRels = slideRelsXml.replace(/<\/Relationships>\s*$/, insertRel + '</Relationships>');
  zip.file(slideRelsPath, updatedSlideRels);

  // Content types Override.
  await mutateContentTypes(zip, (xml) => {
    if (xml.includes(`PartName="/${notesPath}"`)) return xml;
    const insert = `<Override PartName="/${notesPath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
    return xml.replace(/<\/Types>\s*$/, insert + '</Types>');
  });
}

/** Resolve a Target like "../notesSlides/x.xml" against a base zip dir. */
function resolveRelTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseParts = baseDir.split('/').filter(Boolean);
  const tParts = target.split('/');
  const out = [...baseParts];
  for (const p of tParts) {
    if (p === '..') out.pop();
    else if (p === '.' || p === '') continue;
    else out.push(p);
  }
  return out.join('/');
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function encodeXmlEntities(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ── blank pptx template ─────────────────────────────────────────────────
// Minimum parts PowerPoint / LibreOffice need to load + edit a deck. The
// single slide carries a centered title placeholder so the user lands on an
// editable text run rather than a blank canvas.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const BLANK_CONTENT_TYPES =
  XML_DECL +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
  '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
  '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
  '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
  '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
  '</Types>';

const BLANK_ROOT_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
  '</Relationships>';

const BLANK_PRESENTATION =
  XML_DECL +
  '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
  '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
  '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>' +
  '<p:notesSz cx="6858000" cy="9144000"/>' +
  '<p:defaultTextStyle/>' +
  '</p:presentation>';

const BLANK_PRESENTATION_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
  '</Relationships>';

const BLANK_SLIDE =
  XML_DECL +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>' +
  '<p:sp>' +
  '<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="685800" y="2438400"/><a:ext cx="7772400" cy="1143000"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>' +
  '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-TW" sz="4400"/><a:t>標題</a:t></a:r></a:p>' +
  '</p:txBody>' +
  '</p:sp>' +
  '<p:sp>' +
  '<p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="685800" y="3886200"/><a:ext cx="7772400" cy="685800"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>' +
  '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-TW" sz="2400"/><a:t>副標題</a:t></a:r></a:p>' +
  '</p:txBody>' +
  '</p:sp>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sld>';

const BLANK_SLIDE_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
  '</Relationships>';

const BLANK_SLIDE_LAYOUT =
  XML_DECL +
  '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">' +
  '<p:cSld name="Title Slide"><p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
  '</p:sldLayout>';

const BLANK_SLIDE_LAYOUT_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
  '</Relationships>';

const BLANK_SLIDE_MASTER =
  XML_DECL +
  '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld>' +
  '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
  '<p:spTree>' +
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>' +
  '</p:spTree></p:cSld>' +
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
  '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
  '<p:txStyles>' +
  '<p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>' +
  '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>' +
  '<p:otherStyle><a:defPPr><a:defRPr/></a:defPPr></p:otherStyle>' +
  '</p:txStyles>' +
  '</p:sldMaster>';

const BLANK_SLIDE_MASTER_RELS =
  XML_DECL +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
  '</Relationships>';

const BLANK_THEME =
  XML_DECL +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">' +
  '<a:themeElements>' +
  '<a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme>' +
  '<a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
  '</a:fontScheme>' +
  '<a:fmtScheme name="Office">' +
  '<a:fillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:fillStyleLst>' +
  '<a:lnStyleLst>' +
  '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' +
  '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' +
  '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>' +
  '</a:lnStyleLst>' +
  '<a:effectStyleLst>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '</a:bgFillStyleLst>' +
  '</a:fmtScheme>' +
  '</a:themeElements>' +
  '</a:theme>';
