/**
 * Shared font-family list for the Word / Excel / PowerPoint toolbars.
 *
 * Picked to cover the default Office faces + the most common
 * Traditional Chinese / Simplified Chinese / Japanese system fonts. The
 * actual rendering depends on what the user has installed; missing fonts
 * silently fall back to the OS substitute.
 */
export interface FontFamilyOption {
  /** Display label in the dropdown. */
  label: string;
  /** Value written into the docx / xlsx / pptx file. */
  value: string;
}

/**
 * Cross-platform colour-emoji fonts to append to any font-family stack used
 * for editor surfaces. Without these, picking a specific face like "Times
 * New Roman" replaces the cascade entirely and emoji glyphs (😀 🇹🇼 👨‍👩‍👧)
 * render as `.notdef` boxes since none of the Office text fonts include
 * colour emoji.
 *
 * Use via `withEmojiFallback(fontName)` whenever building an inline
 * `fontFamily` style for a textarea / input / editing surface.
 */
export const EMOJI_FONT_FALLBACK = `'Apple Color Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol', 'Noto Color Emoji'`;

/** Quote `value` (single quotes, escaping any embedded ones) and append the
 * emoji fallback chain. Returns a CSS font-family string ready for
 * `style={{ fontFamily: ... }}`. */
export function withEmojiFallback(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  const quoted = `"${trimmed.replace(/"/g, '\\"')}"`;
  return `${quoted}, ${EMOJI_FONT_FALLBACK}`;
}

export const FONT_FAMILIES: FontFamilyOption[] = [
  // Office defaults
  { label: 'Calibri (預設)', value: 'Calibri' },
  { label: 'Calibri Light', value: 'Calibri Light' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Verdana', value: 'Verdana' },
  { label: 'Tahoma', value: 'Tahoma' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Courier New', value: 'Courier New' },
  { label: 'Consolas', value: 'Consolas' },
  // Traditional Chinese
  { label: '微軟正黑體 (Microsoft JhengHei)', value: 'Microsoft JhengHei' },
  { label: '新細明體 (PMingLiU)', value: 'PMingLiU' },
  { label: '細明體 (MingLiU)', value: 'MingLiU' },
  { label: '標楷體 (DFKai-SB)', value: 'DFKai-SB' },
  // Simplified Chinese
  { label: '微軟雅黑 (Microsoft YaHei)', value: 'Microsoft YaHei' },
  { label: '宋體 (SimSun)', value: 'SimSun' },
  { label: '黑體 (SimHei)', value: 'SimHei' },
  { label: '楷體 (KaiTi)', value: 'KaiTi' },
  // macOS / cross-platform CJK
  { label: '蘋方 (PingFang TC)', value: 'PingFang TC' },
  { label: '思源黑體 (Noto Sans TC)', value: 'Noto Sans TC' },
  { label: '思源宋體 (Noto Serif TC)', value: 'Noto Serif TC' },
  // Japanese
  { label: 'Yu Gothic', value: 'Yu Gothic' },
  { label: 'Meiryo', value: 'Meiryo' },
];
