/**
 * Runtime tab state. Distinct from {@link TabDescriptor} (which lives on disk
 * inside `manifest.json`); a Tab is the editor instance loaded into memory.
 */

import type { TabDescriptor, TabType } from './manifest';

/** Generic content envelope for an open tab. */
export interface BaseTab {
  id: string;
  name: string;
  type: TabType;
  /** Source path inside the .gd archive (or empty for unsaved). */
  file: string;
  order: number;
  /** True if there are unsaved edits relative to disk. */
  dirty: boolean;
  /**
   * Absolute disk path when this tab was imported from outside the .gd
   * archive (file explorer click, drag-drop). Lets the explorer activate
   * an already-open tab instead of opening a duplicate when the user
   * clicks the same file twice. Undefined for tabs created blank, loaded
   * from inside the .gd, or restored via reopenClosedTab.
   */
  sourcePath?: string;
}

export interface MarkdownTab extends BaseTab {
  type: 'markdown';
  content: string;
}

/**
 * Plain-HTML tab. Same content-shape as markdown (string in memory, UTF-8 bytes
 * on disk inside the .gd archive) — the editor renders source + sandboxed
 * preview rather than markdown's rendered HTML, but storage-wise they're
 * twins. Kept in its own variant so the editor surface can route to HtmlEditor
 * (CodeMirror with HTML language) instead of MarkdownEditor (markdown
 * language).
 */
export interface HtmlTab extends BaseTab {
  type: 'html';
  content: string;
}

export interface DocxTab extends BaseTab {
  type: 'docx';
  /** Raw bytes; the renderer hands these to Univer Doc. */
  data: Uint8Array;
}

export interface XlsxTab extends BaseTab {
  type: 'xlsx';
  data: Uint8Array;
}

export interface PptxTab extends BaseTab {
  type: 'pptx';
  data: Uint8Array;
}

export type Tab = MarkdownTab | HtmlTab | DocxTab | XlsxTab | PptxTab;

export function tabFromDescriptor(desc: TabDescriptor, payload: Uint8Array): Tab {
  const base: BaseTab = {
    id: desc.id,
    name: desc.name,
    type: desc.type,
    file: desc.file,
    order: desc.order,
    dirty: false,
  };
  switch (desc.type) {
    case 'markdown':
      return { ...base, type: 'markdown', content: new TextDecoder().decode(payload) };
    case 'html':
      return { ...base, type: 'html', content: new TextDecoder().decode(payload) };
    case 'docx':
      return { ...base, type: 'docx', data: payload };
    case 'xlsx':
      return { ...base, type: 'xlsx', data: payload };
    case 'pptx':
      return { ...base, type: 'pptx', data: payload };
    default: {
      // Exhaustiveness check
      const _exhaustive: never = desc.type;
      throw new Error(`Unknown tab type: ${_exhaustive as string}`);
    }
  }
}
