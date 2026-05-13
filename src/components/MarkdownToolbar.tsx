/**
 * Formatting toolbar that drives a CodeMirror 6 EditorView via a ref.
 *
 * Each command either wraps the current selection (bold / italic / code) or
 * mutates the current line's prefix (heading / list / quote). Newly-inserted
 * snippets (table / link / horizontal rule) are placed at the cursor.
 *
 * The toolbar is intentionally schema-light — it works on raw markdown text;
 * round-trip is just bytes in / bytes out via the editor.
 *
 * Same commands are also bound to keyboard shortcuts in `MarkdownEditor` —
 * both code paths share `lib/markdown-commands.ts`.
 */

import {
  Bold,
  Italic,
  Heading1,
  Heading2,
  Heading3,
  List,
  ListOrdered,
  Quote,
  Code,
  Code2,
  Link as LinkIcon,
  Table as TableIcon,
  Minus,
  Image as ImageIcon,
  Columns2,
  FileText,
  Eye,
  ListTree,
  Search,
  Hash,
  FileDown,
} from 'lucide-react';
import type { EditorView } from '@codemirror/view';
import { gotoLine, openSearchPanel } from '@codemirror/search';
import {
  insertCodeBlock,
  insertHr,
  insertTable,
  setLinePrefix,
  wrapSelection,
} from '../lib/markdown-commands';
import { cn } from '../lib/utils';
import { useT } from '../lib/i18n';

export type MarkdownViewMode = 'source' | 'split' | 'preview';

interface Props {
  getView: () => EditorView | null;
  viewMode: MarkdownViewMode;
  onViewModeChange: (m: MarkdownViewMode) => void;
  /** Whether the outline (TOC) sidebar is currently visible. */
  outlineOpen: boolean;
  onToggleOutline: () => void;
  /** Open the floating link / image insertion dialog hosted by MarkdownEditor.
   *  Replaces the previous window.prompt-driven `insertLink`/`insertImage`. */
  onInsertLink: () => void;
  onInsertImage: () => void;
  /** Render the current preview pane to PDF. Disabled until the marked.parse
   *  debounce settles so the user can't kick off an export against an empty
   *  preview seed. */
  onExportPdf: () => void;
  /** True while a PDF export is in flight — avoid double-clicks spawning a
   *  second hidden BrowserWindow + save dialog. */
  exportPdfBusy: boolean;
  /** False before the first preview render lands; export would otherwise
   *  produce a blank PDF. */
  canExportPdf: boolean;
}

export function MarkdownToolbar({
  getView,
  viewMode,
  onViewModeChange,
  outlineOpen,
  onToggleOutline,
  onInsertLink,
  onInsertImage,
  onExportPdf,
  exportPdfBusy,
  canExportPdf,
}: Props): JSX.Element {
  const t = useT();
  const run = (fn: (v: EditorView) => void) => {
    const v = getView();
    if (!v) return;
    fn(v);
    v.focus();
  };

  // Source-only commands (formatting, insertion, search) need a visible source
  // pane to be meaningful. In preview-only mode the CodeMirror view is still
  // mounted (hidden via CSS), so a click would silently dispatch a transaction
  // against an off-screen selection — Bold inserts `****` somewhere the user
  // can't see, then they later switch to source and find phantom artefacts.
  // Disabling these buttons makes the constraint visible and forces an
  // intentional view-mode switch first.
  const sourceHidden = viewMode === 'preview';
  const disabledTitle = t('請先切換回原始碼或分割模式', 'Switch back to source or split mode first');

  return (
    <div className="flex items-center flex-wrap gap-0.5 px-2 py-1 border-b bg-secondary/30">
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('粗體 (Ctrl+B)', 'Bold (Ctrl+B)')}
        disabled={sourceHidden}
        onClick={() => run((v) => wrapSelection(v, '**', '**'))}
      >
        <Bold className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('斜體 (Ctrl+I)', 'Italic (Ctrl+I)')}
        disabled={sourceHidden}
        onClick={() => run((v) => wrapSelection(v, '_', '_'))}
      >
        <Italic className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('行內 code (Ctrl+`)', 'Inline code (Ctrl+`)')}
        disabled={sourceHidden}
        onClick={() => run((v) => wrapSelection(v, '`', '`'))}
      >
        <Code className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('標題 1 (Ctrl+1)', 'Heading 1 (Ctrl+1)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '# '))}
      >
        <Heading1 className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('標題 2 (Ctrl+2)', 'Heading 2 (Ctrl+2)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '## '))}
      >
        <Heading2 className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('標題 3 (Ctrl+3)', 'Heading 3 (Ctrl+3)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '### '))}
      >
        <Heading3 className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('無序列表 (Ctrl+Shift+L)', 'Bulleted list (Ctrl+Shift+L)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '- '))}
      >
        <List className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('有序列表 (Ctrl+Shift+O)', 'Numbered list (Ctrl+Shift+O)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '1. '))}
      >
        <ListOrdered className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* R141 — 引用 was the lone outlier in the line-prefix family: heading
          1/2/3 (Ctrl+1/2/3) and unordered/ordered list (Ctrl+Shift+L/O) all
          advertise their shortcut in tooltip, but blockquote sat as bare
          `'引用'` because no keymap entry existed. Mod-Shift-q now wired at
          MarkdownEditor.tsx (next to the Mod-Shift-l/o pair); tooltip here
          updated to match the Ctrl+Shift+L/Ctrl+Shift+O cousins exactly. */}
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('引用 (Ctrl+Shift+Q)', 'Quote (Ctrl+Shift+Q)')}
        disabled={sourceHidden}
        onClick={() => run((v) => setLinePrefix(v, '> '))}
      >
        <Quote className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('連結 (Ctrl+K)', 'Link (Ctrl+K)')}
        disabled={sourceHidden}
        onClick={onInsertLink}
      >
        <LinkIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* R96 — tooltip honesty. Previous wording 「圖片（外部 URL）」implied
          URL-only, but the dialog this button opens (MarkdownEditor.tsx
          LinkInsertDialog at lines 1116-1138, kind='image') renders a
          「從本機選擇檔案…」button next to the URL field and embeds picked
          files as base64 data URLs into the .md source — the placeholder
          at MarkdownEditor.tsx:1098 already spells both modes literally:
          'https://example.com 或選擇本機檔案'. The toolbar tooltip is
          the entry-point label that ushers users into that dialog, so a
          user trained by 「外部 URL」 with a local PNG wouldn't even reach
          for this button — they'd drag-drop / paste (which works) and
          never learn the base64-embed path exists.
          Sibling editors keep their image-button tooltip aligned with what
          their dialog accepts: DocxEditor.tsx:2157 「插入圖片(PNG / JPG /
          GIF / BMP)」 and PptxEditor.tsx:1083 「插入圖片(PNG / JPG / GIF /
          SVG / WebP)」 both list the formats they actually take. Markdown's
          unique twist is the URL-or-file duality, so the tooltip names
          both modes verbatim from the dialog's own placeholder vocabulary
          rather than spelling formats — keeping one source of truth for
          mode names across the entry-point and the dialog. */}
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('插入圖片（URL 或本機檔案）', 'Insert image (URL or local file)')}
        disabled={sourceHidden}
        onClick={onInsertImage}
      >
        <ImageIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('程式碼區塊', 'Code block')}
        disabled={sourceHidden}
        onClick={() => run(insertCodeBlock)}
      >
        <Code2 className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('表格', 'Table')}
        disabled={sourceHidden}
        onClick={() => run(insertTable)}
      >
        <TableIcon className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('分隔線', 'Horizontal rule')}
        disabled={sourceHidden}
        onClick={() => run(insertHr)}
      >
        <Minus className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <Divider />
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('尋找與取代 (Ctrl+F)', 'Find & Replace (Ctrl+F)')}
        disabled={sourceHidden}
        onClick={() => run((v) => openSearchPanel(v))}
      >
        <Search className="h-3.5 w-3.5" />
      </ToolbarBtn>
      {/* Ctrl+G — open CM6's native "Go to line" panel. The keystroke is wired
          at MarkdownEditor.tsx:450 (`{ key: 'Mod-g', run: gotoLine }`) and the
          doc-comment there at lines 440-449 explicitly frames it as parity
          with the other three editors' jump-to gestures:
            • DocxEditor.tsx:2164      跳至段落… (Ctrl+G) — toolbar button
            • PptxEditor.tsx:1024      跳至投影片… (Ctrl+G) — slide indicator
            • XlsxEditor.tsx:2299      Ctrl+G 可從鍵盤聚焦至此 — name box
          The keymap fix landed but the discoverability fix didn't — Markdown
          is the only editor whose toolbar gives no hint Ctrl+G is bound.
          Hover-only users muscle-trained on jump-to gestures from the other
          three formats reach for Ctrl+G in a .md tab, get the right behavior,
          but a fresh user reading the toolbar to learn shortcuts has no
          clue. CM6's gotoLine panel is its own visible UI once summoned —
          but you have to know to summon it. Same pattern as R64/R68/R71
          welcome-footer additions: shortcut wired across editors, sibling
          surface advertises it, this surface stays silent. Disabled in
          preview mode for the same reason as the other source-only buttons
          above (line 99-100): the CM6 view is hidden, so summoning its
          panel into an off-screen surface is a confusing no-op. Hash icon
          mirrors the line-number connotation (CM6 prompts "Go to line:");
          siblings used Pilcrow / DropdownIndicator-style for paragraph and
          slide respectively, so each format gets its own visual cue while
          sharing the verb "跳至…". */}
      <ToolbarBtn
        title={sourceHidden ? disabledTitle : t('跳至行… (Ctrl+G)', 'Go to line… (Ctrl+G)')}
        disabled={sourceHidden}
        onClick={() => run(gotoLine)}
      >
        <Hash className="h-3.5 w-3.5" />
      </ToolbarBtn>
      <div className="ml-auto flex items-center gap-0.5">
        {/* PDF export drives the rendered preview, not the source — keep it
            enabled in preview-only mode (the one place Word-trained users
            instinctively reach for "save as PDF"). Disabled while a previous
            export is still running, or before the first marked.parse debounce
            has produced any HTML to print. */}
        <ToolbarBtn
          title={
            !canExportPdf
              ? t('預覽尚未產生，請稍候再輸出', 'Preview not ready — please wait before exporting')
              : exportPdfBusy
                ? t('正在輸出 PDF…', 'Exporting PDF…')
                : t('輸出為 PDF', 'Export as PDF')
          }
          disabled={!canExportPdf || exportPdfBusy}
          onClick={onExportPdf}
        >
          <FileDown className="h-3.5 w-3.5" />
        </ToolbarBtn>
        <Divider />
        <ToolbarBtn
          title={outlineOpen ? t('隱藏大綱', 'Hide outline') : t('顯示大綱（標題列表）', 'Show outline (heading list)')}
          active={outlineOpen}
          onClick={onToggleOutline}
        >
          <ListTree className="h-3.5 w-3.5" />
        </ToolbarBtn>
        <Divider />
        <ToolbarBtn
          title={t('只看原始碼', 'Source only')}
          active={viewMode === 'source'}
          onClick={() => onViewModeChange('source')}
        >
          <FileText className="h-3.5 w-3.5" />
        </ToolbarBtn>
        <ToolbarBtn
          title={t('分割：原始碼 + 預覽', 'Split: source + preview')}
          active={viewMode === 'split'}
          onClick={() => onViewModeChange('split')}
        >
          <Columns2 className="h-3.5 w-3.5" />
        </ToolbarBtn>
        <ToolbarBtn
          title={t('只看預覽', 'Preview only')}
          active={viewMode === 'preview'}
          onClick={() => onViewModeChange('preview')}
        >
          <Eye className="h-3.5 w-3.5" />
        </ToolbarBtn>
      </div>
    </div>
  );
}

function ToolbarBtn({
  children,
  title,
  active,
  disabled,
  onClick,
}: {
  children: React.ReactNode;
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      // R153 — surface toggle state to assistive tech via `aria-pressed`. The
      // visual `bg-primary/20 text-primary` highlight at line 350-352 is the
      // sighted user's signal that 大綱 is open / 預覽模式 is on / etc., but
      // SR users had no equivalent — buttons were announced as plain action
      // buttons regardless of whether they were currently 「按下」 or not.
      // React renders `aria-pressed={undefined}` as no attribute, so action
      // buttons (those that don't pass `active`, like 連結 / 程式碼區塊 / 表
      // 格 — see callsites at lines 178-228) keep clean semantics; only the
      // five toggle callsites in this file (大綱 line 291, source/split/
      // preview view-mode trio at 299/306/313) actually pick up the toggle
      // role. Sibling ToolbarBtn definitions in DocxEditor/PptxEditor/Xlsx
      // Editor get the same treatment in this round so the 30+ format-toggle
      // callsites across the four editors all read consistently to SR.
      aria-pressed={active}
      // Suppress focus-shift to the button so the editor caret/selection stays
      // visibly anchored — `run` already calls `v.focus()` at the end, but the
      // brief blur causes the focus ring + selection highlight to flicker.
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={cn(
        'h-7 w-7 inline-flex items-center justify-center rounded transition-colors',
        disabled
          ? 'text-muted-foreground/40 cursor-not-allowed'
          : active
            ? 'bg-primary/20 text-primary'
            : 'text-muted-foreground hover:text-foreground hover:bg-secondary',
      )}
    >
      {children}
    </button>
  );
}

function Divider() {
  return <div className="w-px h-4 bg-border mx-1" />;
}
