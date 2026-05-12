/**
 * Bottom status bar — shows at-a-glance workspace + AI session metrics.
 *
 * Pulls from the workspace and AI stores; everything is read-only display.
 * Sits under the EditorSurface so the user always knows what's loaded and
 * what the model is doing without opening any panels.
 */

import { useEffect, useState } from 'react';
import { Activity, AlertCircle, Check, Download, FileText, Hash, Loader2, Sparkles } from 'lucide-react';
import { useAI } from '../store/ai';
import { useWorkspace } from '../store/workspace';
import { SUPPORTED_MODELS } from '../types/ai';

/**
 * Pretty-print "saved 3s ago" / "saved 2 minutes ago" relative time. We
 * only show seconds and minutes — a save older than an hour stops being
 * an interesting status for this bar.
 */
function formatRelative(ts: number, now: number): string {
  const delta = Math.max(0, Math.floor((now - ts) / 1000));
  if (delta < 5) return '剛剛儲存';
  if (delta < 60) return `${delta} 秒前儲存`;
  const min = Math.floor(delta / 60);
  if (min < 60) return `${min} 分鐘前儲存`;
  return '已儲存';
}

export function StatusBar(): JSX.Element {
  const tabs = useWorkspace((s) => s.tabs);
  const activeTabId = useWorkspace((s) => s.activeTabId);
  const saveState = useWorkspace((s) => s.saveState);
  const lastSavedAt = useWorkspace((s) => s.lastSavedAt);
  const saveError = useWorkspace((s) => s.saveError);
  const exportFlash = useWorkspace((s) => s.exportFlash);
  const usage = useAI((s) => s.usage);
  const model = useAI((s) => s.model);
  const streaming = useAI((s) => s.streaming.requestId !== null);

  // Re-render the "saved Xs ago" label at the granularity it actually shows.
  // The previous version ticked every 15 s while the label promises per-
  // second resolution: the first minute looked frozen on "剛剛儲存" until the
  // tick caught up at 15 s and skipped straight to "15 秒前儲存". We
  // schedule the next wake-up based on which formatRelative branch we're in:
  //   • delta < 5 s   → tick once at the 5 s boundary (label flips to numeric)
  //   • 5–60 s        → 1 s ticks (smooth per-second display)
  //   • 1 min – 1 hr  → 30 s ticks (minute boundary off by ≤30 s)
  //   • ≥ 1 hr        → stop; label settled on "已儲存", further ticks would
  //                     just re-render for no visible change
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    if (lastSavedAt === null) return;
    let h: number | null = null;
    const schedule = () => {
      const elapsed = Date.now() - lastSavedAt;
      let delay: number;
      if (elapsed < 5_000) delay = 5_000 - elapsed;
      else if (elapsed < 60_000) delay = 1_000;
      else if (elapsed < 3_600_000) delay = 30_000;
      else return;
      h = window.setTimeout(() => {
        setNow(Date.now());
        schedule();
      }, delay);
    };
    schedule();
    return () => {
      if (h !== null) window.clearTimeout(h);
    };
  }, [lastSavedAt]);

  const active = tabs.find((t) => t.id === activeTabId);
  // R338 — treat HTML tabs as text-based, like markdown. HtmlTab has
  // `content: string` (not `data: Uint8Array`); the previous branch only
  // tested `type === 'markdown'` and fell through to `'data' in active`,
  // which was false for HTML → charCount stayed 0 → label rendered
  // `0.0 KB` regardless of the editor's actual content. Same fix shape
  // as R337 for FileExplorer's openable-set: extend the "is this text-
  // shaped" predicate so html lands on the markdown branch instead of
  // the binary branch. The label / title strings are slightly
  // generalised — drop the "Markdown 語法" specificity since HTML's
  // syntax characters (tags, attributes) are also counted; "text 字元
  // 數" is accurate for both.
  const isTextTab = active?.type === 'markdown' || active?.type === 'html';
  const charCount = isTextTab
    ? (active as { content: string }).content.length
    : active && 'data' in active
      ? active.data.byteLength
      : 0;
  const charLabel = isTextTab
    ? `${charCount.toLocaleString()} 字元`
    : `${(charCount / 1024).toFixed(1)} KB`;
  // The visible label flips meaning by tab type — `字元` for text tabs
  // (Markdown / HTML, string.length including whitespace + markup syntax)
  // and `KB` for binary tabs (data.byteLength of the serialized OOXML zip
  // payload, kept in sync with edits via XlsxEditor.tsx:530's writeBack /
  // peer DocxEditor / PptxEditor). The number alone is genuinely ambiguous:
  // a user looking at `347.2 KB` next to a `.docx` can't tell if it's
  // on-disk size, in-memory size, or something else; `12,345 字元` doesn't
  // clarify whether whitespace + Markdown syntax are included. Every other
  // dynamic StatusBar slot in this file already carries a title (saveError
  // line 123, lastSavedAt line 132, exportFlash line 151, model line 160,
  // token-usage breakdown line 168) — this slot was the lone outlier. The
  // tooltip vocabulary mirrors the token-usage tooltip's "explain what each
  // bucket means" pattern.
  const charTitle = isTextTab
    ? '目前文件字元數（含空白與標記語法）'
    : '目前檔案內容大小（記憶體中的位元組數）';

  // Cache hit ratio = cacheRead / total-input-tokens. Anthropic splits each
  // request's input into three buckets: regular `inputTokens`,
  // `cacheCreationInputTokens` (written to cache, billed 1.25×), and
  // `cacheReadInputTokens` (served from cache, billed 0.1×). Cache-creation
  // tokens are explicitly *not* hits — we paid full+ price to write them —
  // so they belong in the denominator. Excluding them overstated the ratio,
  // sometimes wildly: input=100 + cacheCreate=200 + cacheRead=300 reads as
  // 75% with the old formula, 50% with the correct one.
  const totalInput =
    usage.inputTokens + usage.cacheCreationInputTokens + usage.cacheReadInputTokens;
  const cacheHitRatio =
    totalInput > 0 ? Math.round((usage.cacheReadInputTokens / totalInput) * 100) : null;

  const modelLabel = SUPPORTED_MODELS.find((m) => m.id === model)?.label ?? model;

  return (
    <div className="flex items-center h-6 px-3 border-t bg-secondary/30 text-[11px] text-muted-foreground gap-4">
      <span className="flex items-center gap-1">
        <FileText className="h-3 w-3" />
        {tabs.length} 個頁籤
      </span>
      {active && (
        <span className="flex items-center gap-1" title={charTitle}>
          <Hash className="h-3 w-3" />
          {charLabel}
        </span>
      )}
      {/* Save status — appears only when there's something to say. The
          toolbar Save button shows the same state as an icon, but the
          status bar adds the human-readable "X 秒前儲存" so the user can
          tell stale state apart from fresh save without waiting for the
          tooltip. */}
      {saveState === 'saving' && (
        <span className="flex items-center gap-1 text-primary">
          <Loader2 className="h-3 w-3 animate-spin" />
          儲存中…
        </span>
      )}
      {saveState === 'error' && (
        <span
          className="flex items-center gap-1 text-destructive"
          title={saveError ?? '未知錯誤'}
        >
          <AlertCircle className="h-3 w-3" />
          儲存失敗
        </span>
      )}
      {saveState !== 'saving' && saveState !== 'error' && lastSavedAt !== null && (
        <span
          className="flex items-center gap-1"
          title={new Date(lastSavedAt).toLocaleString()}
        >
          <Check className={`h-3 w-3 ${saveState === 'success' ? 'text-emerald-500' : ''}`} />
          {formatRelative(lastSavedAt, now)}
        </span>
      )}
      {/* Export flash — auto-clears 5s after the export resolves. The store's
          flashExport() arms a setTimeout so we don't need a local timer here;
          the slot just disappears when exportFlash flips back to null. */}
      {exportFlash !== null && (
        <span
          className="flex items-center gap-1 text-emerald-500"
          // Tooltip surfaces the *full* destination path so users can
          // confirm *where* the file landed without reopening the OS save
          // dialog. The visible text stays compact (basename only) so the
          // StatusBar slot doesn't shove the AI / token-usage indicators
          // off-screen on long paths. Falls back to fileName when no path
          // is recorded (legacy callers, or future flash sources without
          // a known disk location).
          title={`已匯出至 ${exportFlash.filePath ?? exportFlash.fileName}`}
        >
          <Download className="h-3 w-3" />
          已匯出 {exportFlash.fileName}
        </span>
      )}
      <div className="flex-1" />
      <span
        className={`flex items-center gap-1 ${streaming ? 'text-primary' : ''}`}
        title="目前選擇的 AI 模型"
      >
        <Sparkles className={`h-3 w-3 ${streaming ? 'animate-pulse' : ''}`} />
        {modelLabel}
      </span>
      {(totalInput > 0 || usage.outputTokens > 0) && (
        <span
          className="flex items-center gap-1"
          title={`Input: ${usage.inputTokens.toLocaleString()}\nOutput: ${usage.outputTokens.toLocaleString()}\nCache read: ${usage.cacheReadInputTokens.toLocaleString()}\nCache create: ${usage.cacheCreationInputTokens.toLocaleString()}`}
        >
          <Activity className="h-3 w-3" />
          {/* Headline total includes all four buckets — anything else makes
              cache-heavy sessions look dramatically smaller than they are. The
              tooltip already breaks it down by bucket for users who care. */}
          {(totalInput + usage.outputTokens).toLocaleString()} tokens
          {cacheHitRatio !== null && cacheHitRatio > 0 && (
            <span className="text-emerald-500 ml-1">{cacheHitRatio}% cache</span>
          )}
        </span>
      )}
    </div>
  );
}
