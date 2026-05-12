/**
 * Diff preview overlay for a pending AI ChangeSet.
 *
 * v1.0 supports markdown side-by-side line diff. Other op types (word /
 * excel / pptx) only show a textual summary because their write tools are
 * stubbed.
 */

import { useMemo } from 'react';
import { useWorkspace } from '../store/workspace';
import type { PendingChange } from '../store/ai';
import type { ChangeOp } from '../types/changeset';
import type { Tab } from '../types/tab';
import { Button } from './ui/button';
import { Check, X, Pencil, Wand2 } from 'lucide-react';
import { cn } from '../lib/utils';

interface Props {
  pending: PendingChange;
  onApply: () => void;
  onReject: () => void;
  onModify: () => void;
  /** When true, Apply / Reject are visually disabled — used while a turn is
   * streaming, since acting now would orphan the inflight handle. 修改 stays
   * enabled because it only mutates local draft state. */
  actionsDisabled?: boolean;
  /** When true, the keyboard shortcuts (Ctrl+Enter / Ctrl+M / Ctrl+Backspace)
   * route to THIS pending. The AIPanel's keyboard handler always targets the
   * last item in `ai.pending`, so only one card is the live shortcut target
   * at any time. Tooltips advertise the hotkey only when this is true —
   * otherwise users would mash Ctrl+Enter on the top card and watch a
   * different card flash. */
  keyboardActive?: boolean;
}

export function DiffPreview({ pending, onApply, onReject, onModify, actionsDisabled, keyboardActive }: Props): JSX.Element {
  const tabs = useWorkspace((s) => s.tabs);
  const opSummaries = useMemo(
    () => pending.changeset.ops.map((op) => describeOp(op, tabs)),
    [pending, tabs],
  );

  const mdOp = pending.changeset.ops.find((o): o is Extract<ChangeOp, { type: 'md_text' }> => o.type === 'md_text');

  return (
    <div className="border-2 border-amber-500/40 rounded-lg overflow-hidden bg-amber-500/[0.04] shadow-sm">
      {/* Header — clearly framed as a pending decision */}
      <div className="flex items-start justify-between gap-2 px-3 py-2 border-b border-amber-500/20 bg-amber-500/[0.06]">
        <div className="flex items-start gap-2 min-w-0 flex-1">
          <div className="mt-0.5 p-1 rounded-md bg-amber-500/15 text-amber-500 shrink-0">
            <Wand2 className="h-3 w-3" />
          </div>
          <div className="min-w-0 flex-1">
            <div className="text-[10px] uppercase tracking-wider text-amber-500/80 font-medium">
              待你決定
            </div>
            <div className="text-xs font-medium truncate" title={pending.summary}>
              {pending.summary}
            </div>
          </div>
        </div>
      </div>

      {/* Op breakdown */}
      <ul className="px-3 py-2 text-xs text-muted-foreground space-y-0.5">
        {opSummaries.map((s, i) => (
          <li key={i} className="flex gap-1.5">
            <span className="opacity-50">→</span>
            <span className="flex-1">{s}</span>
          </li>
        ))}
      </ul>

      {mdOp && <MarkdownDiff before={mdOp.before} after={mdOp.after} />}

      {/* Action bar — strong CTA hierarchy: Apply > Modify > Reject.
          Visible labels are Chinese to match siblings (修改 / 拒絕) and the
          tooltips ("套用變更 …"). Previously this primary button read
          "Apply" in English while siblings were Chinese, and the tooltip
          even said 套用 — the user saw two languages for the same verb in
          the same row. Comments above stay in English for developer
          legibility (the CTA-hierarchy note is dev-facing). */}
      <div className="flex items-center gap-1 px-3 py-2 border-t border-amber-500/20 bg-secondary/20">
        <Button
          size="sm"
          onClick={onApply}
          disabled={actionsDisabled}
          className="flex-1"
          title={actionsDisabled ? '等待目前回合結束後再套用' : keyboardActive ? '套用變更 (Ctrl+Enter)' : '套用變更'}
        >
          <Check className="h-3.5 w-3.5" />
          套用
        </Button>
        <Button
          size="sm"
          variant="outline"
          onClick={onModify}
          title={keyboardActive ? '修改建議 (Ctrl+M)' : '修改建議'}
        >
          <Pencil className="h-3 w-3" />
          修改
        </Button>
        <Button
          size="sm"
          variant="ghost"
          onClick={onReject}
          disabled={actionsDisabled}
          title={actionsDisabled ? '等待目前回合結束後再拒絕' : keyboardActive ? '拒絕變更 (Ctrl+Backspace)' : '拒絕變更'}
          className="text-muted-foreground hover:text-destructive"
        >
          <X className="h-3 w-3" />
          拒絕
        </Button>
      </div>
    </div>
  );
}

function describeOp(op: ChangeOp, tabs: Tab[]): string {
  const tab = tabs.find((t) => t.id === op.tabId);
  const tabName = tab?.name ?? `(tab ${op.tabId})`;
  switch (op.type) {
    case 'md_text': {
      const delta = op.after.length - op.before.length;
      return `${tabName}：${delta >= 0 ? '+' : ''}${delta} 字元`;
    }
    case 'binary_replace': {
      const delta = op.after.byteLength - op.before.byteLength;
      return `${tabName}：${op.description}（${delta >= 0 ? '+' : ''}${delta} bytes）`;
    }
    case 'word_paragraph':
      return `${tabName}：替換第 ${op.paraIndex} 段（unused legacy op）`;
    case 'excel_cell':
      return `${tabName}：${op.sheet}!${op.address} → ${String(op.after)}（unused legacy op）`;
    case 'pptx_text':
      return `${tabName}：slide ${op.slideIndex} shape ${op.shapeId}（unused legacy op）`;
    case 'tab_create':
      return `新增頁籤：${op.tab.name}`;
    case 'tab_delete':
      return `刪除頁籤：${op.tab.name}`;
    default: {
      const _exhaustive: never = op;
      return `(unknown op ${(_exhaustive as ChangeOp).type})`;
    }
  }
}

function MarkdownDiff({ before, after }: { before: string; after: string }) {
  // Parent DiffPreview subscribes to `tabs`, so every keystroke in any open
  // editor re-renders this component. Without memoization the O(n*m) LCS and
  // line splits run on each keystroke for `before`/`after` strings that
  // never change once the pending diff is constructed — pure waste, and it
  // gets noticeable on long markdown bodies.
  const diff = useMemo(() => {
    const linesBefore = before.split('\n');
    const linesAfter = after.split('\n');
    return simpleLineDiff(linesBefore, linesAfter);
  }, [before, after]);
  return (
    <pre className="px-3 py-2 max-h-56 overflow-auto text-xs font-mono leading-tight">
      {diff.map((row, i) => (
        <div
          key={i}
          className={cn(
            row.kind === 'add' && 'diff-add',
            row.kind === 'del' && 'diff-del',
          )}
        >
          <span className="opacity-50 mr-2">
            {row.kind === 'add' ? '+' : row.kind === 'del' ? '-' : ' '}
          </span>
          {row.text || ' '}
        </div>
      ))}
    </pre>
  );
}

interface DiffRow {
  kind: 'eq' | 'add' | 'del';
  text: string;
}

/** O(n*m) LCS line diff — fine for the small contexts we operate on. */
function simpleLineDiff(a: string[], b: string[]): DiffRow[] {
  const n = a.length;
  const m = b.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }
  const rows: DiffRow[] = [];
  let i = n;
  let j = m;
  while (i > 0 && j > 0) {
    if (a[i - 1] === b[j - 1]) {
      rows.push({ kind: 'eq', text: a[i - 1] });
      i--;
      j--;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      rows.push({ kind: 'del', text: a[i - 1] });
      i--;
    } else {
      rows.push({ kind: 'add', text: b[j - 1] });
      j--;
    }
  }
  while (i > 0) {
    rows.push({ kind: 'del', text: a[--i] });
  }
  while (j > 0) {
    rows.push({ kind: 'add', text: b[--j] });
  }
  return rows.reverse();
}
