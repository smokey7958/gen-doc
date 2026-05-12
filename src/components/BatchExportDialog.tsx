/**
 * BatchExportDialog — checkbox list of all tabs; user picks which to export.
 *
 * Single folder picker downstream (in main/ipc.ts invokeExportTabs) writes
 * every selected tab into the picked folder as its native format (.md /
 * .html / .docx / .xlsx / .pptx). Collision handling (two same-named tabs
 * OR collision with an existing file in the folder) is done in main — see
 * doc-block there.
 *
 * Rationale for the per-tab checkbox UX (vs. "always export every tab"):
 *   • A user with 12 tabs but only 3 finalised drafts wants to publish just
 *     those 3, not litter the destination folder with 9 in-progress files.
 *   • Empty binary tabs (fresh docx/xlsx/pptx that the user opened but
 *     hasn't edited) would write 0-byte files that Office can't open;
 *     they're auto-disabled in the list with a tooltip explaining why.
 *   • Symmetric with the per-tab × Download button on each TabBar entry —
 *     this is "do that, but for many".
 */

import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Code2,
  FileCode2,
  FileSpreadsheet,
  FileText,
  Folder,
  Presentation,
} from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from './ui/dialog';
import { Button } from './ui/button';
import { cn } from '../lib/utils';
import type { Tab } from '../types/tab';
import type { TabType } from '../types/manifest';

const TYPE_ICON: Record<TabType, React.ComponentType<{ className?: string }>> = {
  markdown: FileCode2,
  html: Code2,
  docx: FileText,
  xlsx: FileSpreadsheet,
  pptx: Presentation,
};

const TYPE_COLOR: Record<TabType, string> = {
  markdown: 'text-sky-500',
  html: 'text-rose-500',
  docx: 'text-indigo-500',
  xlsx: 'text-emerald-500',
  pptx: 'text-orange-500',
};

const TYPE_LABEL: Record<TabType, string> = {
  markdown: 'Markdown',
  html: 'HTML',
  docx: 'Word',
  xlsx: 'Excel',
  pptx: 'PowerPoint',
};

/** Maps tab.type → exported file extension (without the dot). */
const TYPE_EXT: Record<TabType, 'md' | 'html' | 'docx' | 'xlsx' | 'pptx'> = {
  markdown: 'md',
  html: 'html',
  docx: 'docx',
  xlsx: 'xlsx',
  pptx: 'pptx',
};

interface Props {
  open: boolean;
  onOpenChange(b: boolean): void;
  tabs: Tab[];
  /** Called with the IDs of the user-checked tabs after they click 匯出. The
   *  parent owns the actual flush + IPC call so error toasts surface the
   *  same shape as single-tab export. */
  onConfirm(selectedIds: string[]): void;
}

export function BatchExportDialog({
  open,
  onOpenChange,
  tabs,
  onConfirm,
}: Props): JSX.Element {
  /**
   * Pre-compute per-tab eligibility. Empty binary tabs (0-byte data) can't
   * be exported as valid Office files — same gate that single-tab export
   * uses (App.tsx handleExportTab line ~730). Markdown / HTML always
   * exportable; an empty .md / .html is still a valid file.
   */
  const rows = useMemo(() => {
    return tabs.map((t) => {
      const canExport =
        t.type === 'markdown' || t.type === 'html' || t.data.byteLength > 0;
      return {
        id: t.id,
        name: t.name,
        type: t.type,
        canExport,
        disabledReason: canExport
          ? null
          : '這個頁籤還沒有內容，先在編輯器中輸入再匯出',
      };
    });
  }, [tabs]);

  // Default: every exportable tab checked, ONLY at the moment the dialog
  // opens. Once open, mutations to the `rows` reference (driven by every
  // Zustand store update — patchTab / saveState / canUndo / error /
  // addUsage all bump the whole-state object and re-render App with a
  // fresh `ws.tabs` reference) MUST NOT wipe the user's checkbox state.
  //
  // R319 — previous version had deps `[open, rows]` so any `tabs`
  // reference change while the dialog was open would reset selection to
  // "all checked", silently overriding the user's deselections. Realistic
  // trigger: user opens dialog with 5 tabs and unchecks 2 (3 selected),
  // then a background AI streaming tick / Ctrl+Enter Apply / undo /
  // autosave dirty-flip fires while they're still looking at the dialog
  // — selection snaps back to all 5 checked, user clicks 匯出 expecting
  // 3 files and gets 5 (the 2 in-progress drafts they specifically
  // deselected now live in the destination folder; recovery requires
  // manual cleanup).
  //
  // Two-effect split closes both axes:
  //   1. `prevOpenRef` records the previous `open` value. Only the
  //      false→true transition seeds a fresh "all exportable" selection.
  //      Subsequent renders with `open===true` don't touch selection.
  //   2. A second effect prunes `selected` whenever `rows` changes to
  //      drop ids that no longer exist (e.g., a tab was closed via a
  //      keybinding while the dialog is open). Pruning is conservative —
  //      we never AUTO-ADD newly-appeared ids, because the user's
  //      "this many checked" intent should be preserved.
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const prevOpenRef = useRef(false);
  useEffect(() => {
    if (open && !prevOpenRef.current) {
      // false → true transition.
      const fresh = new Set<string>();
      for (const r of rows) {
        if (r.canExport) fresh.add(r.id);
      }
      setSelected(fresh);
    }
    prevOpenRef.current = open;
    // `rows` intentionally omitted — we only want this effect to depend
    // on `open`. Including `rows` re-introduces the reset-on-tab-change
    // bug R319 fixes. The current `rows` value is read inside the
    // open-transition branch (closure capture is safe — that branch
    // runs synchronously during the same render that flipped `open`).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);
  // Prune-only effect: drops ids from `selected` that no longer correspond
  // to a row. Doesn't add. Returning `cur` when nothing changed avoids a
  // gratuitous re-render.
  useEffect(() => {
    if (!open) return;
    setSelected((cur) => {
      if (cur.size === 0) return cur;
      const validIds = new Set(rows.map((r) => r.id));
      let changed = false;
      const next = new Set<string>();
      for (const id of cur) {
        if (validIds.has(id)) next.add(id);
        else changed = true;
      }
      return changed ? next : cur;
    });
  }, [rows, open]);

  const toggle = (id: string) => {
    setSelected((cur) => {
      const next = new Set(cur);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const eligibleCount = rows.filter((r) => r.canExport).length;
  const allSelected = eligibleCount > 0 && selected.size === eligibleCount;
  const someSelected = selected.size > 0 && selected.size < eligibleCount;
  const toggleAll = () => {
    if (allSelected) {
      setSelected(new Set());
    } else {
      setSelected(new Set(rows.filter((r) => r.canExport).map((r) => r.id)));
    }
  };

  const confirm = () => {
    if (selected.size === 0) return;
    onConfirm(Array.from(selected));
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>批次匯出頁籤</DialogTitle>
          <DialogDescription>
            勾選要匯出的頁籤，下一步會請你選擇目標資料夾；每個頁籤會以原始格式（.md /
            .html / .docx / .xlsx / .pptx）寫入該資料夾，遇到重名會自動加上 (2)、
            (3) 後綴。
          </DialogDescription>
        </DialogHeader>

        {rows.length === 0 ? (
          <div className="text-sm text-muted-foreground py-4 text-center">
            目前沒有任何頁籤可以匯出
          </div>
        ) : (
          <>
            {/* Header row — 全選 toggle. Indeterminate state when only some
                are selected, mirroring the OS file-explorer convention. */}
            <label className="flex items-center gap-2 px-2 py-1 border-b cursor-pointer hover:bg-secondary/40">
              <input
                type="checkbox"
                checked={allSelected}
                ref={(el) => {
                  if (el) el.indeterminate = someSelected;
                }}
                onChange={toggleAll}
                className="h-3.5 w-3.5"
                aria-label={allSelected ? '取消全選' : '全選'}
              />
              <span className="text-xs text-muted-foreground">
                {allSelected
                  ? '取消全選'
                  : someSelected
                    ? `已選 ${selected.size} / ${eligibleCount} 個`
                    : '全選'}
              </span>
            </label>
            <ul className="max-h-[320px] overflow-y-auto space-y-0.5">
              {rows.map((r) => {
                const Icon = TYPE_ICON[r.type];
                const colorClass = TYPE_COLOR[r.type];
                const checked = selected.has(r.id);
                return (
                  <li key={r.id}>
                    <label
                      className={cn(
                        'flex items-center gap-2 px-2 py-1.5 rounded text-sm',
                        r.canExport
                          ? 'cursor-pointer hover:bg-secondary/40'
                          : 'opacity-50 cursor-not-allowed',
                      )}
                      title={r.disabledReason ?? undefined}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={!r.canExport}
                        onChange={() => toggle(r.id)}
                        className="h-3.5 w-3.5 shrink-0"
                      />
                      <Icon className={cn('h-4 w-4 shrink-0', colorClass)} />
                      <span className="flex-1 truncate">{r.name}</span>
                      <span className="text-[11px] text-muted-foreground shrink-0">
                        {TYPE_LABEL[r.type]}
                      </span>
                    </label>
                  </li>
                );
              })}
            </ul>
          </>
        )}

        <div className="flex items-center justify-between pt-2">
          <span className="text-xs text-muted-foreground">
            {selected.size > 0
              ? `將匯出 ${selected.size} 個檔案`
              : '請至少勾選一個頁籤'}
          </span>
          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => onOpenChange(false)}
            >
              取消
            </Button>
            <Button
              size="sm"
              disabled={selected.size === 0}
              onClick={confirm}
              title={
                selected.size === 0
                  ? '請至少勾選一個頁籤'
                  : `匯出 ${selected.size} 個檔案`
              }
            >
              <Folder className="h-3.5 w-3.5" />
              選擇資料夾並匯出
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** Helper exported for App.tsx's handleBatchExport so the type→ext mapping
 *  is in one place and stays in lockstep with the dialog's eligibility row. */
export function exportExtForTab(type: TabType): 'md' | 'html' | 'docx' | 'xlsx' | 'pptx' {
  return TYPE_EXT[type];
}
