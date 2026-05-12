/**
 * AI Panel — fixed-right sidebar, spec §5.4.
 *
 * Owns: chat history list, input box, model picker, tool toggle, selection
 * badge, pending-change list with Apply / Reject. Streaming UI is driven by
 * the AI store; ChangeSet apply is performed here so that the workspace
 * store update happens in a single place.
 */

import { useEffect, useRef, useState } from 'react';
import { Bot, Check, ChevronDown, Copy, RefreshCw, Send, Wrench, X, AlertTriangle, User, Wand2, FileSearch, MessageSquarePlus } from 'lucide-react';
import { useAI, type PendingChange } from '../store/ai';
import { useWorkspace } from '../store/workspace';
import { SUPPORTED_MODELS, type ChatMessage, type ContentBlock } from '../types/ai';
import { Button } from './ui/button';
import {
  cancelInflight,
  continueAfterToolResult,
  pushSyntheticToolResult,
  retryLastTurn,
  sendUserPrompt,
} from '../ai/orchestrator';
import { applyChangeset } from '../ai/changeset-apply';
import { DiffPreview } from './DiffPreview';
import { notify } from '../store/toast';
import { cn, slicePreview } from '../lib/utils';
import { serializeChangeset } from '../lib/changeset-serialize';

interface Props {
  width: number;
  onWidthChange(w: number): void;
  hasApiKey: boolean;
  onOpenSettings(): void;
}

/**
 * Renderer-lifetime cache for the input box draft. AIPanel is conditionally
 * rendered in App.tsx (`!focusMode && <AIPanel/>`), so toggling focus mode
 * unmounts the whole panel — and with it, any half-typed prompt the user was
 * composing. Same pattern as FindReplaceDialog.lastFindState: module-level
 * mirror, written on every change, seeded into useState on next mount. Not
 * persisted to localStorage since drafts can contain sensitive context the
 * user wouldn't expect to see across an app restart.
 */
let draftMemory = '';

/**
 * Pairs with draftMemory: when the user has scrolled up to read older
 * messages and toggles focus mode (which unmounts the whole panel), the
 * fresh remount would otherwise lose their position. The auto-follow path
 * on mount snaps to bottom, which is correct *only* when the user was
 * pinned there. If they were mid-history we want to land back on the same
 * messages instead of yanking them to the latest stream. Cached fields:
 *   • scrollTop  — absolute pixel offset from the top of the message list.
 *     Stable across remounts because new messages append to the bottom, so
 *     content above the saved offset is byte-identical.
 *   • pinned     — true if the user was within 80px of the bottom at
 *     unmount; lets us short-circuit the restore (default auto-follow does
 *     the right thing) and avoid a stale-content jump if more messages
 *     streamed in while focus mode was on.
 * Same volatility as draftMemory: cleared on app reload, not persisted.
 */
let chatScrollMemory: { scrollTop: number; pinned: boolean } | null = null;

// R259 — workspace-swap wipe at MODULE level so it fires regardless of
// AIPanel's mount state. R174's component-internal useEffect[ws.workspaceId]
// only runs while AIPanel is mounted; in focus mode the panel is unmounted
// (`!focusMode && <AIPanel/>` at App.tsx:~1490), so a Ctrl+O / Ctrl+N
// during focus mode never fires the wipe. Concrete trigger:
//   1. User types「做表格」into the AI prompt in workspace A.
//   2. Toggles focus mode → AIPanel unmounts. `draft` state dies, but the
//      effect at line 70 has already mirrored it into `draftMemory`.
//   3. Ctrl+O loads workspace B. `useWorkspace.workspaceId` changes A → B.
//      No mounted AIPanel to fire R174 useEffect.
//   4. Toggles focus mode off → AIPanel re-mounts. `useState(draftMemory)`
//      seeds with A's「做表格」 — leaking a foreign-workspace draft into
//      B's input. If the user reads it as「I was about to send this」 and
//      hits Enter, the prompt fires against B's tabs / context instead
//      of A's. Same sensitive-leak class R174's doc-block called out.
//
// Module-level Zustand subscribe runs once per renderer process (renderer
// has app-lifetime, no unmount), and fires on every workspace state
// change with prev/next snapshots. We watch sessionEpoch (NOT workspaceId)
// and wipe only when the value actually changes (initial-null → first-non-
// null is not a swap, mirroring R174's `prev === null` skip rationale).
//
// R388 — switched from `workspaceId` to `sessionEpoch` so Save-As doesn't
// trip the wipe. workspaceId changes on BOTH true workspace swaps (open /
// new) AND on Save-As (R385 — to keep SQLite undo / conversation rows
// aligned with the new filePath identity). Save-As from the user's
// perspective is a rename + persist, not a workspace switch: the
// conversation, tabs, draft, scroll position all conceptually belong to
// the same buffer they were just editing in. Pre-R388 this subscriber
// keyed off workspaceId, so a Ctrl+Shift+S on an untitled workspace
// (workspaceId `u-1` → `f-<hash>`) wiped the half-composed draft in the
// AI input box and reset the chat scroll position to the bottom — both
// disruptive surprises with no relation to the save action. sessionEpoch
// is bumped ONLY by loadFromOpened / newWorkspace (store/workspace.ts),
// not by performSave's setState, so this subscriber now ignores Save-As
// noise while still firing for genuine workspace transitions.
//
// Note that this is in addition to (not a replacement for) R174's
// in-component effect: R174 still owns the synchronous `setDraft('')`
// to wipe the visible textarea when the panel IS mounted during the
// swap. The module-level subscriber owns the headless wipe (memories
// only) for the unmounted-during-swap case. Both paths converge on
// the same end-state: empty draft on next mount.
useWorkspace.subscribe((s, prev) => {
  if (s.sessionEpoch === prev.sessionEpoch) return;
  if (prev.sessionEpoch === null) return;
  draftMemory = '';
  chatScrollMemory = null;
});

export function AIPanel({ width, onWidthChange, hasApiKey, onOpenSettings }: Props): JSX.Element {
  const ai = useAI();
  const ws = useWorkspace();
  const [draft, setDraft] = useState(draftMemory);
  useEffect(() => {
    draftMemory = draft;
  }, [draft]);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  /**
   * Synchronous in-flight gate for `send()`. R162 — the existing
   * `ai.streaming.requestId` closure check at line ~209 is stale within the
   * window between two rapid keypresses / clicks: React batches state
   * updates, so a re-render that would expose a fresh `ai` snapshot hasn't
   * committed yet, and `sendUserPrompt`'s own internal `runTurn`-→
   * `resetStreaming` chain only sets `streaming.requestId` AFTER the first
   * `await persistMessage` boundary (sync prelude pushes the user message,
   * then yields). Two rapid sends both pass the gate before either hits its
   * first await, both push `userMsg` to `ai.messages`, both fire
   * `persistMessage` and `runTurn` — the conversation now has duplicate
   * user messages and `inflight` is overwritten so the first IPC stream
   * orphans. The ref is set synchronously before the await and cleared in
   * `finally`, closing the race even when React hasn't re-rendered between
   * keypresses (auto-repeat / rage-click).
   */
  const sendingRef = useRef(false);

  // Only surface a 選取上下文 badge / send selectionContext when the stored
  // selection belongs to the tab the user is currently looking at. Today
  // only MarkdownEditor pushes selection up; switching from md tab A (with
  // a selection) to tab B leaves the badge stale ("選取上下文：[md selection]
  // foo" while the user is staring at a spreadsheet) and any prompt sent
  // from B would silently include A's text. Hide rather than clear so that
  // returning to A still shows the original selection — the editor itself
  // hasn't lost it.
  const activeSelection =
    ws.selection && ws.selection.tabId === ws.activeTabId ? ws.selection : null;

  // Listen for the global "focus AI" command from menu.ts.
  useEffect(() => {
    const handler = () => inputRef.current?.focus();
    window.addEventListener('gendoc:focusAI', handler);
    return () => window.removeEventListener('gendoc:focusAI', handler);
  }, []);

  // R174 — clear cross-mount memories on workspace swap. `draftMemory` and
  // `chatScrollMemory` are module-level caches that survive AIPanel unmount
  // (focus-mode toggle, etc.) so a half-typed prompt or scroll position
  // isn't lost between remounts within the same workspace. But across
  // workspaces the semantics flip: the draft refers to OLD workspace's
  // tabs / files (sensitive context), and the scroll position is keyed
  // against OLD workspace's chat history that's now been wiped by clear()
  // (loadFromOpened / newWorkspace cascade). Without this reset:
  //   • User types "在 sheet1 的 A1 寫入 hello" in workspace A, hits Ctrl+O
  //     to switch to B, sees their A-context draft in B's input — confusing
  //     and a leak (the draft string mentions tab names / structure of A).
  //   • User scrolls A's chat up to read history, swaps to B, toggles focus
  //     mode in B (or anything else that unmounts AIPanel), toggles back —
  //     scrollMemory restores A's scrollTop into B's empty / freshly-loaded
  //     chat, landing arbitrarily off-position.
  // Track the previous sessionEpoch so the cleanup fires only on actual
  // transitions (1 → 2, 2 → 3, …), not on:
  //   • Initial null → first non-null open (the user may have typed a draft
  //     on the empty splash screen, intending it for the workspace they're
  //     about to open — clearing would lose that intent).
  //   • Focus-mode-toggle remount (new component instance, ref re-initializes
  //     to current value so first effect-fire sees prev === current and
  //     skips). Without this guard the first effect run on remount would
  //     wipe the draft we just restored from `draftMemory` at line 67's
  //     `useState(draftMemory)`.
  //   • R388 — Save-As (workspaceId u-1 → f-<hash> with sessionEpoch
  //     unchanged). User pressing Ctrl+Shift+S mid-prompt-composition is
  //     NOT a workspace switch and the in-progress draft must persist.
  //     See module-level subscriber doc-block above for the full Save-As
  //     vs swap distinction.
  const prevSessionEpochRef = useRef(ws.sessionEpoch);
  useEffect(() => {
    const prev = prevSessionEpochRef.current;
    prevSessionEpochRef.current = ws.sessionEpoch;
    if (prev === null || prev === ws.sessionEpoch) return;
    draftMemory = '';
    chatScrollMemory = null;
    setDraft('');
  }, [ws.sessionEpoch]);

  // Auto-scroll to bottom on new content — but only if the user is already
  // pinned near the bottom. If they've scrolled up to re-read mid-stream, we
  // leave them alone instead of yanking the viewport on every chunk.
  // Exception: when the user just sent a prompt, force the scroll regardless
  // — they explicitly acted and expect to see their own message + the reply.
  // (They might have scrolled up to grep history before composing, then sent
  // without scrolling back — without the override the reply would stream in
  // off-screen and look like nothing happened.)
  const lastUserMsgIdx = (() => {
    for (let i = ai.messages.length - 1; i >= 0; i--) {
      if (ai.messages[i].role === 'user') return i;
    }
    return -1;
  })();
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
    // Run only on user-sent transitions: lastUserMsgIdx changes when a new
    // user message lands at the end of the array.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lastUserMsgIdx]);
  // Track "was the user pinned to the bottom *before* this content arrived"
  // via a scroll listener that fires only on the user's own scrolling. The
  // previous version measured the bottom distance INSIDE the post-render
  // useEffect — which sees scrollHeight already grown by the new content,
  // so a single chunk taller than 80px (DiffPreview landing, a multi-line
  // markdown block, a long tool_result) made distanceFromBottom register
  // > 80 even when the user had been glued to the bottom an instant before.
  // The auto-follow then silently failed and the user got stranded mid-
  // stream with no obvious cause. Capturing pinned-state from the last
  // *user-driven* scroll event sidesteps the chunk-size cliff entirely.
  // Programmatic `scrollTop = scrollHeight` also fires `scroll`, which
  // re-pins the ref to true — so the initial mount + manual scroll-to-
  // bottom on user send both keep auto-follow armed for the next chunk.
  // Seed initial pinned-state from chatScrollMemory so a mid-history remount
  // doesn't have to fight an initial true→false race against the auto-follow
  // effect below. When nothing's remembered (first ever mount), default to
  // true (current convention).
  const pinnedToBottomRef = useRef(chatScrollMemory?.pinned ?? true);
  // Re-rendered mirror of pinnedToBottomRef. The ref itself can't drive UI
  // (mutating it never re-renders), but we need a "user has scrolled away"
  // signal to surface the floating Jump-to-latest button. We only flip state
  // when CROSSING the threshold so a continuous scroll gesture doesn't burn
  // a render per pixel — onScroll fires dozens of times per second.
  const [pinnedToBottom, setPinnedToBottom] = useState(chatScrollMemory?.pinned ?? true);
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onScroll = () => {
      const dist = el.scrollHeight - el.scrollTop - el.clientHeight;
      const pinned = dist < 80;
      pinnedToBottomRef.current = pinned;
      // Functional setState avoids a stale-closure re-flip when the threshold
      // straddles repeatedly.
      setPinnedToBottom((cur) => (cur === pinned ? cur : pinned));
    };
    el.addEventListener('scroll', onScroll, { passive: true });
    return () => el.removeEventListener('scroll', onScroll);
  }, []);
  useEffect(() => {
    if (!pinnedToBottomRef.current) return;
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [ai.messages, ai.streaming.content, ai.pending]);

  // Restore-on-mount / capture-on-unmount for chat scroll position. Mirrors
  // FileExplorer's treeScrollMemory pattern. queueMicrotask defers the
  // restore until after the lastUserMsgIdx effect (which unconditionally
  // sets scrollTop=scrollHeight on mount) and the auto-follow effect have
  // run synchronously — both DOM mutations are queued in the same React
  // flush, so the microtask draining after all effects sees the final
  // scrollHeight, then overrides scrollTop with the remembered offset
  // before the browser paints. No intermediate frame at the bottom is
  // visible to the user. Skip the restore when the user was pinned at
  // unmount: in that case the auto-follow path is already correct, and
  // restoring an absolute offset that was "just barely above the bottom"
  // would land them slightly above the latest stream after new messages
  // had arrived during the focus-mode window.
  useEffect(() => {
    const remembered = chatScrollMemory;
    if (remembered && !remembered.pinned) {
      queueMicrotask(() => {
        const el = scrollRef.current;
        if (!el) return;
        el.scrollTop = remembered.scrollTop;
      });
    }
    return () => {
      const el = scrollRef.current;
      if (!el) return;
      chatScrollMemory = {
        scrollTop: el.scrollTop,
        pinned: pinnedToBottomRef.current,
      };
    };
  }, []);

  /**
   * Manual jump-to-latest. Programmatic `scrollTop = scrollHeight` fires the
   * same `scroll` event the user's gestures do, so the listener above will
   * synchronously flip pinnedToBottomRef + state back to true and the button
   * disappears. No need to set the state ourselves.
   */
  const scrollToLatest = () => {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  };

  const send = async () => {
    // R162 — sync ref gate first; see `sendingRef` doc-block above.
    if (sendingRef.current) return;
    if (!draft.trim() || !hasApiKey) return;
    // A turn is already inflight — let the user finish it (or hit 停止) before
    // queueing another. Without this, Enter during streaming starts a second
    // `runTurn` that clobbers the shared `inflight` handle, leaving the first
    // stream un-cancellable and the second response confused about state.
    // Read live store state (not the React snapshot `ai`) so a re-entry that
    // happens after streaming started but before re-render also bounces.
    if (useAI.getState().streaming.requestId !== null) return;
    sendingRef.current = true;
    const prompt = draft;
    setDraft('');
    // R239 — capture workspaceId so the catch handler below can detect
    // mid-send workspace swap. Without this, the post-await catch
    // restores OLD's prompt into NEW workspace's input via setDraft AND
    // writes OLD's exception text into NEW's `ai.error` banner. Both
    // are visible cross-workspace pollution: user opens workspace B
    // post-send, sees an error message about an A-turn they already
    // moved on from, and a prompt suggestion in B's input that's
    // completely irrelevant. Same shape as R236 / R237 / R238 — the
    // catch in this React component is the renderer-side analog to
    // those orchestrator-side guards.
    const sendWorkspaceId = useWorkspace.getState().workspaceId;
    try {
      // R183 — read RunOptions from live store at click-time, not from the
      // React closure `ai`. The closure is rendered-time stable, so a model
      // / tools-enabled change that happens between render and click (user
      // toggles Wrench, hits Enter immediately; user picks a different
      // model in the dropdown then sends without intervening keystrokes
      // letting React commit) would otherwise send the OLD turn's settings.
      // Mirrors the orchestrator's own `currentOpts()` helper at line ~423,
      // which already reads from `useAI.getState()` for its recursive
      // runTurn so handleAssistantToolCalls's follow-up turns honour live
      // settings — same reasoning applies to the user-triggered entry points.
      const a = useAI.getState();
      // R188 — also read selection from live workspace store at click-time.
      // R183 fixed the RunOptions (model / tools / etc.) closure-staleness
      // for the same render-vs-commit race, but the `activeSelection` line
      // still went through the React closure `ws.selection` and could
      // forward an out-of-date selection on the typical「快速 highlight →
      // Ctrl+L → Ctrl+Enter」flow when React hasn't committed the new
      // selection yet. Reading via getState() keeps both axes (RunOptions
      // and selection) consistent at click time.
      const lws = useWorkspace.getState();
      const liveSel =
        lws.selection && lws.selection.tabId === lws.activeTabId
          ? lws.selection
          : null;
      await sendUserPrompt(prompt, {
        toolsEnabled: a.toolsEnabled,
        model: a.model,
        maxTokens: a.maxTokens,
        temperature: a.temperature,
        // Forward the full selection text, not the 60-char badge preview —
        // the model needs to see the whole snippet the user highlighted.
        // Fall back to preview if text is empty (older payloads / host
        // editors that haven't migrated to the typed `text` field yet).
        selectionContext: liveSel?.text || liveSel?.preview,
      });
    } catch (err) {
      // R210 — restore the cleared draft on send-failure. `sendUserPrompt`
      // can throw during its dangling-tool_use cleanup loop (each iteration
      // does `await pushSyntheticToolResult` → `persistMessage` IPC →
      // sqlite append) BEFORE the user's own message is pushed to
      // `ai.messages`. If the DB is locked (OS-level backup snap, antivirus
      // scan, disk full mid-INSERT), the throw escapes here. Without
      // recovery the user sees their typed text vanish from the input,
      // nothing in the chat to indicate what happened, and an
      // unhandledrejection in DevTools they can't act on. Retry doesn't
      // help either — `retryLastTurn` re-runs whatever's at the tail of
      // `ai.messages`, which in this failure window is still the PREVIOUS
      // turn (not the user's just-typed prompt that never made it in).
      //
      // Functional setDraft so we don't clobber text the user typed during
      // the await window. The send IPC chain is sub-second in practice but
      // still a real microtask boundary; whichever text is in the input at
      // catch time is what the user can see and expects to keep. Only seed
      // the lost prompt back when the input is genuinely empty.
      //
      // Surface the error through `setError` (mirrors how onError chunks
      // surface other failures, see orchestrator.ts:439-441 / R203 / R204
      // family) so the chat shows a banner explaining the failure rather
      // than relying on the user to inspect DevTools.
      // R239 — both setDraft restore and setError are gated on workspace
      // identity. Skip both if swapped: the prompt belonged to OLD
      // workspace and shouldn't reappear in NEW's input; the error
      // shouldn't pop in NEW's banner either. R174's swap-cleanup
      // already wiped draftMemory + draft on transition, and R236 /
      // R237 / R238 already cover the orchestrator-side error-banner
      // leaks — R239 closes the renderer-side analog.
      if (useWorkspace.getState().workspaceId === sendWorkspaceId) {
        setDraft((cur) => (cur === '' ? prompt : cur));
        useAI.getState().setError(err instanceof Error ? err.message : String(err));
      }
    } finally {
      sendingRef.current = false;
    }
  };

  const onApply = async (p: PendingChange) => {
    // Refuse while another turn is streaming — `continueAfterToolResult`
    // calls `runTurn`, which overwrites the shared `inflight` handle and
    // `streaming.requestId`. Triggering it on top of a live stream orphans
    // the previous handle and interleaves its chunks into the new buffer.
    // (Easy trigger: an assistant turn with multiple pending changes, or
    // rapid Ctrl+Enter against the most-recent pending.)
    // R265 — read live store, not the React-closure `ai` snapshot. The
    // closure freezes streaming.requestId at the last render-commit
    // moment; if cancelInflight cleared it (e.g., R256 path during the
    // bridge dispatch loop, OR a normal cancel that resolved the runTurn
    // promise) and the user immediately clicks Apply BEFORE React has
    // re-rendered with the new requestId=null, the closure still has
    // the stale bridge uuid and this gate falsely returns. The user
    // clicks again post-render and it works the second time, but the
    // first click is silently swallowed — confusing UX. Mirrors the
    // existing live-read pattern at the gate's siblings: line 409's
    // pending dedupe (`useAI.getState().pending.some(...)`), R163's
    // continueAfterToolResult gate (orchestrator.ts:235), and the
    // keyboard handler at line 668. Reading live closes the
    // sub-millisecond render-commit race without changing the gate's
    // semantic — when streaming truly is in flight, the live read
    // also says non-null and the return still fires.
    if (useAI.getState().streaming.requestId !== null) return;
    // Guard against rapid double-fire (mouse double-click, or Ctrl+Enter
    // mashed before the streaming flag flips). Without this, two clicks
    // before the first await yields can both pass the streaming check;
    // the second run then applies the changeset a second time (doubling
    // the edit) and pushes a duplicate tool_result for the same
    // tool_use_id — which the next assistant turn rejects with
    // Anthropic 400 "tool_use_id ... appeared more than once". Reading
    // current store state via getState() (not the React-snapshot `ai`)
    // ensures the second click sees the removePending below.
    if (!useAI.getState().pending.some((x) => x.id === p.id)) return;
    // R180 — capture workspaceId so the post-await mutations don't leak into
    // a swapped workspace. `await window.gendoc.undo.push(...)` correctly
    // targets the OLD workspace's undo stack via its `workspaceId` arg
    // (closure-bound, captured pre-await), but the optimistic
    // `setCanUndo(true)` / `clearAiRedo()` that follow read the LIVE store
    // — and a workspace swap during the push IPC would write OLD's "we
    // just added an undo row" optimism onto NEW's store. NEW's actual
    // canUndo is governed by its own SQLite stack (which R167's useEffect
    // refreshes on workspaceId change), so the stray setCanUndo(true)
    // creates a brief lying-button window until refreshCanUndo settles.
    // continueAfterToolResult itself has R179's workspace guard, but the
    // toolbar-flip lines between them don't — fix here.
    const applyWorkspaceId = useWorkspace.getState().workspaceId;
    // R189 — bail if the React closure `ws` is from a workspace that's
    // already been swapped before the click handler ran (swap → click
    // before React commit). Without this, the rest of onApply executes
    // in a tangled state: `applyChangeset(NEW.tabs, OLD.changeset)` is a
    // no-op (R186 reads NEW tabs, but p.changeset references OLD tab ids
    // that don't exist in NEW), the setState below sets NEW.dirty=true
    // for no visible change, the undo.push targets `ws.workspaceId` (OLD
    // closure) so it lands in OLD's stack, then R180's check passes
    // (applyWorkspaceId === current === NEW) and `setCanUndo(true)` is
    // written to NEW even though NEW's stack didn't gain a row. The
    // pending p references OLD's flow; the user's click intent was for
    // OLD, but they swapped before the handler ran. Silently no-op is
    // the cleanest semantics: their stale click on a no-longer-visible
    // pending shouldn't push DB rows or flip UI state in NEW.
    if (ws.workspaceId !== applyWorkspaceId) return;
    // R186 — read tabs live (not from closure `ws.tabs`). Rapid Apply on
    // two PendingChange (P1 → P2) without an intervening React re-render
    // both reach this line with the SAME closure `ws.tabs` (the snapshot
    // before either click landed). P2's applyChangeset on stale tabs
    // produces "OLD + P2" and the subsequent setState overwrites P1's
    // already-committed "OLD + P1" — P1's changes silently vanish. Most
    // visible when an assistant turn emits multiple write tools and the
    // user mashes Ctrl+Enter twice (the keyboard shortcut targets the
    // last pending each time, so two presses Apply two different cards).
    // useAI's pending dedupe at line 303 catches double-click on SAME p,
    // but distinct ps on the same render slip through. Reading live tabs
    // composes each apply on top of the previous setState's result.
    const liveTabs = useWorkspace.getState().tabs;
    const { tabs } = applyChangeset(liveTabs, p.changeset);
    // R250 — fix activeTabId if the changeset removed the active tab.
    // tab_delete in p.changeset filters out the targeted tab from
    // `tabs`; if that tab happened to be activeTabId, the post-setState
    // state has activeTabId pointing at a non-existent tab. EditorSurface
    // .tsx:43 `tabs.find((t) => t.id === activeTabId)` returns
    // undefined, EmptyState renders, and Ctrl+W silently no-ops via
    // R208 (workspace.ts:507 `if (removedIdx < 0) return s`). Match
    // removeTab's neighbor-fallback semantics (workspace.ts:516-521):
    // if the removed tab had a right neighbor, jump there; else fall
    // back to the last remaining tab; else null. Computed in setState
    // callback so we read live `activeTabId` (race-safe against any
    // setActiveTab that might land between the getState read above
    // and the setState write — currently zero such call sites in this
    // sync block, but the form is robust). Same shape applied at
    // handleUndo / handleRedo for parity.
    useWorkspace.setState((s) => {
      const stillValid = tabs.some((t) => t.id === s.activeTabId);
      // R276 — mirror R202's "every non-null selection points at a live tab"
      // invariant. R202 set this up for the workspace.removeTab path; the
      // tab_delete forward changeset path lands tabs in the same shape
      // (one fewer tab) but went through here (setState) instead of
      // removeTab, so it bypassed R202's selection-clear. After Apply of
      // tab_delete on a tab that had the selection, `s.selection` would
      // orphan to a vanished tab id; AIPanel's `activeSelection` filter
      // hides the stale badge but the store invariant is broken. Reconcile
      // selection in the same atomic setState as activeTabId.
      const selStillValid =
        !s.selection || tabs.some((t) => t.id === s.selection!.tabId);
      const selPatch = selStillValid ? null : { selection: null };
      if (stillValid) return { tabs, dirty: true, ...selPatch };
      const removedIdx = liveTabs.findIndex((t) => t.id === s.activeTabId);
      const fallback =
        removedIdx >= 0 && removedIdx < tabs.length
          ? tabs[removedIdx].id
          : tabs[tabs.length - 1]?.id ?? null;
      return { tabs, activeTabId: fallback, dirty: true, ...selPatch };
    });
    // Drop the pending BEFORE the first await so the dedupe gate above
    // catches a second click during the undo.push IPC window. (If we
    // dropped after the await, a click landing while undo.push was in
    // flight would still see the pending and re-enter applyChangeset.)
    ai.removePending(p.id);
    // Persist undo entry. Use `applyWorkspaceId` (live) instead of
    // closure `ws.workspaceId` — R189's guard above ensures they match
    // here, but using the live value makes the data-flow explicit and
    // robust to any future render-vs-click race that might slip past
    // R189 (e.g., if the gate is later relaxed or restructured).
    // R284 — wrap the undo.push IPC. Sibling sweep R245 / R249 / R271
    // already added try/catch to handleUndo / handleRedo / handleClearRecent
    // / SettingsDialog patch; R211 covered continueAfterToolResult right
    // below; the undo.push here was the last unguarded user-triggered IPC
    // in this handler. Without this guard, an SQLite reject (backup tool
    // pinning chat.sqlite, disk full INSERT, IPC bridge異常) escapes onApply
    // entirely — applyChangeset already mutated tabs (line ~483) and
    // removePending already fired (line ~509), but setCanUndo /
    // clearAiRedo / continueAfterToolResult all skip. Most critically,
    // continueAfterToolResult NEVER pushes the tool_result message; the
    // assistant's tool_use stays unanswered in ai.messages and the next
    // user prompt 400s on "tool_use ids were found without tool_result
    // blocks". The visual Apply succeeded but the chat is silently broken.
    //
    // Failure recovery: notify the user that this step won't be undoable
    // (the disk write didn't land — refreshCanUndo will reconcile the
    // toolbar from real state), then STILL fall through to
    // continueAfterToolResult below so the Anthropic conversation
    // invariant stays intact. Missing an undo entry is a recoverable
    // degradation; missing a tool_result breaks the chat round-trip and
    // is the higher-priority invariant to preserve.
    let pushOk = true;
    try {
      await window.gendoc.undo.push({
        // R216 — use Uint8Array-aware serializer so binary_replace /
        // tab_create / tab_delete ops survive the SQLite round-trip.
        // Naive JSON.stringify mangles Uint8Array into a plain object
        // and undo silently corrupts the tab's bytes.
        changesetJson: serializeChangeset(p.changeset),
        workspaceId: applyWorkspaceId,
      });
    } catch (err) {
      pushOk = false;
      // Workspace-swap guard — don't notify on the NEW workspace for
      // OLD's failed push. Same shape as R240 / R271's catch arms.
      if (useWorkspace.getState().workspaceId === applyWorkspaceId) {
        const msg = err instanceof Error ? err.message : String(err);
        notify(`儲存復原紀錄失敗：${msg}（此步驟無法以 Ctrl+Z 復原）`, 'error');
      }
    }
    // R180 — only flip the toolbar mirror / clear redo when we're still in
    // the originating workspace. NEW workspace's undo / redo stacks are
    // its own and shouldn't be touched by OLD's Apply flow.
    if (useWorkspace.getState().workspaceId !== applyWorkspaceId) return;
    if (pushOk) {
      // The push guarantees an undoable row exists; flip the toolbar mirror
      // optimistically so the Undo button enables in the same render where
      // the change becomes visible (no extra IPC round-trip needed).
      useWorkspace.getState().setCanUndo(true);
      // Branching the timeline (a fresh apply after some undos) invalidates
      // the redo stack — same convention as text editors / IDEs.
      useWorkspace.getState().clearAiRedo();
    }
    // R284 — proceed to continueAfterToolResult REGARDLESS of pushOk.
    // The tool_result message is what keeps Anthropic happy; preserving
    // that invariant matters more than the (already-degraded) undo trail.
    // Tell the model what happened so it can continue.
    // R183 — live RunOptions read; same rationale as `send` above.
    {
      const a = useAI.getState();
      // R211 — wrap continueAfterToolResult in try/catch so a
      // persistMessage rejection (DB locked, disk full, IPC reject)
      // doesn't escape onApply as unhandledrejection. Same shape as R210
      // for sendUserPrompt and R203 / R204 for the void-call siblings.
      // The Apply has already mutated workspace state (line 396 setState
      // tabs/dirty, line 407 undo.push, line 418 setCanUndo, line 421
      // clearAiRedo) BEFORE this await — the disk side is consistent;
      // only the model-continuation side fails. Surfacing via setError
      // gives the user a banner (with Retry) to resume the chat round-
      // trip; the applied change itself stays on screen and undoable.
      try {
        await continueAfterToolResult(
          p.toolUseId,
          { content: `Applied: ${p.summary}` },
          { toolsEnabled: a.toolsEnabled, model: a.model, maxTokens: a.maxTokens, temperature: a.temperature },
        );
      } catch (err) {
        // R240 — workspace guard, mirrors R236 / R237 / R238 / R239.
        // continueAfterToolResult's await window can outlast a Ctrl+O;
        // setError without this check would write OLD's exception
        // text into NEW workspace's `ai.error`. Reuse the
        // `applyWorkspaceId` captured at line 363 instead of a fresh
        // capture — that's already the OLD-workspace anchor that the
        // post-IPC undo flow uses (R180), so re-using it keeps the
        // workspace-pinning invariant uniform across this handler.
        if (useWorkspace.getState().workspaceId !== applyWorkspaceId) return;
        useAI.getState().setError(err instanceof Error ? err.message : String(err));
      }
    }
  };

  const onReject = async (p: PendingChange) => {
    // R265 — live read, see onApply sibling for full doc-block on the
    // closure-snapshot stale-uuid race that can swallow a click during
    // the post-cancel pre-rerender window.
    if (useAI.getState().streaming.requestId !== null) return;
    // Same dedupe gate as onApply — a second click before
    // continueAfterToolResult flips streaming.requestId would push a
    // duplicate tool_result for the same tool_use_id.
    if (!useAI.getState().pending.some((x) => x.id === p.id)) return;
    // R240 — capture workspaceId pre-mutation so the catch handler
    // below can detect mid-reject workspace swap. Without this, an
    // OLD-turn continueAfterToolResult failure that arrives after the
    // user already Ctrl+O'd would write OLD's error into NEW's banner.
    // Same shape as onApply's R180 + R240 capture-and-check, just
    // adapted: onApply already had `applyWorkspaceId` for its undo
    // push; onReject didn't have a corresponding capture because R211
    // assumed errors only fire same-workspace.
    const rejectWorkspaceId = useWorkspace.getState().workspaceId;
    ai.removePending(p.id);
    // R183 — live RunOptions read; same rationale as `send` above.
    const a = useAI.getState();
    // R211 — same try/catch shape as onApply above. The pending card has
    // already been removed sync (line 440); a thrown rejection here would
    // leave the user with the card gone but no model response and no
    // visible error. Surface via setError so the chat banner explains the
    // failure and the Retry button can re-issue the tool_result.
    try {
      await continueAfterToolResult(
        p.toolUseId,
        { content: `Rejected by user — propose a different approach.`, isError: true },
        { toolsEnabled: a.toolsEnabled, model: a.model, maxTokens: a.maxTokens, temperature: a.temperature },
      );
    } catch (err) {
      // R240 — workspace guard before setError. See onApply sibling.
      if (useWorkspace.getState().workspaceId !== rejectWorkspaceId) return;
      useAI.getState().setError(err instanceof Error ? err.message : String(err));
    }
  };

  const onModify = (p: PendingChange) => {
    // Dedupe gate — pushSyntheticToolResult would otherwise emit a second
    // tool_result for the same tool_use_id on a double-click, and the
    // user's next sendUserPrompt would 400.
    if (!useAI.getState().pending.some((x) => x.id === p.id)) return;
    // R243 — capture workspaceId so the void pushSyntheticToolResult.catch
    // below doesn't leak OLD workspace's error text into NEW's `ai.error`
    // banner across a workspace swap. Same shape as R236-R240 (orchestrator
    // and AIPanel send/apply/reject already covered) and the
    // workspace.removeTab sibling fixed in the same R243 round. Trigger:
    // user clicks 修改 in workspace A, the void IPC fires, user
    // immediately Ctrl+O to workspace B; persistMessage rejects (DB locked
    // / disk full / sqlite-wal contention during swap window); the .catch
    // reads LIVE useAI (now NEW after clear()) and writes OLD's exception
    // text into NEW's banner. R174's swap-cleanup already wiped draft /
    // chatScroll memories; this completes the renderer-side
    // workspace-pinning invariant for the last unguarded setError site.
    const modifyWorkspaceId = useWorkspace.getState().workspaceId;
    // Preserve a draft the user is already composing. Ctrl+M is global
    // (see header comment above the keydown effect) and advertised on
    // every DiffPreview's 修改 button, so the user can fire it while
    // mid-typing a follow-up in the AI textarea — the previous version
    // unconditionally overwrote `draft` with the placeholder template
    // and discarded their carefully-worded feedback. Only seed the
    // template when the textarea is empty; if they already have text,
    // that text *is* their modify request, leave it alone.
    const hasDraft = draft.trim().length > 0;
    if (!hasDraft) {
      setDraft(`修改建議：${p.summary}\n\n（請描述要怎麼調整）`);
    }
    ai.removePending(p.id);
    // Anthropic's API requires every `tool_use` block to be followed by a
    // matching `tool_result` before the next user message. The assistant's
    // tool_use is already persisted in `ai.messages` (finalizeStreaming put
    // it there); just dropping the pending leaves it orphaned, and the
    // user's next prompt would 400 with "tool_use ids were found without
    // tool_result blocks". Push a synthetic tool_result now to close the
    // pair — but DO NOT trigger a model turn yet. The user is mid-compose;
    // their actual feedback lands in the next sendUserPrompt and the model
    // sees the full sequence (tool_use → "user is modifying" → user prompt)
    // in a single round-trip.
    // R204 — attach .catch so a persistMessage rejection inside
    // pushSyntheticToolResult (IPC fail, sqlite locked, disk full) doesn't
    // surface as an unhandledrejection that Electron would crash the
    // renderer on. Same rationale as R203's persistMessage void-call: the
    // chat UI already shows the synthetic tool_result (pushMessage is sync,
    // ran before the await), so the user sees no behavioral change. The
    // failure is real — next sendUserPrompt loads from db and the
    // tool_result will be missing — but a visible error toast plus a
    // graceful retry beats a renderer crash mid-compose.
    void pushSyntheticToolResult(
      p.toolUseId,
      'User chose to modify this proposal. They will describe the desired changes in their next message.',
      true,
    ).catch((err) => {
      // R243 — workspace guard, see modifyWorkspaceId capture above.
      if (useWorkspace.getState().workspaceId !== modifyWorkspaceId) return;
      useAI.getState().setError(err instanceof Error ? err.message : String(err));
    });
    // Defer focus + select to the next frame: setDraft hasn't propagated
    // through React yet so the textarea's value is still the old one.
    // After paint, the new draft is mounted and we can select the whole
    // template — the user immediately overtypes with their actual ask.
    // When we kept the user's existing draft (hasDraft branch), DON'T
    // select-all: that would highlight their typed text and the next
    // keystroke would replace the very content we just preserved.
    requestAnimationFrame(() => {
      const el = inputRef.current;
      if (!el) return;
      el.focus();
      if (!hasDraft) el.select();
    });
  };

  // Keyboard shortcuts for the most recent pending change. Buttons in
  // DiffPreview already advertise these in their tooltips (Ctrl+Enter /
  // Ctrl+M / Ctrl+Backspace) but were never wired — users had to mouse to
  // act on every AI suggestion. We target the *last* pending so the user
  // can rapid-fire through several proposals without picking by hand.
  //
  // Per-shortcut focus rules:
  //   • Ctrl+Backspace is "delete previous word" in every text input on
  //     the planet — hijacking it while the user is mid-thought composing
  //     a follow-up question is a footgun, so we skip it inside any
  //     editable surface.
  //   • Ctrl+Enter and Ctrl+M have no native text-editing meaning, so we
  //     fire them everywhere. The high-leverage case: user just hit Enter
  //     to send a prompt, AI proposes a change, focus is still in the AI
  //     textarea — they read the tooltip, hit Ctrl+Enter, and Apply runs.
  //     Previously this was silently swallowed because the textarea
  //     check fired on all three keys uniformly.
  useEffect(() => {
    if (ai.pending.length === 0) return;
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      // R217 — require Ctrl/Cmd ALONE; reject Shift / Alt combos. The
      // tooltip the DiffPreview shows on each card promises exactly
      // Ctrl+Enter / Ctrl+M / Ctrl+Backspace; binding to "any modifier
      // combo that includes Ctrl/Cmd" surprises users in two concrete
      // ways:
      //
      //   1. Ctrl+Shift+Enter inside the AI textarea. The textarea's
      //      onKeyDown (line ~1416-1421 in InputBar) rejects Enter with
      //      any modifier and falls through to default — the user gets
      //      a newline AS WELL AS an Apply firing on the latest pending.
      //      Visible symptom: typing「下一輪我想…」hitting Ctrl+Shift+
      //      Enter for "newline + start fresh paragraph" silently
      //      applies the still-visible AI card under the input.
      //   2. Ctrl+Shift+M is the macOS system shortcut "Minimize All
      //      Windows" in many IMEs / window managers; Cmd+Alt+M is the
      //      Mac standard "minimize and hide". Hijacking either fires
      //      onModify on a card the user wasn't even looking at, and
      //      the global system action stops working as long as a
      //      pending card exists.
      //
      // The textarea's send-Enter handler already uses the strict form
      // (`!e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey`); this
      // guard brings the global counterpart into the same convention.
      // Backspace's input-skip rule below is preserved so Ctrl+
      // Backspace still defers inside text fields — the modifier
      // tightness is orthogonal.
      if (e.shiftKey || e.altKey) return;
      // While a turn is streaming, Apply / Reject would orphan the inflight
      // handle (see comment on onApply). 修改 is purely local-state and
      // safe to allow.
      const streaming = useAI.getState().streaming.requestId !== null;
      const pendingTarget = ai.pending[ai.pending.length - 1];
      if (e.key === 'Enter') {
        if (streaming) return;
        e.preventDefault();
        void onApply(pendingTarget);
      } else if (e.key.toLowerCase() === 'm') {
        e.preventDefault();
        onModify(pendingTarget);
      } else if (e.key === 'Backspace') {
        // Only Ctrl+Backspace defers to the input — see header comment.
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          target?.isContentEditable ||
          !!target?.closest?.('.cm-editor')
        ) {
          return;
        }
        if (streaming) return;
        e.preventDefault();
        void onReject(pendingTarget);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
    // onApply / onReject / onModify close over `ai` & `ws`; refresh
    // listener whenever the pending list changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ai.pending]);

  return (
    <div
      className="relative h-full border-l flex flex-col bg-background"
      style={{ width }}
    >
      <ResizeHandle width={width} onChange={onWidthChange} />
      <Header hasApiKey={hasApiKey} onOpenSettings={onOpenSettings} />

      {activeSelection && (() => {
        // The badge previously showed only `preview` (already capped at 60
        // chars upstream), then `truncate` clipped that further whenever
        // the panel was narrow — at default ~360 px width we land around
        // 30-35 visible chars. The user's actual selection (forwarded as
        // `selectionContext`) can be hundreds or thousands of characters,
        // so they had no way to verify what was attached to their prompt
        // before sending. Two additions:
        //   • Char count on the right — at-a-glance "is this 50 chars or
        //     5000?" sanity-check, scoped to a fixed-width tabular-nums
        //     span so it doesn't fight the truncate ellipsis.
        //   • Hover tooltip on the preview span carrying the full text,
        //     capped at 800 chars (browsers render very long titles
        //     awkwardly — line wraps vary, some platforms ellipsize).
        //     Anything past the cap is suffixed with the trailing-char
        //     count so users see "this snippet exists, here's the first
        //     800 chars, and another 4,200 are attached".
        // Falls back to `preview` for hosts whose `text` is empty (the
        // SelectionInfo contract permits that for spreadsheet-style
        // payloads — see workspace.ts SelectionInfo.text).
        const fullText = activeSelection.text;
        const charCount = fullText.length;
        const TOOLTIP_CAP = 800;
        const tooltip = !fullText
          ? activeSelection.preview
          : fullText.length <= TOOLTIP_CAP
            ? fullText
            : `${fullText.slice(0, TOOLTIP_CAP)}…\n\n（已截斷，全部 ${fullText.length.toLocaleString()} 字元）`;
        return (
          <div className="px-3 py-1.5 border-b bg-accent/40 text-xs flex items-center gap-2">
            {/* R104 — translation-oversight cleanup. "Context:" was the lone
                English label in this row while both sibling strings beside
                it were Traditional Chinese:
                  • char-count title (line 470) — 「選取片段字元數（送給模型
                    的完整內容）」
                  • clear-X title (line 477) — 「清除選取上下文」
                The X button's title is the strongest precedent because it
                explicitly names the thing this label labels — "清除選取上
                下文" establishes 「選取上下文」 as the canonical noun for
                this concept, used to describe the action targeting THIS
                label. Reusing the same noun verbatim ("選取上下文：") makes
                one row read in one voice instead of three. Same R-class
                fix as SettingsDialog.tsx:373-380 where the "Test connection"
                button label — also the lone English string in an otherwise
                fully-Chinese dialog — was renamed to「測試連線」 and the
                comment there explicitly framed it as a translation
                oversight rather than an intentional choice. Identical
                pattern, same codebase, same fix shape.

                R143 — refreshed line numbers (435/442 → 470/477, both off
                by +35 from R104's original cite). R104's wording is
                structured around quoting peer titles in the same JSX
                `<div>`: each cited line should land on a `title=`
                attribute *inside the same element this comment annotates*.
                The drift broke that locality — `(line 435)` would now
                point into the middle of the truncate-clamp logic above
                the JSX return, not a title attribute at all. Same
                R136/R137/R140/R142 stale-line-ref paradigm, with the
                largest per-comment offset seen so far because all the
                inserts happened in a contiguous block above the row
                JSX rather than scattered. (R143's own anchor block
                added ~13 lines, so the first internal R143 self-cite
                of `line 453/460` was already obsolete by the time the
                file saved — corrected in-place to the post-anchor
                positions, the same R142 same-round-self-drift loop. */}
            <span className="opacity-70 shrink-0">選取上下文：</span>
            <span className="truncate flex-1" title={tooltip}>
              {activeSelection.preview}
            </span>
            {charCount > 0 && (
              <span
                className="opacity-60 shrink-0 tabular-nums"
                title="選取片段字元數（送給模型的完整內容）"
              >
                {charCount.toLocaleString()} 字
              </span>
            )}
            <button
              onClick={() => ws.setSelection(null)}
              title="清除選取上下文"
              className="opacity-50 hover:opacity-100 hover:bg-secondary rounded p-0.5 transition-colors shrink-0"
            >
              <X className="h-3 w-3" />
            </button>
          </div>
        );
      })()}

      {/* Wrap the scroll container so the Jump-to-latest button can be
          absolute-positioned at its bottom-right edge without scrolling
          along with the messages. `min-h-0` is the standard fix for the
          flex-overflow trap (a `flex-1` child with overflow-auto can't
          shrink below its intrinsic content height without it). */}
      <div className="flex-1 min-h-0 relative">
      <div ref={scrollRef} className="absolute inset-0 overflow-auto px-3 py-3 space-y-3 text-sm">
        {ai.messages.length === 0 && !ai.streaming.content.length && (
          <Welcome
            hasApiKey={hasApiKey}
            onOpenSettings={onOpenSettings}
            // Click-to-fill: drops the example into the input and focuses
            // it so the user can tweak before hitting Enter. First-time
            // users otherwise had to read + retype the sample prompt
            // verbatim, which defeats the whole "here's what to try" point.
            //
            // R94 — preserve an in-progress draft. The Welcome panel shows
            // only when `ai.messages.length === 0`, but `draftMemory`
            // (line 61 + 66) keeps the textarea content across mounts: a
            // returning user with a half-composed prompt who clicks an
            // example for reference would otherwise have their text silently
            // destroyed by the unconditional `setDraft(text)`. This is the
            // exact failure mode onModify (line 296-299) already guards
            // against — its comment ("the previous version unconditionally
            // overwrote `draft` with the placeholder template and discarded
            // their carefully-worded feedback") names this bug class. Same
            // file, same store; aligning the two writers' policies. Surface
            // the skip via notify so the click feels intentional — onModify
            // can stay silent because it has other observable effects
            // (removePending, synthetic tool result), but the example click
            // exists *only* to write the textarea, so a silent no-op would
            // read as a broken button.
            onPickExample={(text) => {
              if (draft.trim()) {
                notify('輸入框已有內容，請先清空再點選範例', 'info');
                inputRef.current?.focus();
                return;
              }
              setDraft(text);
              inputRef.current?.focus();
            }}
          />
        )}
        {ai.messages.map((m, i) => (
          <MessageBubble key={i} message={m} />
        ))}
        {/* Show the streaming bubble as soon as the request is in flight, even
            before any chunk has landed. The bubble is the user's "AI is
            working on it" affordance inside the chat itself; tying it to
            `requestId` (set by resetStreaming) instead of `content.length > 0`
            closes the TTFB gap between Send and the first text_delta /
            tool_use_start — that gap is hundreds of ms (longer on cold
            cache / large prompts) and previously left the chat area
            silent under the user's just-sent prompt with no in-bubble
            feedback. The 停止 button + textarea placeholder cover the
            input side, but a user looking at the chat had no signal.
            MessageBubble renders a "思考中…" placeholder when content is
            empty and no tool_use input is mid-stream. */}
        {ai.streaming.requestId !== null && (
          <MessageBubble
            message={{ role: 'assistant', content: ai.streaming.content }}
            streaming
            inProgressToolUse={ai.streaming.inProgressToolUse}
          />
        )}
        {ai.pending.map((p, i, arr) => (
          <DiffPreview
            key={p.id}
            pending={p}
            onApply={() => void onApply(p)}
            onReject={() => void onReject(p)}
            onModify={() => onModify(p)}
            // Apply / Reject would race the live stream — keep them visibly
            // disabled until the current turn finishes.
            actionsDisabled={ai.streaming.requestId !== null}
            // Keyboard shortcuts (Ctrl+Enter / Ctrl+M / Ctrl+Backspace) target
            // the LAST pending; advertise the hotkey only on that card so
            // intermediate cards don't show a shortcut they ignore.
            keyboardActive={i === arr.length - 1}
          />
        ))}
        {ai.error && (
          <div className="border border-destructive/40 bg-destructive/10 text-destructive px-3 py-2 rounded-md text-xs flex gap-2 items-start">
            <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
            <div className="flex-1 min-w-0">
              <div className="break-words">{ai.error}</div>
              {/* Retry only when the last message is a user/tool_result —
                  otherwise there's nothing to re-run against. Disable while
                  another stream is already inflight to avoid double-fire. */}
              {ai.messages.length > 0 && ai.streaming.requestId === null && (
                <button
                  type="button"
                  onClick={() => {
                    // R183 — live RunOptions read at click-time; same
                    // rationale as `send` / `onApply` / `onReject` above.
                    const a = useAI.getState();
                    // R287 — attach .catch. retryLastTurn's sync prelude
                    // (resetStreaming(''), setError(null), drop trailing
                    // assistant message) commits state changes immediately;
                    // a subsequent `await runTurn` reject (most realistically:
                    // sendChatTurn's preload listener registration throwing
                    // when the IPC bridge is mid-rebind / desynced) escapes
                    // via `void` as unhandledrejection AND leaves
                    // `streaming.requestId = ''` permanently — AIPanel's
                    // send / Apply / Reject gates all read non-null and
                    // refuse forever, locking the chat with no visible
                    // error. Catch + finalizeStreaming() clears the
                    // sentinel; notify surfaces the cause. Same idiom as
                    // R210 / R211 / R284 / R286 for sibling user-triggered
                    // async paths.
                    retryLastTurn({
                      toolsEnabled: a.toolsEnabled,
                      model: a.model,
                      maxTokens: a.maxTokens,
                      temperature: a.temperature,
                    }).catch((err) => {
                      // Clear the streaming sentinel so the input gates
                      // re-open and the user can try a fresh prompt or
                      // re-attempt retry.
                      useAI.getState().finalizeStreaming();
                      const msg = err instanceof Error ? err.message : String(err);
                      notify(`重試失敗：${msg}`, 'error');
                    });
                  }}
                  className="mt-1.5 inline-flex items-center gap-1 px-2 py-0.5 rounded border border-destructive/40 bg-background hover:bg-destructive/10"
                >
                  <RefreshCw className="h-3 w-3" />
                  重試
                </button>
              )}
            </div>
            {/* Dismiss × — the error otherwise lingers until the next
                successful turn clears it (orchestrator setError(null) at
                turn start). Common path: user reads "API key invalid",
                opens Settings to fix it, comes back — the angry red tile
                is still here visually claiming something is broken. Or
                they read "rate limited", decide to take a coffee break,
                and the tile sits there for the rest of the session.
                Mirrors the Toaster × pattern (Toaster.tsx:130-136); the
                store's setError(null) is the canonical clear path.
                R121 — Toaster + FileExplorer error banner caught up to
                this button's scope-disclosure: Toaster.tsx:132 now reads
                「關閉通知」 (multi-variant) and FileExplorer.tsx:318 now
                reads 「關閉錯誤訊息」 verbatim (variant-fixed error). */}
            <button
              type="button"
              onClick={() => ai.setError(null)}
              title="關閉錯誤訊息"
              aria-label="關閉錯誤訊息"
              className="shrink-0 -mt-0.5 -mr-1 opacity-60 hover:opacity-100 hover:bg-destructive/15 rounded p-0.5 transition-colors"
            >
              <X className="h-3 w-3" />
            </button>
          </div>
        )}
      </div>
        {/* Jump-to-latest. Visible only when the user has scrolled away from
            the bottom — the auto-pin handles every "I'm already at the
            bottom" case, so the button would be a no-op at best and visual
            noise at worst when shown there. Highlighted with a primary
            background while a stream is in flight so users can see "AI is
            still typing — click here to follow along". */}
        {!pinnedToBottom && (
          <button
            type="button"
            onClick={scrollToLatest}
            title={ai.streaming.requestId !== null ? 'AI 仍在回應，點擊跳到最新內容' : '跳到最新訊息'}
            className={cn(
              'absolute bottom-3 right-3 z-10 h-7 px-2.5 rounded-full shadow-md border text-xs inline-flex items-center gap-1 transition-colors',
              ai.streaming.requestId !== null
                ? 'bg-primary text-primary-foreground border-primary hover:bg-primary/90 animate-pulse'
                : 'bg-background hover:bg-accent border-border',
            )}
          >
            <ChevronDown className="h-3.5 w-3.5" />
            跳到最新
          </button>
        )}
      </div>

      <InputBar
        inputRef={inputRef}
        draft={draft}
        setDraft={setDraft}
        onSend={send}
        disabled={!hasApiKey}
        streaming={ai.streaming.requestId !== null}
        onCancel={cancelInflight}
      />
    </div>
  );
}

function Header({ hasApiKey, onOpenSettings }: { hasApiKey: boolean; onOpenSettings: () => void }) {
  const ai = useAI();
  // Two-step confirm for "新對話" — destructive (drops messages, pending,
  // streaming partial, conversationId; no undo) and the button is small and
  // sits next to non-destructive controls. Same shape as Round 29's Replace
  // All guard: first click stages, button flips to amber + confirm copy,
  // 4-second auto-cancel timer. Without confirmation, an accidental click
  // mid-conversation discards potentially-long context users were curating.
  const [confirmClear, setConfirmClear] = useState(false);
  const confirmTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (confirmTimer.current) clearTimeout(confirmTimer.current);
    },
    [],
  );
  // Cancel staged confirm whenever the conversation state shifts under the
  // button — a fresh streamed token / new pending mid-confirm would mean
  // the user's "click to confirm" is no longer aimed at the same chat they
  // saw a second ago.
  useEffect(() => {
    if (!confirmClear) return;
    setConfirmClear(false);
    if (confirmTimer.current) clearTimeout(confirmTimer.current);
    // confirmClear intentionally not in deps — see Round 29 rationale.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ai.messages, ai.pending, ai.streaming.requestId]);

  const hasContent =
    ai.messages.length > 0 || ai.pending.length > 0 || ai.streaming.content.length > 0;
  const streaming = ai.streaming.requestId !== null;
  // Disable while streaming — the in-flight request would still resolve,
  // try to land tokens / a tool_use into a freshly-cleared store, and orphan
  // a conversationId. Same defence as workspace.ts cancel-then-clear.
  const clearDisabled = !hasContent || streaming;

  const onClearClick = () => {
    if (clearDisabled) return;
    if (!confirmClear) {
      setConfirmClear(true);
      if (confirmTimer.current) clearTimeout(confirmTimer.current);
      confirmTimer.current = setTimeout(() => setConfirmClear(false), 4000);
      return;
    }
    setConfirmClear(false);
    if (confirmTimer.current) clearTimeout(confirmTimer.current);
    cancelInflight({ persistPartial: false });
    ai.clear();
  };

  return (
    <div className="flex items-center justify-between px-3 h-10 border-b">
      {/* Ctrl+L is bound at menu.ts:176-177 → 'menu:focusAI' →
          dispatchEvent('gendoc:focusAI') (App.tsx:583 / 609) → focus the
          input textarea (AIPanel.tsx:87-89). The keystroke jumps the caret
          straight into the prompt box from anywhere — including the editor
          surface, FileExplorer, or even when the AI panel itself doesn't
          have focus — but currently the only place this shortcut surfaces
          is the native AI submenu, which most users never open. Sibling
          window-level shortcuts (Ctrl+B explorer toggle at App.tsx:1184,
          Ctrl+Shift+F focus mode at App.tsx:1296, Ctrl+, settings at
          App.tsx:1307) all have toolbar tooltips advertising them — Ctrl+L
          is the only one without any in-app hint. Surface it on the panel
          title so users hovering "AI 助手" to confirm what this pane is
          discover the shortcut at the same time, the same way the section
          headers in our other panes (FileExplorer's project label, etc.)
          double as documentation surfaces. */}
      <div
        className="flex items-center gap-2 text-sm font-medium"
        title="AI 助手 — 從任何位置按 Ctrl+L 即可聚焦至下方輸入框"
      >
        <Bot className="h-4 w-4" />
        <span>AI 助手</span>
      </div>
      <div className="flex items-center gap-2">
        {/* Tooltip parity with sibling controls — the Wrench (line 667) and
            Clear (line 681) buttons in this same row both carry `title`, but
            the model `<select>` previously had none, so a hover-to-discover
            user got no hint that this dropdown is the model switcher (the
            label inside reads e.g. "Sonnet 4.6" — a model name, not a verb).
            setModel only flips state.model (store/ai.ts:102); it does NOT
            reset conversationId or messages, so switching mid-thread is
            safe — the next turn just goes to the new model with the same
            history. The tooltip says "切換模型" rather than "重新開始對話"
            to reflect that. */}
        <select
          value={ai.model}
          onChange={(e) => ai.setModel(e.target.value)}
          title="切換 AI 模型（不會清空目前的對話）"
          className="text-xs bg-transparent border rounded px-1 py-0.5"
        >
          {SUPPORTED_MODELS.map((m) => (
            <option key={m.id} value={m.id}>
              {m.label}
            </option>
          ))}
        </select>
        <button
          onClick={() => ai.setToolsEnabled(!ai.toolsEnabled)}
          className={cn(
            'p-1 rounded',
            ai.toolsEnabled ? 'text-primary' : 'text-muted-foreground',
          )}
          // Verb-first tooltip to match every other toggle in the app:
          // DocxEditor:1960 / PptxEditor:1137 / MarkdownToolbar:234 大綱、
          // SettingsDialog:241 API key、App.tsx:1171 檔案總管 — they all
          // describe what clicking *does*, not what state the button is
          // currently in. The previous "工具呼叫已啟用 / 已停用" was a
          // pure declarative state read; users hovering to confirm "what
          // happens if I click this?" got told their current state instead.
          // The icon's colour already communicates state (text-primary
          // when on, text-muted-foreground when off — see className above)
          // so the tooltip's job is the click-action, not a redundant
          // state echo. Parenthetical context kept on both sides so a
          // user new to the feature understands what tool calls are.
          title={
            ai.toolsEnabled
              ? '停用工具呼叫（改為純對話模式）'
              : '啟用工具呼叫（讓 AI 透過工具直接修改檔案）'
          }
          aria-label={ai.toolsEnabled ? '停用工具呼叫' : '啟用工具呼叫'}
          // R153 — toggle-state SR exposure. The dynamic `aria-label` above
          // already tells SR users「what clicking will do next」, but a
          // sighted user also reads the icon's text-primary tint as「目前
          // 開啟」 — and SR users had no parallel for the *current* state.
          // `aria-pressed` is the canonical toggle semantic; engines announce
          // 「按下」 / 「未按下」 alongside the label, completing the parity
          // with the sighted colour-state cue. Same pattern landed on the
          // four ToolbarBtn definitions across MarkdownToolbar / DocxEditor /
          // PptxEditor / XlsxEditor in this round, plus the檔案總管 toggle
          // in App.tsx — making this the canonical app-wide toggle attribute.
          aria-pressed={ai.toolsEnabled}
        >
          <Wrench className="h-3.5 w-3.5" />
        </button>
        <button
          onClick={onClearClick}
          disabled={clearDisabled}
          className={cn(
            'p-1 rounded transition-colors',
            confirmClear
              ? 'text-amber-600 bg-amber-500/15 ring-1 ring-amber-500/40'
              : 'text-muted-foreground hover:text-foreground',
            clearDisabled && 'opacity-30 pointer-events-none',
          )}
          title={
            clearDisabled
              ? streaming
                ? '請先等串流完成（或按下方停止）再清空'
                : '對話已是空的'
              : confirmClear
                ? '再次點擊以確認清空對話歷史（含未套用變更）'
                : '新對話：清空目前的對話歷史'
          }
          // R153 — same icon-only-button accessible-name parity that R152
          // closed for the four close-X buttons. The Wrench sibling at
          // line 752-778 (just above) already pairs `title` with
          // `aria-label`, with the exact convention「title 帶 parenthetical
          // 與動態副本，aria-label 只承載動詞」 documented at lines 758-768.
          // This 新對話 button was the lone holdout in the same Header row.
          // Static label here (vs. the four-branch title above) because SR
          // users can pick up the disabled-state and confirm-staged state
          // from `aria-disabled` (`disabled` on line 781 propagates) and
          // the visual amber ring's text equivalent is the toast cue from
          // ai.clear() — leading aria-label with a single canonical verb
          // matches AIPanel.tsx:920「複製訊息」 and SettingsDialog.tsx:317
          // -318「顯示/隱藏 API key」 dynamism-elsewhere conventions.
          aria-label="新對話"
        >
          <MessageSquarePlus className="h-3.5 w-3.5" />
        </button>
        {!hasApiKey && (
          <Button size="sm" variant="outline" onClick={onOpenSettings}>
            設定 API key
          </Button>
        )}
      </div>
    </div>
  );
}

/**
 * Concatenate the plain-text portion of a message's content blocks for the
 * copy-to-clipboard button. We deliberately skip `tool_use` / `tool_result`
 * blocks: their JSON / truncated payloads are debugging fodder, not what a
 * user pasting an AI reply elsewhere expects to land on the clipboard.
 * Returns the joined string trimmed of trailing whitespace; empty string
 * means there's nothing meaningful to copy and the caller should hide the
 * button entirely.
 */
function extractCopyableText(message: ChatMessage): string {
  return message.content
    .filter((b): b is Extract<ContentBlock, { type: 'text' }> => b.type === 'text')
    .map((b) => b.text)
    .join('')
    .replace(/\s+$/, '');
}

function MessageBubble({
  message,
  streaming,
  inProgressToolUse,
}: {
  message: ChatMessage;
  streaming?: boolean;
  inProgressToolUse?: { id: string; name: string } | null;
}) {
  const isUser = message.role === 'user';
  const isTool = message.role === 'tool_result';
  const [copied, setCopied] = useState(false);
  // Only expose a copy affordance once a stable text payload exists. During
  // streaming the partial text is unstable (mid-sentence, model may revise);
  // tool_result rows are decorative traces, not messages; pure tool_use
  // bubbles have no copyable plain-text. extractCopyableText handles the
  // last filter — empty result = hide.
  const copyableText = !streaming && !isTool ? extractCopyableText(message) : '';
  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(copyableText);
    } catch {
      // Same fallback shape as ErrorBoundary.copyDiagnostics — clipboard API
      // can fail in rare unfocused-document edge cases. Silent ok, the Check
      // icon below would lie if the OS rejected the copy, so guard it.
      const ta = document.createElement('textarea');
      ta.value = copyableText;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand('copy');
      } catch {
        document.body.removeChild(ta);
        return;
      }
      document.body.removeChild(ta);
    }
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1500);
  };

  if (isTool) {
    // Tool results render as a small inline trace; not really a "message".
    return (
      <div className="flex gap-2 text-[11px] text-muted-foreground italic px-1">
        <FileSearch className="h-3 w-3 mt-0.5 shrink-0" />
        <div className="flex-1 min-w-0 truncate">
          {message.content.map((b, i) => (
            <BlockRender key={i} block={b} />
          ))}
        </div>
      </div>
    );
  }

  return (
    <div className={cn('group flex gap-2', isUser && 'flex-row-reverse')}>
      <div
        className={cn(
          'shrink-0 h-6 w-6 rounded-full flex items-center justify-center text-[11px]',
          isUser ? 'bg-primary text-primary-foreground' : 'bg-sky-500/15 text-sky-600',
        )}
      >
        {isUser ? <User className="h-3 w-3" /> : <Bot className="h-3 w-3" />}
      </div>
      <div
        className={cn(
          'relative min-w-0 max-w-[calc(100%-2.5rem)] px-3 py-2 rounded-lg whitespace-pre-wrap break-words text-sm',
          isUser
            ? 'bg-primary text-primary-foreground rounded-tr-sm'
            : 'bg-secondary/60 rounded-tl-sm',
        )}
      >
        {message.content.map((b, i) => (
          <BlockRender key={i} block={b} />
        ))}
        {inProgressToolUse && <InProgressToolBadge name={inProgressToolUse.name} />}
        {/* Hover-revealed copy affordance. Anchored to the bubble's outer
            corner away from the avatar (top-right for assistant, top-left
            for user since the row is reversed). Without this, copying an AI
            reply meant manual select-drag — common ask now that the model
            often emits paste-worthy snippets (regex, formulas, prose). The
            opacity-0 → group-hover:opacity-100 transition matches the tab
            close button's "appears when you mean to act on it" pattern, so
            it doesn't visually crowd the read-only chat history. */}
        {copyableText && (
          <button
            type="button"
            onClick={handleCopy}
            title={copied ? '已複製' : '複製訊息文字'}
            aria-label="複製訊息"
            className={cn(
              'absolute top-1 h-5 w-5 rounded-md flex items-center justify-center opacity-0 group-hover:opacity-100 focus:opacity-100 transition-opacity',
              isUser
                ? 'left-1 bg-primary-foreground/15 hover:bg-primary-foreground/25 text-primary-foreground'
                : 'right-1 bg-background/70 hover:bg-background border border-border/50 text-muted-foreground hover:text-foreground',
            )}
          >
            {copied ? <Check className="h-3 w-3" /> : <Copy className="h-3 w-3" />}
          </button>
        )}
        {/* Pre-first-chunk "thinking" placeholder. Only when the bubble has
            literally nothing else to show — once any text or tool input
            arrives we hand off to the cursor pulse below so the user can
            read what's landing. Three bouncing dots is the genre-standard
            "AI is preparing a reply" affordance (Claude.ai / ChatGPT). */}
        {streaming && message.content.length === 0 && !inProgressToolUse && (
          <span className="inline-flex items-center gap-1.5 text-muted-foreground italic text-xs">
            <span className="inline-flex gap-0.5">
              <span className="h-1 w-1 rounded-full bg-muted-foreground/70 animate-bounce [animation-delay:-0.3s]" />
              <span className="h-1 w-1 rounded-full bg-muted-foreground/70 animate-bounce [animation-delay:-0.15s]" />
              <span className="h-1 w-1 rounded-full bg-muted-foreground/70 animate-bounce" />
            </span>
            思考中…
          </span>
        )}
        {streaming && (message.content.length > 0 || inProgressToolUse) && (
          <span className="opacity-50 animate-pulse">▌</span>
        )}
      </div>
    </div>
  );
}

/**
 * Live placeholder for a tool_use whose `input` JSON is still streaming in.
 * Anthropic emits `tool_use_start` (name only) → many `tool_use_input_delta`
 * chunks → `tool_use_complete` (full input). For tools with large inputs the
 * gap between start and complete can be several seconds, during which the
 * normal stream cursor would just blink with no other movement. This badge
 * tells the user "the model is still typing the tool args" so the stream
 * doesn't look frozen.
 */
function InProgressToolBadge({ name }: { name: string }) {
  return (
    <div className="mt-2 inline-flex items-center gap-1.5 text-[11px] font-mono px-2 py-1 rounded-md bg-background/50 border border-amber-500/40 ring-1 ring-amber-500/30 animate-pulse max-w-full">
      <Wand2 className="h-3 w-3 text-amber-500 shrink-0" />
      <span className="text-foreground/80 shrink-0">{name}</span>
      <span className="text-muted-foreground border-l border-border/50 pl-1.5">準備中…</span>
    </div>
  );
}

/**
 * Compact "what did the AI just do?" string for a tool_use block. Picks the
 * single most location-defining argument so the badge reads like
 * `excel_set_cell · Sheet1!B7` instead of just `excel_set_cell`. The full
 * input JSON is still exposed via `title` for debugging. Returns null when
 * the input has no recognizable identifier (rare — usually means a stub tool
 * with only `tabId`).
 */
function summarizeToolInput(input: Record<string, unknown>): string | null {
  const sheet = typeof input.sheet === 'string' ? input.sheet : '';
  const withSheet = (loc: string) => (sheet ? `${sheet}!${loc}` : loc);
  if (typeof input.address === 'string') return withSheet(input.address);
  if (typeof input.startAddress === 'string') return withSheet(input.startAddress);
  if (typeof input.range === 'string') return withSheet(input.range);
  // slideIndex / paragraphIndex / afterIndex are all schema-described as
  // 0-based (see ai/tools/index.ts), but every user-facing surface — the
  // dispatcher's post-Apply description (`第 1 段`), PPTX's slide rail
  // (`Slide 1`), markdown line numbers — presents 1-based. Mirror the
  // slideIndex+1 convention here so the tool_use badge agrees with the
  // "Applied:" message that follows it; otherwise the same edit reads as
  // `¶0` in the badge and `第 1 段` in the description, looking like a
  // mismatch the user has to mentally reconcile.
  if (typeof input.slideIndex === 'number') return `Slide ${input.slideIndex + 1}`;
  if (typeof input.heading === 'string') return `#${input.heading}`;
  if (typeof input.paragraphIndex === 'number') return `¶${input.paragraphIndex + 1}`;
  if (typeof input.line === 'number') return `L${input.line}`;
  if (typeof input.afterIndex === 'number') return `after ¶${input.afterIndex + 1}`;
  if (typeof input.rowIndex === 'number')
    return sheet ? `${sheet} row ${input.rowIndex}` : `row ${input.rowIndex}`;
  if (Array.isArray(input.sourceTabIds))
    return `${input.sourceTabIds.length} tabs → ${
      typeof input.destTabId === 'string' ? input.destTabId : '?'
    }`;
  if (typeof input.sourceTabId === 'string') return input.sourceTabId;
  return null;
}

function BlockRender({ block }: { block: ContentBlock }) {
  if (block.type === 'text') return <span>{block.text}</span>;
  if (block.type === 'tool_use') {
    const summary = summarizeToolInput(block.input);
    // Full input on hover for debugging; pretty-printed so it's readable in
    // the native tooltip (browsers wrap on \n).
    const fullInput = JSON.stringify(block.input, null, 2);
    return (
      <div
        className="mt-2 inline-flex items-center gap-1.5 text-[11px] font-mono px-2 py-1 rounded-md bg-background/50 border border-border/50 max-w-full"
        title={fullInput}
      >
        <Wand2 className="h-3 w-3 text-amber-500 shrink-0" />
        <span className="text-foreground/80 shrink-0">{block.name}</span>
        {summary && (
          <span className="text-muted-foreground border-l border-border/50 pl-1.5 truncate">
            {summary}
          </span>
        )}
      </div>
    );
  }
  if (block.type === 'tool_result') {
    // R379 — color-differentiate is_error tool_results. See doc-block at
    // bottom of this branch for full enumeration of error sources.
    // R383 — code-point-aware truncation via slicePreview helper. Inline
    // `.slice(0, 200)` with a separate `length > 200 ? '…' : ''` check
    // could split an emoji's surrogate pair at the 200-char boundary;
    // the displayed result rendered an orphan high surrogate as「�」.
    // slicePreview folds the「truncate + maybe append …」 logic into one
    // call that's also surrogate-safe.
    const text = slicePreview(block.content, 200);
    const isError = block.is_error === true;
    // R379 — `is_error?: boolean` (types/ai.ts:58) is set true by every
    // dispatcher error path (R228 dispatch_threw / R297 / R313 / R314 /
    // R316 collision / R349 substring-bail / R359 block-kind-mismatch /
    // every NOT_IMPLEMENTED / R357 + R358 sibling-redirect / R366
    // pptx_empty), by orchestrator's dangling-cleanup「User moved on」
    // synth, by AIPanel's reject-pending「Rejected by user」 synth, and
    // by workspace.removeTab's orphan「Target tab was closed」 synth.
    return (
      <span
        className={isError ? 'text-destructive/80' : 'opacity-70'}
        title={block.content}
      >
        {isError ? '⊘ ' : '↳ '}
        {text}
      </span>
    );
  }
  return null;
}

function Welcome({
  hasApiKey,
  onOpenSettings,
  onPickExample,
}: {
  hasApiKey: boolean;
  onOpenSettings: () => void;
  onPickExample?: (text: string) => void;
}) {
  const examples = [
    { tag: 'Markdown', text: '把『產品介紹』那節改成更精煉的版本' },
    { tag: 'Excel', text: '在 Sheet1 的 B7 寫入 1234' },
    { tag: 'Word', text: '把第 3 段改成「本季業績超越目標 20%」' },
    { tag: 'PowerPoint', text: '把第 2 張投影片的「TBD」改成「2026 Q1 上線」' },
    { tag: '跨檔', text: '讀 budget.xlsx 然後在 notes.md 新增摘要章節' },
  ];
  return (
    <div className="space-y-3">
      <div className="space-y-1.5">
        <div className="flex items-center gap-1.5 text-sm font-medium text-foreground">
          <Bot className="h-4 w-4 text-sky-500" />
          AI 助手
        </div>
        <p className="text-xs text-muted-foreground">
          描述你想要的修改，AI 會用工具呼叫提出變更，按下『套用』才會生效。
        </p>
      </div>
      <div className="space-y-1.5">
        <div className="text-[10px] uppercase tracking-wider text-muted-foreground">
          範例（點擊填入輸入框）
        </div>
        <ul className="space-y-1.5">
          {examples.map((ex, i) => (
            <li key={i}>
              <button
                type="button"
                onClick={() => onPickExample?.(ex.text)}
                // Whole row clickable; left-aligned text mirrors the
                // textarea below so users see "click → fill → tweak" as
                // one continuous gesture. No `disabled` even without an
                // API key — they can still preview the prompt and the
                // settings CTA right below covers the actual gate.
                className="w-full text-left text-xs px-2.5 py-1.5 rounded-md bg-secondary/60 border border-border/50 hover:bg-secondary hover:border-border transition-colors"
                // R113 — tooltip carries info beyond the section heading
                // AND warns about the draft-protection precondition.
                // Prior shape was `title="點擊填入輸入框"` — verbatim
                // duplicate of the section heading's parenthetical 14
                // lines above (line 1050: 「範例（點擊填入輸入框）」).
                // Hovering returned exactly what the eye had already
                // read — same zero-info-on-hover reading the R108 fix
                // closed for ShapePicker (tooltip = visible label) and
                // R112 closed for 新增文字框, but here the redundancy is
                // with a section heading instead of the visible label.
                //
                // Tooltip-vs-actual-behavior honesty: onPickExample at
                // line 501-509 conditionally rejects the click. If
                // `draft.trim()` is non-empty, setDraft is never called
                // and the user gets a toast「輸入框已有內容，請先清空
                // 再點選範例」 instead. That's a load-bearing
                // precondition the codebase already protects — the
                // comment at line 488-500 explicitly cites the parallel
                // `onModify` writer (line 296-299) and frames this as
                // "aligning the two writers' policies" against an
                // earlier bug-class where unconditional overwrite
                // discarded users' carefully-worded feedback. Yet the
                // tooltip never propagated that rule, leaving "click is
                // a no-op when input has content" as an empirical-
                // discovery surprise.
                //
                // Wording mirrors the toast's vocabulary verbatim
                // (「輸入框已有內容」 → 「若輸入框已有內容，會提示先
                // 清空」) so a user who saw the toast once recognises the
                // tooltip phrasing on later hovers, and shifts the verb
                // from generic「填入輸入框」to specific「填入此範例到
                // 輸入框」 — disambiguating between many same-section
                // buttons (the section heading uses the generic phrasing
                // because it's a one-time legend; per-button tooltips
                // benefit from the「此」demonstrative). */
                title="填入此範例到輸入框（若輸入框已有內容，會提示先清空）"
              >
                <span className="text-[10px] font-mono mr-1.5 px-1 rounded bg-background/60 text-muted-foreground">
                  {ex.tag}
                </span>
                {ex.text}
              </button>
            </li>
          ))}
        </ul>
      </div>
      {!hasApiKey && (
        <Button size="sm" variant="outline" className="w-full mt-2" onClick={onOpenSettings}>
          先去設定 Anthropic API key
        </Button>
      )}
    </div>
  );
}

function InputBar({
  inputRef,
  draft,
  setDraft,
  onSend,
  disabled,
  streaming,
  onCancel,
}: {
  inputRef: React.RefObject<HTMLTextAreaElement>;
  draft: string;
  setDraft: (s: string) => void;
  onSend: () => void;
  disabled: boolean;
  streaming: boolean;
  onCancel: () => void;
}) {
  // Auto-grow the textarea up to ~8 lines while composing. Two static rows
  // were fine for one-liners but anything pasted (a stack trace, a multi-
  // paragraph instruction) gets clipped to 2 visible lines and forces the
  // user to scroll inside the input — easy way to lose track of what they're
  // about to send. We measure on every change: reset height, then snap to
  // scrollHeight clamped to a sane max.
  //
  // Floor-via-CSS: the auto-grow's `el.style.height = 'auto' → scrollHeight`
  // pass overrides the textarea's `rows` attribute, so an empty draft would
  // collapse the input to a single line — visible as a vertical jump every
  // time the user hits Enter to send (textarea shrinks, then grows back as
  // they type the next prompt). The CSS `min-h-` class wins over inline
  // height at layout time, so we let inline height carry the upper bound
  // (auto-grow) while the floor stays anchored at 2 rows.
  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    el.style.height = 'auto';
    const max = 200; // px — ≈8 rows at the current font size
    el.style.height = `${Math.min(el.scrollHeight, max)}px`;
  }, [draft, inputRef]);

  return (
    <div className="border-t p-2 flex gap-2 items-end">
      <textarea
        ref={inputRef}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        // R156 — explicit accessible name. Same rationale as R155's
        // FindReplaceDialog 尋找 / 取代 inputs: dynamic placeholder
        // (line 1244-1250 flips between「請先設定 API key…」/「AI 回應中…」/
        // the default prompt hint based on disabled / streaming state)
        // disappears the moment the user types, leaving SR users a blank
        // edit field with no anchor for "what do I type here?". The
        // aria-label stays static and content-anchored (「AI 提示詞」),
        // mirroring the role this textarea plays in the panel — the
        // same semantic the parent Header「AI 助手」 announces, but here
        // labeling the *input* of that subsystem rather than the panel.
        aria-label="AI 提示詞"
        onKeyDown={(e) => {
          // Plain Enter sends; Shift+Enter inserts a newline. Skip when
          // any modifier is held — Ctrl+Enter is reserved for "Apply
          // pending change" globally, and Alt+Enter is a future-proof
          // hook. Skip during IME composition so CJK candidate-confirm
          // (Enter) doesn't accidentally fire the prompt.
          //
          // While streaming, `send()` silently no-ops (mid-turn Enter
          // would orphan the inflight handle — see the early-return at
          // the top of `send`). Fall through to the textarea's default
          // behaviour so Enter inserts a newline as users keep drafting
          // their next prompt — preventing default here would swallow
          // the keystroke, leaving the user staring at a frozen input
          // with no feedback.
          if (
            e.key === 'Enter' &&
            !e.shiftKey &&
            !e.ctrlKey &&
            !e.metaKey &&
            !e.altKey &&
            !e.nativeEvent.isComposing &&
            !streaming
          ) {
            e.preventDefault();
            onSend();
          }
          // Esc-to-cancel while streaming. Mirrors Claude.ai / ChatGPT — the
          // user is typically already typing the next prompt while waiting,
          // so the keystroke is right under their hands; mousing to the 停止
          // button on the side mid-thought is the awkward path. Scoped to
          // the textarea (vs. document-level) so it doesn't conflict with
          // editor Esc handlers (cell-edit cancel, find-dialog dismiss,
          // present-mode exit) and only fires when the user has explicit
          // focus here. IME-composition guard mirrors the Enter path: Esc
          // is the standard CJK candidate-cancel keystroke, so during
          // composition we leave the IME alone and fall through.
          if (
            e.key === 'Escape' &&
            streaming &&
            !e.nativeEvent.isComposing
          ) {
            e.preventDefault();
            onCancel();
          }
        }}
        disabled={disabled}
        // R124 — placeholder punctuation harmonized across the streaming
        // / normal pair. Previously the streaming branch used full-width
        // 「（）」 and 「、」 while the normal branch used half-width
        // 「()」 and 「,」 — same input, same kind of compound shortcut
        // hint, but a user who fires a turn watches the punctuation style
        // flip mid-session as the ternary swaps placeholders. The full-
        // width form is the dominant convention for CJK-context hints in
        // this same component (line 752 「停用工具呼叫（改為純對話模式）」,
        // line 1099 「填入此範例到輸入框（若輸入框已有內容，會提示先
        // 清空）」, and the streaming sibling above), so the normal branch
        // is the lone half-width holdout. Aligning to full-width also lets
        // the keystroke separator 「、」 match the streaming branch
        // verbatim, so two adjacent hints read in one voice.
        placeholder={
          disabled
            ? '請先設定 API key…'
            : streaming
              ? 'AI 回應中… Esc 停止、Enter 換行（亦可按右側停止）'
              : '描述你想要的修改（Enter 送出、Shift+Enter 換行）'
        }
        className="flex-1 resize-none bg-secondary/40 rounded-md px-2 py-1.5 text-sm outline-none focus:ring-1 focus:ring-ring min-h-[3.25rem]"
      />
      {streaming ? (
        // Mirror of the Send button below: same input row, same pixel slot,
        // mutually exclusive with Send via the `streaming` ternary. The Send
        // half carries `title="送出 (Enter)"` (line ~1162) advertising the
        // keystroke its keymap binds (Enter at line 1102-1113); the Stop
        // half also has a keystroke wired — Esc at line 1124-1131 fires
        // onCancel() — but previously surfaced no tooltip, so the binding
        // was reachable only via the placeholder string at line 1138-1139
        // ("AI 回應中… Esc 取消") which the user only sees when the input
        // is empty *and* focused. Once they type a draft (the common case
        // — users keep composing the next prompt while AI is responding),
        // the placeholder is replaced by the draft text and the Esc hint
        // vanishes; hovering the visible 停止 button at that point would
        // be the natural fallback discovery surface, but it was silent.
        // Tooltip mirrors Send's "verb (key)" shape exactly. The label
        // already says "停止" so we don't need an aria-label for screen
        // readers — the title's only job here is the keystroke surface.
        <Button
          size="sm"
          variant="outline"
          // R391 — explicit no-arg invocation. `onClick={onCancel}` (where
          // onCancel === cancelInflight) silently passes React's MouseEvent
          // as the `opts: { persistPartial?: boolean }` parameter. It works
          // today only by coincidence: MouseEvent has no `persistPartial`
          // property, so `opts.persistPartial !== false` evaluates true and
          // the default persist-partial branch fires — same as `cancelInflight()`
          // with no args. But that's a landmine for any future field added
          // to the opts shape: a hypothetical `opts.preserveError?: boolean`
          // (or anything whose name happens to alias a MouseEvent property)
          // would silently flip behavior based on whether the user clicked
          // 停止 vs invoked the function programmatically. The Esc-key path
          // at AIPanel.tsx:1742 already uses `onCancel()` (explicit no-args
          // — Esc handlers don't have a useful event to forward); the click
          // path should match. Same idiom as line 1159's
          // `cancelInflight({ persistPartial: false })` — explicit option
          // shape, no React-event leakage. Pure defensive refactor, no
          // behavior change today.
          onClick={() => onCancel()}
          title="停止 (Esc)"
        >
          停止
        </Button>
      ) : (
        // The neighbouring 停止 button has visible label text so it's
        // self-describing; the Send button is icon-only and was previously
        // surfacing nothing to screen readers / hover-tooltip users. Adding
        // both `title` (mouse / native tooltip) and `aria-label` (assistive
        // tech) brings it in line with the toolbar buttons elsewhere.
        //
        // R88 — disabled-state tooltip. The button has TWO disable reasons
        // (no API key OR empty draft) but the static title「送出 (Enter)」
        // explained neither. The sibling textarea at line 1134-1140 already
        // does the right thing: its `placeholder` flips to「請先設定 API
        // key…」on the no-key path. Mirror that exact wording when
        // disabled-by-no-key so the row speaks with one voice; for
        // disabled-by-empty-draft show a hint that pairs with the textarea's
        // normal-state placeholder「描述你想要的修改 (Enter 送出, …)」 —
        // i.e. tell the user *why* Enter isn't firing right now. The
        // streaming case isn't a third branch because this button isn't
        // rendered when streaming (Stop button at line 1160 takes the slot).
        <Button
          size="sm"
          onClick={onSend}
          disabled={disabled || !draft.trim()}
          // Same verb as the textarea placeholder ("Enter 送出") — mixing
          // 送出 / 發送 in the same row reads as two different actions to
          // CJK users at a glance, even though they're the same intent.
          // orchestrator.ts:201 also uses 送出, so this is the dominant
          // verb across the AI flow's user-facing strings.
          title={
            disabled
              ? '請先設定 API key…'
              : !draft.trim()
                ? '輸入提示後即可送出 (Enter)'
                : '送出 (Enter)'
          }
          aria-label="送出"
        >
          <Send className="h-3.5 w-3.5" />
        </Button>
      )}
    </div>
  );
}

function ResizeHandle({ onChange }: { width: number; onChange: (w: number) => void }) {
  return (
    <div
      className="absolute -left-0.5 top-0 bottom-0 w-1 cursor-col-resize hover:bg-primary/40 active:bg-primary/60 z-10 transition-colors"
      onMouseDown={(e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = (e.currentTarget.parentElement as HTMLElement).offsetWidth;
        // Lock the body cursor + disable text selection for the duration
        // of the drag. The handle itself is only 1 px wide and a fast
        // drag (or one that hits the min/max clamp) routinely takes the
        // pointer outside the handle — without this lock, the cursor
        // flickers to the I-beam over editor text or the default pointer
        // over buttons, and text under the cursor can become selected.
        // Save/restore the previous values so we don't trample any host
        // styles set elsewhere.
        const prevCursor = document.body.style.cursor;
        const prevUserSelect = document.body.style.userSelect;
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        const onMove = (ev: MouseEvent) => {
          const next = Math.max(280, Math.min(600, startWidth + (startX - ev.clientX)));
          onChange(next);
        };
        const onUp = () => {
          document.body.style.cursor = prevCursor;
          document.body.style.userSelect = prevUserSelect;
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      }}
    />
  );
}
