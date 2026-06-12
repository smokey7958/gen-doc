/**
 * High-level orchestrator: drives one user prompt → assistant turn(s) loop
 * with optional tool calls, until the assistant says message_stop without
 * any pending tool_use that the user hasn't resolved.
 *
 * Spec §6.2 step [1]–[11]. Apply / Reject is user-driven via the AIPanel
 * buttons; this module just emits PendingChange entries.
 */

import { v4 as uuid } from 'uuid';
import type {
  ChatMessage,
  ContentBlock,
  ToolDefinition,
  ToolResultBlock,
  ToolUseBlock,
} from '../types/ai';
import { useAI, type PendingChange } from '../store/ai';
import { useWorkspace } from '../store/workspace';
import { sendChatTurn, type InflightHandle } from './provider';
// R415 — dispatcher is loaded on demand (first tool call) instead of
// statically: it pulls the whole OOXML adapter graph (docx, xlsx /
// xlsx-js-style, jszip, mammoth) into its chunk, which would otherwise
// sit in the renderer's startup bundle. Only types are imported eagerly.
import type { DispatchResult } from './dispatcher';
import { TOOLS } from './tools';
import { SYSTEM_PROMPT } from './system-prompt';

interface RunOptions {
  /** When false, tools are not advertised to the model. */
  toolsEnabled: boolean;
  model: string;
  maxTokens: number;
  temperature: number;
  /** Optional selection summary appended to the user message as context. */
  selectionContext?: string;
}

/**
 * Send a single user prompt and resolve when the model has stopped emitting
 * tokens (one full assistant turn — which may include tool_use blocks that
 * become PendingChange entries the user must resolve).
 */
export async function sendUserPrompt(prompt: string, opts: RunOptions): Promise<void> {
  const ai = useAI.getState();
  const ws = useWorkspace.getState();
  // R179 — capture workspaceId for swap-detection across the dangling-cleanup
  // loop and the userMsg persist. Same shape as R178 (handleAssistantToolCalls)
  // and continueAfterToolResult above. Trigger: user sends a new prompt that
  // fires dangling cleanup of N stale tool_uses (each is its own
  // pushSyntheticToolResult → persistMessage IPC), then Ctrl+O during the
  // loop. Without this guard, the loop's later iterations push OLD turn's
  // synthetic tool_results into NEW workspace's chat, and the userMsg push
  // + runTurn that follow apply OLD's prompt to NEW's outline.
  const wsAtStart = ws.workspaceId;
  const stillSameWs = () => useWorkspace.getState().workspaceId === wsAtStart;

  // Clear any leftover error banner from the previous turn the moment the
  // user commits to a new prompt. `runTurn` → `resetStreaming` would zero
  // it eventually, but only after we await pushSyntheticToolResult (per
  // dangling tool_use) and persistMessage — that window leaves the
  // "Retry" button alongside the user's just-sent prompt, suggesting the
  // *old* failure is still actionable while a fresh prompt is already in
  // flight. Same rationale as retryLastTurn's explicit setError(null).
  if (ai.error) ai.setError(null);

  // Auto-resolve any tool_uses still hanging from the previous assistant
  // turn before we push the user's new message. Anthropic rejects the next
  // call with 400 "tool_use ids were found without tool_result blocks" if
  // the assistant message contains tool_uses that no later message answers
  // — and the user has already moved on (they're typing a new prompt), so
  // expecting them to click Apply/Reject first is the wrong UX. Push a
  // synthetic tool_result for every dangling id and drop the matching
  // PendingChange cards so stale DiffPreviews don't outlive the turn.
  const dangling = unresolvedToolUseIds(ai.messages);
  if (dangling.length > 0) {
    // R244 — drop the matching PendingChange cards SYNCHRONOUSLY before the
    // synth-tool_result push loop, not after it. Original order pushed
    // synthetic tool_results in an `await` loop, then evicted the pendings
    // in a single setState only after the loop fully completed; if the
    // loop threw partway through (persistMessage IPC reject during
    // sqlite-wal contention / disk full / OS-level antivirus lock —
    // rare but real, and the family R203 / R204 / R210 / R211 already
    // covered for sibling callsites), state was left dangerously
    // inconsistent:
    //   • k synthetic tool_results in `ai.messages` (pushMessage is sync,
    //     fired before the failing await)
    //   • all original pendings still visible as DiffPreview cards
    //     (because the post-loop setState never ran)
    // The user sees an error banner ("DB locked"), retries, but in the
    // meantime can click Apply on any of those stale cards — onApply's
    // dedupe gate (`pending.some(x => x.id === p.id)`) still passes
    // because the pending wasn't removed, and continueAfterToolResult
    // pushes a SECOND tool_result for a tool_use_id that already has a
    // synthetic result. The next assistant turn now has duplicate
    // tool_results for one tool_use_id, and Anthropic rejects with
    // "tool_use_id ... appeared more than once" — a confusing 400 the
    // user can't act on. (Trigger requires multi-pending turn + persist
    // failure mid-cleanup + user clicks Apply on a stale card; rare,
    // but each pre-condition is realistic individually.)
    //
    // Reordering closes the window: removing pendings sync UP FRONT
    // means a stale click during the failing-loop window bounces at
    // onApply's dedupe gate (pending no longer present in store), so
    // no duplicate tool_result can be created. If the loop later
    // throws after k successful pushes, the next sendUserPrompt's
    // own `unresolvedToolUseIds` re-detects the un-synth'd remainder
    // and re-runs the loop for them — partial progress is preserved
    // and recovery is automatic.
    const danglingSet = new Set(dangling);
    useAI.setState({
      pending: useAI.getState().pending.filter((p) => !danglingSet.has(p.toolUseId)),
    });
    for (const id of dangling) {
      if (!stillSameWs()) return;
      await pushSyntheticToolResult(
        id,
        'User moved on to a new request without responding to this proposal.',
        true,
      );
    }
    if (!stillSameWs()) return;
  }

  if (!stillSameWs()) return;

  const userBlocks: ContentBlock[] = [];
  if (opts.selectionContext) {
    userBlocks.push({ type: 'text', text: `[Context]\n${opts.selectionContext}` });
  }
  // Always include the active tab outline so the model has something to work with.
  const outline = buildOutline(ws.tabs, ws.activeTabId);
  if (outline) {
    userBlocks.push({ type: 'text', text: `[Active workspace]\n${outline}` });
  }
  userBlocks.push({ type: 'text', text: prompt });

  const userMsg: ChatMessage = { role: 'user', content: userBlocks };
  useAI.getState().pushMessage(userMsg);

  await persistMessage(userMsg);
  if (!stillSameWs()) return;

  await runTurn(opts);
}

/**
 * Push a synthetic tool_result without running a new model turn. Used by the
 * 修改 (Modify) flow: when the user dismisses a pending change to compose a
 * modification request, we still owe Anthropic a `tool_result` for the
 * assistant's `tool_use` block — the API rejects the next user turn otherwise
 * with "tool_use ids were found without tool_result blocks". We satisfy the
 * invariant immediately, then let the user's next `sendUserPrompt` carry
 * their actual feedback, all in one round-trip to the model.
 */
export async function pushSyntheticToolResult(
  toolUseId: string,
  content: string,
  isError = false,
): Promise<void> {
  const ai = useAI.getState();
  const msg: ChatMessage = {
    role: 'tool_result',
    content: [
      {
        type: 'tool_result',
        tool_use_id: toolUseId,
        content,
        is_error: isError,
      },
    ],
  };
  ai.pushMessage(msg);
  await persistMessage(msg);
}

/** Re-enter the loop after the user has applied / rejected a pending tool result. */
export async function continueAfterToolResult(
  toolUseId: string,
  result: { content: string; isError?: boolean },
  opts: RunOptions,
): Promise<void> {
  const ai = useAI.getState();
  // R179 — capture workspaceId so we can abort if the user swaps mid-flow.
  // Same shape as R178's handleAssistantToolCalls guard. After pushMessage
  // (sync) + persistMessage (async IPC) + runTurn (async IPC), each microtask
  // boundary is a swap window. If the user clicks Apply on workspace A's
  // PendingChange and then immediately Ctrl+O to workspace B before the
  // persistMessage IPC resolves, the post-await `unresolvedToolUseIds` /
  // streaming.requestId checks read NEW workspace's empty (post-clear)
  // state, both gates pass, runTurn fires for OLD's tool_result on NEW's
  // store. Capture pre-IPC, abort post-IPC if swap.
  const wsAtStart = useWorkspace.getState().workspaceId;
  // Apply / Reject is a forward action — the previous turn's error banner
  // (if any) belonged to a turn that's now superseded. Clear it before the
  // tool_result lands so users don't see "Retry" sitting next to the new
  // round-trip. See sendUserPrompt for the full rationale.
  if (ai.error) ai.setError(null);
  const msg: ChatMessage = {
    role: 'tool_result',
    content: [
      {
        type: 'tool_result',
        tool_use_id: toolUseId,
        content: result.content,
        is_error: result.isError,
      },
    ],
  };
  ai.pushMessage(msg);
  await persistMessage(msg);
  // R179 — abort if workspace swapped during persistMessage IPC. The push
  // above already wrote to OLD's messages (or NEW's, if swap landed before
  // pushMessage); either way the unresolved-toolUseIds gate below must
  // read against the same workspace whose tool_use we're answering.
  if (useWorkspace.getState().workspaceId !== wsAtStart) return;
  // Multi-pending case: the assistant emitted ≥2 tool_uses in one turn and
  // ≥2 became PendingChanges. Apply on the first one used to runTurn here
  // immediately, but the sibling pending's tool_use was still unanswered
  // and Anthropic 400'd. Gate the turn on every tool_use from the last
  // assistant message having a matching tool_result somewhere after it —
  // the remaining Apply/Reject calls will trip this same gate and only
  // the final one actually fires runTurn.
  if (unresolvedToolUseIds(useAI.getState().messages).length > 0) return;
  // R163 — re-entrancy gate, sibling to R162's send / retry guards.
  // Realistic trigger: the user rapidly Applies P1 then P2 from the same
  // assistant turn. Both onApply flows run sync up to `await undo.push`,
  // both continueAfterToolResult calls push their tool_result and await
  // persistMessage; by the time both persist resolves, both pushMessages
  // are already in `messages`, so both pass the unresolved-toolUseIds
  // check above and both reach this line. Without a runTurn-side gate,
  // both fire `await runTurn`, both sync preludes set their own uuid into
  // `streaming.requestId` and overwrite `inflight` — the first IPC stream
  // is orphaned, two parallel responses interleave their token chunks
  // into the shared `streaming.content`. Reading live store state here
  // closes the window: whichever continueAfterToolResult reaches `await
  // runTurn` first runs its sync prelude (sets requestId=uuid), the
  // straggler reads uuid via `getState()` and silently returns.
  if (useAI.getState().streaming.requestId !== null) return;
  await runTurn(opts);
}

/**
 * Returns the tool_use ids from the most recent assistant message that
 * have not yet been paired with a tool_result block in any later message.
 * An empty array means the conversation is ready for the next runTurn —
 * Anthropic's API requires every tool_use to be answered before the next
 * assistant turn.
 */
function unresolvedToolUseIds(messages: ChatMessage[]): string[] {
  let lastAssistantIdx = -1;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === 'assistant') {
      lastAssistantIdx = i;
      break;
    }
  }
  if (lastAssistantIdx === -1) return [];
  const expected = new Set<string>();
  for (const b of messages[lastAssistantIdx].content) {
    if (b.type === 'tool_use') expected.add((b as ToolUseBlock).id);
  }
  if (expected.size === 0) return [];
  for (let j = lastAssistantIdx + 1; j < messages.length; j++) {
    for (const b of messages[j].content) {
      if (b.type === 'tool_result') expected.delete((b as ToolResultBlock).tool_use_id);
    }
  }
  return [...expected];
}

/**
 * In-flight slot for the active runTurn. Wraps the IPC handle alongside the
 * runTurn Promise's `resolve` so cancelInflight can unblock awaiters — see
 * R164 doc-block in cancelInflight for the full leaked-Promise scenario.
 */
interface InflightSlot {
  handle: InflightHandle;
  resolve: () => void;
}
let inflight: InflightSlot | null = null;

export function cancelInflight(opts: { persistPartial?: boolean } = {}): void {
  if (!inflight) {
    // R256 — also handle the bridge window between onDone (which sets
    // inflight=null) and the next runTurn (which sets a new inflight).
    // During that window handleAssistantToolCalls runs its `await
    // pushSyntheticToolResult` loop for inline tool_uses. R161's bridge
    // has set streaming.requestId to a fresh uuid so the 停止 button
    // is visible (the InputBar gates `streaming = requestId !== null`),
    // but the early return above silently no-ops every cancel click in
    // that window. Concrete failure: assistant emits a chain of read
    // tool_uses (e.g. "讀 sheet1 的 5 個 cell"), each pushSynthetic
    // ToolResult takes 50-500ms (DB persist). Total dispatch loop
    // window: 250ms-2.5s. User clicks 停止 mid-loop, sees nothing
    // happen, mashes the button — and the dispatch keeps iterating
    // until the loop naturally ends and runTurn fires for the
    // follow-up assistant turn (which DOES respond to cancel because
    // by then inflight is non-null again).
    //
    // Clearing streaming via finalizeStreaming resets requestId to
    // null. handleAssistantToolCalls captures the bridge requestId
    // at entry and checks `stillSameStream()` at each iteration top
    // (R256 sibling) — when requestId changes, the loop bails before
    // pushing more synth results AND before falling through to the
    // post-loop runTurn call. The .then(resolve, ...) chain wired
    // around handleAssistantToolCalls in onDone fires resolve when
    // the function returns (whether normally or via the new bail),
    // so the runTurn Promise still settles per R164's invariant.
    //
    // No persistPartial branch needed: at this point the assistant
    // message was already committed in onDone (line ~474), so
    // streaming.content is empty (the bridge reset it). finalize
    // Streaming with empty content returns null and just resets the
    // slot — same effect as the explicit reset above's normal path.
    if (useAI.getState().streaming.requestId !== null) {
      useAI.getState().finalizeStreaming();
    }
    return;
  }
  // Capture before nulling so we can resolve the runTurn Promise after the
  // store cleanup below.
  const { handle, resolve } = inflight;
  // R384 — defensive try/catch around handle.cancel(). The current provider
  // (provider.ts:83-86) does `void window.gendoc.ai.cancel(requestId); off();`
  // — both nominally non-throwing today. But ANY throw escaping here would
  // skip the four critical lines below:
  //   • `inflight = null` — the slot stays marked-busy forever, the next
  //     `cancelInflight` call falls into the early-return branch (which
  //     also no-ops on a non-null handle that's effectively dead), and
  //     the input bar's「停止」 button stays visible
  //   • `finalizeStreaming()` — partial assistant content isn't committed,
  //     streaming.requestId stays non-null → InputBar's send/retry stay
  //     disabled → the user can't continue
  //   • the .catch/then chain that calls `resolve()` on the runTurn
  //     Promise — the outer `await runTurn(...)` in sendUserPrompt /
  //     retryLastTurn never settles, the calling React component's
  //     try/catch never runs, sendingRef stays true forever
  // All three are progress-blocking. Wrap cancel in try/catch so a
  // future provider impl that DOES throw (network race, IPC bridge
  // refactor, sandboxed-provider rejection) still lets the cleanup
  // path run. Same defensive posture as R210 / R211 / R245 added to
  // user-triggered async paths — convert thrown exceptions into
  // controlled-fall-through so the「app stays usable」 invariant
  // holds even on unexpected internal errors.
  try {
    handle.cancel();
  } catch (err) {
    // Log to console so DevTools shows the unexpected error; do NOT
    // setError on the AI store (the user already clicked 停止, popping
    // an error banner about「停止失敗」 is more confusing than helpful
    // — they probably don't care WHY cancel had a hiccup, they care
    // that the input bar unblocks and they can keep working).
    console.error('cancelInflight: handle.cancel() threw, continuing cleanup', err);
  }
  inflight = null;
  // The provider's cancel() removes the chunk listener before main's
  // abort can deliver the synthetic `message_stop`, so onDone never
  // fires for this turn. Finalize here so:
  //   • any partial assistant text the user already saw is preserved
  //     as a regular message instead of vanishing,
  //   • streaming.requestId clears, so the input bar unblocks (送出 /
  //     重試 / Enter all gated on `streaming.requestId === null`).
  // If nothing streamed yet, finalizeStreaming returns null and just
  // resets the streaming slot — same effect, no empty bubble.
  const ai = useAI.getState();
  const finalized = ai.finalizeStreaming();
  // `persistPartial` defaults to true — keep the partial in the user-
  // facing 停止 path, where the user has already seen the half-reply on
  // screen and would be confused if it vanished from history. Workspace
  // switches pass `false`: the AI store is about to be `clear()`ed
  // synchronously, which races with `persistMessage`'s first `await` and
  // would otherwise spawn a phantom conversation in the *new* workspace
  // (conversationId reads as null after clear, persistMessage falls back
  // to creating a brand-new convo for the new workspaceId).
  if (finalized && opts.persistPartial !== false) {
    // R203 — catch handler matches R165's pattern at handleAssistantToolCalls's
    // `.then(resolve, ...)`. Without this, a DB write failure (disk full,
    // SQLite locked by an OS-level backup tool, IPC reject) would surface
    // as an unhandledrejection on the void-promise — visible in DevTools
    // but with no useful action for the user. Pushing into the AI error
    // banner gives the user a recoverable signal: their just-cancelled
    // partial is in-memory but not persisted; same Retry button (for the
    // last turn) is still available. setError mirrors how onError chunks
    // surface other failures (line 397).
    // R237 — capture workspaceId at cancel time so the `.catch` handler
    // doesn't leak OLD workspace's error onto NEW. Trigger: user clicks
    // 停止 (cancel inflight, persistPartial defaults to true), then
    // immediately Ctrl+O to load a different .gd. The persistMessage IPC
    // is in flight; sqlite-wal lock during the swap window rejects;
    // setError fires AFTER swap → writes OLD's exception text into
    // NEW workspace's `ai.error` (NEW's clear() set error to null,
    // R237's late setError overwrites). Same shape as R236's outer-
    // catch guard in handleAssistantToolCalls. The persist itself
    // still tries (the doc-block above acknowledges this races with
    // workspace swap and can spawn a phantom conversation in NEW —
    // existing concern, not addressed by R237); but at least the
    // user-visible error banner stays workspace-correct.
    const wsAtCancel = useWorkspace.getState().workspaceId;
    void persistMessage(finalized).catch((err) => {
      if (useWorkspace.getState().workspaceId !== wsAtCancel) return;
      useAI.getState().setError(
        err instanceof Error ? err.message : String(err),
      );
    });
  }
  // R164 — resolve the runTurn Promise so awaiters unblock. Without this,
  // `runTurn`'s Promise had no path to settle on cancel: provider.cancel()
  // calls `off()` to unregister the chunk listener, so main's synthetic
  // `message_stop` never reaches `onDone`, and `onDone` was the only place
  // that called `resolve()`. The Promise leaked, which silently broke any
  // `try { await runTurn(...) } finally {…}` cleanup at higher layers —
  // most concretely R162's `sendingRef` in AIPanel.send: pressing 停止
  // mid-turn left `sendingRef.current = true` permanently, blocking every
  // subsequent send for the rest of the session. Idempotent — Promise
  // resolve is a no-op on second call (e.g. if a late chunk slipped past
  // off() and fired `onDone`'s closure-bound resolve before the listener
  // was actually removed by Electron's event emitter).
  resolve();
}

/**
 * Re-run the last turn after a transient failure. The user's prompt is still
 * at the tail of `ai.messages` — we only need to discard partial streaming
 * state and re-issue the model call. No new user message is appended; this
 * isn't a fresh send.
 */
export async function retryLastTurn(opts: RunOptions): Promise<void> {
  const ai = useAI.getState();
  if (ai.messages.length === 0) return;
  // R162 — re-entrancy gate. Same-shape race as AIPanel.send: the retry
  // button rerenders away once we set requestId, but a rapid double-click
  // can fire two retryLastTurn calls in close succession (browser dispatches
  // queued click events sequentially after each handler's first await).
  // The second call would: (a) overwrite the assistant-message-drop logic
  // below — its `ai.messages.slice(0, -1)` reads a now-shorter messages
  // array and chops the user prompt the first call preserved, (b) start a
  // second sendChatTurn whose `onChunk` handler appends into the same
  // shared `streaming.content` slot as the first turn, mixing tokens from
  // two parallel responses. Read live store state (not the React snapshot
  // at the click site) so the gate stays correct even if the button's
  // closure was stale. `requestId === ''` after resetStreaming('') below
  // counts as "in flight" — '!== null' covers both the empty-string sentinel
  // and a real uuid.
  if (ai.streaming.requestId !== null) return;
  // Clear any half-streamed partial that the failed turn left behind so the
  // next turn doesn't double up.
  ai.resetStreaming('');
  ai.setError(null);
  // When the failing turn streamed any text before erroring, `onDone` ran
  // `finalizeStreaming` and persisted that broken half as a real assistant
  // message. Re-issuing the turn now would feed the model its own truncated
  // reply (after the user prompt / tool_result that originally triggered
  // it) and produce confused output. Drop the trailing assistant message
  // first — "retry" means "redo from the last input forward".
  // R307 — walk backwards to find the most recent assistant index instead of
  // only checking `messages[last]`. Previous code drop-tail only fired when
  // the very last message was assistant; but `handleAssistantToolCalls` may
  // have pushed synthetic tool_results (inline read_tab_content, or R228's
  // `dispatch_threw:` error tool_result) AFTER finalizeStreaming committed
  // the assistant message and BEFORE the dispatch loop hit R230's persist
  // failure that surfaced the error banner. In that state the message tail
  // is `[..., assistant{read, write}, tool_result(read_inline)]` — last is
  // tool_result, not assistant, so the original check skipped the drop. The
  // unresolved `write` tool_use remained in messages with no matching
  // tool_result, runTurn re-issued the API call, and Anthropic 400'd with
  // "tool_use ids were found without tool_result blocks" — turning Retry
  // into a second error banner the user couldn't escape (every click 400'd
  // identically). Finding the most-recent assistant and slicing from THAT
  // index drops the failed assistant AND its trailing synth tool_results
  // together, restoring the messages tail to the last user/tool_result_
  // from_user input — the canonical "redo from last input forward" state.
  let lastAssistantIdx = -1;
  for (let i = ai.messages.length - 1; i >= 0; i--) {
    if (ai.messages[i].role === 'assistant') {
      lastAssistantIdx = i;
      break;
    }
  }
  if (lastAssistantIdx >= 0) {
    // Any pending changes whose toolUseId points into this dropped
    // assistant message are now orphaned. If the user clicks Apply / Reject
    // on a stale pending, `continueAfterToolResult` pushes a tool_result
    // for a tool_use_id that no longer exists in any message — Anthropic
    // rejects the next turn with "tool_use_id was not referenced in any
    // tool_use block in the previous assistant message". Evict them in the
    // same setState so the orphaned DiffPreview cards disappear with the
    // dropped message. (Realistic trigger: a tool_use streamed to
    // completion, the dispatcher created a Pending, then the network
    // dropped before message_stop; the user retries to get a clean reply.)
    const droppedToolUseIds = new Set(
      ai.messages[lastAssistantIdx].content
        .filter((b): b is ToolUseBlock => b.type === 'tool_use')
        .map((b) => b.id),
    );
    useAI.setState({
      messages: ai.messages.slice(0, lastAssistantIdx),
      pending: ai.pending.filter((p) => !droppedToolUseIds.has(p.toolUseId)),
    });
  }
  await runTurn(opts);
}

async function runTurn(opts: RunOptions): Promise<void> {
  const requestId = uuid();
  const ai = useAI.getState();
  ai.resetStreaming(requestId);

  const messages = ai.messages;
  const cacheBreakpoints = computeCacheBreakpoints(messages);
  const tools: ToolDefinition[] = opts.toolsEnabled ? TOOLS : [];
  // R238 — capture workspaceId at runTurn entry so the async `.catch`
  // handlers below (persistMessage in onDone, handleAssistantToolCalls's
  // .then errFn) don't leak OLD workspace's error message into NEW's
  // banner. Trigger: assistant message finalises, void persistMessage
  // IPC fires, user immediately Ctrl+O. The persist's reject lands AFTER
  // swap has wiped useAI; setError reads LIVE store (now NEW) and writes
  // OLD's exception text into NEW's `ai.error`. Same shape as R236
  // (handleAssistantToolCalls outer catch) and R237 (cancelInflight's
  // partial-persist catch). The chunk listener (`onChunk` / `onError`)
  // doesn't need this guard because cancelInflight's `handle.cancel()`
  // off()'s the IPC listener at swap time; chunks arriving post-swap
  // never reach those handlers. The `.catch` / `.then(errFn)` chains
  // here are the post-listener async tail that survives off() and
  // can fire after swap.
  const runWorkspaceId = useWorkspace.getState().workspaceId;

  return new Promise<void>((resolve) => {
    let toolUseInProgress: { id: string; name: string; partial: string } | null = null;

    inflight = {
      resolve,
      handle: sendChatTurn(
      {
        model: opts.model,
        system: SYSTEM_PROMPT,
        messages,
        tools,
        cacheBreakpoints,
        maxTokens: opts.maxTokens,
        temperature: opts.temperature,
        // Read from the live AI store rather than RunOptions: promptCache is a
        // global cost-mode toggle (not a per-turn knob like model / tools), so
        // it doesn't belong on the RunOptions surface every caller has to
        // populate. Mirrors how the store also owns model / maxTokens /
        // temperature for the same reason — see store/ai.ts:46-53.
        promptCache: ai.promptCache,
      },
      {
        onChunk: (chunk) => {
          const s = useAI.getState();
          switch (chunk.type) {
            case 'text_delta':
              s.appendStreamText(chunk.text);
              break;
            case 'tool_use_start':
              toolUseInProgress = { id: chunk.id, name: chunk.name, partial: '' };
              // Surface a "preparing…" badge while the tool's input JSON streams
              // in. Without this, long inputs (excel_set_range 2D arrays, big
              // markdown blobs) make the stream appear frozen between
              // tool_use_start and tool_use_complete.
              s.setInProgressToolUse({ id: chunk.id, name: chunk.name });
              break;
            case 'tool_use_input_delta':
              if (toolUseInProgress && toolUseInProgress.id === chunk.id) {
                toolUseInProgress.partial += chunk.partialJson;
              }
              break;
            case 'tool_use_complete': {
              const block: ToolUseBlock = {
                type: 'tool_use',
                id: chunk.id,
                name: chunk.name,
                input: chunk.input,
              };
              // Clear the placeholder BEFORE appending the finalized block so
              // the renderer doesn't briefly show both the badge and the
              // committed tool_use side-by-side.
              s.setInProgressToolUse(null);
              s.appendStreamToolUse(block);
              toolUseInProgress = null;
              break;
            }
            case 'usage':
              s.addUsage({
                inputTokens: chunk.inputTokens,
                outputTokens: chunk.outputTokens,
                cacheReadInputTokens: chunk.cacheReadInputTokens,
                cacheCreationInputTokens: chunk.cacheCreationInputTokens,
              });
              break;
            case 'message_stop':
              break;
          }
        },
        onError: (msg) => {
          // R274 — workspace guard for the IPC-reject path. handlers.onError
          // is called from TWO sources: (a) the chunk listener when main
          // sends a `{type: 'error', message}` chunk (provider.ts:47-51),
          // and (b) the chat IPC's `.catch` when the IPC promise itself
          // rejects (provider.ts:75-79). R238's existing argument
          // ("chunks arriving post-swap never reach these handlers
          // because cancelInflight off()'s the listener") covers (a)
          // only — off() removes the chunk listener, but the chat
          // promise's `.catch` is a separate handler tied to the chat
          // call itself, NOT to the chunk listener registry. So a
          // post-cancel IPC reject (main crash, bridge error, future SDK
          // wrapping that surfaces upstream errors as promise rejects
          // rather than `error` chunks) still fires this onError, and
          // setError lands on whichever useAI is currently live — which
          // may be NEW workspace's freshly-cleared error slot. Same
          // workspace-id capture pattern as R236 / R237 / R238 closes
          // the last unguarded setError site in this turn's tail.
          if (useWorkspace.getState().workspaceId !== runWorkspaceId) return;
          useAI.getState().setError(msg);
        },
        onDone: () => {
          inflight = null;
          const finalized = useAI.getState().finalizeStreaming();
          if (finalized) {
            // R203 — same .catch coverage as cancelInflight's persistMessage
            // call (line ~282) and handleAssistantToolCalls's R165 .then's
            // rejection arg. DB persist failures get surfaced via setError
            // instead of leaking as unhandledrejection.
            // R238 — workspace-id guard, see runTurn-top comment.
            void persistMessage(finalized).catch((err) => {
              if (useWorkspace.getState().workspaceId !== runWorkspaceId) return;
              useAI.getState().setError(
                err instanceof Error ? err.message : String(err),
              );
            });
            // R161 — bridge `streaming.requestId` across the inline-tool
            // processing window. `finalizeStreaming` zeroes the requestId,
            // and `handleAssistantToolCalls` may then dispatch read-only
            // tools synchronously, push synthetic tool_results, and call
            // another `runTurn` (which sets its own fresh requestId). In
            // between those two state changes — microseconds, but real —
            // the AIPanel's send gate (`if (ai.streaming.requestId !==
            // null) return;` at line 209) reads null, so a user Enter at
            // exactly that moment can fire `sendUserPrompt` concurrent
            // with `handleAssistantToolCalls`'s recursive `runTurn`. Both
            // overwrite the shared `inflight` handle and stream into the
            // same `streaming.content`; two assistant messages get
            // appended and one IPC stream is orphaned. Re-setting a
            // sentinel id here keeps the gate closed; the next runTurn
            // overwrites it with its own uuid, and the pending-exit
            // branch in `handleAssistantToolCalls` clears it explicitly
            // so Apply/Reject buttons unblock. Skip the bridge when no
            // tool_uses are present — that turn ended cleanly with text
            // and the gate should genuinely open.
            const hasToolUses = finalized.content.some(
              (b) => b.type === 'tool_use',
            );
            if (hasToolUses) {
              // R263 — set streaming directly instead of via resetStreaming,
              // because resetStreaming(requestId) ALSO clears `error: null`
              // (it's shared with fresh-turn and retry callsites that
              // legitimately want a clean slate). Calling it here from the
              // bridge would wipe an `error` that onError just set.
              //
              // Trigger: provider.ts's chunk listener handles `error` chunks
              // by firing onError(msg) THEN onDone() in immediate sequence
              // (provider.ts:47-52). orchestrator's onError calls
              // setError(msg). Then onDone runs (this code path); if the
              // partial assistant message had any tool_use blocks, the
              // bridge resetStreaming clobbers `ai.error` to null before
              // React even renders. User never sees「Rate limit exceeded」
              // / 「Network error」 / 「API authentication failed」 — just
              // a turn that "stopped working" mysteriously, with the
              // partial content visible but no diagnostic banner.
              //
              // Use a direct setState that ONLY touches streaming, not
              // error. Same effect as resetStreaming for the bridge's
              // purposes (content empty, requestId set to fresh uuid,
              // inProgressToolUse cleared) but doesn't tread on the
              // independent error channel. resetStreaming itself stays
              // unchanged so the runTurn-entry / retryLastTurn callers
              // continue to clear errors as they should.
              useAI.setState({
                streaming: {
                  content: [],
                  requestId: uuid(),
                  inProgressToolUse: null,
                },
              });
            }
            // R165 — second `.then` arg covers the rejection path. The
            // dispatch loop awaits `pushSyntheticToolResult`, which awaits
            // `persistMessage` → `appendMessage` IPC; a DB write failure
            // (disk full, locked SQLite during a backup, IPC rejected)
            // would otherwise propagate as an unhandled rejection on the
            // void promise here, and `resolve` would never fire — exact
            // same Promise-leak shape R164 just closed for the cancel
            // path. Surface the error through the AI store's setError
            // (same vocabulary the on-error chunk path uses, line 396)
            // and resolve so awaiters can unwind.
            void handleAssistantToolCalls(finalized).then(resolve, (err) => {
              // R238 — workspace-id guard before setError, see runTurn-top
              // comment. handleAssistantToolCalls's own R178 / R236 guards
              // make it return early on swap (no rethrow), so this errFn
              // mostly catches genuine internal exceptions; but defensive
              // gate matches the persistMessage catch above and keeps the
              // setError invariant uniform.
              if (useWorkspace.getState().workspaceId === runWorkspaceId) {
                useAI.getState().setError(
                  err instanceof Error ? err.message : String(err),
                );
              }
              resolve();
            });
          } else {
            resolve();
          }
        },
      },
      ),
    };
  });
}

async function handleAssistantToolCalls(msg: ChatMessage): Promise<void> {
  const toolUses = msg.content.filter((b): b is ToolUseBlock => b.type === 'tool_use');
  if (toolUses.length === 0) return;

  // Anthropic requires every tool_use in an assistant message to be paired
  // with a tool_result before the next assistant turn — partial pairing
  // produces a 400 "tool_use ids were found without tool_result blocks".
  // The previous loop fired `continueAfterToolResult` (which kicks off a
  // fresh runTurn) on the FIRST inline/error result and `return`ed, leaving
  // any remaining tool_uses in the same assistant message unanswered. The
  // most common bad case: assistant emits a read tool + a write tool, the
  // read inline-completes, the loop short-circuits, the write tool_use is
  // silently dropped, no PendingChange is created, and the next user prompt
  // 400s.
  //
  // The fix is to walk every tool_use, push synthetic tool_results for the
  // inline/error cases (no runTurn — that would re-enter the API mid-loop
  // with the rest of the tool_uses still unpaired), and queue the rest as
  // PendingChanges. Only after the loop do we decide who triggers the next
  // turn: if everything was inline we runTurn ourselves; if any pending
  // remains the user's Apply/Reject closes the chain via
  // continueAfterToolResult.
  const ws = useWorkspace.getState();
  // R178 — capture workspaceId at start so the loop can abort if the user
  // swaps workspace mid-dispatch. Without this guard, the dispatch loop
  // happily writes the OLD turn's tool_results / pending changes into the
  // NEW workspace's AI store. Realistic trigger: a long inline-tool chain
  // (e.g. "讀 sheet1 的 5 個 cell") streams its tool_uses, onDone fires,
  // R161 sets the bridge requestId, this loop starts dispatching; user
  // hits Ctrl+O during dispatch (each `await pushSyntheticToolResult →
  // persistMessage → IPC` is a microtask boundary). loadFromOpened →
  // cancelInflight (no-op, inflight=null already) → `useAI.clear()` wipes
  // messages/pending/conversationId. The loop's next iteration then
  // pushes OLD's tool_result into NEW's empty messages, addPending writes
  // OLD's PendingChange to NEW's pending list — so the user lands in
  // NEW workspace looking at chat artifacts that have nothing to do with
  // its tabs. cancelInflight ALONE doesn't fix this because handleAssistant
  // ToolCalls runs from the `void ...then(resolve)` outside `inflight`'s
  // lifecycle (see onDone above).
  const wsAtStart = ws.workspaceId;
  const stillSameWs = () => useWorkspace.getState().workspaceId === wsAtStart;
  // R256 — capture the bridge requestId set by onDone (R161) so the loop
  // can detect a user 停止 click during the inline-dispatch window.
  // cancelInflight clears streaming.requestId via finalizeStreaming when
  // invoked while inflight is null but a dispatch is active; we compare
  // against the captured value at each iteration top and bail before
  // pushing more synth results AND before falling through to runTurn.
  // Without this, the 停止 button is silently broken from onDone until
  // the post-dispatch runTurn re-arms inflight — a 200-2500ms window
  // for typical inline-tool chains. See cancelInflight doc-block for
  // the full failure trace.
  const initialRequestId = useAI.getState().streaming.requestId;
  const stillSameStream = () =>
    useAI.getState().streaming.requestId === initialRequestId;
  let anyPending = false;
  // R316 — track which tabs already have a pending op from THIS turn so we
  // can reject a second tool_use that would target the same tab. The
  // dispatcher captures `before = tab.data` from `ws.tabs` (the snapshot
  // taken at handleAssistantToolCalls entry, line 685), so every iteration
  // in this loop sees the SAME `before` bytes. If two write tool_uses
  // (e.g., AI: "insert 前言 at idx 3" + "insert 結語 at idx 8") both
  // dispatch successfully and addPending, they produce two binary_replace
  // ops with op_1.after = D0+前言 and op_2.after = D0+結語 — both derived
  // from the same D0 base.
  //
  // applyChangeset (changeset-apply.ts:82-89) is a pure overwrite
  // (`tab.data = op.after`) — it doesn't compose ops. So when the user
  // sequentially Applies P1 then P2:
  //   • Apply P1: tab.data = D0+前言. Renders OK.
  //   • Apply P2: tab.data = D0+結語. The 前言 paragraph vanishes — P2's
  //     `after` was computed from D0, not from D0+前言.
  // User sees both PendingChange cards, clicks Apply on both, expects both
  // edits to land. Only the second sticks; the first is silently lost.
  // Ctrl+Z can't recover the "both edits applied" state because that state
  // was never computed.
  //
  // Cross-turn doesn't trigger this — Apply 1's continueAfterToolResult
  // fires a fresh runTurn, the new turn's handleAssistantToolCalls
  // captures a NEW `ws.tabs` snapshot reflecting the post-Apply 1 state,
  // and the next AI op's `before` is the right baseline. Only within ONE
  // turn does the shared snapshot defeat compose.
  //
  // Reject the second tool_use upfront with a typed error so AI knows to
  // either (a) combine multiple edits into a single tool_use (if the
  // available tools support it — they mostly don't today, e.g.,
  // word_insert_paragraph is single-paragraph only), or (b) tell the
  // user to apply the first proposal before requesting another change.
  // Same fail-fast-with-actionable-error idiom as R214 (orphanedOp) and
  // R313 (empty oldText).
  const tabsWithPending = new Set<string>();
  for (const block of toolUses) {
    if (!stillSameWs()) return;
    if (!stillSameStream()) return;
    // R228 / R230 — two-layer per-block exception guard. R228 catches
    // dispatchTool throws (malformed OOXML triggering parseXlsx /
    // parseDocx / parsePptx exceptions, TextEncoder edge cases, future
    // tool helpers that haven't been audited for the typed-result
    // contract). R230 wraps the whole iteration body so any of the
    // FOUR `await pushSyntheticToolResult` callsites in this loop
    // (dispatch-threw / typed-failure / inline-result / R214 orphan)
    // can reject — most realistic via persistMessage IPC failure (DB
    // locked, disk full, sqlite-wal corruption) — without aborting
    // every later tool_use in the same assistant message. R228 alone
    // wasn't enough: its catch path itself does
    // `await pushSyntheticToolResult(...)` and that await can throw
    // too, escaping the inner catch and aborting the loop the same
    // way R228 was supposed to prevent.
    //
    // The synthetic tool_result is in `ai.messages` already (pushMessage
    // inside pushSyntheticToolResult is sync, ran before the failing
    // persist await), so the in-memory invariant Anthropic depends on
    // (every tool_use paired with a tool_result) is intact. Only the
    // persistent SQLite record has a gap until next session — which is
    // an acceptable degradation since chat.sqlite is per-instance
    // (cleared on app reload). setError surfaces the failure via the
    // banner, same channel as R210 / R211 / R203 / R204 use.
    try {
      let result: DispatchResult;
      try {
        // R415 — lazy import inside the R228 try: a (highly unlikely)
        // chunk-load failure routes through the same dispatch_threw
        // synthetic-tool-result path as a dispatcher exception.
        const { dispatchTool } = await import('./dispatcher');
        result = await dispatchTool(block, { tabs: ws.tabs, activeTabId: ws.activeTabId });
      } catch (err) {
        // R262 — bail before pushing synth on dispatch-threw if user
        // cancelled mid-dispatchTool. Without this, cancelInflight's
        // R256 finalizeStreaming clears the bridge requestId, but the
        // inner catch is mid-execution and goes on to push synth into
        // OLD's now-cleared streaming context. The synth is still
        // technically valid (Anthropic-consistency: every tool_use
        // wants a tool_result), but it lands AFTER the user's
        // explicit cancel — the next sendUserPrompt's dangling-
        // cleanup loop would have synthesized exactly the same
        // tool_result anyway. Returning here defers consistency to
        // dangling-cleanup, matching the cancel-as-clean-stop
        // semantic R256 / R261 build toward.
        if (!stillSameStream()) return;
        const msg = err instanceof Error ? err.message : String(err);
        await pushSyntheticToolResult(block.id, `dispatch_threw: ${msg}`, true);
        continue;
      }
      if (!stillSameWs()) return;
      // R262 — also short-circuit post-dispatchTool if user cancelled
      // during the dispatchTool await. Without this:
      //   • result.ok=false branch pushes synth post-cancel.
      //   • result.ok=true + inlineResult branch pushes synth post-cancel.
      //   • orphanedOp branch pushes synth post-cancel.
      //   • Worst: addPending fires post-cancel → user sees a NEW
      //     PendingChange card pop into the AIPanel AFTER they clicked
      //     停止. The cancel-as-clean-stop UX expectation is that no
      //     new cards / messages appear after the user explicitly
      //     stopped. The next sendUserPrompt's R244 dangling cleanup
      //     covers Anthropic API consistency (synthesizes tool_results
      //     for the unanswered tool_uses on next user turn), so we
      //     can safely punt synth for THIS iteration.
      if (!stillSameStream()) return;
      if (!result.ok) {
        await pushSyntheticToolResult(block.id, result.error, true);
        continue;
      }
      // R302 — collapse the two "no ops" branches into one. Original code
      // only handled `ops.length === 0 && inlineResult !== undefined`
      // (inline-result tools like read_tab_content). If a future tool
      // returns `{ok: true, ops: [], inlineResult: undefined}` — successful
      // no-op — it would fall through to addPending(pending with empty
      // ops) and the user sees an empty DiffPreview card whose Apply
      // dirties the workspace for no visible reason. Treat any
      // zero-ops success as inline-resolved, falling back to a
      // 'no_changes' synth tool_result if the tool didn't supply one.
      if (result.changeset.ops.length === 0) {
        await pushSyntheticToolResult(
          block.id,
          result.inlineResult ?? 'no_changes',
          false,
        );
        continue;
      }
      // R214 — verify every non-tab_create op still has a live target tab
      // BEFORE adding the pending. The dispatch above used the captured
      // `ws.tabs` snapshot from line 528; if the user closed a target tab
      // between the snapshot and now (most concretely: the dispatch loop
      // awaited persistMessage IPC for an earlier block, the user
      // impatiently × the tab in question, removeTab's orphan-eviction at
      // workspace.ts:446 found an empty pending list because we hadn't
      // addPending'd yet — so no synthetic was pushed there either),
      // addPending would create a PendingChange referencing a closed tab.
      // The user clicks Apply, applyChangeset's `t.id === op.tabId` filter
      // never matches, the changeset effectively no-ops but `setState({
      // tabs, dirty: true })` still flips the workspace dirty flag — so
      // the user sees the card disappear, no visible change, and a「未
      // 儲存」 dot they didn't earn. Push a synthetic error tool_result
      // instead so the model knows the proposal is moot, mirroring the
      // shape removeTab uses for the post-addPending orphan path (R204).
      const liveTabIds = new Set(useWorkspace.getState().tabs.map((t) => t.id));
      const orphanedOp = result.changeset.ops.find(
        (op) => op.type !== 'tab_create' && !liveTabIds.has(op.tabId),
      );
      if (orphanedOp) {
        await pushSyntheticToolResult(
          block.id,
          'Target tab was closed before the change could be queued.',
          true,
        );
        continue;
      }
      // R316 — reject a second pending op on a tab that's already pending
      // from THIS turn. The shared `ws.tabs` snapshot makes any second op
      // see a stale `before`, and applyChangeset's overwrite semantics
      // silently drops the first Apply's effect when the second Apply
      // lands. See `tabsWithPending` declaration above for the full bug
      // trace. The error string names the tab id so AI can attribute the
      // collision to a specific tool_use in its plan.
      const collidingOp = result.changeset.ops.find(
        (op) => op.type !== 'tab_create' && tabsWithPending.has(op.tabId),
      );
      if (collidingOp) {
        await pushSyntheticToolResult(
          block.id,
          `multi_edit_same_tab: tab ${collidingOp.tabId} already has a pending edit from earlier in this turn — applyChangeset cannot compose two binary_replace / md_text ops sharing the same base bytes (the second Apply would silently overwrite the first). Combine into a single tool_use, or wait for the user to Apply the first proposal before issuing another change to this tab.`,
          true,
        );
        continue;
      }
      const pending: PendingChange = {
        id: uuid(),
        toolUseId: block.id,
        toolName: block.name,
        changeset: result.changeset,
        summary: result.summary,
      };
      useAI.getState().addPending(pending);
      anyPending = true;
      // Record every non-tab_create op's target tab so the next iteration's
      // collidingOp check sees them. tab_create is excluded because its
      // tabId names a NEW tab to be created — collisions are impossible
      // unless two tab_creates share an id (essentially zero with uuid).
      for (const op of result.changeset.ops) {
        if (op.type !== 'tab_create') tabsWithPending.add(op.tabId);
      }
    } catch (err) {
      // R230 — pushSyntheticToolResult rejection (or any other
      // un-typed throw) inside any of the branches above lands here
      // instead of aborting the whole loop. Surface to setError so
      // the user sees the failure; continue so remaining tool_uses
      // get processed (and their synthetic tool_results pushed in-
      // memory, even if persistence is failing for the moment).
      // R236 — also short-circuit if the workspace swapped during
      // the failed await. Without this guard, an OLD-turn error that
      // arrived after the user already Ctrl+O'd into NEW workspace
      // would write OLD's exception message into NEW's `ai.error`,
      // showing a banner about a turn from a workspace that's no
      // longer visible. Same shape as the four other `if (!still
      // SameWs()) return;` checks scattered through this loop —
      // this is the catch-arm equivalent that completes the
      // workspace-swap-mid-iteration coverage. Returning (instead
      // of `continue`) is correct: if workspace swapped, every
      // remaining iteration would also bail at the iteration-top
      // stillSameWs check, so we may as well exit now.
      if (!stillSameWs()) return;
      // R261 — also short-circuit if the user fired 停止 (R256 cleared
      // the bridge requestId via finalizeStreaming inside cancelInflight).
      // Without this guard, a persistMessage rejection on the iteration
      // that happened to be in-flight when the user cancelled would
      // surface a「Database locked: …」 banner AFTER the user's
      // explicit cancel — looking like the cancel itself caused the
      // error. Mirrors R236's stillSameWs-suppress philosophy: the
      // user's primary intent (cancel) supersedes secondary noise
      // (DB error from a now-cancelled iteration). The DB write WAS
      // a real failure, but the iteration is already abandoned —
      // surfacing it confuses the cancel UX and the in-memory state
      // is consistent (sync pushMessage already wrote the synth
      // tool_result; only persistence failed for one orphaned row).
      // Same `return` semantics as the stillSameWs branch — the
      // next iteration's top-level stillSameStream check would
      // bail anyway, so we exit now.
      if (!stillSameStream()) return;
      useAI.getState().setError(err instanceof Error ? err.message : String(err));
      continue;
    }
  }
  // R178 — final guard before either continuing the chain (runTurn) or
  // clearing the bridge requestId. Both branches mutate the AI store; if
  // workspace swapped, the OLD turn shouldn't influence NEW state.
  if (!stillSameWs()) return;
  // R256 — same bail for user 停止 between the loop's last iteration and
  // the runTurn call below. Without this, a cancel that landed AFTER the
  // last iteration's await but BEFORE this post-loop section would slip
  // through into runTurn and start a fresh streaming turn against the
  // user's clear cancel intent. The bridge requestId mismatch is the
  // signal; finalizeStreaming inside cancelInflight already cleared it.
  if (!stillSameStream()) return;
  // All tool_uses produced inline/error results — the synthetic tool_results
  // are already in `messages`; let the model continue. If any PendingChange
  // remains, defer: the user's Apply/Reject is what closes the loop.
  if (!anyPending) {
    await runTurn(currentOpts());
  } else {
    // R161 — clear the bridge requestId set by onDone (see comment there).
    // The bridge kept the AIPanel's send / Apply / Reject gates closed
    // during the dispatch loop; with at least one PendingChange now
    // awaiting user action, those gates must reopen so the user can click
    // Apply/Reject (their handlers in AIPanel.tsx:232 / 271 explicitly
    // require `streaming.requestId === null`). `finalizeStreaming` with
    // empty content (the assistant message was already committed in
    // onDone) is the canonical "reset streaming slot" — returns null
    // and zeroes requestId without re-touching messages.
    useAI.getState().finalizeStreaming();
  }
}

function currentOpts(): RunOptions {
  const ai = useAI.getState();
  return {
    toolsEnabled: ai.toolsEnabled,
    model: ai.model,
    // Read live values from the store — Settings dialog edits are mirrored
    // here by App.tsx, so a self-recurse turn (post-tool_result follow-up)
    // honors the user's current maxTokens/temperature instead of the old
    // 4096/0.3 hardcoded literals.
    maxTokens: ai.maxTokens,
    temperature: ai.temperature,
  };
}

/**
 * Selection of cache breakpoints — keep the boundary between the historical
 * messages and the most recent turn (spec §6.3.2 strategy: cache the prefix).
 *
 * We mark the message preceding the most-recent user message as the breakpoint.
 */
function computeCacheBreakpoints(messages: ChatMessage[]): number[] {
  if (messages.length < 2) return [];
  // Find the last user message; cache up to (but not including) it.
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === 'user') {
      if (i > 0) return [i - 1];
      return [];
    }
  }
  return [];
}

function buildOutline(tabs: ReturnType<typeof useWorkspace.getState>['tabs'], activeId: string | null): string {
  // R400 — emit an explicit「workspace is empty」 signal rather than empty
  // string. Pre-R400, an empty workspace caused this function to return ''
  // → sendUserPrompt's `if (outline)` skipped the [Active workspace] block
  // entirely → AI saw a prompt with NO workspace context whatsoever.
  // Concrete failure mode: user is on the EmptyState screen, types「幫我
  // 寫一個 Q1 報告」, AI responds without knowing whether tabs exist.
  // Some plausible AI outputs:
  //   • Generic「我可以幫您撰寫，請告訴我要包含什麼章節」 ignoring that
  //     they could be using addTab + word_insert_paragraph
  //   • Tries word_insert_paragraph / md_replace_section anyway, hitting
  //     no_active_tab (R310), wastes a tool-use round trip
  //   • Asks「請問您要在哪個 tab 寫」, asking the user to confirm
  //     something the system could have told it
  // Surfacing「（目前工作區沒有頁籤）」 closes the gap: AI sees that
  // tabs.length===0 is a real state (not a missing-context bug), can
  // suggest「請從工具列「+」新增 markdown 頁籤」 directly, and avoids
  // wasteful tool-use attempts. Same posture as R279's「outline accurately
  // describes the binary tabs' write capability」 — give AI the truth, AI
  // plans correctly.
  if (tabs.length === 0) return '（目前工作區沒有頁籤；可請使用者透過工具列「+」新增 markdown / html / docx / xlsx / pptx 頁籤後再操作）';
  const lines: string[] = [];
  for (const t of tabs) {
    const marker = t.id === activeId ? '★' : ' ';
    if (t.type === 'markdown' || t.type === 'html') {
      // R315 — track ``` fences so `# comment` inside a code block isn't
      // surfaced as a heading. The MarkdownEditor's parseOutline at
      // MarkdownEditor.tsx:70-93 already does this; the AI outline was the
      // lone fence-blind surface. Symptoms without the guard: AI sees
      // fake headings, asks user to confirm operations against them, and
      // dispatcher's replaceMdSection (R315 sibling fix below) blows away
      // doc content from the fake heading to EOF when AI calls
      // md_replace_section with the fake title. Same one-pass scan as
      // parseOutline; cap at 12 headings to match the prior slice.
      let headings: string[] = [];
      if (t.type === 'markdown') {
        const mdLines = t.content.split('\n');
        let inFence = false;
        // R342 — use the SAME heading regex parseOutline / replaceMdSection
        // (post-R329) use, so closing ATX `#` markers (`# Title #`,
        // `## Section ##`) are stripped from the captured text. Old code
        // pushed the raw line `l` — so AI saw `# Title #` in the outline,
        // stripped the leading `# ` per the tool description's「不含 # 前
        // 綴」hint, and passed `Title #` to md_replace_section. The
        // dispatcher (post-R329) compares against the clean text `Title`,
        // gets `Title # !== Title`, returns heading_not_found — exactly
        // the failure mode R329 was supposed to close. R329 fixed half of
        // the round-trip (the dispatcher side); buildOutline was the other
        // half that still emitted dirty text. Three-way consistency
        // restored: user sees `Title` in the outline pane, AI sees
        // `# Title` in this outline, dispatcher matches against `Title`
        // — all three use the same canonical heading text per CommonMark
        // §4.2's ATX heading definition.
        const headingRe = /^(#{1,6})\s+(.+?)\s*#*\s*$/;
        // R387 — accept tilde fences (~~~) too; see dispatcher.ts:1275 doc-block
        // for the data-loss scenario this closes. Keeps the three outline
        // scanners (this, parseOutline at MarkdownEditor.tsx:77, replaceMdSection
        // at dispatcher.ts:1302) in lockstep so the AI sees, the user sees, and
        // the dispatcher matches the SAME set of "real" headings.
        const fenceRe = /^\s*(?:```|~~~)/;
        for (const l of mdLines) {
          if (fenceRe.test(l)) {
            inFence = !inFence;
            continue;
          }
          if (inFence) continue;
          const m = headingRe.exec(l);
          if (m) {
            headings.push(`${m[1]} ${m[2]}`);
            if (headings.length >= 12) break;
          }
        }
      } else {
        // HTML: extract <h1>…</h6> visible text via a stateless regex sweep.
        // Simple regex is intentional — DOMParser would import a whole
        // module worth of dependency just for an outline-extraction helper
        // that runs once per turn. We pull the FIRST text content between
        // matching open/close tags and strip nested tags + entities lightly.
        // Same 12-heading cap as markdown so the outline budget stays
        // bounded for AI consumption.
        //
        // R321 — strip `<script>` / `<style>` bodies BEFORE the heading
        // scan. Their inner text is a single text node from DOM's view
        // and frequently contains heading-shaped substrings — JS string
        // literals like `const sample = '<h1>X</h1>'`, CSS comments
        // documenting heading styles, generated-HTML templates inside
        // `<script type="text/template">`. Without this pre-strip the
        // regex below picks those up as if they were real document
        // headings and AI plans operations / writes responses referring
        // to them — confusing for the user (those "headings" are
        // invisible in the rendered preview). Mirrors R317's SKIPPED_TAGS
        // family for html-to-docx and R315's fence tracking for markdown
        // heading scans; both closed the same shape on their respective
        // paths. The strip regex is non-greedy and case-insensitive so
        // it doesn't run past the first `</script>` / `</style>` and
        // handles `<SCRIPT>` / `<Style>` uppercase variants.
        const cleaned = t.content
          .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
          .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '');
        const re = /<h([1-6])[^>]*>([\s\S]*?)<\/h\1>/gi;
        let m: RegExpExecArray | null;
        while ((m = re.exec(cleaned)) !== null && headings.length < 12) {
          const level = Number(m[1]);
          // R345 — decode order: `&amp;` MUST go LAST, not first. The previous
          // order decoded `&amp;` → `&` upfront, then re-scanned for `&lt;`
          // etc. — so an authored `&amp;lt;` (HTML source for the literal
          // text 「&lt;」, used in docs about HTML tags or escaped CMS
          // content) round-tripped to `<` instead of `&lt;`. Step-by-step:
          //   input:    &amp;lt;
          //   &amp;→&   becomes  &lt;
          //   &lt;→<    becomes  <            ✗ over-decoded by one round
          // Reversing the order:
          //   input:    &amp;lt;
          //   &lt;→<    no match (the `&lt;` substring's `&` is preceded by
          //             `&amp` not standalone)
          //   &gt;→>    no match
          //   &amp;→&   becomes  &lt;          ✔ correct
          // This matches HTML's canonical single-pass parser semantics
          // (it tokenizes left-to-right and never re-scans the output).
          // For the common case `Q&amp;A` both orders produce「Q&A」 —
          // only the `&amp;<named-entity>` pattern was affected. Realistic
          // trigger: technical docs whose heading reads「The &lt; Operator」
          // (authored as `<h2>The &amp;lt; Operator</h2>` to display the
          // literal「&lt;」) — outline showed「The < Operator」, misleading
          // AI's understanding of the heading text.
          const inner = m[2]
            .replace(/<[^>]+>/g, '') // strip nested tags
            .replace(/&nbsp;/g, ' ')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .trim();
          if (!inner) continue;
          headings.push(`${'#'.repeat(level)} ${inner}`);
        }
      }
      const typeLabel = t.type === 'markdown' ? 'md' : 'html';
      lines.push(`${marker} [${typeLabel}] ${t.name} (id=${t.id})`);
      headings.forEach((h) => lines.push(`     ${h}`));
    } else {
      // R279 — describe actual write capabilities instead of the stale
      // "read-only in v1.0" claim. The dispatcher implements basic text
      // writes for all three binary formats (word_replace_paragraph /
      // word_insert_paragraph for docx, excel_set_cell / excel_set_range
      // for xlsx, pptx_replace_text for pptx); only the advanced ops
      // (styles / charts / new slides / new rows) return
      // not_implemented_in_mvp. Telling the AI a format is "read-only"
      // makes it refuse perfectly valid write tools — the outline is
      // the line the AI reads on every turn, so accuracy here directly
      // controls whether the AI even attempts the right tool. Surface
      // the tool family by type so the AI's planning step has a
      // concrete pointer to the correct entry point.
      const writeHint =
        t.type === 'docx'
          ? 'word_replace_paragraph / word_insert_paragraph'
          : t.type === 'xlsx'
            ? 'excel_set_cell / excel_set_range'
            : 'pptx_replace_text';
      lines.push(
        `${marker} [${t.type}] ${t.name} (id=${t.id}, ${t.data.byteLength} bytes; 基本寫入：${writeHint})`,
      );
    }
  }
  return lines.join('\n');
}

/**
 * Dedupe concurrent createConversation calls + handle workspace-swap mid-IPC.
 *
 * R159 closed the same-workspace dedupe race: two persistMessage calls in
 * the same microtask both saw `conversationId === null`, both awaited
 * their own createConversation IPC, and the second `setConversation`
 * overwrote the first; the first call's appendMessage landed in an orphan
 * conversation row.
 *
 * R160 closes two sibling races R159 left open:
 *   1. **Workspace swap during the create IPC.** `loadFromOpened` /
 *      `newWorkspace` call `clear()` which resets `conversationId` to null,
 *      but the in-flight createConversation still resolves with an id keyed
 *      to the OLD workspace. The previous fix would `setConversation(idA)`
 *      on the NEW workspace's store, so every subsequent persistMessage in
 *      the new workspace would append into the old workspace's convo row —
 *      cross-workspace history bleed.
 *   2. **Cross-workspace dedupe collision.** The R159 pendingCreate had no
 *      workspaceId tag, so a persistMessage from the new workspace would
 *      reuse a pending create that was started by the old workspace, get
 *      back the old workspace's convo id, and append its (new-workspace)
 *      message into the old convo.
 *
 * Tag pendingCreate with its `wsId` so cross-workspace callers don't share;
 * after the create IPC resolves, only `setConversation` if the workspace
 * still matches. Old-workspace callers still get their id back so their
 * own message appends to the right (old) convo — we just don't pollute the
 * new workspace's store pointer.
 *
 * The `finally` only clears `pendingCreate` if it's still pointing at our
 * own promise — a mid-flight workspace swap that started a new pendingCreate
 * for the new wsId mustn't be wiped by an earlier IIFE's late completion.
 */
let createGen = 0;
let pendingCreate:
  | { wsId: string | null; gen: number; promise: Promise<string> }
  | null = null;

async function ensureConversation(): Promise<string> {
  const cur = useAI.getState().conversationId;
  if (cur) return cur;
  // workspaceId is `string | null` in the store (a fresh app start before
  // any open/new shows null); the persistence layer's createConversation
  // accepts null too. We tag pendingCreate with whatever the value is so
  // the cross-workspace dedupe still works for the null case.
  const wsId = useWorkspace.getState().workspaceId;
  if (pendingCreate && pendingCreate.wsId === wsId) return pendingCreate.promise;
  // Capture our own generation number BEFORE the IIFE starts so the finally
  // block has a stable identity to compare against — using the `promise`
  // variable itself would be a TDZ forward-reference (the IIFE's body
  // begins executing synchronously before `const promise = …` finishes
  // binding).
  const myGen = ++createGen;
  const promise = (async (): Promise<string> => {
    try {
      const created = await window.gendoc.history.createConversation({
        title: 'Conversation',
        workspaceId: wsId,
      });
      // Workspace may have swapped during the IPC. Only bind the new id to
      // the live AI store if we're still in the originating workspace; the
      // returned id is still correct for the caller's own appendMessage
      // (their message belongs to wsId's convo, not whichever workspace
      // happens to be active now).
      if (useWorkspace.getState().workspaceId === wsId) {
        useAI.getState().setConversation(created.id);
      }
      return created.id;
    } finally {
      // Only null out pendingCreate if it's still us — a workspace swap
      // mid-IPC may have already replaced it with a fresh {wsId, gen, promise}
      // for the new workspace; clearing then would erase the new workspace's
      // legitimate in-flight create.
      if (pendingCreate?.gen === myGen) {
        pendingCreate = null;
      }
    }
  })();
  pendingCreate = { wsId, gen: myGen, promise };
  return promise;
}

async function persistMessage(msg: ChatMessage): Promise<void> {
  // R308 — capture workspaceId before ensureConversation's await so a
  // mid-IPC swap doesn't cause OLD's message to land in NEW's history.
  // Concrete trigger: user clicks 修改 on a pending in workspace A →
  // pushSyntheticToolResult fires → pushMessage adds the synth tool_result
  // to A's ai.messages (sync) → await persistMessage starts → ensureConversation
  // awaits createConversation IPC (or returns A's cached conversationId) →
  // user immediately Ctrl+O loads workspace B → loadFromOpened calls
  // useAI.clear() which resets conversationId to null AND changes
  // workspace.workspaceId to B. If clear() landed BEFORE ensureConversation
  // read conversationId, `cur = null` → ensureConversation uses live
  // `workspaceId = B` → createConversation makes a NEW conversation for B
  // → appendMessage writes A's modify-synth tool_result into that brand-new
  // B conversation. Next time the user opens B, chat history shows an
  // orphaned `tool_result` whose `tool_use_id` doesn't match any tool_use
  // in B's messages — confusing UI with no recovery path.
  //
  // The workspace.ts R243 doc-block at workspace.ts:531-533 explicitly
  // flags this exact case as "an existing dropped-result concern, not
  // addressed here". This guard closes it: if workspaceId no longer
  // matches wsAtStart by the time ensureConversation resolves, abort
  // before appendMessage. The msg is already gone from OLD's in-memory
  // `ai.messages` (clear() wiped it) so dropping the persist matches
  // the in-memory state. ensureConversation may have created an empty
  // conversation row for NEW as a side effect of the race — that's a
  // minor SQLite stray (empty rows are harmless to UI) and strictly
  // less bad than writing OLD's tool_result text into NEW's real
  // history. Same workspace-pinning capture-and-check shape as R236
  // (handleAssistantToolCalls outer catch), R237 (cancelInflight's
  // partial persist), R238 (runTurn onDone catch), R239 (AIPanel send),
  // R240 (AIPanel onApply/onReject), R243 (workspace removeTab's synth).
  const wsAtStart = useWorkspace.getState().workspaceId;
  const convoId = await ensureConversation();
  if (useWorkspace.getState().workspaceId !== wsAtStart) return;
  // Map our ChatMessage role to the persistence row.
  const role = msg.role === 'tool_result' ? 'tool_result' : msg.role;
  // R296 — optional-chain the unsafe cast. Original `(msg.content[0] as {
  // tool_use_id?: string }).tool_use_id` accesses `.tool_use_id` directly
  // on the cast value; if msg.content is empty (`[]`) when role is
  // 'tool_result' — a contract all current callers respect but the type
  // system doesn't enforce — the cast yields undefined and the property
  // read throws TypeError. Today no callsite produces empty content for
  // tool_result, but a future tool-result constructor / test mock /
  // corrupt-history replay would surface the bug as a cryptic
  // "Cannot read properties of undefined" via the .catch chain; cleaner
  // to fall back to null (same as non-tool_result messages) so the row
  // persists with an unlinked toolUseId rather than failing the whole
  // appendMessage.
  const firstBlock = msg.content[0] as { tool_use_id?: string } | undefined;
  const toolUseId = msg.role === 'tool_result' ? firstBlock?.tool_use_id ?? null : null;
  await window.gendoc.history.appendMessage(convoId, {
    role,
    content: JSON.stringify(msg.content),
    toolUseId,
    tokenInput: null,
    tokenOutput: null,
    cacheRead: null,
    cacheCreation: null,
  });
}
