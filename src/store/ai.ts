/**
 * AI session store — chat messages, current model, in-flight stream state,
 * pending ChangeSet awaiting user Apply/Reject.
 */

import { create } from 'zustand';
import { v4 as uuid } from 'uuid';
import type { ChatMessage, ContentBlock } from '../types/ai';
import type { ChangeSet } from '../types/changeset';

export interface PendingChange {
  id: string;
  toolUseId: string;
  toolName: string;
  changeset: ChangeSet;
  /** Human readable summary line. */
  summary: string;
}

export interface UsageTotals {
  inputTokens: number;
  outputTokens: number;
  cacheReadInputTokens: number;
  cacheCreationInputTokens: number;
}

interface AIState {
  conversationId: string | null;
  messages: ChatMessage[];
  /**
   * Currently streaming assistant message under construction.
   *
   * `inProgressToolUse` mirrors a tool_use whose name has arrived
   * (`tool_use_start`) but whose `input` is still streaming in via
   * `tool_use_input_delta` chunks. Kept OUT of `content` so a mid-stream
   * cancel/finalize doesn't persist a tool_use with empty input (Anthropic
   * would 400 on the next turn). Renderer surfaces a "preparing…" badge
   * sourced from this field to give live feedback during long tool inputs.
   */
  streaming: {
    content: ContentBlock[];
    requestId: string | null;
    inProgressToolUse: { id: string; name: string } | null;
  };
  model: string;
  /**
   * Per-turn knobs hydrated from `UserConfig` on app start and on every
   * `gendoc:configChanged` broadcast. Previously the AIPanel and the
   * orchestrator's self-recurse path hardcoded 4096 / 0.3 in five places, so
   * adjusting these in the Settings dialog persisted to disk but never took
   * effect at runtime — surprise. Mirroring them into the AI store keeps
   * one source of truth for every code path that builds a RunOptions.
   */
  maxTokens: number;
  temperature: number;
  toolsEnabled: boolean;
  /**
   * Mirrors UserConfig.promptCache. When false, the IPC chat call signals
   * main/ai/anthropic.ts to skip every `cache_control` marker so the request
   * bills as if the cache feature didn't exist. Lives here for the same
   * reason maxTokens / temperature do (see comment block above): the
   * SettingsDialog persists the value to disk via window.gendoc.config.set,
   * but the orchestrator runs from this store on every turn — without a
   * mirror, the toggle was previously a complete no-op and Anthropic kept
   * billing cache writes regardless of the user's choice.
   */
  promptCache: boolean;
  pending: PendingChange[];
  usage: UsageTotals;
  error: string | null;

  setConversation(id: string): void;
  setModel(id: string): void;
  setMaxTokens(n: number): void;
  setTemperature(n: number): void;
  setToolsEnabled(b: boolean): void;
  setPromptCache(b: boolean): void;
  pushMessage(m: ChatMessage): void;
  resetStreaming(requestId: string): void;
  appendStreamText(text: string): void;
  appendStreamToolUse(block: ContentBlock): void;
  setInProgressToolUse(v: { id: string; name: string } | null): void;
  finalizeStreaming(): ChatMessage | null;
  addPending(p: PendingChange): void;
  removePending(id: string): void;
  addUsage(u: Partial<UsageTotals>): void;
  setError(msg: string | null): void;
  clear(): void;
}

export const useAI = create<AIState>((set, get) => ({
  conversationId: null,
  messages: [],
  streaming: { content: [], requestId: null, inProgressToolUse: null },
  model: 'claude-sonnet-4-6',
  // Mirrored from UserConfig.maxTokens / temperature; defaults match
  // config.ts DEFAULTS so untouched config still produces identical
  // behavior to the previous hardcoded values.
  maxTokens: 4096,
  temperature: 0.3,
  toolsEnabled: true,
  // Default true matches config.ts DEFAULTS.promptCache so untouched config
  // and store defaults agree before the App.tsx hydrate fires.
  promptCache: true,
  pending: [],
  usage: {
    inputTokens: 0,
    outputTokens: 0,
    cacheReadInputTokens: 0,
    cacheCreationInputTokens: 0,
  },
  error: null,

  setConversation(id) {
    set({ conversationId: id });
  },
  setModel(id) {
    set({ model: id });
  },
  setMaxTokens(n) {
    set({ maxTokens: n });
  },
  setTemperature(n) {
    set({ temperature: n });
  },
  setToolsEnabled(b) {
    set({ toolsEnabled: b });
  },
  setPromptCache(b) {
    set({ promptCache: b });
  },
  pushMessage(m) {
    set((s) => ({ messages: [...s.messages, m] }));
  },
  resetStreaming(requestId) {
    set({
      streaming: { content: [], requestId, inProgressToolUse: null },
      error: null,
    });
  },
  appendStreamText(text) {
    // R304 — short-circuit on empty delta. Some SDK versions emit
    // `text: ''` text_delta chunks at stream boundaries (especially
    // around thinking-block start/stop pairs). Without this guard, each
    // empty delta still triggers a Zustand setState with a fresh
    // streaming.content array reference — Zustand notifies all
    // subscribers (AIPanel content render, StatusBar token tally,
    // FileExplorer indicators that share useAI selectors) and React
    // re-renders the whole AI tree for no visible change. For a long
    // turn with hundreds of chunks, even a few empty deltas multiply
    // into dozens of unnecessary workspace-wide re-renders. Same
    // "return-prev-on-no-actual-change" idiom as markTabDirty
    // (workspace.ts:697).
    if (text === '') return;
    set((s) => {
      const blocks = [...s.streaming.content];
      const last = blocks[blocks.length - 1];
      if (last && last.type === 'text') {
        blocks[blocks.length - 1] = { ...last, text: last.text + text };
      } else {
        blocks.push({ type: 'text', text });
      }
      return { streaming: { ...s.streaming, content: blocks } };
    });
  },
  appendStreamToolUse(block) {
    // R305 — defensive dedupe by tool_use id. Anthropic's wire protocol
    // guarantees uniqueness within an assistant message, but a buggy SDK
    // / future reconnect path / mock could deliver the same
    // tool_use_complete twice. Without dedupe, the finalized assistant
    // message carries duplicate tool_use blocks, and the next user turn
    // request 400s on「tool_use_id ... appeared more than once」 — a
    // failure mode that surfaces far from its cause (Anthropic API
    // round-trip later, not the duplicate-emit moment). Cheap O(n)
    // check on streaming.content (typically 1-5 blocks per turn).
    if (block.type === 'tool_use') {
      const id = block.id;
      set((s) =>
        s.streaming.content.some((b) => b.type === 'tool_use' && b.id === id)
          ? s
          : { streaming: { ...s.streaming, content: [...s.streaming.content, block] } },
      );
      return;
    }
    set((s) => ({
      streaming: { ...s.streaming, content: [...s.streaming.content, block] },
    }));
  },
  setInProgressToolUse(v) {
    set((s) => ({ streaming: { ...s.streaming, inProgressToolUse: v } }));
  },
  finalizeStreaming() {
    const s = get();
    if (s.streaming.content.length === 0) {
      set({ streaming: { content: [], requestId: null, inProgressToolUse: null } });
      return null;
    }
    const msg: ChatMessage = { role: 'assistant', content: s.streaming.content };
    set((cur) => ({
      messages: [...cur.messages, msg],
      streaming: { content: [], requestId: null, inProgressToolUse: null },
    }));
    return msg;
  },
  addPending(p) {
    set((s) => ({ pending: [...s.pending, p] }));
  },
  removePending(id) {
    set((s) => ({ pending: s.pending.filter((p) => p.id !== id) }));
  },
  addUsage(u) {
    set((s) => ({
      usage: {
        inputTokens: s.usage.inputTokens + (u.inputTokens ?? 0),
        outputTokens: s.usage.outputTokens + (u.outputTokens ?? 0),
        cacheReadInputTokens:
          s.usage.cacheReadInputTokens + (u.cacheReadInputTokens ?? 0),
        cacheCreationInputTokens:
          s.usage.cacheCreationInputTokens + (u.cacheCreationInputTokens ?? 0),
      },
    }));
  },
  setError(msg) {
    set({ error: msg });
  },
  clear() {
    // Reset usage totals alongside messages/pending/streaming. The StatusBar
    // surfaces these as headline `X tokens · Y% cache` which users naturally
    // read as "spend on the current conversation". When clear() runs from a
    // workspace swap (loadFromOpened / newWorkspace) or the 新對話 button,
    // the previous conversation is gone — leaving its tokens accumulated on
    // top of the new conversation's mis-reads as "spend so far on this
    // chat", which it isn't (tooltips break it down by bucket but the
    // headline number is what most users see). The clear() callsites all
    // run cancelInflight first, so no in-flight `addUsage` chunk can race
    // and leave a stale partial behind.
    set({
      messages: [],
      streaming: { content: [], requestId: null, inProgressToolUse: null },
      pending: [],
      error: null,
      conversationId: null,
      usage: {
        inputTokens: 0,
        outputTokens: 0,
        cacheReadInputTokens: 0,
        cacheCreationInputTokens: 0,
      },
    });
  },
}));

export function makePendingId(): string {
  return uuid();
}
