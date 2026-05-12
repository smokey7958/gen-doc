/**
 * Anthropic SDK adapter — runs in main so the API key never crosses to renderer.
 * Streaming chunks are forwarded over IPC channel `ai:chunk` keyed by requestId.
 *
 * Cache control (spec §6.3.2):
 *   - System prompt + tools schema → ephemeral
 *   - First N messages flagged via `cacheBreakpoints`
 *   - Final user turn + selection context → uncached
 */

import Anthropic from '@anthropic-ai/sdk';
import { BrowserWindow } from 'electron';
import { readApiKey } from '../storage/config';
import { IPC } from '../../types/ipc';
import type { StreamChunk } from '../../types/ai';

interface ChatRequest {
  requestId: string;
  model: string;
  system: string;
  messages: Array<{ role: 'user' | 'assistant' | 'tool'; content: unknown }>;
  tools: Array<{ name: string; description: string; input_schema: object }>;
  maxTokens: number;
  temperature: number;
  cacheBreakpoints: number[];
  /** When false, withCache() short-circuits — no `cache_control` markers are
   * attached to system / tools / messages, so every input token bills at the
   * non-cache rate. Mirrors UserConfig.promptCache via the store/ai.ts mirror
   * pattern. Previously this knob was orphaned: SettingsDialog persisted the
   * value to disk, but no read path consulted it, so toggling it OFF had zero
   * effect on Anthropic billing — exactly the same bug class the store/ai.ts
   * doc-comment at lines 46-53 calls out for the older maxTokens / temperature
   * leak. */
  promptCache: boolean;
}

const inflight = new Map<string, AbortController>();

function send(requestId: string, chunk: StreamChunk): void {
  const win = BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0];
  win?.webContents.send(IPC.ai.chunk, { requestId, chunk });
}

/**
 * Decorate system + tools + messages with cache_control markers (§6.3.2).
 * The bundled SDK version's type defs predate cache_control, so we shape
 * everything as plain objects and cast at the call site.
 */
function withCache(req: ChatRequest): {
  system: unknown;
  tools: unknown[];
  messages: unknown[];
} {
  // Short-circuit when the user has explicitly opted out. Returning the raw
  // system / tools / messages without ANY cache_control markers means the
  // Anthropic API bills every input token at standard rates (no 1.25× cache-
  // write surcharge, no 0.1× cache-read discount). The previous behaviour
  // applied markers unconditionally regardless of UserConfig.promptCache,
  // making the SettingsDialog checkbox a no-op — see ChatRequest.promptCache
  // doc-comment above for the orphaned-setting story. The shape returned here
  // matches the post-decoration shape below: a system [{type, text}] block
  // (no cache_control), tools as plain {name, description, input_schema}
  // objects, and messages with their content untouched.
  if (!req.promptCache) {
    return {
      system: [{ type: 'text', text: req.system }],
      tools: req.tools.map((t) => ({
        name: t.name,
        description: t.description,
        input_schema: t.input_schema,
      })),
      messages: req.messages.map((m) => ({ role: m.role, content: m.content })),
    };
  }
  const system = [
    {
      type: 'text',
      text: req.system,
      cache_control: { type: 'ephemeral' },
    },
  ];

  const tools = req.tools.map((t, i) => {
    const tool: Record<string, unknown> = {
      name: t.name,
      description: t.description,
      input_schema: t.input_schema,
    };
    if (i === req.tools.length - 1) {
      tool.cache_control = { type: 'ephemeral' };
    }
    return tool;
  });

  const breakpointSet = new Set(req.cacheBreakpoints);
  const messages = req.messages.map((m, idx) => {
    const content = m.content as Array<Record<string, unknown>> | string;
    if (!breakpointSet.has(idx) || typeof content === 'string') {
      return { role: m.role, content };
    }
    // Append a cache_control marker to the last block of this turn.
    const blocks = content.map((b, i) => {
      if (i === content.length - 1) {
        return { ...b, cache_control: { type: 'ephemeral' } };
      }
      return b;
    });
    return { role: m.role, content: blocks };
  });

  return { system, tools, messages };
}

export async function chat(req: ChatRequest): Promise<void> {
  const apiKey = await readApiKey();
  if (!apiKey) {
    send(req.requestId, {
      type: 'error',
      message: 'Anthropic API key not configured. Open Settings to add one.',
    });
    return;
  }
  const client = new Anthropic({ apiKey });
  // R264 — compute `withCache(req)` BEFORE registering the AbortController
  // in `inflight`. The original order was:
  //
  //   inflight.set(req.requestId, ctrl);
  //   const { system, tools, messages } = withCache(req);
  //   try { ... } finally { inflight.delete(req.requestId); }
  //
  // `withCache` walks `req.messages` and indexes into `cacheBreakpoints`;
  // a malformed payload (renderer bug, future preload refactor sending
  // an unexpected shape, SDK schema drift) can throw — and that throw
  // happens BEFORE the `try` block, so the `finally` clause never fires
  // and the AbortController leaks into the module-level Map permanently.
  // No realistic recovery: subsequent `cancel(requestId)` calls find the
  // stale entry and abort a never-attached controller (harmless but
  // misleading), and `cancelAll()` on window-all-closed iterates and
  // aborts the leaked entries (also harmless). Memory leak is bounded
  // by the rate of malformed requests, which is small but unbounded in
  // a long-running session.
  //
  // Moving withCache before the `inflight.set` line means any synchronous
  // throw from it propagates without ever touching the Map. The IPC
  // handler's reject path delivers the error to the renderer just as
  // before, but main stays clean.
  const { system, tools, messages } = withCache(req);

  const ctrl = new AbortController();
  inflight.set(req.requestId, ctrl);

  try {
    const stream = client.messages.stream(
      {
        model: req.model,
        max_tokens: req.maxTokens,
        temperature: req.temperature,
        system,
        tools: tools.length ? tools : undefined,
        messages,
      } as unknown as Anthropic.Messages.MessageStreamParams,
      { signal: ctrl.signal },
    );

    // Track tool_use events (deltas arrive as input_json_delta).
    const partialInputs = new Map<string, string>();
    const toolNames = new Map<string, string>();
    // R195 — track whether we've already sent a `message_stop` chunk to the
    // renderer. We send it on `message_delta` with `stop_reason` set
    // (canonical path), but if the upstream stream ends without ever
    // emitting that delta (network drop after partial response that the
    // SDK swallows, server returning a degenerate empty turn, future SDK
    // versions changing event ordering), the renderer's listener never
    // fires `onDone` (it gates on `chunk.type === 'message_stop'`), the
    // runTurn Promise leaks, and AIPanel.send's R162 sendingRef stays
    // true. Fallback at end-of-loop sends a synthetic message_stop so
    // every successful (non-throwing) chat() invocation terminates the
    // renderer chain. Only fires when the canonical path didn't.
    let stopSent = false;

    for await (const event of stream) {
      if (event.type === 'message_start') {
        // R227 — capture initial usage from message_start. The Anthropic
        // streaming protocol splits usage across two events: input_tokens
        // / cache_read_input_tokens / cache_creation_input_tokens land on
        // `message_start.message.usage`, while output_tokens accumulates
        // on `message_delta.usage` (per the SDK type at messages.d.ts:
        // `MessageDeltaUsage` has ONLY `output_tokens`). We were only
        // processing message_delta, so the entire input + cache picture
        // was silently dropped — `inputTokens ?? 0`, `cacheReadInputTokens
        // ?? 0`, `cacheCreationInputTokens ?? 0` always read undefined →
        // 0 from MessageDeltaUsage. The StatusBar's「N tokens · X%
        // cache」row showed output-only counts and forever 0% cache hit
        // ratio, even when prompt-cache was clearly working (the API
        // billing showed cache reads, but our display didn't).
        // The cast covers both SDK versions: older types have `usage:
        // Usage` (input + output only), newer types may add cache fields,
        // and Anthropic's wire protocol has had cache_* in the response
        // long before the SDK types caught up. Reading via optional cast
        // captures whatever is actually present at runtime.
        const m = event.message as { usage?: {
          input_tokens?: number;
          output_tokens?: number;
          cache_read_input_tokens?: number;
          cache_creation_input_tokens?: number;
        } };
        if (m.usage) {
          send(req.requestId, {
            type: 'usage',
            inputTokens: m.usage.input_tokens ?? 0,
            outputTokens: m.usage.output_tokens ?? 0,
            cacheReadInputTokens: m.usage.cache_read_input_tokens ?? 0,
            cacheCreationInputTokens: m.usage.cache_creation_input_tokens ?? 0,
          });
        }
      } else if (event.type === 'content_block_start') {
        const block = event.content_block;
        if (block.type === 'tool_use') {
          toolNames.set(block.id, block.name);
          partialInputs.set(block.id, '');
          send(req.requestId, { type: 'tool_use_start', id: block.id, name: block.name });
        }
      } else if (event.type === 'content_block_delta') {
        const delta = event.delta;
        if (delta.type === 'text_delta') {
          send(req.requestId, { type: 'text_delta', text: delta.text });
        } else if (delta.type === 'input_json_delta') {
          // The block id we need is on the parent content_block_start event
          // but Anthropic SDK gives us index → look up in partialInputs by
          // tracking insertion order, simpler: use the most recently-started.
          const lastId = [...partialInputs.keys()].pop();
          if (lastId !== undefined) {
            partialInputs.set(lastId, (partialInputs.get(lastId) ?? '') + delta.partial_json);
            send(req.requestId, {
              type: 'tool_use_input_delta',
              id: lastId,
              partialJson: delta.partial_json,
            });
          }
        }
      } else if (event.type === 'content_block_stop') {
        // Resolve any tool_use we accumulated input for.
        const lastId = [...partialInputs.keys()].pop();
        if (lastId !== undefined && toolNames.has(lastId)) {
          let parsed: Record<string, unknown> = {};
          try {
            parsed = JSON.parse(partialInputs.get(lastId) || '{}') as Record<string, unknown>;
          } catch {
            parsed = {};
          }
          send(req.requestId, {
            type: 'tool_use_complete',
            id: lastId,
            name: toolNames.get(lastId)!,
            input: parsed,
          });
          // Clear so the next tool_use reuses the same lookup.
          partialInputs.delete(lastId);
          toolNames.delete(lastId);
        }
      } else if (event.type === 'message_delta') {
        // R223 — emit `usage` BEFORE `message_stop`. The renderer's
        // sendChatTurn listener (provider.ts) treats `message_stop` as
        // the terminal chunk: it calls `handlers.onDone()` then `off()`,
        // unregistering the IPC listener. If `usage` is sent AFTER
        // `message_stop`, it lands in the IPC queue, the renderer
        // dispatches `message_stop` first (listener fires → off), then
        // dispatches `usage` (listener list now empty → chunk dropped).
        // The orchestrator's `addUsage` for the FINAL turn was being
        // silently lost — most visibly, the StatusBar's「N tokens · X%
        // cache」row missed the input/output/cache counts of the very
        // last assistant turn (which is also the most relevant — it's
        // the one the user just finished waiting on). Anthropic's
        // message_delta event itself carries both fields together when
        // the turn wraps up, so reordering at the send-site preserves
        // the bundle's intent: addUsage happens first, then onDone
        // closes the listener.
        if (event.usage) {
          const usage = event.usage as {
            input_tokens?: number;
            output_tokens?: number;
            cache_read_input_tokens?: number;
            cache_creation_input_tokens?: number;
          };
          send(req.requestId, {
            type: 'usage',
            inputTokens: usage.input_tokens ?? 0,
            outputTokens: usage.output_tokens ?? 0,
            cacheReadInputTokens: usage.cache_read_input_tokens ?? 0,
            cacheCreationInputTokens: usage.cache_creation_input_tokens ?? 0,
          });
        }
        // stop_reason on delta indicates the message is wrapping up.
        if (event.delta.stop_reason) {
          send(req.requestId, { type: 'message_stop', stopReason: event.delta.stop_reason });
          stopSent = true;
        }
      }
    }
    // R195 — fallback message_stop if the loop completed without one.
    // See `stopSent` doc-block. `end_turn` is the conventional non-error
    // stopReason, matching what the model would emit on a clean wrap-up.
    if (!stopSent) {
      send(req.requestId, { type: 'message_stop', stopReason: 'end_turn' });
    }
  } catch (err) {
    if ((err as { name?: string }).name === 'AbortError') {
      send(req.requestId, { type: 'message_stop', stopReason: 'user_cancelled' });
    } else {
      send(req.requestId, {
        type: 'error',
        message: err instanceof Error ? err.message : String(err),
      });
    }
  } finally {
    inflight.delete(req.requestId);
  }
}

export function cancel(requestId: string): void {
  const ctrl = inflight.get(requestId);
  if (ctrl) ctrl.abort();
  inflight.delete(requestId);
}

/**
 * R222 — abort every in-flight Anthropic stream. Used when the renderer
 * window goes away (`window-all-closed`) so we stop spending tokens on
 * responses no one will ever read. On Windows / Linux the process exits
 * anyway via `app.quit()` → `before-quit` → `app.exit(0)`, so the abort
 * is mostly redundant there; on macOS the standard convention keeps the
 * process alive after the last window closes (so the Dock icon can re-
 * launch a fresh window) — without this hook, an in-flight chat just
 * keeps streaming chunks into a `BrowserWindow.getAllWindows()` that's
 * now empty (`win?.webContents.send` no-ops), the Anthropic SDK reads
 * the rest of the SSE response, and the user gets billed for tokens
 * the renderer never received. Aborting the AbortController short-
 * circuits the SDK's stream iterator: it throws `AbortError`, the catch
 * in `chat()` sends a `message_stop` chunk (also a no-op since no
 * window), and the process settles back to idle.
 */
export function cancelAll(): void {
  for (const ctrl of inflight.values()) {
    ctrl.abort();
  }
  inflight.clear();
}

export async function ping(model: string): Promise<{ ok: boolean; error?: string; model?: string }> {
  const apiKey = await readApiKey();
  if (!apiKey) return { ok: false, error: 'no_api_key' };
  try {
    const client = new Anthropic({ apiKey });
    await client.messages.create({
      model,
      max_tokens: 16,
      messages: [{ role: 'user', content: 'ping' }],
    });
    return { ok: true, model };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}
