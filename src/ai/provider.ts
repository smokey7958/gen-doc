/**
 * Provider abstraction — see spec §6.1.
 *
 * We don't ship a real second provider in v1.0, but every renderer call site
 * goes through `sendChatTurn`, so swapping in OpenAI / Ollama later means
 * adding a new switch arm in main and a new IPC backend.
 */

import { v4 as uuid } from 'uuid';
import type { ChatMessage, StreamChunk, ToolDefinition } from '../types/ai';

export interface SendChatRequest {
  model: string;
  system: string;
  messages: ChatMessage[];
  tools: ToolDefinition[];
  cacheBreakpoints: number[];
  maxTokens: number;
  temperature: number;
  /** Mirrors UserConfig.promptCache via the AI store; main/ai/anthropic.ts's
   * withCache() short-circuits when false. See store/ai.ts:46-53 for the
   * mirror-into-store rationale that this field follows verbatim. */
  promptCache: boolean;
}

export interface ChatTurnHandlers {
  onChunk(chunk: StreamChunk): void;
  onError(message: string): void;
  onDone(): void;
}

export interface InflightHandle {
  requestId: string;
  cancel(): void;
}

/**
 * Forward a chat turn over IPC and stream chunks back to the caller's
 * handlers. Each call subscribes its own listener; the cleanup happens
 * automatically on `message_stop` / `error`.
 */
export function sendChatTurn(req: SendChatRequest, handlers: ChatTurnHandlers): InflightHandle {
  const requestId = uuid();
  const off = window.gendoc.ai.onChunk((evt) => {
    if (evt.requestId !== requestId) return;
    const chunk = evt.chunk as StreamChunk;
    if (chunk.type === 'error') {
      handlers.onError(chunk.message);
      handlers.onDone();
      off();
      return;
    }
    handlers.onChunk(chunk);
    if (chunk.type === 'message_stop') {
      handlers.onDone();
      off();
    }
  });

  void window.gendoc.ai
    .chat({
      requestId,
      model: req.model,
      system: req.system,
      messages: req.messages.map((m) => ({
        role: m.role === 'tool_result' ? 'user' : m.role,
        content: m.content,
      })),
      tools: req.tools,
      maxTokens: req.maxTokens,
      temperature: req.temperature,
      cacheBreakpoints: req.cacheBreakpoints,
      promptCache: req.promptCache,
    })
    .catch((err) => {
      handlers.onError(err instanceof Error ? err.message : String(err));
      handlers.onDone();
      off();
    });

  return {
    requestId,
    cancel: () => {
      void window.gendoc.ai.cancel(requestId);
      off();
    },
  };
}
