/**
 * AI layer types — Provider abstraction (spec §6.1) and message / tool
 * shapes shared between the renderer UI, the tool dispatcher, and the
 * Anthropic provider implementation in main.
 */

export interface ModelDescriptor {
  id: string;
  /** Human-readable label shown in the model picker. */
  label: string;
  /** Approximate input pricing per 1M tokens, for the in-app cost estimator. */
  inputPricePerMTok: number;
  outputPricePerMTok: number;
  cachedInputPricePerMTok: number;
  /** Max input tokens for the model. */
  contextWindow: number;
}

export const SUPPORTED_MODELS: ModelDescriptor[] = [
  {
    id: 'claude-sonnet-4-6',
    label: 'Claude Sonnet 4.6（預設）',
    inputPricePerMTok: 3,
    outputPricePerMTok: 15,
    cachedInputPricePerMTok: 0.3,
    contextWindow: 200_000,
  },
  {
    id: 'claude-opus-4-7',
    label: 'Claude Opus 4.7（深度思考）',
    inputPricePerMTok: 15,
    outputPricePerMTok: 75,
    cachedInputPricePerMTok: 1.5,
    contextWindow: 200_000,
  },
];

/**
 * R389 — coerce a candidate model id into one the current build actually
 * supports. SUPPORTED_MODELS changes between app versions (deprecated ids
 * removed, new ones added) but `config.json` on disk persists across
 * upgrades; a user who selected `claude-haiku-3` in v1.0 then upgrades to a
 * build whose SUPPORTED_MODELS dropped haiku-3 would otherwise load that
 * stale id straight into the AI store. Downstream symptoms:
 *   • AIPanel's `<select value={ai.model}>` finds no matching `<option>` and
 *     renders inconsistently (browsers usually show the first option
 *     visually but the underlying value stays stale, so the picker LIES
 *     about what the next prompt will hit)
 *   • Every chat IPC sends the unsupported id to Anthropic; the API rejects
 *     with HTTP 404 `model_not_found`. The user sees「呼叫失敗：model not
 *     found」 banners with no hint that it's a model-id mismatch from a
 *     previous app version
 *   • SettingsDialog's「測試 API key」 ping (SettingsDialog.tsx:249) reuses
 *     `config.defaultModel` for its test request — also 404s, making the
 *     user think their API key is invalid when it's actually fine
 * Fall back to the first SUPPORTED_MODELS entry (the canonical "default"
 * for the current build, by ordering convention) when the candidate is
 * absent / null / unrecognized. Same pattern as config.ts's `{...DEFAULTS,
 * ...parsed}` for missing fields, just one level up at the value-validity
 * layer instead of presence layer. We intentionally do NOT persist the
 * corrected id back to disk: keeping the stale id in config.json means
 * downgrading the app (or restoring an older config.json from backup)
 * still presents the user's original choice if that build supported it;
 * round-tripping through fallback every load is cheap and consistent.
 */
export function resolveSupportedModelId(
  candidate: string | undefined | null,
): string {
  if (candidate && SUPPORTED_MODELS.some((m) => m.id === candidate)) {
    return candidate;
  }
  return SUPPORTED_MODELS[0].id;
}

export type Role = 'user' | 'assistant' | 'tool_result';

export interface TextBlock {
  type: 'text';
  text: string;
  /** Optional cache breakpoint marker (passed through to Anthropic). */
  cacheControl?: 'ephemeral';
}

export interface ToolUseBlock {
  type: 'tool_use';
  id: string;
  name: string;
  input: Record<string, unknown>;
}

export interface ToolResultBlock {
  type: 'tool_result';
  tool_use_id: string;
  content: string;
  is_error?: boolean;
}

export type ContentBlock = TextBlock | ToolUseBlock | ToolResultBlock;

export interface ChatMessage {
  role: Role;
  content: ContentBlock[];
}

export interface ToolDefinition {
  name: string;
  description: string;
  input_schema: {
    type: 'object';
    properties: Record<string, unknown>;
    required?: string[];
  };
}

/**
 * Streamed chunk shape we forward over IPC. Keeps a stable wire format so the
 * renderer doesn't depend on the Anthropic SDK's internal event types.
 */
export type StreamChunk =
  | { type: 'text_delta'; text: string }
  | { type: 'tool_use_start'; id: string; name: string }
  | { type: 'tool_use_input_delta'; id: string; partialJson: string }
  | { type: 'tool_use_complete'; id: string; name: string; input: Record<string, unknown> }
  | { type: 'message_stop'; stopReason: string }
  | {
      type: 'usage';
      inputTokens: number;
      outputTokens: number;
      cacheReadInputTokens: number;
      cacheCreationInputTokens: number;
    }
  | { type: 'error'; message: string };
