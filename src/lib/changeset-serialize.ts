/**
 * R216 — JSON-safe ChangeSet serialization for the persistent undo stack.
 *
 * `BinaryReplaceOp` / `TabCreateOp` / `TabDeleteOp` carry `Uint8Array` payloads
 * (the pre/post bytes of a docx/xlsx/pptx tab — typically tens to hundreds of
 * KB). The persistent undo stack stores changesets as JSON strings in
 * SQLite (`undo_entries.changeset_json`), and a naive `JSON.stringify(cs)` /
 * `JSON.parse(json)` round-trip silently corrupts the bytes:
 *
 *   `JSON.stringify(new Uint8Array([1,2,3]))` → `'{"0":1,"1":2,"2":3}'`
 *
 * — Uint8Array isn't an Array (Array.isArray returns false), and it lacks a
 * `toJSON` method, so the default object-property enumerator emits
 * numeric-string keys. Parsing that back yields a *plain object* with
 * `{0:1,1:2,2:3}` shape, NOT a Uint8Array — `byteLength` is undefined, the
 * iterator protocol is gone, and all the OOXML zip parsers (parseXlsx,
 * parseDocx, parsePptx) implicitly assume Uint8Array.
 *
 * Concrete failure: user asks AI to set cell A1 → AI emits a
 * `binary_replace` op → user clicks Apply → workspace updates fine (the
 * in-memory ChangeSet still holds real Uint8Arrays) → user clicks Ctrl+Z →
 * `handleUndo` pops from SQLite, JSON.parses, undoChangeset assigns the
 * mangled plain object to `tab.data` → XlsxEditor re-mounts, parseXlsx
 * fails or produces garbage. The user's spreadsheet is now in a broken
 * intermediate state with no recovery short of closing without saving.
 *
 * This module provides a stable wire format that survives the round-trip:
 * Uint8Array values are wrapped as `{ __u8: number[] }` during stringify,
 * and the matching shape is unwrapped back into a fresh Uint8Array during
 * parse. The wrapper key is unlikely to collide with anything in the
 * ChangeOp schema (none of the typed ops carry a field named `__u8`), and
 * the reviver only triggers on objects that carry exactly that signature
 * with an array payload.
 *
 * Why number[] over base64:
 *   • base64 via `btoa(String.fromCharCode(...u8))` blows the call-stack
 *     argument limit (~65 K bytes in V8) for typical document-sized blobs;
 *     a chunked workaround is more code than the array form costs.
 *   • The on-disk size penalty is comparable: each byte renders as 2-4
 *     chars in the JSON array form (`,255`), versus the broken object
 *     form's 6-12 chars per byte (`,"4096":255`) — array is actually
 *     SMALLER than what we were already writing.
 *   • Same Array.from path also handles `TabCreateOp.data` /
 *     `TabDeleteOp.data` for free; no per-op-type dispatch needed.
 */

import type { ChangeSet } from '../types/changeset';

interface Uint8Wrapper {
  __u8: number[];
}

function isUint8Wrapper(value: unknown): value is Uint8Wrapper {
  return (
    value !== null &&
    typeof value === 'object' &&
    Array.isArray((value as { __u8?: unknown }).__u8)
  );
}

export function serializeChangeset(cs: ChangeSet): string {
  return JSON.stringify(cs, (_key, value) => {
    if (value instanceof Uint8Array) {
      return { __u8: Array.from(value) } satisfies Uint8Wrapper;
    }
    return value;
  });
}

export function deserializeChangeset(json: string): ChangeSet {
  return JSON.parse(json, (_key, value) => {
    if (isUint8Wrapper(value)) {
      return new Uint8Array(value.__u8);
    }
    return value;
  }) as ChangeSet;
}
