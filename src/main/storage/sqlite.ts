/**
 * SQLite backing store for chat history + undo stack — see spec §7.3.
 *
 * `better-sqlite3` is a synchronous native module; we keep all calls in main
 * (where the binary lives) and expose async-looking IPC handlers above it.
 * The DB file lives in the per-instance temp dir; persistence into the .gd
 * archive is handled by the workspace save flow.
 */

import Database from 'better-sqlite3';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { promises as fs } from 'node:fs';
import { ensureSessionTempRoot } from './gd';
import type { ChatRow, PersistChatRow, UndoRow } from '../../types/ipc';

let db: Database.Database | null = null;
/**
 * Dedupe concurrent first-time openDatabase calls. R193 — same shape as
 * R160's `pendingCreate` for ensureConversation: the existing
 * `if (db) return db;` early-return is correct after init completes, but
 * during the very first init two concurrent callers both see `db === null`,
 * both `await ensureSessionTempRoot()` (no dedupe there either, but it's
 * idempotent), both reach `db = new Database(file)`, and the second `new
 * Database` overwrites the reference — the first instance's underlying
 * file handle leaks (better-sqlite3 holds a native handle that's only
 * released on explicit `.close()` or process exit). Realistic trigger on
 * app startup: renderer's auto-open IPC fires `persistMessage` (→
 * `appendMessage` → openDatabase) concurrently with R167's `refreshCanUndo`
 * useEffect (→ `listUndo` → openDatabase) before either's mkdir/Database
 * sync portion runs.
 *
 * Module-scope is correct (singleton DB per main process); the chained
 * promise lets all callers await the same init.
 */
let openPromise: Promise<Database.Database> | null = null;

export async function openDatabase(): Promise<Database.Database> {
  if (db) return db;
  if (openPromise) return openPromise;
  openPromise = (async () => {
    try {
      const root = await ensureSessionTempRoot();
      const file = path.join(root, 'chat.sqlite');
      await fs.mkdir(path.dirname(file), { recursive: true });
      const next = new Database(file);
      // R246 — close `next` if pragma/exec throws so the native file handle
      // doesn't leak. better-sqlite3's Database constructor opens the
      // OS-level file handle eagerly; if either init step fails after the
      // handle is open, `next` falls out of scope but the native handle
      // stays until process exit (the JS GC can't release it — only an
      // explicit `.close()` does). Concrete failure modes:
      //   • pragma('journal_mode = WAL') — locked by another better-sqlite3
      //     instance against the same file (concurrent renderer reload
      //     during dev), read-only mounted volume.
      //   • exec(SCHEMA) — disk full mid-CREATE TABLE, EACCES on the temp
      //     dir tightened by an antivirus product, future schema typo
      //     during a migration.
      // The leaked handle pins the file open. R194's `closeDatabase()`
      // only inspects the module-level `db`, which was never assigned
      // here (the throw happens before line `db = next`), so shutdown
      // can't close it. cleanupSessionTempRoot then hits EBUSY on
      // chat.sqlite (and -wal / -shm), the .catch swallows, and the
      // temp dir leaks for sweepStaleTempRoots's 7-day TTL — exact
      // same disk-leak class R194 closed for the close-before-cleanup
      // ordering bug. Explicit try/catch closes the symmetric init-time
      // gap. better-sqlite3's `.close()` is idempotent and synchronous,
      // safe to call before re-throwing.
      try {
        next.pragma('journal_mode = WAL');
        next.exec(SCHEMA);
        db = next;
        return next;
      } catch (err) {
        next.close();
        throw err;
      }
    } finally {
      openPromise = null;
    }
  })();
  return openPromise;
}

/**
 * R194 — close DB on shutdown so temp-dir cleanup can remove chat.sqlite.
 * Windows enforces file locks while a handle is open; without this, the
 * subsequent `fs.rm(sessionTempRoot, ...)` in `cleanupSessionTempRoot`
 * fails with EBUSY on `chat.sqlite` (and `chat.sqlite-wal` / `-shm`),
 * the cleanup `.catch(() => undefined)` swallows the error, and the
 * temp dir stays on disk until `sweepStaleTempRoots` removes it 7 days
 * later. Disk-space leak across every Windows session for the lifetime
 * of that 7-day window. better-sqlite3's `close()` is idempotent and
 * synchronously releases the native handle plus flushes WAL — safe to
 * call from the main-process shutdown path.
 */
export function closeDatabase(): void {
  if (db) {
    db.close();
    db = null;
  }
}

const SCHEMA = `
CREATE TABLE IF NOT EXISTS conversations (
  id TEXT PRIMARY KEY,
  title TEXT,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL,
  workspace_id TEXT
);

CREATE TABLE IF NOT EXISTS messages (
  id TEXT PRIMARY KEY,
  conversation_id TEXT NOT NULL REFERENCES conversations(id),
  role TEXT NOT NULL,
  content TEXT NOT NULL,
  tool_use_id TEXT,
  created_at INTEGER NOT NULL,
  token_input INTEGER,
  token_output INTEGER,
  cache_read INTEGER,
  cache_creation INTEGER
);

CREATE INDEX IF NOT EXISTS idx_messages_conv ON messages(conversation_id, created_at);

CREATE TABLE IF NOT EXISTS undo_entries (
  id TEXT PRIMARY KEY,
  changeset_json TEXT NOT NULL,
  applied_at INTEGER NOT NULL,
  workspace_id TEXT
);

CREATE INDEX IF NOT EXISTS idx_undo_workspace ON undo_entries(workspace_id, applied_at);
`;

// ── conversations ────────────────────────────────────────────────────────

export async function createConversation(opts: {
  id?: string;
  title: string;
  workspaceId: string | null;
}): Promise<{ id: string; createdAt: number; updatedAt: number }> {
  const d = await openDatabase();
  const id = opts.id ?? randomUUID();
  const now = Date.now();
  d.prepare(
    `INSERT INTO conversations (id, title, created_at, updated_at, workspace_id)
     VALUES (?, ?, ?, ?, ?)`,
  ).run(id, opts.title, now, now, opts.workspaceId);
  return { id, createdAt: now, updatedAt: now };
}

export async function listConversations(
  workspaceId: string | null,
): Promise<Array<{ id: string; title: string; createdAt: number; updatedAt: number }>> {
  const d = await openDatabase();
  const rows = d
    .prepare(
      `SELECT id, title, created_at as createdAt, updated_at as updatedAt
       FROM conversations
       WHERE workspace_id IS ?
       ORDER BY updated_at DESC`,
    )
    .all(workspaceId) as Array<{ id: string; title: string; createdAt: number; updatedAt: number }>;
  return rows;
}

// ── messages ─────────────────────────────────────────────────────────────

export async function appendMessage(
  conversationId: string,
  row: PersistChatRow,
): Promise<ChatRow> {
  const d = await openDatabase();
  const id = row.id ?? randomUUID();
  const createdAt = row.createdAt ?? Date.now();
  // R190 — wrap the INSERT message + UPDATE conversation in a SQLite
  // transaction so the pair is atomic. Without this, an uncommon but real
  // failure mode (DB locked by an OS-level antivirus scan, disk-full
  // mid-statement, IPC abort between the two .run calls) leaves the message
  // row inserted but the conversation.updated_at not bumped — listConversations
  // sorts DESC by updated_at, so the just-touched conversation drops below
  // its actual last-activity position. better-sqlite3's `d.transaction(fn)`
  // returns a callable that runs fn inside `BEGIN ... COMMIT` and rolls back
  // on any throw, so either both writes land or neither does.
  const tx = d.transaction(() => {
    d.prepare(
      `INSERT INTO messages
       (id, conversation_id, role, content, tool_use_id, created_at,
        token_input, token_output, cache_read, cache_creation)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    ).run(
      id,
      conversationId,
      row.role,
      row.content,
      row.toolUseId,
      createdAt,
      row.tokenInput,
      row.tokenOutput,
      row.cacheRead,
      row.cacheCreation,
    );
    d.prepare(`UPDATE conversations SET updated_at = ? WHERE id = ?`).run(createdAt, conversationId);
  });
  tx();
  return {
    id,
    conversationId,
    role: row.role,
    content: row.content,
    toolUseId: row.toolUseId ?? null,
    createdAt,
    tokenInput: row.tokenInput ?? null,
    tokenOutput: row.tokenOutput ?? null,
    cacheRead: row.cacheRead ?? null,
    cacheCreation: row.cacheCreation ?? null,
  };
}

export async function listMessages(conversationId: string): Promise<ChatRow[]> {
  const d = await openDatabase();
  return d
    .prepare(
      `SELECT id, conversation_id as conversationId, role, content,
              tool_use_id as toolUseId, created_at as createdAt,
              token_input as tokenInput, token_output as tokenOutput,
              cache_read as cacheRead, cache_creation as cacheCreation
       FROM messages
       WHERE conversation_id = ?
       ORDER BY created_at ASC`,
    )
    .all(conversationId) as ChatRow[];
}

// ── undo entries ────────────────────────────────────────────────────────

export async function pushUndo(entry: {
  changesetJson: string;
  workspaceId: string | null;
}): Promise<UndoRow> {
  const d = await openDatabase();
  const id = randomUUID();
  const appliedAt = Date.now();
  // R190 — same INSERT-then-DELETE atomicity concern as appendMessage
  // above. If the trim DELETE fails after the INSERT lands, the workspace's
  // undo stack temporarily exceeds 50 rows; not visually broken (next push
  // re-tries the trim) but worth keeping atomic for consistency with the
  // appendMessage transaction pattern. better-sqlite3's `d.transaction(fn)`
  // wraps both statements in BEGIN/COMMIT; either both land or neither.
  const tx = d.transaction(() => {
    d.prepare(
      `INSERT INTO undo_entries (id, changeset_json, applied_at, workspace_id)
       VALUES (?, ?, ?, ?)`,
    ).run(id, entry.changesetJson, appliedAt, entry.workspaceId);
    // Trim to last 50 per workspace (spec §5.4.5).
    d.prepare(
      `DELETE FROM undo_entries
       WHERE workspace_id IS ? AND id NOT IN (
         SELECT id FROM undo_entries
         WHERE workspace_id IS ?
         ORDER BY applied_at DESC LIMIT 50
       )`,
    ).run(entry.workspaceId, entry.workspaceId);
  });
  tx();
  return { id, changesetJson: entry.changesetJson, appliedAt, workspaceId: entry.workspaceId };
}

export async function popUndo(workspaceId: string | null): Promise<UndoRow | null> {
  const d = await openDatabase();
  // R190 — wrap SELECT + DELETE in a transaction so the pop is atomic. If
  // the DELETE failed after the SELECT (DB locked, disk full mid-statement),
  // the row would remain in the stack and the next popUndo would return the
  // same row again — caller (App.tsx::handleUndo) re-applies the same
  // changeset and re-pushes to aiRedo, effectively duplicating the undo.
  // The transaction reverts the SELECT's read on DELETE failure (caller
  // gets a thrown error instead of a silent re-popable row), and the
  // outer App.tsx await catches via its existing try/catch.
  const tx = d.transaction((): UndoRow | null => {
    const row = d
      .prepare(
        `SELECT id, changeset_json as changesetJson, applied_at as appliedAt,
                workspace_id as workspaceId
         FROM undo_entries
         WHERE workspace_id IS ?
         ORDER BY applied_at DESC LIMIT 1`,
      )
      .get(workspaceId) as UndoRow | undefined;
    if (!row) return null;
    d.prepare(`DELETE FROM undo_entries WHERE id = ?`).run(row.id);
    return row;
  });
  return tx();
}

export async function listUndo(workspaceId: string | null, limit: number): Promise<UndoRow[]> {
  const d = await openDatabase();
  return d
    .prepare(
      `SELECT id, changeset_json as changesetJson, applied_at as appliedAt,
              workspace_id as workspaceId
       FROM undo_entries
       WHERE workspace_id IS ?
       ORDER BY applied_at DESC LIMIT ?`,
    )
    .all(workspaceId, limit) as UndoRow[];
}

export async function clearUndo(workspaceId: string | null): Promise<void> {
  const d = await openDatabase();
  d.prepare(`DELETE FROM undo_entries WHERE workspace_id IS ?`).run(workspaceId);
}

// R386 — re-tag persistent rows when a workspace's identity changes mid-session
// (Save-As that mints a new filePath → workspaceIdFor produces a new hash, see
// store/workspace.ts:232). Without this, every undo_entries / conversations row
// pushed BEFORE the Save-As stays tagged to the OLD workspaceId; the
// post-Save-As workspace queries SQLite under the NEW workspaceId, finds zero
// rows, and the user observes:
//   • App.tsx:970 refreshCanUndo returns 0 rows → toolbar Undo button greys out
//     despite the user clearly having a stack of undoable edits
//   • AIPanel's conversation history list (history.listConversations) shows the
//     workspace as having no prior chats, even though the user ran a long AI
//     session in the same buffer pre-Save-As
// R385 fixed the in-memory workspaceId mapping; this completes the move by
// migrating the persistent rows in the same atomic step. Both UPDATEs run in
// a single SQLite transaction so we don't observe a half-relinked state — if
// either statement fails (DB locked mid-write), neither lands and the OLD ids
// remain intact for retry. No-op when oldId === newId so plain Ctrl+S (no path
// change) doesn't churn the DB.
export async function relinkWorkspaceId(
  oldId: string | null,
  newId: string,
): Promise<void> {
  if (oldId === newId) return;
  const d = await openDatabase();
  const tx = d.transaction(() => {
    d.prepare(`UPDATE undo_entries SET workspace_id = ? WHERE workspace_id IS ?`).run(
      newId,
      oldId,
    );
    d.prepare(`UPDATE conversations SET workspace_id = ? WHERE workspace_id IS ?`).run(
      newId,
      oldId,
    );
  });
  tx();
}
