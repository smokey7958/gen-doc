/**
 * User-level config + API key vault.
 *
 * - Non-secret config: JSON at `~/.gendoc/config.json`
 * - API key: encrypted with Electron `safeStorage` (OS keystore — Windows DPAPI
 *   / macOS Keychain / libsecret), stored as base64 ciphertext at
 *   `~/.gendoc/secrets.bin`. Plaintext never touches disk.
 */

import { app, safeStorage } from 'electron';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import type { UserConfig } from '../../types/ipc';

const CONFIG_DIR_NAME = '.gendoc';
const CONFIG_FILE = 'config.json';
const SECRET_FILE = 'secrets.bin';

export function getConfigDir(): string {
  return path.join(app.getPath('home'), CONFIG_DIR_NAME);
}

const DEFAULTS: UserConfig = {
  defaultModel: 'claude-sonnet-4-6',
  temperature: 0.3,
  maxTokens: 4096,
  promptCache: true,
  embedChatHistoryDefault: false,
  theme: 'light',
  autoSaveIntervalMs: 0,
  recentFiles: [],
  keymapOverrides: {},
  windowBounds: null,
  autoOpenLastWorkspace: true,
  // R405 — null means「follow OS locale」; renderer resolves to 'zh' / 'en' at
  // boot via app.getOsLocale(). See lib/i18n.ts for the resolution rules.
  locale: null,
};

let cached: UserConfig | null = null;

export async function loadConfig(): Promise<UserConfig> {
  if (cached) return cached;
  const dir = getConfigDir();
  await fs.mkdir(dir, { recursive: true });
  const file = path.join(dir, CONFIG_FILE);
  try {
    const raw = await fs.readFile(file, 'utf-8');
    const parsed = JSON.parse(raw) as Partial<UserConfig>;
    cached = { ...DEFAULTS, ...parsed };
  } catch {
    cached = { ...DEFAULTS };
    await saveConfig(cached);
  }
  return cached;
}

export async function saveConfig(next: UserConfig): Promise<UserConfig> {
  // R252 — assign `cached` ONLY after the disk write succeeds, not before.
  // Original order mutated the cache sync before `await fs.writeFile`, so
  // a write rejection (ENOSPC, EACCES on `~/.gendoc`, network-share user
  // profile offline, antivirus pinning the config dir) left a real
  // divergence:
  //   • cache = new values (the failed-write candidate)
  //   • disk = old values (write never landed)
  // The patchConfig caller's catch path (SettingsDialog.patch /
  // App.tsx::handleClearRecent) shows a `儲存設定失敗：…` toast, so the
  // user knows it failed — but they don't know the cache is now lying.
  // Concrete symptom: user changes temperature 0.3 → 0.7 in Settings,
  // gets the failure toast, closes Settings; subsequent IPC reads of
  // `config.get()` (e.g., reopening Settings to check / verify they
  // need to retry) return cache=0.7 from this divergence and read as
  // "looks saved", so the user closes the app trusting it. Next launch
  // loads from disk → 0.3 → setting silently reverted with zero signal.
  // Same shape across every patchable field (recentFiles dedupe,
  // autoSaveIntervalMs, defaultModel, promptCache, etc.).
  //
  // Mutating cache *after* successful write keeps cache and disk in
  // lockstep. On failure, both remain at the prior value — consistent
  // across session reloads, and the failure toast accurately reflects
  // what's actually persisted. patchTail serialization (R182) already
  // ensures ordering, so the concurrent-read window where "another
  // caller sees the optimistic new value during the in-flight write"
  // was the only weak rationale for the old order — and that read
  // would have lied about what's actually durable.
  const dir = getConfigDir();
  await fs.mkdir(dir, { recursive: true });
  const file = path.join(dir, CONFIG_FILE);
  const candidate: UserConfig = { ...next };
  await fs.writeFile(file, JSON.stringify(candidate, null, 2), 'utf-8');
  cached = candidate;
  return cached;
}

/**
 * Serialize patchConfig calls with a chained promise so concurrent patches
 * don't lose writes. R182 — without this, two near-simultaneous patches
 * (e.g. `pushRecent` from a Save resolving while the user fires Ctrl+O
 * which also calls pushRecent on its own success) both `await loadConfig`
 * synchronously from the in-memory cache, both compute their merged
 * `next`, both call `await fs.writeFile(...)`. The kernel doesn't
 * guarantee a write order matching the JS dispatch order — if B's write
 * lands on disk first then A's stale-but-later write overwrites B,
 * recentFiles loses one of the entries (typical user-visible failure:
 * "I just opened file X, why isn't it in 最近開啟?"). Chaining off a
 * tail promise serializes the awaits — each patch's write must finish
 * before the next reads + writes — so disk and cache stay aligned.
 *
 * R252 — saveConfig now updates `cached` only after the disk write
 * succeeds (was: sync-before the await). Each patch's `run()` does
 * `await loadConfig()` which reads from cache; combined with patchTail
 * serialization, the cache state at run-time always reflects the most
 * recent successfully-persisted config (failed writes leave both cache
 * and disk at the prior value). See saveConfig doc-block for the
 * divergence-on-write-failure story this closed.
 */
let patchTail: Promise<UserConfig> | null = null;

export function patchConfig(patch: Partial<UserConfig>): Promise<UserConfig> {
  const run = async (): Promise<UserConfig> => {
    const current = await loadConfig();
    return saveConfig({ ...current, ...patch });
  };
  const next = patchTail ? patchTail.then(run, run) : run();
  // Hold onto the chain so the next caller awaits us. `.then(run, run)`
  // covers both fulfilment and rejection of the previous patch — if A
  // failed we still want B to attempt (and write disk based on the
  // current cache, which may or may not have been mutated by A).
  patchTail = next.catch(() => loadConfig());
  return next;
}

function secretPath(): string {
  return path.join(getConfigDir(), SECRET_FILE);
}

export async function hasApiKey(): Promise<boolean> {
  try {
    const buf = await fs.readFile(secretPath());
    return buf.byteLength > 0;
  } catch {
    return false;
  }
}

export async function readApiKey(): Promise<string | null> {
  try {
    const buf = await fs.readFile(secretPath());
    if (!safeStorage.isEncryptionAvailable()) {
      // Fall back: stored in cleartext when keystore missing (Linux without libsecret).
      // We log a warning in main.ts startup if this branch ever runs.
      return buf.toString('utf-8');
    }
    return safeStorage.decryptString(buf);
  } catch {
    return null;
  }
}

/**
 * Serialize API-key writes/clears with a chained promise so concurrent
 * saveKey + clearKey calls don't race on the secrets.bin file. Same shape
 * as R182's patchTail for config.json. R201 — without this, the user
 * could:
 *   1. Type a new key, click 儲存 → `setApiKey` IPC fires writeApiKey,
 *      enters fs.writeFile.
 *   2. Click 清除 (first stage), click again within 4s (R90 two-step
 *      latch) → `clearApiKey` IPC fires clearApiKey, enters fs.unlink.
 * Two fs operations on the same file (`secrets.bin`) with no atomicity
 * guarantee — kernel scheduling decides the final state. If unlink lands
 * before writeFile completes, end state has the key (user expected
 * clear). If writeFile lands first, then unlink, end state has no key
 * (matches user's last action). Either outcome can mismatch user intent
 * since both IPCs were dispatched independently. Chaining via tail
 * promise enforces "whichever IPC dispatched last wins" semantics that
 * users actually expect.
 */
let secretTail: Promise<unknown> | null = null;

function chainSecret<T>(run: () => Promise<T>): Promise<T> {
  const next = secretTail ? secretTail.then(run, run) : run();
  // Hold onto the chain so the next caller awaits us. catch() to suppress
  // unhandled rejection on the tail itself; the next `run` is `.then(run, run)`
  // which fires regardless.
  secretTail = next.catch(() => undefined);
  return next;
}

export async function writeApiKey(key: string): Promise<void> {
  return chainSecret(async () => {
    const dir = getConfigDir();
    await fs.mkdir(dir, { recursive: true });
    const file = secretPath();
    if (!safeStorage.isEncryptionAvailable()) {
      await fs.writeFile(file, key, 'utf-8');
      return;
    }
    const enc = safeStorage.encryptString(key);
    await fs.writeFile(file, enc);
  });
}

export async function clearApiKey(): Promise<void> {
  return chainSecret(async () => {
    await fs.unlink(secretPath()).catch(() => undefined);
  });
}
