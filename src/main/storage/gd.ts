/**
 * `.gd` archive read/write — see spec §5.3.
 *
 * On open: unzip into a temp directory, parse manifest, return tab payloads.
 * On save: re-pack from the in-memory state passed by renderer, atomic write
 * (write to .tmp, fsync, rename).
 */

import { promises as fs } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import JSZip from 'jszip';
import type {
  Manifest,
  TabDescriptor,
} from '../../types/manifest';
import { MANIFEST_FILENAME, MANIFEST_VERSION } from '../../types/manifest';
import type { OpenedWorkspace, SaveWorkspaceRequest, TabPayload } from '../../types/ipc';

/** Per-instance temp dir; cleaned on app quit. */
const sessionTempRoot = path.join(tmpdir(), `gendoc-${randomUUID()}`);

export function getSessionTempRoot(): string {
  return sessionTempRoot;
}

export async function ensureSessionTempRoot(): Promise<string> {
  await fs.mkdir(sessionTempRoot, { recursive: true });
  return sessionTempRoot;
}

export async function cleanupSessionTempRoot(): Promise<void> {
  await fs.rm(sessionTempRoot, { recursive: true, force: true }).catch(() => undefined);
}

export async function sweepStaleTempRoots(maxAgeMs = 7 * 24 * 60 * 60 * 1000): Promise<void> {
  const root = tmpdir();
  const entries = await fs.readdir(root).catch(() => [] as string[]);
  const now = Date.now();
  await Promise.all(
    entries
      .filter((name) => name.startsWith('gendoc-'))
      .map(async (name) => {
        const full = path.join(root, name);
        if (full === sessionTempRoot) return;
        const stat = await fs.stat(full).catch(() => null);
        if (!stat) return;
        if (now - stat.mtimeMs > maxAgeMs) {
          await fs.rm(full, { recursive: true, force: true }).catch(() => undefined);
        }
      }),
  );
}

function defaultManifest(title: string): Manifest {
  const now = new Date().toISOString();
  return {
    version: MANIFEST_VERSION,
    title,
    createdAt: now,
    modifiedAt: now,
    tabs: [],
    settings: { embedChatHistory: false, defaultModel: 'claude-sonnet-4-6' },
    metadata: { appVersion: '1.0.0' },
  };
}

export function createEmptyWorkspace(title = '未命名'): OpenedWorkspace {
  return { filePath: '', manifest: defaultManifest(title), tabs: [] };
}

// R376 — tab type union for runtime validation. Kept in sync with
// TabType in types/manifest.ts (also `'markdown' | 'html' | 'docx' |
// 'xlsx' | 'pptx'`). Use Set for O(1) membership check.
const VALID_TAB_TYPES = new Set<string>([
  'markdown',
  'html',
  'docx',
  'xlsx',
  'pptx',
]);

// R416 — safe-relative-path pattern for manifest `file` entries. Legitimate
// writers only ever emit `doc/<uuid>.<ext>` (workspace.ts addTab /
// openExternalFile); a hand-crafted .gd could carry `../x`, `/etc/x`,
// `C:\x`, or backslash paths. The char class excludes backslash and colon
// (drive letters), and `^[\w]` rejects a leading separator or dot; `..`
// inside the path is screened per-segment since the class allows '.'.
const SAFE_ARCHIVE_PATH = /^[\w][\w\-./ ]*$/;

function isSafeArchivePath(p: string): boolean {
  if (!SAFE_ARCHIVE_PATH.test(p)) return false;
  return p.split('/').every((seg) => seg.length > 0 && seg !== '.' && seg !== '..');
}

function validateManifest(raw: unknown): Manifest {
  if (!raw || typeof raw !== 'object') {
    throw new Error('manifest.json missing or not an object');
  }
  const m = raw as Partial<Manifest>;
  if (m.version !== MANIFEST_VERSION) {
    throw new Error(`Unsupported manifest version: ${String(m.version)}`);
  }
  if (!Array.isArray(m.tabs)) {
    throw new Error('manifest.tabs is not an array');
  }
  // R377 — track ids seen so far to catch duplicates. The dispatcher's
  // `findTab`, workspace's `setActiveTab` / `patchTab` / `removeTab` all
  // route through `array.find((t) => t.id === id)` — first match wins. If
  // two tabs share the same id (corrupted .gd, manual edit, third-party
  // tool emitting bad manifests), every operation against that id
  // silently targets the FIRST one in the array, the second is
  // effectively orphaned UI-wise: visible in the tab strip, can't be
  // closed (× click looks up by id → hits first), can't be activated
  // (clicking the row sets activeTabId=id → setActiveTab → both rows
  // highlight in the strip's rendered list because both match the
  // active filter). Confusing & inactionable — the only escape is to
  // close the workspace.
  const seenIds = new Set<string>();
  // R390 — also track `file` paths seen so far. Two tabs sharing `file` is
  // a SILENT DATA-LOSS bug at save time: writeGdArchive at gd.ts:194-196
  // iterates `req.tabs` and does `zip.file(descriptor.file, bytes)` for
  // each; JSZip overwrites the entry when the same path is written twice.
  // So the SECOND tab's bytes win, and the FIRST tab's content is
  // permanently lost the moment the user hits Ctrl+S — including any
  // edits they just made. ReadGdArchive at line 167 maps each descriptor
  // to `zip.file(desc.file)`, so on load both tabs DO get bytes (both
  // resolve to the same zip entry's content), masking the issue until
  // the next save overwrites. Realistic triggers parallel R377's id
  // collision list: corrupted .gd (concurrent-save race in an older
  // version), manual edit of manifest.json, third-party tooling
  // emitting bad manifests, or — astronomically rare but possible — a
  // uuid-collision in `addTab` / `openExternalFile` constructing the
  // same `doc/${id}.${ext}` path twice. Reject upfront with the same
  //「fail loudly on load」 posture R377 set: better a clear「duplicate
  // file path」 error the .gd creator can fix than silent data loss
  // on the next Ctrl+S.
  const seenFiles = new Set<string>();
  for (const t of m.tabs as TabDescriptor[]) {
    if (!t.id || !t.name || !t.type || !t.file) {
      throw new Error(`Invalid tab descriptor: ${JSON.stringify(t)}`);
    }
    // R416 — see isSafeArchivePath doc-block above. Validated before any
    // zip.file(desc.file) lookup in readGdArchive.
    if (!isSafeArchivePath(t.file)) {
      throw new Error(
        `Invalid tab descriptor (unsafe file path "${t.file}"; archive paths must be relative — no "..", leading separators, backslashes, or drive letters): ${JSON.stringify(t)}`,
      );
    }
    if (seenIds.has(t.id)) {
      throw new Error(
        `Invalid tab descriptor (duplicate id "${t.id}"; tab ids must be unique within a manifest): ${JSON.stringify(t)}`,
      );
    }
    seenIds.add(t.id);
    // R390 — see seenFiles doc-block above.
    if (seenFiles.has(t.file)) {
      throw new Error(
        `Invalid tab descriptor (duplicate file path "${t.file}"; each tab must map to a distinct archive entry to avoid silent overwrite on save): ${JSON.stringify(t)}`,
      );
    }
    seenFiles.add(t.file);
    // R376 — extend validation to two fields the original checker missed:
    //
    // 1. `type` must be in TabType union. The old check only verified
    //    truthy, so `type: 'pdf'` (corrupted .gd / future format
    //    speculation / typo) passed validation, made it into the Tab
    //    union via `loadFromOpened`'s `as Tab` cast, then EditorSurface
    //    fell through all five `active.type === '…'` branches and
    //    rendered EmptyState — with no error message telling the user
    //    why their pdf file is showing as blank.
    //
    // 2. `order` must be a finite non-negative number. Both call sites
    //    that read it (workspace.ts:304 / gd.ts:103) do
    //    `tabs.sort((a, b) => a.order - b.order)`. If `order` is
    //    undefined / NaN / string, the subtraction produces NaN and the
    //    sort comparator becomes non-deterministic — V8's TimSort
    //    happens to be stable but the comparator returning NaN is
    //    spec-undefined; in practice the tab order ends up arbitrary,
    //    different sessions might display tabs in different orders for
    //    the same .gd, and side-effects like「Ctrl+1 jumps to first
    //    tab」 become unstable. Reject upfront with a clear error so
    //    the .gd creator knows what to fix instead of seeing「tabs
    //    re-arrange on every open」.
    if (!VALID_TAB_TYPES.has(t.type as string)) {
      // R390 — stray `}` removed (was `)}: ${JSON.stringify(t)}`). The typo
      // produced「…not one of markdown/html/docx/xlsx/pptx)}: {raw}」 with
      // an inexplicable `}` between the close paren and the colon. Sibling
      // descriptor errors (line 109's missing-field, line 113's duplicate-id,
      // line 145's invalid-order) all use the clean `): …` form; this branch
      // was the lone outlier.
      throw new Error(
        `Invalid tab descriptor (type "${String(t.type)}" not one of ${[...VALID_TAB_TYPES].join('/')}): ${JSON.stringify(t)}`,
      );
    }
    if (typeof t.order !== 'number' || !Number.isFinite(t.order) || t.order < 0) {
      throw new Error(
        `Invalid tab descriptor (order must be a finite non-negative number, got ${String(t.order)}): ${JSON.stringify(t)}`,
      );
    }
  }
  return raw as Manifest;
}

export async function readGdArchive(filePath: string): Promise<OpenedWorkspace> {
  const buf = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);
  const manifestEntry = zip.file(MANIFEST_FILENAME);
  if (!manifestEntry) {
    throw new Error(`${filePath} is not a valid .gd archive (missing manifest.json)`);
  }
  const manifest = validateManifest(JSON.parse(await manifestEntry.async('string')));

  const tabs: TabPayload[] = [];
  // Preserve declared order rather than zip enumeration order.
  const sorted = [...manifest.tabs].sort((a, b) => a.order - b.order);
  for (const desc of sorted) {
    const entry = zip.file(desc.file);
    if (!entry) {
      throw new Error(`Manifest references missing file: ${desc.file}`);
    }
    const bytes = await entry.async('uint8array');
    tabs.push({ descriptor: desc, bytes });
  }

  return { filePath, manifest, tabs };
}

export async function writeGdArchive(req: SaveWorkspaceRequest): Promise<{
  filePath: string;
  modifiedAt: string;
}> {
  if (!req.filePath) {
    throw new Error('writeGdArchive requires filePath; use Save-As flow upstream.');
  }
  const zip = new JSZip();
  const modifiedAt = new Date().toISOString();
  const manifest: Manifest = {
    ...req.manifest,
    modifiedAt,
    version: MANIFEST_VERSION,
  };
  zip.file(MANIFEST_FILENAME, JSON.stringify(manifest, null, 2));

  for (const { descriptor, bytes } of req.tabs) {
    zip.file(descriptor.file, bytes);
  }

  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    platform: 'UNIX', // forward-slash paths, UTF-8
  });

  await atomicWrite(req.filePath, out);
  // R212 — backup rotation is post-save housekeeping; its failure must NOT
  // surface to the renderer as a save error. atomicWrite has already
  // delivered the user's bytes to their .gd path; if the .gendoc/backups
  // directory is unreachable (HOME unset, USERPROFILE rejecting mkdir,
  // tmpdir permission flip on a hardened Linux configuration, network-
  // share user-profile gone offline), rotateBackups's `fs.mkdir(dir,
  // { recursive: true })` at gd.ts:181 throws — and that throw propagates
  // out of writeGdArchive → the IPC reply rejects → renderer sees「儲存
  // 失敗」 toast and the saveState flips to 'error', even though the .gd
  // is sitting on disk exactly where the user asked. The user retries,
  // hits the same mkdir error, eventually gives up assuming nothing was
  // saved — meanwhile their data has been on disk the whole time.
  // Internal-only catch fits the "best-effort" posture rotateBackups
  // already adopts internally (every fs op inside is `.catch(() =>
  // undefined)`); the OUTER mkdir was the lone unswallowed throw point.
  // Backup loss is a degraded-but-acceptable outcome in this failure
  // window; misreporting save success as failure is not.
  try {
    await rotateBackups(req.filePath, manifest.title);
  } catch {
    /* swallow — see R212 doc-block above. */
  }

  return { filePath: req.filePath, modifiedAt };
}

// R251 — exported so ipc.ts can reuse the write-tmp → fsync → rename →
// cleanup-on-fail invariant for the two export paths (workspace.exportTab
// + markdown.exportPdf). Both overwrite a user-picked destination on
// disk; a non-atomic `fs.writeFile` would leave the target truncated on
// partial-write failure (disk full mid-write, antivirus EBUSY, network
// drive disconnect), destroying the user's prior version of the file
// AND failing the export. Same data-loss shape R191 / R206 closed for
// `.gd` save; these two callsites were the lone holdouts.
export async function atomicWrite(target: string, data: Buffer): Promise<void> {
  const tmp = `${target}.tmp-${randomUUID()}`;
  await fs.mkdir(path.dirname(target), { recursive: true });
  const handle = await fs.open(tmp, 'w');
  // R206 — extend R191's clean-tmp-on-failure policy to the write/sync
  // legs. R191 cleans tmp only when `fs.rename` throws; the earlier
  // `handle.writeFile` and `handle.sync` calls have the same "tmp file is
  // already on disk" property (fs.open with 'w' truncates+creates before
  // the writeFile starts), so a failure there leaves `*.tmp-<uuid>`
  // littered next to the user's .gd just like a rename failure would.
  // Realistic triggers parallel R191's list:
  //   • ENOSPC mid-writeFile (disk fills between open and write)
  //   • EIO from a flaky drive surfacing in the middle of a large pptx
  //   • Antivirus quarantine snap landing between open and writeFile
  //     (Windows Defender / endpoint AV inspect new files synchronously
  //     and may yank the handle by returning EBUSY on the next op)
  //   • Network-drive disconnect mid-fsync
  // sweepStaleTempRoots only sweeps `gendoc-*` dirs in os.tmpdir(); the
  // user's Documents folder gets none of that hygiene, so dozens of
  // mystery `Report.gd.tmp-<uuid>` files would accumulate after a series
  // of failed saves on a flaky disk until the user manually cleans up.
  // The structure here mirrors R191's catch+unlink+throw on the rename
  // path; we still close the handle in finally before unlinking so the
  // unlink doesn't race with an open writer (Windows lock semantics —
  // unlink while a handle is open returns EBUSY on Windows even though
  // POSIX would allow it).
  try {
    try {
      await handle.writeFile(data);
      await handle.sync();
    } finally {
      await handle.close();
    }
  } catch (err) {
    await fs.unlink(tmp).catch(() => undefined);
    throw err;
  }
  // R191 — clean up tmp on rename failure. Without this, a rename failure
  // (target locked by AV scan, EACCES from a recently-revoked write
  // permission, ENOSPC mid-rename, network-drive disconnect) leaves a
  // visible `*.tmp-<uuid>` file alongside the user's .gd in their Documents
  // folder. The sweepStaleTempRoots pass at line 37 only cleans
  // `gendoc-*` dirs in os.tmpdir(), it doesn't touch user-picked save
  // destinations. Repeated save failures (slow flaky disk, network drive
  // dropping mid-write) would litter dozens of mystery files until the
  // user manually finds them. unlink failure is best-effort — if even
  // unlink can't run, the user already has bigger filesystem issues.
  try {
    await fs.rename(tmp, target);
  } catch (err) {
    await fs.unlink(tmp).catch(() => undefined);
    throw err;
  }
}

async function rotateBackups(filePath: string, title: string): Promise<void> {
  // Spec §10.4 — keep the most recent 3 backups in ~/.gendoc/backups/.
  const home = process.env['HOME'] ?? process.env['USERPROFILE'] ?? tmpdir();
  const dir = path.join(home, '.gendoc', 'backups');
  await fs.mkdir(dir, { recursive: true });
  const safeTitle = title.replace(/[^\w\u4e00-\u9fa5-]+/g, '_').slice(0, 64) || 'workspace';
  // R192 \u2014 include a short hash of the absolute filePath so two .gd files
  // with the SAME manifest title don't share the same rotation pool.
  // Without this, a user with `~/work/Report.gd` and `~/personal/Report.gd`
  // (both titled\u300cReport\u300din the manifest) had their backups intermingled:
  // the rotation kept "the most recent 3 across both files", so saving
  // the work copy 3 times in a row evicted ALL of the personal copy's
  // backups even though the personal file was untouched. Hashing filePath
  // gives each .gd its own scoped 3-backup pool. Filter pattern includes
  // the hash so rotation only sees this file's own backups; older
  // legacy-format backups (without hash) remain in dir untouched \u2014
  // they're stale by definition (the file's been re-saved since R192) and
  // don't accumulate indefinitely because each saved file generates its
  // own hashed pool that's bounded at 3.
  const pathHash = createHash('sha256').update(filePath).digest('hex').slice(0, 8);
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = path.join(dir, `${safeTitle}-${pathHash}-${ts}.gd`);
  await fs.copyFile(filePath, backupPath).catch(() => undefined);

  const entries = (await fs.readdir(dir).catch(() => [] as string[]))
    .filter((n) => n.startsWith(`${safeTitle}-${pathHash}-`) && n.endsWith('.gd'))
    .sort()
    .reverse();
  const stale = entries.slice(3);
  await Promise.all(stale.map((n) => fs.unlink(path.join(dir, n)).catch(() => undefined)));
}
