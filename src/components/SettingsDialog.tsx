/**
 * Settings dialog. Owns API-key entry (Test connection → Anthropic ping),
 * default model, theme, embed-chat-history toggle. API key never leaves
 * the renderer in plaintext — main encrypts via safeStorage.
 */

import { useEffect, useRef, useState } from 'react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from './ui/dialog';
import { Button } from './ui/button';
import { cn } from '../lib/utils';
import { notify } from '../store/toast';
import { Check, Eye, EyeOff, Loader2, ShieldCheck } from 'lucide-react';
import type { UserConfig } from '../types/ipc';
import { SUPPORTED_MODELS, resolveSupportedModelId } from '../types/ai';

interface Props {
  open: boolean;
  onOpenChange(b: boolean): void;
  onApiKeyChange?(hasKey: boolean): void;
}

export function SettingsDialog({ open, onOpenChange, onApiKeyChange }: Props): JSX.Element {
  /**
   * Wrap close so we flush any in-flight NumberField draft (Temperature /
   * Max tokens) before unmounting. Those inputs commit on `onBlur`, but
   * Radix's Esc / click-outside / X-button paths all go straight to
   * `onOpenChange(false)` and unmount the input — native blur never fires
   * on element removal, so the user's typed value silently disappears.
   * Reopening then shows the *old* persisted value, which reads as "the
   * dialog ignored my change". Force-blurring the focused INPUT here lets
   * its onBlur handler run synchronously (parse → clamp → patch → IPC →
   * configChanged broadcast), so the value is committed before the dialog
   * tears down. Safe for the API-key input too: its blur has no committing
   * side effect (saveKey is gated behind Enter / button click), and the
   * useEffect on `open` resets keyDraft on next reopen anyway.
   */
  const handleOpenChange = (next: boolean) => {
    if (!next) {
      const ae = document.activeElement as HTMLElement | null;
      if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA')) {
        ae.blur();
      }
    }
    onOpenChange(next);
  };

  const [config, setConfig] = useState<UserConfig | null>(null);
  const [hasKey, setHasKey] = useState(false);
  const [keyDraft, setKeyDraft] = useState('');
  const [showKey, setShowKey] = useState(false);
  const [pingState, setPingState] = useState<
    | { status: 'idle' }
    | { status: 'pending' }
    | { status: 'ok'; model: string }
    | { status: 'error'; message: string }
  >({ status: 'idle' });
  /**
   * Monotonic counter that invalidates in-flight ping promises. The dialog
   * is permanently mounted (App.tsx renders it unconditionally and only
   * toggles `open`), so a Test-connection click that takes 2-3s can outlive
   * a close/reopen, a key save, a key clear, or a model switch — each of
   * which resets pingState back to 'idle' to retract the stale OK badge.
   * Without this guard the original ping's `setPingState({status:'ok'})`
   * paints right back over the reset, vouching for a config the user has
   * already moved past. Bumping the epoch at every reset point and
   * comparing on resolve drops the stale write on the floor.
   */
  const pingEpochRef = useRef(0);

  // R90 — two-step confirm for 清除 (API key wipe). The action is
  // *destructive and irreversible*: clearKey() at line ~149 calls
  // window.gendoc.config.clearApiKey() which removes the OS-keystore-
  // encrypted blob with no recovery path — the user has to fetch a new key
  // from the Anthropic console. The doc-comment on the 清除 button itself
  // (line ~204) already flagged this as a foot-gun ("a user who misreads
  // 清除 as 'clear field' instead of 'wipe the saved key' loses it
  // silently") but the previous mitigation was a tooltip only — clicking
  // still went straight through. Mirrors the established two-click pattern
  // for destructive actions:
  //   AIPanel.tsx:601-641   新對話 (drops messages / pending / streaming)
  //   FindReplaceDialog.tsx 全部取代 confirmAll latch
  // Both stage on first click with amber visual + tooltip flip, auto-cancel
  // after 4s. Same shape here so the dialog speaks one voice across the
  // app's "are you sure" surface.
  const [confirmClearKey, setConfirmClearKey] = useState(false);
  const confirmClearTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (confirmClearTimer.current) clearTimeout(confirmClearTimer.current);
    },
    [],
  );
  // Reset the staged confirm when the dialog closes — re-opening should not
  // resurface a half-armed clear that the user has had time to forget about.
  useEffect(() => {
    if (open) return;
    setConfirmClearKey(false);
    if (confirmClearTimer.current) clearTimeout(confirmClearTimer.current);
  }, [open]);

  useEffect(() => {
    if (!open) return;
    let alive = true;
    pingEpochRef.current++;
    void (async () => {
      // R273 — wrap the initial-load IPC pair. config.get's underlying
      // loadConfig() awaits fs.mkdir + fs.readFile + fs.writeFile; any of
      // those can reject (disk full, EACCES on `~/.gendoc`, antivirus
      // pinning the config dir, network user-profile offline). Without
      // this guard the rejection escapes as unhandledrejection AND the
      // dialog sits at the loading spinner forever with no way to recover.
      // Mirrors the patch helper's reportConfigError pattern (line ~133)
      // and the R249 / R270 / R271 sweep for user-triggered async paths.
      try {
        const [cfg, has] = await Promise.all([
          window.gendoc.config.get(),
          window.gendoc.config.hasApiKey(),
        ]);
        if (!alive) return;
        setConfig(cfg);
        setHasKey(has);
        setKeyDraft('');
        setPingState({ status: 'idle' });
      } catch (err) {
        if (!alive) return;
        const msg = err instanceof Error ? err.message : String(err);
        notify(`載入設定失敗：${msg}`, 'error');
        // Close the dialog so the user isn't stuck at the spinner. The
        // outer parent owns `open`; calling onOpenChange(false) routes
        // through the same close path Esc / ✕ uses.
        onOpenChange(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, [open, onOpenChange]);

  // Surface IPC write failures (disk full, AppData permission denied, file
  // locked by AV scanner, etc.) as a toast. Without this layer the awaited
  // promise rejected silently — the toggle's visual state didn't update
  // (because the post-await `setConfig(next)` never ran), so the user saw
  // their setting "snap back" with no explanation. Returning a boolean lets
  // each call site know whether to run its post-success state updates.
  const reportConfigError = (action: string, err: unknown): void => {
    const msg = err instanceof Error ? err.message : String(err);
    notify(`${action}失敗：${msg}`, 'error');
  };

  const patch = async (p: Partial<UserConfig>) => {
    // Switching the default model invalidates any prior Test-connection
    // badge — that "✓ <model> OK" was for the model the user just left.
    // Clearing here forces a re-test against the new model so the badge
    // never misleads users into shipping with an unverified model.
    const modelChanged =
      p.defaultModel !== undefined && p.defaultModel !== config?.defaultModel;
    let next: UserConfig;
    try {
      next = await window.gendoc.config.set(p);
    } catch (err) {
      reportConfigError('儲存設定', err);
      return;
    }
    setConfig(next);
    if (modelChanged) {
      pingEpochRef.current++;
      setPingState({ status: 'idle' });
    }
    // Notify App.tsx so live consumers (auto-save interval, theme) refresh
    // without waiting for the next app launch.
    window.dispatchEvent(new CustomEvent('gendoc:configChanged'));
  };

  const saveKey = async () => {
    if (!keyDraft.trim()) return;
    try {
      await window.gendoc.config.setApiKey(keyDraft.trim());
    } catch (err) {
      reportConfigError('儲存 API key', err);
      return;
    }
    setKeyDraft('');
    setHasKey(true);
    // Any prior "Test connection ✓" badge belonged to the *previous* key —
    // we must not let it carry over to the new key, otherwise the user sees
    // a green check that vouches for a key the app has never actually
    // pinged. Mirrors the model-change reset above.
    pingEpochRef.current++;
    setPingState({ status: 'idle' });
    onApiKeyChange?.(true);
  };

  const clearKey = async () => {
    try {
      await window.gendoc.config.clearApiKey();
    } catch (err) {
      reportConfigError('清除 API key', err);
      return;
    }
    setHasKey(false);
    // Same rationale as saveKey: an "OK" badge surviving the clear could
    // resurface if the user immediately pastes another key without closing
    // the dialog.
    pingEpochRef.current++;
    setPingState({ status: 'idle' });
    onApiKeyChange?.(false);
  };

  // R90 — first click stages, second click within 4s actually wipes. See
  // confirmClearKey doc-block above (~line 73) for the full rationale and
  // sibling pattern citations (AIPanel 新對話 / FindReplaceDialog 全部取代).
  const onClearKeyClick = () => {
    if (!confirmClearKey) {
      setConfirmClearKey(true);
      if (confirmClearTimer.current) clearTimeout(confirmClearTimer.current);
      confirmClearTimer.current = setTimeout(() => setConfirmClearKey(false), 4000);
      return;
    }
    setConfirmClearKey(false);
    if (confirmClearTimer.current) clearTimeout(confirmClearTimer.current);
    void clearKey();
  };

  const test = async () => {
    // Capture the epoch *before* the await; if anything (close/reopen, model
    // change, save/clear key) bumps the ref while the ping is in flight,
    // this resolution belongs to a config the user already left and must
    // not overwrite the freshly-reset 'idle' badge.
    const epoch = ++pingEpochRef.current;
    setPingState({ status: 'pending' });
    // R286 — wrap the IPC await. main-side ping() already catches Anthropic
    // API errors and returns {ok:false, error} for the application-level
    // failures, but the IPC bridge itself can still reject on main crash /
    // preload re-bind / non-serializable payload corner cases. Without this
    // guard the reject escapes via `onClick={test}` (React doesn't catch),
    // setPingState never fires the resolved state, and the「測試連線」
    // button stays disabled with the spinner forever — only escape is close
    // + reopen the dialog. Catch + map to status:'error' so the button
    // unlocks and the user sees why. Same try/catch idiom as sibling
    // saveKey / clearKey above.
    let result: Awaited<ReturnType<typeof window.gendoc.ai.ping>>;
    try {
      // R389 — coerce against SUPPORTED_MODELS. A stale defaultModel from
      // an older app version's config.json would otherwise propagate into
      // the ping, returning model_not_found and masquerading as an
      // API-key failure to the user. See resolveSupportedModelId doc-block
      // in types/ai.ts. The literal-fallback below was the original
      // null-guard; replaced by resolveSupportedModelId which handles
      // both undefined AND unrecognized values in one shot.
      result = await window.gendoc.ai.ping(resolveSupportedModelId(config?.defaultModel));
    } catch (err) {
      if (epoch !== pingEpochRef.current) return;
      const msg = err instanceof Error ? err.message : String(err);
      setPingState({ status: 'error', message: msg });
      return;
    }
    if (epoch !== pingEpochRef.current) return;
    if (result.ok) setPingState({ status: 'ok', model: result.model ?? '' });
    else setPingState({ status: 'error', message: result.error ?? 'unknown' });
  };

  if (!config) {
    return (
      <Dialog open={open} onOpenChange={handleOpenChange}>
        <DialogContent>
          <Loader2 className="h-4 w-4 animate-spin" />
        </DialogContent>
      </Dialog>
    );
  }

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>設定</DialogTitle>
          <DialogDescription>
            API key 透過 OS keystore（Windows DPAPI / macOS Keychain）加密儲存，明文不落地。
          </DialogDescription>
        </DialogHeader>

        <section className="space-y-3">
          <h3 className="text-sm font-medium">Anthropic API key</h3>
          {hasKey ? (
            <div className="flex items-center gap-2 text-sm">
              <ShieldCheck className="h-4 w-4 text-emerald-500" />
              <span>已儲存（加密）</span>
              {/* 清除 was the lone titleless button left in this dialog after
                  the systematic title pass that landed on its siblings: the
                  visibility toggle (line ~241), 儲存 (line ~258), and 測試連線
                  (line ~294) all carry tooltips. The asymmetry stings extra
                  here because clearKey() at line 149-163 is *destructive* and
                  *irreversible* — it removes the OS-keystore-encrypted key
                  with no confirm dialog, so a user who misreads "清除" as
                  "clear field" (instead of "wipe the saved key") loses it
                  silently. Mirrors DocxEditor.tsx:3204's secondary 清除連結
                  button (`title="移除目前連結"`), which set the verb pattern
                  for destructive clear buttons across the codebase.
                  R90 — tooltip alone wasn't enough; the destructive path now
                  routes through onClearKeyClick (~line 165) which gates the
                  IPC call behind a 4-second two-click confirm and an amber
                  visual flip. Same shape as AIPanel 新對話 (line ~601). */}
              <Button
                size="sm"
                variant="outline"
                className={cn(
                  'ml-auto transition-colors',
                  confirmClearKey &&
                    'bg-amber-500/15 border-amber-500/60 text-amber-700 hover:bg-amber-500/25',
                )}
                onClick={onClearKeyClick}
                title={
                  confirmClearKey
                    ? '再次點擊以確認移除 API key（不可復原）'
                    : '移除已儲存的 API key'
                }
              >
                {confirmClearKey ? '再次點擊以確認' : '清除'}
              </Button>
            </div>
          ) : (
            <div className="flex gap-2">
              <input
                type={showKey ? 'text' : 'password'}
                value={keyDraft}
                onChange={(e) => setKeyDraft(e.target.value)}
                // Pasted keys land here — select-all on focus lets the user
                // overtype without manually clearing first. Enter saves so
                // they don't have to take their hands off the keyboard.
                onFocus={(e) => e.currentTarget.select()}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && keyDraft.trim()) {
                    e.preventDefault();
                    saveKey();
                  }
                }}
                placeholder="sk-ant-..."
                className="flex-1 bg-secondary/40 rounded px-2 py-1 text-sm outline-none focus:ring-1 focus:ring-ring"
              />
              {/* Visibility toggle was icon-only with no title / aria-label —
                  the lone icon-only Button in this dialog without one, and
                  inconsistent with every <Button size="icon"> in App.tsx
                  (新專案 / 開啟 / 儲存 / 匯出 all carry a title). Worse: the
                  icon flips with state (Eye ↔ EyeOff), so users hovering for
                  a hint to confirm "does this reveal or mask?" got nothing.
                  Title + aria-label both reflect the *action*, not the
                  current state — clicking when showKey is true masks the key
                  (so the label says "隱藏"), and vice versa. Mirrors the
                  AIPanel error-close pattern (title + aria-label paired). */}
              <Button
                size="icon"
                variant="ghost"
                onClick={() => setShowKey((b) => !b)}
                title={showKey ? '隱藏 API key' : '顯示 API key'}
                aria-label={showKey ? '隱藏 API key' : '顯示 API key'}
              >
                {showKey ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
              </Button>
              {/* Surface the Enter shortcut the input's onKeyDown (line ~294)
                  binds. The doc-comment on the input ("Enter saves so they
                  don't have to take their hands off the keyboard") states
                  the intent, but with no tooltip / footer hint anywhere the
                  binding was invisible — users who paste a key had to mouse
                  over to this button or guess.
                  R103 — disabled-state tooltip. The button is `disabled` when
                  `keyDraft.trim()` is empty, AND the input's onKeyDown at
                  line 294-299 also gates Enter on the same condition — so a
                  user staring at a freshly-opened settings dialog with no
                  key typed sees a greyed button advertising「儲存 API key
                  (Enter)」, presses Enter, nothing happens. The sibling
                  AIPanel Send button at AIPanel.tsx:1218-1224 (R88) had the
                  exact same pre-fix shape and was upgraded to a state-aware
                  3-way tooltip; this Save button's previous comment even
                  cited「AIPanel.tsx 送出 (Enter)」as its precedent, but
                  that precedent has since moved on to the recovery-pointing
                  pattern「輸入提示後即可送出 (Enter)」 — leaving this Save
                  button as the lone same-codebase holdout still using the
                  pre-R88 static-string shape. Wording mirrors AIPanel R88's
                  "輸入提示後即可送出 (Enter)" verbatim in shape: name the
                  recovery action ("輸入 API key"), then surface that the
                  shortcut comes alive once that's done. Single disable
                  reason here, so a 2-way tooltip suffices (R88 needed 3
                  branches because Send had two disable causes). */}
              <Button
                size="sm"
                onClick={saveKey}
                disabled={!keyDraft.trim()}
                title={!keyDraft.trim()
                  ? '輸入 API key 後即可儲存 (Enter)'
                  : '儲存 API key (Enter)'}
              >
                儲存
              </Button>
            </div>
          )}
          {/* Soft-warn on non-Anthropic-shaped keys before the user saves and
              hits a confusing 401 on Test connection. We don't hard-block —
              Anthropic could legitimately issue new prefixes — but most of the
              time a key without `sk-ant-` is either an OpenAI / OpenRouter
              key the user pasted into the wrong box, or a partial paste that
              dropped the prefix. Catching it pre-save saves a round-trip
              through Save → Test → 401 → "what was wrong with my key?". */}
          {!hasKey && keyDraft.trim() && !keyDraft.trim().startsWith('sk-ant-') && (
            <div className="text-[11px] text-amber-500">
              這看起來不像 Anthropic 金鑰（通常以 <code className="font-mono">sk-ant-</code> 開頭）。確認沒貼錯後仍可儲存。
            </div>
          )}
          {hasKey && (
            <div className="flex items-center gap-2 text-xs">
              {/* Button label was the only English UI string in this otherwise
                  fully Traditional-Chinese dialog (siblings: 已儲存（加密）/
                  清除 / 儲存 / 預設模型 / 啟用 prompt cache / 自動儲存 …).
                  The mixed-language outlier read as a translation oversight
                  rather than an intentional choice. Code comments and the
                  doc-comment header still reference "Test connection" — those
                  are dev-facing and stay in English; only the visible label
                  changes. Tooltip hints at what the button actually does
                  (sends a tiny ping against the currently selected model)
                  since the section header is just "Anthropic API key" and
                  doesn't itself say "test". */}
              {/* R123 — disabled-state-aware tooltip. The button at line
                  388 reads `disabled={pingState.status === 'pending'}` so
                  while a ping is in-flight the user sees a spinning Loader2
                  (line 391) and can't click — but the tooltip stayed static,
                  describing what the button *does* when active rather than
                  why it's currently unclickable. Same-dialog sibling at line
                  351-353 (the 儲存 button) already honors its `disabled`
                  prop with a two-branch tooltip `!keyDraft.trim() ? '輸入
                  API key 後即可儲存 (Enter)' : '儲存 API key (Enter)'`, so
                  this lone always-on tooltip was the dialog's only
                  state-blind hold-out. The pending branch reuses the
                  project's canonical busy-state shape from MarkdownToolbar
                  .tsx:274 (`正在輸出 PDF…`); 「正在測試連線…」 mirrors the
                  visible label so a busy-spinning user reads the same
                  noun on both surfaces. */}
              <Button
                size="sm"
                variant="outline"
                onClick={test}
                disabled={pingState.status === 'pending'}
                title={
                  pingState.status === 'pending'
                    ? '正在測試連線…'
                    : '向 Anthropic 發送一次測試請求，確認 API key 與目前模型可用'
                }
              >
                {pingState.status === 'pending' && <Loader2 className="h-3 w-3 animate-spin" />}
                測試連線
              </Button>
              {/* R154 — SR-readable result of the test-connection button.
                  Both states materialize asynchronously after the user clicks
                  測試連線; without aria-live, SR users hear the click but get
                  no feedback when the result lands seconds later. The OK
                  branch is `role="status"` (polite) — successful confirmation
                  shouldn't interrupt; the error branch is `role="alert"`
                  (assertive) so SR users hear failures immediately, matching
                  the urgency tier the Toaster uses (Toaster.tsx:109). */}
              {pingState.status === 'ok' && (
                <span role="status" aria-live="polite" className="text-emerald-500 flex items-center gap-1">
                  <Check className="h-3 w-3" /> {pingState.model} OK
                </span>
              )}
              {pingState.status === 'error' && (
                <span role="alert" className="text-destructive truncate">{pingState.message}</span>
              )}
            </div>
          )}
        </section>

        <section className="space-y-2">
          <h3 className="text-sm font-medium">預設模型</h3>
          {/* The h3 reads "預設模型", which a literal-minded user parses as
              "applies only to *new* sessions — won't touch the chat I have
              open right now". App.tsx:202-225 actually mirrors this patch
              live into the running AI store (setModel on configChanged), so
              changing it here switches the model for the current turn too,
              without clearing messages or the conversation id. The sibling
              AIPanel model picker (AIPanel.tsx:681) advertises the same
              non-destructive guarantee — "切換 AI 模型（不會清空目前的對話）"
              — because hovering a model name (e.g. "Sonnet 4.6") doesn't
              answer the user's real question ("will this nuke my chat?").
              Same picker semantics, same tooltip dialect; the gap was just
              that the section header here gave a false sense of "default-
              only" scope. */}
          <select
            value={config.defaultModel}
            onChange={(e) => patch({ defaultModel: e.target.value })}
            title="切換預設模型，並立即套用至目前對話（不會清空對話內容）"
            className="w-full bg-secondary/40 rounded px-2 py-1 text-sm"
          >
            {SUPPORTED_MODELS.map((m) => (
              <option key={m.id} value={m.id}>
                {m.label}
              </option>
            ))}
          </select>
        </section>

        <section className="grid grid-cols-2 gap-3">
          {/* Range-in-label pattern: "(0 – 1)" / "(256 – 16384)" makes the
              valid range visible at all times, not just on hover. Without
              this, a user who types e.g. 5 for Temperature blurs out, sees
              the input snap to 1, and has no in-dialog explanation of *why*
              the snap happened — `min`/`max` constrain the spinner buttons
              but not typed values, mirroring the same gap GoToDialog called
              out (GoToDialog.tsx:107-113). The toast inside NumberField
              (added below) covers the post-blur "we adjusted this" cue;
              the label range covers the pre-blur "what should I type?" cue.
              Two complementary signals matching the GoToDialog placeholder
              ("1 – {max}") + toast pair. */}
          <Field label="Temperature (0 – 1)">
            <NumberField
              label="Temperature"
              value={config.temperature}
              step="0.1"
              min={0}
              max={1}
              parse={parseFloat}
              onCommit={(v) => patch({ temperature: v })}
            />
          </Field>
          <Field label="Max tokens (256 – 16384)">
            <NumberField
              label="Max tokens"
              value={config.maxTokens}
              step="256"
              min={256}
              max={16384}
              parse={(s) => parseInt(s, 10)}
              onCommit={(v) => patch({ maxTokens: v })}
            />
          </Field>
        </section>

        <section className="space-y-2">
          {/* R86 tooltips: the three labels here are the most jargon-heavy
              in the dialog (`prompt cache`、`.gd`)，沒有 hover 解釋使用者
              很難判斷該不該勾。Tooltip 風格沿用 line 311 / 346 的「動詞
              開頭 + 副作用括註」── 把行為與代價講清楚。 */}
          <CheckboxRow
            label="啟用 prompt cache"
            checked={config.promptCache}
            onChange={(v) => patch({ promptCache: v })}
            title="多輪對話中重複使用 system prompt 與工具定義的快取，可大幅降低 token 費用；首次寫入以 1.25× 計價，關閉後一律以標準費率計算"
          />
          <CheckboxRow
            label="預設將 AI 對話歷史內嵌至 .gd"
            checked={config.embedChatHistoryDefault}
            onChange={(v) => patch({ embedChatHistoryDefault: v })}
            title="新建檔案時預設將 AI 對話歷史一併儲存進 .gd 容器；每個檔案可在工作區設定中個別覆蓋"
          />
          <CheckboxRow
            label="啟動時自動開啟上次的檔案"
            checked={config.autoOpenLastWorkspace}
            onChange={(v) => patch({ autoOpenLastWorkspace: v })}
            title="啟動 Gen-Doc 時自動還原最近一次開啟的檔案；關閉後每次啟動都以空白工作區開始"
          />
        </section>

        <section className="space-y-2">
          <h3 className="text-sm font-medium">自動儲存</h3>
          {/* Off / 30s / 1min / 5min — common cadences. Auto-save only writes
              when the workspace already has a path on disk; new untitled
              workspaces won't pop a Save-As dialog behind your back. */}
          {/* R153 — same-dialog hover-explanation parity with the 預設模型
              `<select>` at line 440-451 above (which carries `title="切換預
              設模型，並立即套用至目前對話（不會清空對話內容）"`). Both pickers
              live in the same dialog, both flip a single config key via
              `patch`, but only the model side advertised what clicking does
              on hover — leaving the auto-save side as the only top-level
              interactive control without a hover-discoverable tooltip. The
              section header「自動儲存」names the feature, and the inline
              warning at line 540-544 surfaces the「先 Ctrl+S 命名」rule, but
              both fire AFTER the user has committed to a non-zero option;
              a user about to pick「30 秒」 to confirm what auto-save does
              gets no signal until they've already selected a value. The
              tooltip風格 mirrors line 443's verb-first shape (「切換」起手)
              and exposes the silent-noop guard at App.tsx:450 (`if
              (!ws.filePath) return`) up-front, instead of relegating the
              caveat to an inline post-selection hint. */}
          <select
            value={String(config.autoSaveIntervalMs)}
            onChange={(e) => patch({ autoSaveIntervalMs: parseInt(e.target.value, 10) })}
            title="切換自動儲存間隔（僅對已儲存到磁碟的檔案生效，未命名檔案請先按 Ctrl+S）"
            className="w-full bg-secondary/40 rounded px-2 py-1 text-sm"
          >
            <option value="0">關閉</option>
            <option value="30000">30 秒</option>
            <option value="60000">1 分鐘</option>
            <option value="300000">5 分鐘</option>
          </select>
          {/* Surface the silent-noop guard at App.tsx:450 (`if (!ws.filePath)
              return`). Without this hint, a user who enables 自動儲存 on a
              fresh untitled workspace sees nothing happen — no toast, no
              Save-As dialog — and reasonably concludes the feature is
              broken. The guard exists on purpose (we don't want auto-save to
              pop a modal behind the user's back), but the requirement must
              be visible at the point of configuration, not buried in a
              source-code comment. Only show the hint when auto-save is
              actually enabled — when "關閉" is selected the requirement is
              moot and the line would just be noise. */}
          {config.autoSaveIntervalMs > 0 && (
            <p className="text-[11px] text-muted-foreground leading-snug">
              提示：自動儲存僅對已儲存過的檔案生效；新建的未命名檔案請先按 Ctrl+S 命名後才會自動儲存。
            </p>
          )}
        </section>

        {/* 主題 picker intentionally removed — the app is pinned to a
            light/white-based UI by design. See `ensureLightTheme()` in
            App.tsx and the doc-comment above `.work-canvas` in
            renderer/index.css. */}
      </DialogContent>
    </Dialog>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="space-y-1 text-xs">
      <span className="text-muted-foreground">{label}</span>
      {children}
    </label>
  );
}

/**
 * Number input that decouples the displayed string from the committed numeric
 * value. The previous version bound the input directly to `config.temperature`
 * and patched on every keystroke. That made decimal entry impossible: typing
 * "0.5" went "0" → patch(0) → re-render value="0" → "." discarded → "0.5"
 * unreachable via straight typing. Here the user freely edits a local draft
 * string; we only parse/clamp/commit on blur, and revert the draft if the
 * input is left non-numeric. External changes still win via the `value` effect.
 */
function NumberField({
  label,
  value,
  step,
  min,
  max,
  parse,
  onCommit,
}: {
  /** Human-readable name of the field, used in the clamp-notification toast.
   *  Optional so the legacy callsite (none, after R41) still compiles, but
   *  current callers always pass it so the toast can name *which* setting
   *  was adjusted (Temperature vs Max tokens). */
  label?: string;
  value: number;
  step: string;
  min: number;
  max: number;
  parse: (s: string) => number;
  onCommit: (n: number) => void;
}) {
  const [draft, setDraft] = useState(String(value));
  useEffect(() => {
    setDraft(String(value));
  }, [value]);
  return (
    <input
      type="number"
      step={step}
      min={min}
      max={max}
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        const v = parse(draft);
        if (Number.isNaN(v)) {
          setDraft(String(value));
          return;
        }
        const clamped = Math.max(min, Math.min(max, v));
        setDraft(String(clamped));
        if (clamped !== value) onCommit(clamped);
        // Match the GoToDialog clamp-notify pattern (GoToDialog.tsx:133-135)
        // that R32's comment explicitly justified: "without the toast, the
        // permissive-clamp behaviour silently lies about where the user
        // ended up". Same gap was sitting in this dialog: typing 5 for
        // Temperature blurred out, the input visibly snapped to 1, but
        // nothing told the user *why* — they couldn't tell if the spinner
        // had a hidden cap, if the field was buggy, or if their typed
        // value was rejected. Fires only when actual clamping occurred
        // (clamped !== v), so the in-range fast path stays quiet. Uses
        // `label` (e.g. "Temperature") rather than the with-range Field
        // label ("Temperature (0 – 1)") so the toast doesn't double-print
        // the range — the parenthetical 「有效範圍 …」 already carries it.
        if (clamped !== v) {
          notify(
            `${label ?? '數值'} 已調整為 ${clamped}（輸入 ${v} 超出有效範圍 ${min} – ${max}）`,
            'info',
          );
        }
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
      }}
      className="w-full bg-secondary/40 rounded px-2 py-1 text-sm"
    />
  );
}

function CheckboxRow({
  label,
  checked,
  onChange,
  title,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
  /** Hover hint surfaced on the wrapping <label>. R86: every other
   * interactive control in this file (lines 220 / 258 / 275 / 311 / 346)
   * already carries an explanatory `title=`, but the three CheckboxRow
   * call-sites at 394 / 399 / 404 had nothing — leaving the most jargon-
   * heavy controls in the dialog (`啟用 prompt cache`, `內嵌至 .gd`) with
   * zero discoverable explanation. Optional so future plain-language
   * checkboxes don't have to fabricate a tooltip. */
  title?: string;
}) {
  return (
    <label className="flex items-center gap-2 text-sm cursor-pointer" title={title}>
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{label}</span>
    </label>
  );
}
