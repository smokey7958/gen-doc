/**
 * Lightweight i18n — 2-locale (Chinese / English) inline translator.
 *
 * Design rationale: the codebase originally hardcoded Traditional Chinese
 * everywhere; retrofitting a full i18next-style key/catalog system would
 * require inventing keys for ~500+ user-facing strings and a multi-day
 * refactor. The `t(zh, en)` inline form keeps the original Chinese visible
 * at the call site as living source of truth, with the English pair right
 * next to it for easy review.
 *
 *   useT() — React hook; re-renders the component on locale change.
 *   tImp() — imperative; reads the live locale at call time. Use outside
 *            React (toast helpers, IPC error builders, main process menu).
 *
 * Locale resolution:
 *   1. User preference (UserConfig.locale, persisted to ~/.gendoc/config.json)
 *   2. If null → OS locale via app.getLocale() ('zh-TW' / 'en-US' / etc.)
 *   3. Anything starting with "zh" → 'zh'; everything else → 'en'.
 *
 * Initial store value is 'en' — App.tsx's first-load effect hydrates the
 * real locale before any user-facing render. The brief flash window (a few
 * ms between mount and the hydrate IPC's resolve) is acceptable on a
 * Chinese OS because the boot splash is invariant text.
 */

import { create } from 'zustand';

export type Locale = 'zh' | 'en';

interface LocaleStore {
  locale: Locale;
  setLocale: (l: Locale) => void;
}

export const useLocale = create<LocaleStore>((set) => ({
  locale: 'en',
  setLocale: (l) => set({ locale: l }),
}));

/**
 * React hook — returns a translator function bound to the current locale.
 * The hook subscribes to `useLocale`, so calling components re-render on
 * locale toggle and `t()` produces the new string on the same render.
 */
export function useT(): (zh: string, en: string) => string {
  const locale = useLocale((s) => s.locale);
  return (zh, en) => (locale === 'zh' ? zh : en);
}

/**
 * Imperative translator for non-React call sites (toast builders, IPC
 * error message construction, main-process menu labels). Reads the live
 * locale via getState() each call — no subscription, so the caller is
 * responsible for being re-invoked on locale change (e.g., main rebuilds
 * the menu on every config patch that touches `locale`).
 */
export function tImp(zh: string, en: string): string {
  return useLocale.getState().locale === 'zh' ? zh : en;
}

/**
 * Map a Chromium-style OS locale string ('zh-TW' / 'zh-CN' / 'en-US' / ...)
 * to our 2-locale union. Anything starting with "zh" (case-insensitive) is
 * Chinese; everything else falls back to English. Covers zh-TW / zh-HK /
 * zh-CN / zh-SG / zh-Hant / zh-Hans without enumerating each.
 */
export function resolveOsLocale(raw: string | undefined | null): Locale {
  if (raw && raw.toLowerCase().startsWith('zh')) return 'zh';
  return 'en';
}
