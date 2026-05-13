/**
 * R405 — main-process locale state shared between main.ts (close-prompt
 * dialog) and ipc.ts (config.set handler that updates the locale on user
 * toggle). Kept in its own module so neither file has to import the
 * other — main.ts and ipc.ts already have a circular dependency via
 * `registerIpcHandlers`, and threading the setter through there would
 * widen the surface for no real benefit. Plain mutable export is the
 * simplest shape — single writer (ipc.ts on config patch + main.ts on
 * boot), multiple readers (main.ts's close-prompt handler).
 */

let effectiveMainLocale: 'zh' | 'en' = 'en';

export function getMainLocale(): 'zh' | 'en' {
  return effectiveMainLocale;
}

export function setMainLocale(l: 'zh' | 'en'): void {
  effectiveMainLocale = l;
}
