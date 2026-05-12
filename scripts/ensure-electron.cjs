#!/usr/bin/env node
/**
 * Idempotent guard for the Electron binary.
 *
 * `npm install --ignore-scripts` skips Electron's own postinstall (which
 * downloads the platform binary), leaving the package present but the
 * binary missing — `electron-vite dev` then fails with `Error: Electron
 * uninstall`. This script checks for the binary and runs Electron's
 * install.js only when needed, so subsequent `npm install`s self-heal.
 */

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const electronDir = path.join(__dirname, '..', 'node_modules', 'electron');
const pathTxt = path.join(electronDir, 'path.txt');
const installJs = path.join(electronDir, 'install.js');

if (!fs.existsSync(electronDir)) {
  console.log('[ensure-electron] electron package not installed; skipping.');
  process.exit(0);
}

if (fs.existsSync(pathTxt)) {
  // path.txt only exists once Electron's install.js has succeeded.
  process.exit(0);
}

if (!fs.existsSync(installJs)) {
  console.warn('[ensure-electron] install.js missing — package layout changed?');
  process.exit(0);
}

console.log('[ensure-electron] downloading Electron binary…');
const r = spawnSync(process.execPath, [installJs], { stdio: 'inherit' });
process.exit(r.status ?? 0);
