// Generates logo/logo.ico from logo/logo.png so electron-builder / NSIS can use it.
// Run via: node scripts/gen-ico.cjs
const path = require('node:path');
const fs = require('node:fs');
const pngToIco = require('png-to-ico').default;

const src = path.resolve(__dirname, '..', 'logo', 'logo.png');
const dst = path.resolve(__dirname, '..', 'logo', 'logo.ico');

(async () => {
  const buf = await pngToIco(src);
  fs.writeFileSync(dst, buf);
  console.log('Wrote', dst, '(', buf.length, 'bytes )');
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
