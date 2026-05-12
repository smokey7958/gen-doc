// This file is mostly a placeholder; the canonical config lives in
// electron.vite.config.ts. We keep this so editors / tools that look for
// vite.config.ts find a valid config too.
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
      '@components': resolve(__dirname, 'src/components'),
      '@store': resolve(__dirname, 'src/store'),
      '@types': resolve(__dirname, 'src/types'),
      '@ai': resolve(__dirname, 'src/ai'),
    },
  },
});
