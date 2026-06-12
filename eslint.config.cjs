/**
 * R411 — flat-config migration. The repo ships ESLint 9 (package.json:
 * `eslint ^9.14.0`) but only had a legacy `.eslintrc.cjs`; ESLint 9 looks
 * for `eslint.config.*` by default, so `npm run lint` died with "couldn't
 * find an eslint.config file" before linting anything. This file is a
 * 1:1 translation of the old `.eslintrc.cjs` (kept for reference /
 * editors pinned to legacy mode):
 *   env browser+node+es2022  → languageOptions.globals
 *   parser/parserOptions     → languageOptions.parser/parserOptions
 *   extends recommended sets → spread configs from each plugin
 *   ignorePatterns           → global `ignores` entry
 * Rule overrides are byte-identical.
 */
const js = require('@eslint/js');
const tsPlugin = require('@typescript-eslint/eslint-plugin');
const tsParser = require('@typescript-eslint/parser');
const reactPlugin = require('eslint-plugin-react');
const reactHooks = require('eslint-plugin-react-hooks');

module.exports = [
  { ignores: ['dist/**', 'dist-electron/**', 'release/**', 'node_modules/**'] },
  js.configs.recommended,
  {
    files: ['src/**/*.{ts,tsx}', 'scripts/**/*.{js,cjs}', '*.{js,cjs,ts}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 'latest',
        sourceType: 'module',
        ecmaFeatures: { jsx: true },
      },
      globals: {
        // env: { browser: true, node: true } equivalent — the renderer and
        // main process share this config, same as the legacy file.
        window: 'readonly',
        document: 'readonly',
        navigator: 'readonly',
        console: 'readonly',
        setTimeout: 'readonly',
        clearTimeout: 'readonly',
        setInterval: 'readonly',
        clearInterval: 'readonly',
        requestAnimationFrame: 'readonly',
        cancelAnimationFrame: 'readonly',
        fetch: 'readonly',
        URL: 'readonly',
        URLSearchParams: 'readonly',
        Blob: 'readonly',
        File: 'readonly',
        FileReader: 'readonly',
        FormData: 'readonly',
        AbortController: 'readonly',
        TextEncoder: 'readonly',
        TextDecoder: 'readonly',
        crypto: 'readonly',
        performance: 'readonly',
        process: 'readonly',
        Buffer: 'readonly',
        __dirname: 'readonly',
        __filename: 'readonly',
        module: 'writable',
        require: 'readonly',
        exports: 'writable',
        global: 'readonly',
        HTMLElement: 'readonly',
        HTMLInputElement: 'readonly',
        HTMLTextAreaElement: 'readonly',
        HTMLDivElement: 'readonly',
        HTMLButtonElement: 'readonly',
        HTMLSelectElement: 'readonly',
        HTMLCanvasElement: 'readonly',
        HTMLImageElement: 'readonly',
        HTMLAnchorElement: 'readonly',
        KeyboardEvent: 'readonly',
        MouseEvent: 'readonly',
        DragEvent: 'readonly',
        ClipboardEvent: 'readonly',
        Event: 'readonly',
        CustomEvent: 'readonly',
        Node: 'readonly',
        Element: 'readonly',
        Range: 'readonly',
        Selection: 'readonly',
        MutationObserver: 'readonly',
        ResizeObserver: 'readonly',
        IntersectionObserver: 'readonly',
        Image: 'readonly',
        localStorage: 'readonly',
        sessionStorage: 'readonly',
        getComputedStyle: 'readonly',
        DOMParser: 'readonly',
        XMLSerializer: 'readonly',
        atob: 'readonly',
        btoa: 'readonly',
        structuredClone: 'readonly',
        queueMicrotask: 'readonly',
        MessageChannel: 'readonly',
        Worker: 'readonly',
        FocusEvent: 'readonly',
        InputEvent: 'readonly',
        CompositionEvent: 'readonly',
        WheelEvent: 'readonly',
        PointerEvent: 'readonly',
        TouchEvent: 'readonly',
        DataTransfer: 'readonly',
        CSSStyleDeclaration: 'readonly',
        NodeJS: 'readonly',
        React: 'readonly',
        JSX: 'readonly',
      },
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
      react: reactPlugin,
      'react-hooks': reactHooks,
    },
    settings: { react: { version: 'detect' } },
    rules: {
      ...tsPlugin.configs.recommended.rules,
      ...reactPlugin.configs.recommended.rules,
      ...reactHooks.configs.recommended.rules,
      'react/react-in-jsx-scope': 'off',
      'react/prop-types': 'off',
      // TS-aware unused-vars replaces the core rule (same as the legacy
      // `plugin:@typescript-eslint/recommended` behaviour).
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/no-explicit-any': 'warn',
      // Core no-undef is unreliable on TypeScript sources (type-only
      // globals); tsc --noEmit owns undefined-identifier checking.
      'no-undef': 'off',
    },
  },
];
