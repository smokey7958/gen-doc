# Gen Doc — Locked Tech Stack（所有 sub-agent 必讀）

This file is the SHARED CONTRACT for all parallel agents working on the
Gen Doc project. Do NOT diverge from these choices. If a choice seems
wrong, surface it back to the main session for review — do NOT make a
unilateral substitution.

## Project root
- `C:\Gen-Doc\`

## Identity
- Product name: **Gen Doc**
- File extension: `.gd`
- Short tagline: 統一筆記 / 文書 / 表格 / 簡報，內建 AI 編輯助手

## Runtime
- **Electron** (latest stable) — cross-platform shell
- **Node.js** 20.x for build tooling
- **TypeScript** 5.x

## UI layer
- **React** 18 + functional components only
- **Vite** for renderer bundling
- **Tailwind CSS** for styling
- **shadcn/ui** + Radix primitives for base components (tabs, dialog, etc.)
- **Zustand** for client-side state

## Document engines
- **Univer** SDK (https://univer.ai) — primary engine for doc / sheet / slide
- **CodeMirror 6** + Markdown extension — for the Markdown tab (lighter than Univer for plain MD notes)
- **JSZip** — for `.gd` archive read/write

## AI layer
- Default provider: **Anthropic Claude** (Sonnet 4.6 default, Opus 4.7 for heavy thinking)
- SDK: `@anthropic-ai/sdk` with prompt caching enabled
- Provider abstraction allows swapping in OpenAI / local Ollama later
- **Tool use / function calling** is the mechanism by which AI mutates documents

## Storage
- `.gd` archive layout (zip):
  ```
  manifest.json        — { version, title, tabs: [{name, type, file}] }
  doc/<files>          — actual md/docx/xlsx/pptx files
  assets/images/       — embedded resources
  ```
- **better-sqlite3** for chat history + undo stack (per-session local DB)

## Folder layout (locked)
```
C:\Gen-Doc\
├── docs/                  — owned by Agent 1 (spec writer)
├── src/
│   ├── main/              — Electron main process (Agent 2)
│   ├── renderer/          — React app entry (Agent 2)
│   ├── components/
│   │   ├── AIPanel.tsx    — AI sidebar (Agent 4)
│   │   ├── AIPanel.css    — AI sidebar styles (Agent 4)
│   │   └── …other UI…     — owned by Agent 2
│   └── ai/
│       ├── tools/         — tool-calling JSON schemas (Agent 3)
│       └── …provider…     — owned by Agent 2 stub, Agent 4 can refine
├── package.json           — Agent 2 owns
├── tsconfig.json          — Agent 2 owns
├── vite.config.ts         — Agent 2 owns
└── README.md              — Agent 2 stub
```

## Cross-agent invariants
- Use **TypeScript** strict mode. No `any` unless absolutely necessary.
- All source files end with newline.
- Default to **Tailwind classnames**, not custom CSS, unless animation requires it.
- Comments only when the WHY is non-obvious.
- Prefer composition over inheritance.
- All cross-module types live in `src/types/` (Agent 2 sets this up; others import).
- **繁體中文**：UI 文字、user-facing message、文檔（spec / README）一律繁中
- 錯誤訊息 / log: 英文（方便除錯）
