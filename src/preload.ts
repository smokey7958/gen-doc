/**
 * Preload script — runs in an isolated context with Node enabled and exposes
 * a typed bridge to the renderer via `contextBridge`. The renderer cannot
 * import Node modules directly (sandbox + contextIsolation: true).
 */

import { contextBridge, ipcRenderer, webUtils } from 'electron';
import { IPC, type GenDocBridge } from './types/ipc';

const bridge: GenDocBridge = {
  app: {
    info: () => ipcRenderer.invoke(IPC.app.info),
    setDirty: (dirty) => {
      // Use `send` (not `invoke`) — main caches the value and we don't need
      // a response. Keeping it sync-style fits the renderer's `useEffect`
      // dependency on ws.dirty without dragging in a Promise.
      ipcRenderer.send(IPC.app.setDirty, dirty);
    },
    confirm: (message) => ipcRenderer.invoke(IPC.app.confirm, message),
    saveAndQuitResult: (ok) => {
      ipcRenderer.send(IPC.app.saveAndQuitResult, ok);
    },
    getOsLocale: () => ipcRenderer.invoke(IPC.app.getOsLocale),
  },
  workspace: {
    open: () => ipcRenderer.invoke(IPC.workspace.open),
    openPath: (filePath) => ipcRenderer.invoke(IPC.workspace.openPath, filePath),
    save: (req) => ipcRenderer.invoke(IPC.workspace.save, req),
    saveAs: (req) => ipcRenderer.invoke(IPC.workspace.saveAs, req),
    exportTab: (req) => ipcRenderer.invoke(IPC.workspace.exportTab, req),
    exportTabs: (req) => ipcRenderer.invoke(IPC.workspace.exportTabs, req),
  },
  markdown: {
    exportPdf: (req) => ipcRenderer.invoke(IPC.markdown.exportPdf, req),
  },
  fs: {
    pickDirectory: () => ipcRenderer.invoke(IPC.fs.pickDirectory),
    listDirectory: (dirPath) => ipcRenderer.invoke(IPC.fs.listDirectory, dirPath),
    readFile: (filePath) => ipcRenderer.invoke(IPC.fs.readFile, filePath),
  },
  config: {
    get: () => ipcRenderer.invoke(IPC.config.get),
    set: (patch) => ipcRenderer.invoke(IPC.config.set, patch),
    hasApiKey: () => ipcRenderer.invoke(IPC.config.hasApiKey),
    setApiKey: (key) => ipcRenderer.invoke(IPC.config.setApiKey, key),
    clearApiKey: () => ipcRenderer.invoke(IPC.config.clearApiKey),
  },
  ai: {
    ping: (model) => ipcRenderer.invoke(IPC.ai.ping, model),
    chat: (req) => ipcRenderer.invoke(IPC.ai.chat, req),
    cancel: (requestId) => ipcRenderer.invoke(IPC.ai.cancel, requestId),
    onChunk: (handler) => {
      const listener = (_event: Electron.IpcRendererEvent, payload: unknown) => {
        handler(payload as { requestId: string; chunk: unknown });
      };
      ipcRenderer.on(IPC.ai.chunk, listener);
      return () => {
        ipcRenderer.removeListener(IPC.ai.chunk, listener);
      };
    },
  },
  history: {
    listConversations: (workspaceId) =>
      ipcRenderer.invoke(IPC.history.listConversations, workspaceId),
    listMessages: (conversationId) =>
      ipcRenderer.invoke(IPC.history.listMessages, conversationId),
    appendMessage: (conversationId, row) =>
      ipcRenderer.invoke(IPC.history.appendMessage, conversationId, row),
    createConversation: (opts) => ipcRenderer.invoke(IPC.history.createConversation, opts),
  },
  undo: {
    push: (entry) => ipcRenderer.invoke(IPC.undo.push, entry),
    pop: (workspaceId) => ipcRenderer.invoke(IPC.undo.pop, workspaceId),
    list: (workspaceId, limit) => ipcRenderer.invoke(IPC.undo.list, workspaceId, limit),
    clear: (workspaceId) => ipcRenderer.invoke(IPC.undo.clear, workspaceId),
    relink: (oldId, newId) => ipcRenderer.invoke(IPC.undo.relink, oldId, newId),
  },
  webUtils: {
    getPathForFile: (file) => webUtils.getPathForFile(file),
  },
};

contextBridge.exposeInMainWorld('gendoc', bridge);

// Menu commands flow main → renderer over per-action channels. We don't
// re-expose ipcRenderer to the renderer (security), so we relay them as
// CustomEvents on `window` instead — App.tsx subscribes to those.
const MENU_CHANNELS = [
  'menu:newProject',
  'menu:open',
  'menu:openFolder',
  'menu:save',
  'menu:saveAs',
  'menu:saveAndQuit',
  'menu:exportTab',
  'menu:batchExport',
  'menu:undo',
  'menu:redo',
  'menu:focusAI',
  'menu:openSettings',
  'menu:openRecent',
  'menu:clearRecent',
];
for (const channel of MENU_CHANNELS) {
  // Forward any args (e.g. the path argument on `menu:openRecent`) on
  // CustomEvent.detail so the renderer can act on them. detail is the
  // raw arg array — single-arg consumers grab `detail[0]`.
  ipcRenderer.on(channel, (_event, ...args: unknown[]) => {
    window.dispatchEvent(new CustomEvent(channel, { detail: args }));
  });
}
