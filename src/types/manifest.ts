/**
 * manifest.json schema for `.gd` archive — see docs/MVP-SPEC.md §7.1.
 * The `tabs` field controls which editor renders for each entry.
 */

export type TabType = 'markdown' | 'html' | 'docx' | 'xlsx' | 'pptx';

export interface TabDescriptor {
  id: string;
  name: string;
  type: TabType;
  /** Relative path inside the .gd archive, e.g. "doc/notes.md". */
  file: string;
  /** Display order; ascending. */
  order: number;
}

export interface ManifestSettings {
  embedChatHistory: boolean;
  defaultModel: string;
}

export interface ManifestMetadata {
  appVersion: string;
  /** Reserved for v2.0 multi-user prep. */
  lastEditedBy?: string;
}

export interface Manifest {
  version: '1.0';
  title: string;
  /** ISO-8601. */
  createdAt: string;
  /** ISO-8601. */
  modifiedAt: string;
  tabs: TabDescriptor[];
  settings: ManifestSettings;
  metadata: ManifestMetadata;
}

export const MANIFEST_VERSION = '1.0' as const;
export const MANIFEST_FILENAME = 'manifest.json' as const;
