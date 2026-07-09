/**
 * TaskSweep Hub — shared types
 *
 * Every task the app knows about flows through these shapes.
 * Connectors produce RawInput → extraction produces ExtractedTask →
 * dedup merges into Task (stored in IndexedDB).
 */

/** Where a task originally came from */
export type TaskSource =
  | 'paste'
  | 'file'
  | 'm365-todo'
  | 'm365-outlook'
  | 'm365-onenote'
  | 'm365-teams'
  | 'beacon'
  | 'manual'
  | 'export'

export type TaskStatus = 'open' | 'completed' | 'snoozed'

export type TaskPriority = 'low' | 'normal' | 'high' | 'urgent'

/** A chunk of raw text or metadata from any connector */
export interface RawInput {
  id: string
  source: TaskSource
  content: string
  /** Optional link back to the original item (email, To Do item, etc.) */
  sourceUrl?: string
  /** Connector-specific metadata (message id, list name, etc.) */
  metadata?: Record<string, string>
  receivedAt: string
}

/** Task shape after AI or local extraction (before dedup) */
export interface ExtractedTask {
  title: string
  dueDate?: string
  priority: TaskPriority
  notes?: string
  source: TaskSource
  sourceUrl?: string
  sourceId?: string
  metadata?: Record<string, string>
  /** Tags inferred from context (Cedar Ridge, nonprofit, family, etc.) */
  tags?: string[]
}

/** Final task stored in IndexedDB and shown in the UI */
export interface Task {
  id: string
  title: string
  dueDate?: string
  priority: TaskPriority
  notes?: string
  status: TaskStatus
  source: TaskSource
  sourceUrl?: string
  sourceId?: string
  metadata?: Record<string, string>
  tags: string[]
  /** Hash for exact dedup */
  contentHash: string
  /** Fingerprint for fuzzy similarity dedup */
  similarityKey: string
  createdAt: string
  updatedAt: string
  completedAt?: string
  snoozedUntil?: string
  /** Result of last sync-back attempt to M365 / primary tool */
  syncStatus?: 'synced' | 'failed' | 'skipped' | 'local'
  syncMessage?: string
}

/** Graph-backed connectors available per M365 account */
export type M365ConnectorSource = 'todo' | 'outlook' | 'onenote'

/** A signed-in Microsoft 365 account (personal or work) */
export interface M365Account {
  /** MSAL home account ID — stable key across sessions */
  homeAccountId: string
  username: string
  name?: string
  tenantId?: string
  /** User label, e.g. "Personal" or "Work" */
  label?: string
  /** Which Graph sources to sweep/sync for this account (default: smart by tenant) */
  enabledSources?: M365ConnectorSource[]
}

/** User preferences — all local, no cloud required */
export interface AppSettings {
  /** Which AI providers the user has enabled */
  enabledAiProviders: AiProviderId[]
  /** Default AI for extraction */
  primaryAi: AiProviderId
  /** M365 tenant/client config (optional) */
  m365ClientId?: string
  /** Signed-in M365 accounts discovered via MSAL */
  m365Accounts?: M365Account[]
  /** Default account for pushing new tasks to Microsoft To Do */
  m365ActiveAccountId?: string
  /** Which accounts to include in M365 sweeps (default: all signed-in) */
  m365SweepAccountIds?: string[]
  /** xAI Grok API key (optional, stored locally on device) */
  grokApiKey?: string
  /** Where completed tasks should sync when possible */
  primaryTaskTool: PrimaryTaskTool
  /** Custom beacon marker text (default: [TaskSweep-Beacon]) */
  beaconMarker: string
  /** Context tags the AI should watch for */
  contextTags: string[]
  /** True after manual setup or importing a device setup pack */
  setupCompleted?: boolean
}

export type AiProviderId =
  | 'local'
  | 'copilot'
  | 'claude'
  | 'kiro'
  | 'grok'
  | 'siri'

export type PrimaryTaskTool =
  | 'hub-only'
  | 'ms-todo'
  | 'todoist'
  | 'apple-reminders'
  | 'jira'
  | 'export-csv'

/** Connector interface — add new sources by implementing this */
export interface Connector {
  id: string
  name: string
  description: string
  /** Whether this connector needs auth (M365, etc.) */
  requiresAuth: boolean
  /** Is the connector configured and ready? */
  isAvailable: () => boolean | Promise<boolean>
  /** Pull raw inputs for sweeping */
  sweep: () => Promise<RawInput[]>
}

/** Result of beacon scan */
export interface BeaconHit {
  source: TaskSource
  marker: string
  context: string
  suggestedConnector?: string
}