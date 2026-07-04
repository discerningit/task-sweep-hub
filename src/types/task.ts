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
}

/** User preferences — all local, no cloud required */
export interface AppSettings {
  /** Which AI providers the user has enabled */
  enabledAiProviders: AiProviderId[]
  /** Default AI for extraction */
  primaryAi: AiProviderId
  /** M365 tenant/client config (optional) */
  m365ClientId?: string
  /** Where completed tasks should sync when possible */
  primaryTaskTool: PrimaryTaskTool
  /** Custom beacon marker text (default: [TaskSweep-Beacon]) */
  beaconMarker: string
  /** Context tags the AI should watch for */
  contextTags: string[]
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