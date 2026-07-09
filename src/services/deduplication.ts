/**
 * Deduplication — prevents the same task appearing twice
 * when you sweep from email, paste, and M365 To Do at once.
 *
 * Two layers:
 * 1. contentHash — exact match on normalized title + due date
 * 2. similarityKey — basic fuzzy match (shared words in title)
 */

import type { ExtractedTask, Task } from '../types/task'

/** Normalize text for hashing: lowercase, trim, collapse spaces */
function normalize(text: string): string {
  return text.toLowerCase().trim().replace(/\s+/g, ' ')
}

/** Simple string hash (djb2) — good enough for client-side dedup */
export function hashString(input: string): string {
  let hash = 5381
  for (let i = 0; i < input.length; i++) {
    hash = (hash * 33) ^ input.charCodeAt(i)
  }
  return (hash >>> 0).toString(36)
}

export function buildContentHash(
  task: Pick<ExtractedTask, 'title' | 'dueDate' | 'metadata'>,
): string {
  const accountId = task.metadata?.m365HomeAccountId ?? ''
  const key = `${normalize(task.title)}|${task.dueDate ?? ''}|${accountId}`
  return hashString(key)
}

/** Extract significant words for similarity grouping */
export function buildSimilarityKey(title: string): string {
  const stopWords = new Set([
    'a', 'an', 'the', 'to', 'for', 'and', 'or', 'on', 'in', 'at', 'by',
    'is', 'it', 'of', 'with', 'from', 'as', 'be', 'this', 'that',
  ])
  const words = normalize(title)
    .replace(/[^\w\s]/g, '')
    .split(/\s+/)
    .filter((w) => w.length > 2 && !stopWords.has(w))
    .sort()
  return words.slice(0, 6).join('-') || normalize(title)
}

/** How similar two titles are (0–1) using word overlap */
export function titleSimilarity(a: string, b: string): number {
  const wordsA = new Set(normalize(a).split(/\s+/).filter((w) => w.length > 2))
  const wordsB = new Set(normalize(b).split(/\s+/).filter((w) => w.length > 2))
  if (wordsA.size === 0 || wordsB.size === 0) return 0

  let overlap = 0
  for (const w of wordsA) {
    if (wordsB.has(w)) overlap++
  }
  return overlap / Math.max(wordsA.size, wordsB.size)
}

const SIMILARITY_THRESHOLD = 0.75

/**
 * Merge extracted tasks into existing tasks, skipping duplicates.
 * Returns only NEW tasks to add.
 */
export function deduplicateAgainstExisting(
  extracted: ExtractedTask[],
  existing: Task[],
): Task[] {
  const now = new Date().toISOString()
  const hashIndex = new Map(existing.map((t) => [t.contentHash, t]))
  const newTasks: Task[] = []

  for (const item of extracted) {
    const contentHash = buildContentHash(item)
    if (hashIndex.has(contentHash)) continue

    const similarityKey = buildSimilarityKey(item.title)
    const fuzzyMatch = existing.some(
      (t) =>
        t.similarityKey === similarityKey ||
        titleSimilarity(t.title, item.title) >= SIMILARITY_THRESHOLD,
    )
    if (fuzzyMatch) continue

    const task: Task = {
      id: crypto.randomUUID(),
      title: item.title,
      dueDate: item.dueDate,
      priority: item.priority,
      notes: item.notes,
      status: 'open',
      source: item.source,
      sourceUrl: item.sourceUrl,
      sourceId: item.sourceId,
      metadata: item.metadata,
      tags: item.tags ?? [],
      contentHash,
      similarityKey,
      createdAt: now,
      updatedAt: now,
    }
    newTasks.push(task)
    hashIndex.set(contentHash, task)
    existing.push(task)
  }

  return newTasks
}