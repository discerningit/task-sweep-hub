import { describe, expect, it } from 'vitest'
import {
  buildContentHash,
  buildSimilarityKey,
  deduplicateAgainstExisting,
  titleSimilarity,
} from './deduplication'
import type { ExtractedTask, Task } from '../types/task'

function makeTask(overrides: Partial<Task> = {}): Task {
  return {
    id: 'existing-1',
    title: 'Call Cedar Ridge contractor',
    priority: 'normal',
    status: 'open',
    source: 'paste',
    tags: [],
    contentHash: buildContentHash({ title: 'Call Cedar Ridge contractor' }),
    similarityKey: buildSimilarityKey('Call Cedar Ridge contractor'),
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z',
    ...overrides,
  }
}

describe('deduplication', () => {
  it('builds stable content hashes for the same title', () => {
    const a = buildContentHash({ title: 'Review client ticket #4521' })
    const b = buildContentHash({ title: 'Review client ticket #4521' })
    expect(a).toBe(b)
  })

  it('treats different due dates as different tasks', () => {
    const a = buildContentHash({ title: 'Board meeting prep', dueDate: '3/15' })
    const b = buildContentHash({ title: 'Board meeting prep', dueDate: '3/16' })
    expect(a).not.toBe(b)
  })

  it('detects high title similarity', () => {
    const score = titleSimilarity(
      'Call Cedar Ridge contractor',
      'Call contractor Cedar Ridge',
    )
    expect(score).toBeGreaterThanOrEqual(0.75)
  })

  it('skips exact duplicates', () => {
    const existing = [makeTask()]
    const extracted: ExtractedTask[] = [
      { title: 'Call Cedar Ridge contractor', priority: 'normal', source: 'paste' },
    ]
    const result = deduplicateAgainstExisting(extracted, existing)
    expect(result).toHaveLength(0)
  })

  it('skips fuzzy duplicates', () => {
    const existing = [makeTask({ title: 'Call Cedar Ridge contractor' })]
    const extracted: ExtractedTask[] = [
      { title: 'Call contractor Cedar Ridge', priority: 'normal', source: 'paste' },
    ]
    const result = deduplicateAgainstExisting(extracted, existing)
    expect(result).toHaveLength(0)
  })

  it('adds genuinely new tasks', () => {
    const existing = [makeTask()]
    const extracted: ExtractedTask[] = [
      { title: 'Submit nonprofit grant report', priority: 'high', source: 'paste' },
    ]
    const result = deduplicateAgainstExisting(extracted, existing)
    expect(result).toHaveLength(1)
    expect(result[0].title).toBe('Submit nonprofit grant report')
    expect(result[0].id).toBeTruthy()
  })
})