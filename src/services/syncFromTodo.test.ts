import { describe, expect, it, vi, beforeEach } from 'vitest'

vi.mock('../db/indexedDb', () => ({
  getSettings: vi.fn(),
  getAllTasks: vi.fn(),
  saveTasks: vi.fn(),
}))

vi.mock('./connectors/m365', () => ({
  isM365SignedIn: vi.fn(),
  sweepM365TodoOnly: vi.fn(),
}))

vi.mock('./aiOrchestrator', () => ({
  orchestrateExtraction: vi.fn(),
}))

vi.mock('./deduplication', () => ({
  deduplicateAgainstExisting: vi.fn(),
}))

vi.mock('./primaryToolPush', () => ({
  reconcileToDoCompletions: vi.fn(),
}))

import { getAllTasks, getSettings } from '../db/indexedDb'
import { isM365SignedIn, sweepM365TodoOnly } from './connectors/m365'
import { orchestrateExtraction } from './aiOrchestrator'
import { deduplicateAgainstExisting } from './deduplication'
import { reconcileToDoCompletions } from './primaryToolPush'
import { runSyncFromTodo } from './syncFromTodo'

import type { AppSettings } from '../types/task'

const settings: AppSettings = {
  enabledAiProviders: ['local'],
  primaryAi: 'local',
  primaryTaskTool: 'ms-todo',
  beaconMarker: '[TaskSweep-Beacon]',
  contextTags: [],
  m365ClientId: 'client-1',
}

describe('syncFromTodo', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    vi.mocked(getSettings).mockResolvedValue(settings)
    vi.mocked(isM365SignedIn).mockReturnValue(true)
    vi.mocked(getAllTasks).mockResolvedValue([])
    vi.mocked(sweepM365TodoOnly).mockResolvedValue([])
    vi.mocked(orchestrateExtraction).mockResolvedValue({
      tasks: [],
      status: {
        provider: 'local',
        requestedProvider: 'local',
        usedFallback: false,
        extractedCount: 0,
      },
    })
    vi.mocked(deduplicateAgainstExisting).mockReturnValue([])
    vi.mocked(reconcileToDoCompletions).mockResolvedValue({ updated: [], completedCount: 0 })
  })

  it('requires M365 sign-in', async () => {
    vi.mocked(isM365SignedIn).mockReturnValue(false)
    const result = await runSyncFromTodo()
    expect(result.message).toContain('Sign in')
    expect(sweepM365TodoOnly).not.toHaveBeenCalled()
  })

  it('reports new tasks and completions', async () => {
    vi.mocked(sweepM365TodoOnly).mockResolvedValue([
      {
        id: 'in-1',
        source: 'm365-todo',
        content: 'New from todo',
        receivedAt: '2026-01-01T00:00:00.000Z',
      },
    ])
    vi.mocked(orchestrateExtraction).mockResolvedValue({
      tasks: [{ title: 'New from todo', priority: 'normal', source: 'm365-todo' }],
      status: {
        provider: 'local',
        requestedProvider: 'local',
        usedFallback: false,
        extractedCount: 1,
      },
    })
    vi.mocked(deduplicateAgainstExisting).mockReturnValue([
      {
        id: 't1',
        title: 'New from todo',
        priority: 'normal',
        status: 'open',
        source: 'm365-todo',
        tags: [],
        contentHash: 'h',
        similarityKey: 'k',
        createdAt: '2026-01-01T00:00:00.000Z',
        updatedAt: '2026-01-01T00:00:00.000Z',
      },
    ])
    vi.mocked(getAllTasks).mockResolvedValue([])
    vi.mocked(reconcileToDoCompletions).mockResolvedValue({ updated: [], completedCount: 2 })

    const result = await runSyncFromTodo()
    expect(result.newTaskCount).toBe(1)
    expect(result.completedCount).toBe(2)
    expect(result.message).toContain('1 new from To Do')
    expect(result.message).toContain('2 marked done')
  })
})