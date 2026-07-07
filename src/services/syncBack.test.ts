import { describe, expect, it, vi, beforeEach } from 'vitest'
import type { AppSettings, Task } from '../types/task'
import { completeTask } from './syncBack'

vi.mock('./connectors/m365', () => ({
  isM365SignedIn: vi.fn(() => true),
  completeM365TodoTask: vi.fn(),
  clearM365OutlookFlag: vi.fn(),
}))

import {
  clearM365OutlookFlag,
  completeM365TodoTask,
  isM365SignedIn,
} from './connectors/m365'

const settings: AppSettings = {
  enabledAiProviders: ['local'],
  primaryAi: 'local',
  primaryTaskTool: 'hub-only',
  beaconMarker: '[TaskSweep-Beacon]',
  contextTags: [],
  m365ClientId: 'test-client',
}

function baseTask(overrides: Partial<Task> = {}): Task {
  return {
    id: '1',
    title: 'Test task',
    priority: 'normal',
    status: 'completed',
    source: 'paste',
    tags: [],
    contentHash: 'abc',
    similarityKey: 'test',
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z',
    ...overrides,
  }
}

describe('syncBack', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    vi.mocked(isM365SignedIn).mockReturnValue(true)
    vi.mocked(completeM365TodoTask).mockResolvedValue(undefined)
    vi.mocked(clearM365OutlookFlag).mockResolvedValue(undefined)
  })

  it('syncs m365-todo completion to Graph', async () => {
    const result = await completeTask(
      baseTask({
        source: 'm365-todo',
        sourceId: 'todo-99',
        metadata: { id: 'todo-99', listId: 'list-1' },
      }),
      settings,
    )
    expect(completeM365TodoTask).toHaveBeenCalledWith(settings, 'todo-99', 'list-1')
    expect(result.syncStatus).toBe('synced')
    expect(result.success).toBe(true)
  })

  it('clears Outlook flag for m365-outlook tasks', async () => {
    const result = await completeTask(
      baseTask({
        source: 'm365-outlook',
        metadata: { id: 'msg-42' },
      }),
      settings,
    )
    expect(clearM365OutlookFlag).toHaveBeenCalledWith(settings, 'msg-42')
    expect(result.syncStatus).toBe('synced')
  })

  it('skips M365 sync when not signed in', async () => {
    vi.mocked(isM365SignedIn).mockReturnValue(false)
    const result = await completeTask(
      baseTask({ source: 'm365-todo', metadata: { id: 't1', listId: 'l1' } }),
      settings,
    )
    expect(completeM365TodoTask).not.toHaveBeenCalled()
    expect(result.syncStatus).toBe('skipped')
  })

  it('marks local-only for paste tasks', async () => {
    const result = await completeTask(baseTask({ source: 'paste' }), settings)
    expect(result.syncStatus).toBe('local')
  })

  it('completes pushed To Do tasks when primary tool is ms-todo', async () => {
    const result = await completeTask(
      baseTask({
        source: 'paste',
        sourceId: 'todo-pushed',
        metadata: { id: 'todo-pushed', listId: 'list-1', pushedToMsTodo: 'true' },
      }),
      { ...settings, primaryTaskTool: 'ms-todo' },
    )
    expect(completeM365TodoTask).toHaveBeenCalledWith(
      expect.objectContaining({ primaryTaskTool: 'ms-todo' }),
      'todo-pushed',
      'list-1',
    )
    expect(result.syncStatus).toBe('synced')
  })
})