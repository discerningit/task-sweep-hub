import { describe, expect, it, vi, beforeEach } from 'vitest'
import type { AppSettings, Task } from '../types/task'
import { pushNewTasksToPrimaryTool, reconcileToDoCompletions } from './primaryToolPush'

vi.mock('./connectors/m365', () => ({
  isM365SignedIn: vi.fn(() => true),
  createM365TodoTask: vi.fn(),
  getM365TodoTaskStatus: vi.fn(),
}))

vi.mock('./remindersPush', () => ({
  pushToAppleReminders: vi.fn(),
}))

import { pushToAppleReminders } from './remindersPush'

import { createM365TodoTask, getM365TodoTaskStatus, isM365SignedIn } from './connectors/m365'
const settings: AppSettings = {
  enabledAiProviders: ['local'],
  primaryAi: 'local',
  primaryTaskTool: 'ms-todo',
  beaconMarker: '[TaskSweep-Beacon]',
  contextTags: [],
  m365ClientId: 'test-client',
}

function task(overrides: Partial<Task> = {}): Task {
  return {
    id: '1',
    title: 'New pasted task',
    priority: 'normal',
    status: 'open',
    source: 'paste',
    tags: [],
    contentHash: 'x',
    similarityKey: 'new',
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z',
    ...overrides,
  }
}

describe('primaryToolPush', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    vi.mocked(isM365SignedIn).mockReturnValue(true)
    vi.mocked(createM365TodoTask).mockResolvedValue({ id: 'todo-new', listId: 'list-1' })
  })

  it('pushes paste tasks to Microsoft To Do', async () => {
    const result = await pushNewTasksToPrimaryTool([task()], settings)
    expect(createM365TodoTask).toHaveBeenCalled()
    expect(result.pushedCount).toBe(1)
    expect(result.tasks[0].metadata?.pushedToMsTodo).toBe('true')
    expect(result.tasks[0].syncStatus).toBe('synced')
  })

  it('skips tasks already from m365-todo', async () => {
    const result = await pushNewTasksToPrimaryTool(
      [task({ source: 'm365-todo', metadata: { id: 't1', listId: 'l1' } })],
      settings,
    )
    expect(createM365TodoTask).not.toHaveBeenCalled()
    expect(result.pushedCount).toBe(0)
  })

  it('does nothing when primary tool is hub-only', async () => {
    const result = await pushNewTasksToPrimaryTool(
      [task()],
      { ...settings, primaryTaskTool: 'hub-only' },
    )
    expect(createM365TodoTask).not.toHaveBeenCalled()
    expect(result.pushedCount).toBe(0)
  })

  it('delegates to Apple Reminders push when configured', async () => {
    vi.mocked(pushToAppleReminders).mockResolvedValue({
      tasks: [task({ metadata: { pushedToReminders: 'true' }, syncStatus: 'synced' })],
      pushedCount: 1,
      failedCount: 0,
      message: 'Shared 1 task(s)',
    })

    const result = await pushNewTasksToPrimaryTool(
      [task()],
      { ...settings, primaryTaskTool: 'apple-reminders' },
    )

    expect(pushToAppleReminders).toHaveBeenCalled()
    expect(createM365TodoTask).not.toHaveBeenCalled()
    expect(result.pushedCount).toBe(1)
    expect(result.tasks[0].metadata?.pushedToReminders).toBe('true')
  })
})

describe('reconcileToDoCompletions', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    vi.mocked(isM365SignedIn).mockReturnValue(true)
  })

  it('marks hub tasks complete when done in To Do', async () => {
    vi.mocked(getM365TodoTaskStatus).mockResolvedValue('completed')

    const result = await reconcileToDoCompletions(
      [
        task({
          metadata: { id: 'todo-1', listId: 'list-1', pushedToMsTodo: 'true' },
          sourceId: 'todo-1',
        }),
      ],
      settings,
    )

    expect(result.completedCount).toBe(1)
    expect(result.updated[0].status).toBe('completed')
    expect(result.updated[0].syncMessage).toContain('To Do')
  })

  it('leaves open tasks when still open in To Do', async () => {
    vi.mocked(getM365TodoTaskStatus).mockResolvedValue('notStarted')

    const result = await reconcileToDoCompletions(
      [task({ metadata: { id: 'todo-1', pushedToMsTodo: 'true' } })],
      settings,
    )

    expect(result.completedCount).toBe(0)
    expect(result.updated).toHaveLength(0)
  })
})