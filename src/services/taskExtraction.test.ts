import { describe, expect, it } from 'vitest'
import { extractTasksLocally } from './taskExtraction'
import type { RawInput } from '../types/task'

function pasteInput(content: string): RawInput {
  return {
    id: 'test-1',
    source: 'paste',
    content,
    receivedAt: '2026-01-01T00:00:00.000Z',
  }
}

describe('taskExtraction', () => {
  it('extracts bullet-list tasks', () => {
    const tasks = extractTasksLocally(pasteInput(`
- Call Cedar Ridge contractor re: permit
- Board meeting prep (nonprofit) — due 3/15
- Review client ticket #4521 — urgent
    `))

    expect(tasks.length).toBeGreaterThanOrEqual(3)
    expect(tasks[0].title).toContain('Cedar Ridge')
    expect(tasks.some((t) => t.tags?.includes('nonprofit'))).toBe(true)
    expect(tasks.some((t) => t.priority === 'urgent')).toBe(true)
  })

  it('extracts due dates', () => {
    const tasks = extractTasksLocally(pasteInput('- File taxes — due 4/15/2026'))
    expect(tasks[0]?.dueDate).toBeTruthy()
  })

  it('tags Cedar Ridge home project context', () => {
    const tasks = extractTasksLocally(pasteInput('- Pick up lumber for Cedar Ridge deck'))
    expect(tasks[0]?.tags).toContain('Cedar Ridge')
  })

  it('ignores empty lines and noise', () => {
    const tasks = extractTasksLocally(pasteInput(`


Hello team,

- Actually do the thing

Thanks,
Bob
    `))
    expect(tasks.some((t) => t.title.includes('Actually do'))).toBe(true)
    expect(tasks.some((t) => t.title.includes('Hello team'))).toBe(false)
  })
})