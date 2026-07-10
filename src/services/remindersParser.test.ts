import { describe, expect, it } from 'vitest'
import {
  formatRemindersShareText,
  parseRemindersJson,
  parseRemindersText,
} from './remindersParser'
import type { Task } from '../types/task'

describe('remindersParser', () => {
  it('parses TaskSweep JSON export', () => {
    const items = parseRemindersJson(
      JSON.stringify({
        app: 'tasksweep-reminders-export',
        version: 1,
        reminders: [
          { title: 'Call contractor', dueDate: '7/15', list: 'Cedar Ridge', completed: false },
          { title: 'Done item', completed: true },
        ],
      }),
    )
    expect(items).toHaveLength(1)
    expect(items[0].title).toBe('Call contractor')
    expect(items[0].list).toBe('Cedar Ridge')
  })

  it('parses plain-text bullet lists', () => {
    const items = parseRemindersText(`
Groceries:
- Buy milk
- Eggs due 7/10

- Board prep
`)
    expect(items).toHaveLength(3)
    expect(items[0].title).toBe('Buy milk')
    expect(items[0].list).toBe('Groceries')
  })

  it('formats share text for Shortcuts', () => {
    const tasks: Task[] = [
      {
        id: '1',
        title: 'Call client',
        dueDate: '7/12',
        priority: 'normal',
        status: 'open',
        source: 'paste',
        tags: [],
        contentHash: 'a',
        similarityKey: 'b',
        createdAt: '',
        updatedAt: '',
      },
    ]
    const text = formatRemindersShareText(tasks, 'Work')
    expect(text).toContain('List: Work')
    expect(text).toContain('Call client')
    expect(text).toContain('due 7/12')
  })
})