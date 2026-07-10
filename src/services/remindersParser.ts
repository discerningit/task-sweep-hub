/**
 * Apple Reminders import/export — works with iOS Shortcuts and file exports.
 * No web API exists; this supports file import and Share-sheet push formats.
 */

import type { Task } from '../types/task'

export const REMINDERS_EXPORT_APP = 'tasksweep-reminders-export'
export const REMINDERS_EXPORT_VERSION = 1

export interface ReminderItem {
  title: string
  notes?: string
  dueDate?: string
  list?: string
  completed?: boolean
  id?: string
}

export interface RemindersExportPack {
  app: typeof REMINDERS_EXPORT_APP
  version: number
  exportedAt: string
  reminders: ReminderItem[]
}

/** Parse JSON export from TaskSweep or an iOS Shortcut */
export function parseRemindersJson(text: string): ReminderItem[] {
  const parsed = JSON.parse(text.trim()) as unknown

  if (Array.isArray(parsed)) {
    return parsed.map(normalizeReminder).filter((r) => r.title && !r.completed)
  }

  if (parsed && typeof parsed === 'object') {
    const pack = parsed as Partial<RemindersExportPack>
    if (Array.isArray(pack.reminders)) {
      return pack.reminders.map(normalizeReminder).filter((r) => r.title && !r.completed)
    }
  }

  throw new Error('Reminders JSON must be an array or { "reminders": [...] }')
}

/** Parse plain-text exports (Shortcuts, copy-paste from Reminders lists) */
export function parseRemindersText(text: string): ReminderItem[] {
  const lines = text.replace(/\r\n/g, '\n').split('\n')
  const items: ReminderItem[] = []
  let currentList: string | undefined

  for (const raw of lines) {
    const line = raw.trim()
    if (!line) continue

    const listHeading = line.match(/^#{1,3}\s+(.+)$/) ?? line.match(/^(.+):$/)
    if (listHeading && !TASK_LINE.test(line)) {
      currentList = listHeading[1].trim()
      continue
    }

    if (!TASK_LINE.test(line)) continue

    const title = cleanReminderLine(line)
    if (title.length < 2) continue

    items.push({
      title,
      dueDate: extractInlineDue(title) ?? undefined,
      list: currentList,
    })
  }

  return items
}

const TASK_LINE =
  /^(?:[-*•]\s*|\[\s?[xX ]?\]\s*|\d+[.)]\s*|TODO[:\s]|REMINDER[:\s])/i

function cleanReminderLine(line: string): string {
  return line
    .replace(TASK_LINE, '')
    .replace(/\s*[-–—]\s*(due|by)\s*.*/i, '')
    .trim()
}

function extractInlineDue(title: string): string | undefined {
  const match = title.match(/\b(?:due|by)\s+(\S.{2,30})/i)
  return match?.[1]?.trim()
}

function normalizeReminder(raw: unknown): ReminderItem {
  if (!raw || typeof raw !== 'object') {
    return { title: '' }
  }
  const item = raw as Record<string, unknown>
  const title = String(item.title ?? item.name ?? '').trim()
  return {
    title,
    notes: item.notes ? String(item.notes) : item.body ? String(item.body) : undefined,
    dueDate: item.dueDate
      ? String(item.dueDate)
      : item.due
        ? String(item.due)
        : undefined,
    list: item.list ? String(item.list) : item.listName ? String(item.listName) : undefined,
    completed: Boolean(item.completed ?? item.isCompleted),
    id: item.id ? String(item.id) : item.reminderId ? String(item.reminderId) : undefined,
  }
}

export function parseRemindersFile(text: string, filename: string): ReminderItem[] {
  const lower = filename.toLowerCase()
  if (lower.endsWith('.json')) {
    return parseRemindersJson(text)
  }
  return parseRemindersText(text)
}

/** Plain text for iOS Share sheet → Shortcut → Add Reminder */
export function formatRemindersShareText(
  tasks: Task[],
  defaultList?: string,
): string {
  const header = defaultList ? `List: ${defaultList}\n` : ''
  const lines = tasks.map((task) => {
    const parts = [task.title]
    if (task.dueDate) parts.push(`(due ${task.dueDate})`)
    if (task.notes) parts.push(`— ${task.notes}`)
    return parts.join(' ')
  })
  return `${header}${lines.join('\n')}`.trim()
}

export function createRemindersExportPack(tasks: Task[]): RemindersExportPack {
  return {
    app: REMINDERS_EXPORT_APP,
    version: REMINDERS_EXPORT_VERSION,
    exportedAt: new Date().toISOString(),
    reminders: tasks.map((task) => ({
      title: task.title,
      notes: task.notes,
      dueDate: task.dueDate,
      list: task.metadata?.remindersList,
      completed: task.status === 'completed',
      id: task.sourceId ?? task.metadata?.reminderId,
    })),
  }
}

export function reminderItemToRawContent(item: ReminderItem): string {
  const lines = [item.title]
  if (item.notes) lines.push(item.notes)
  if (item.dueDate) lines.push(`Due: ${item.dueDate}`)
  if (item.list) lines.push(`List: ${item.list}`)
  return lines.join('\n')
}