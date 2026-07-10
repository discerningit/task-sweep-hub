/**
 * Push tasks to Apple Reminders via Share sheet or file download.
 * iOS has no browser API for Reminders — user completes via Shortcut or manual import.
 */

import {
  createRemindersExportPack,
  formatRemindersShareText,
} from './remindersParser'
import type { AppSettings, Task } from '../types/task'

export interface RemindersPushResult {
  tasks: Task[]
  pushedCount: number
  failedCount: number
  message: string
}

export function alreadyLinkedToReminders(task: Task): boolean {
  return (
    task.source === 'apple-reminders' ||
    task.metadata?.pushedToReminders === 'true'
  )
}

function downloadFile(filename: string, body: string, mime: string): void {
  const blob = new Blob([body], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function markRemindersPushed(
  all: Task[],
  pushed: Task[],
  settings: AppSettings,
  now: string,
): Task[] {
  const pushedIds = new Set(pushed.map((t) => t.id))
  return all.map((task) => {
    if (!pushedIds.has(task.id)) return task
    return {
      ...task,
      syncStatus: 'synced' as const,
      syncMessage: 'Sent to Apple Reminders (via Share or download)',
      metadata: {
        ...task.metadata,
        pushedToReminders: 'true',
        ...(settings.remindersDefaultList
          ? { remindersList: settings.remindersDefaultList }
          : {}),
      },
      updatedAt: now,
    }
  })
}

/** Share or download open tasks for Apple Reminders */
export async function pushToAppleReminders(
  tasks: Task[],
  settings: AppSettings,
): Promise<RemindersPushResult> {
  const pending = tasks.filter((t) => t.status === 'open' && !alreadyLinkedToReminders(t))
  if (pending.length === 0) {
    return {
      tasks,
      pushedCount: 0,
      failedCount: 0,
      message: 'No new tasks to send to Reminders.',
    }
  }

  const shareText = formatRemindersShareText(pending, settings.remindersDefaultList)
  const json = JSON.stringify(createRemindersExportPack(pending), null, 2)
  const now = new Date().toISOString()

  if (typeof navigator.share === 'function') {
    try {
      const files: File[] = [
        new File([json], 'tasksweep-reminders.json', { type: 'application/json' }),
      ]
      const shareData: ShareData = {
        title: 'TaskSweep → Reminders',
        text: shareText,
      }
      if (navigator.canShare?.({ files })) {
        shareData.files = files
      }
      await navigator.share(shareData)

      return {
        tasks: markRemindersPushed(tasks, pending, settings, now),
        pushedCount: pending.length,
        failedCount: 0,
        message: `Shared ${pending.length} task(s) — choose your Add to Reminders Shortcut.`,
      }
    } catch (err) {
      if (err instanceof Error && err.name === 'AbortError') {
        return {
          tasks,
          pushedCount: 0,
          failedCount: 0,
          message: 'Share cancelled.',
        }
      }
    }
  }

  downloadFile('tasksweep-reminders.json', json, 'application/json')
  downloadFile('tasksweep-reminders.txt', shareText, 'text/plain')

  return {
    tasks: markRemindersPushed(tasks, pending, settings, now),
    pushedCount: pending.length,
    failedCount: 0,
    message: `Downloaded ${pending.length} task(s) as .json and .txt for Reminders import.`,
  }
}