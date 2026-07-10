/**
 * Apple Reminders connector — import from iOS Shortcut / file exports.
 *
 * Export from iOS: run a Shortcut that saves open reminders as JSON or text,
 * then use Sweep Reminders in TaskSweep Hub.
 */

import type { AppSettings, Connector, RawInput } from '../../types/task'
import {
  parseRemindersFile,
  reminderItemToRawContent,
} from '../remindersParser'

let pendingFiles: File[] = []

export function setRemindersFiles(files: FileList | File[]): void {
  pendingFiles = Array.from(files).filter((file) => {
    const name = file.name.toLowerCase()
    return name.endsWith('.json') || name.endsWith('.txt') || name.endsWith('.csv')
  })
}

export function clearRemindersFiles(): void {
  pendingFiles = []
}

export function createAppleRemindersConnector(getSettings: () => AppSettings): Connector {
  return {
    id: 'apple-reminders',
    name: 'Apple Reminders',
    description: 'Import JSON or text exports from iOS Reminders / Shortcuts',
    requiresAuth: false,
    isAvailable: () => getSettings().remindersEnabled !== false,
    async sweep() {
      if (pendingFiles.length === 0) return []

      const inputs: RawInput[] = []
      const now = new Date().toISOString()

      for (const file of pendingFiles) {
        try {
          const text = await file.text()
          const items = parseRemindersFile(text, file.name)
          for (const item of items) {
            inputs.push({
              id: crypto.randomUUID(),
              source: 'apple-reminders',
              content: reminderItemToRawContent(item),
              receivedAt: now,
              metadata: {
                filename: file.name,
                ...(item.id ? { reminderId: item.id } : {}),
                ...(item.list ? { remindersList: item.list } : {}),
                ...(item.dueDate ? { dueDate: item.dueDate } : {}),
              },
            })
          }
        } catch (err) {
          console.warn(`Reminders import skipped ${file.name}:`, err)
        }
      }

      clearRemindersFiles()
      return inputs
    },
  }
}