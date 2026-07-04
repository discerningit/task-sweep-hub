/**
 * Sync-back — when you complete or edit a task in the hub,
 * push changes to the source (M365) or primary tool when possible.
 *
 * MVP: logs intent and updates local state. M365 write-back is stubbed.
 */

import type { AppSettings, Task } from '../types/task'

export interface SyncResult {
  success: boolean
  message: string
}

/** Mark complete in hub and attempt source sync */
export async function completeTask(
  task: Task,
  settings: AppSettings,
): Promise<SyncResult> {
  if (settings.primaryTaskTool === 'export-csv') {
    return { success: true, message: 'Marked complete locally. Use Export for backup.' }
  }

  if (task.source.startsWith('m365-')) {
    // TODO: PATCH via Graph API when M365 auth is active
    return {
      success: true,
      message: `Marked complete in hub. M365 sync-back coming soon (source: ${task.source}).`,
    }
  }

  if (settings.primaryTaskTool === 'ms-todo' && task.source !== 'm365-todo') {
    return {
      success: true,
      message: 'Marked complete in hub. MS To Do push requires M365 connector.',
    }
  }

  return { success: true, message: 'Marked complete in TaskSweep Hub.' }
}

export async function snoozeTask(
  task: Task,
  until: string,
  _settings: AppSettings,
): Promise<SyncResult> {
  void task
  return {
    success: true,
    message: `Snoozed until ${until}. Source sync not yet implemented.`,
  }
}

/** Export tasks as CSV for Todoist / Reminders import */
export function exportTasksCsv(tasks: Task[]): string {
  const header = 'title,dueDate,priority,status,source,tags,notes'
  const rows = tasks.map((t) =>
    [
      csvEscape(t.title),
      csvEscape(t.dueDate ?? ''),
      t.priority,
      t.status,
      t.source,
      csvEscape(t.tags.join(';')),
      csvEscape(t.notes ?? ''),
    ].join(','),
  )
  return [header, ...rows].join('\n')
}

function csvEscape(value: string): string {
  if (value.includes(',') || value.includes('"') || value.includes('\n')) {
    return `"${value.replace(/"/g, '""')}"`
  }
  return value
}