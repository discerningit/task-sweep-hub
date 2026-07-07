/**
 * Sync-back — when you complete a task in the hub,
 * push changes to the source (M365) where possible.
 */

import {
  clearM365OutlookFlag,
  completeM365TodoTask,
  isM365SignedIn,
} from './connectors/m365'
import type { AppSettings, Task } from '../types/task'

export interface SyncResult {
  success: boolean
  message: string
  syncStatus: Task['syncStatus']
}

/** Mark complete in hub and attempt source sync */
export async function completeTask(
  task: Task,
  settings: AppSettings,
): Promise<SyncResult> {
  if (task.source === 'm365-todo') {
    return syncM365TodoComplete(task, settings)
  }

  if (task.source === 'm365-outlook') {
    return syncM365OutlookComplete(task, settings)
  }

  const pushedTodoId = task.metadata?.pushedToMsTodo === 'true'
    ? (task.sourceId ?? task.metadata?.id)
    : undefined
  if (pushedTodoId && settings.primaryTaskTool === 'ms-todo') {
    return syncM365TodoComplete(
      {
        ...task,
        source: 'm365-todo',
        sourceId: pushedTodoId,
        metadata: { ...task.metadata, id: pushedTodoId },
      },
      settings,
    )
  }

  return {
    success: true,
    message: 'Marked complete in TaskSweep Hub.',
    syncStatus: 'local',
  }
}

async function syncM365TodoComplete(
  task: Task,
  settings: AppSettings,
): Promise<SyncResult> {
  const taskId = task.sourceId ?? task.metadata?.id
  if (!taskId) {
    return {
      success: true,
      message: 'Marked complete locally. Missing To Do task ID — re-sweep from M365.',
      syncStatus: 'skipped',
    }
  }

  if (!settings.m365ClientId || !isM365SignedIn()) {
    return {
      success: true,
      message: 'Marked complete locally. Sign in to M365 in Settings to sync To Do.',
      syncStatus: 'skipped',
    }
  }

  try {
    await completeM365TodoTask(settings, taskId, task.metadata?.listId)
    return {
      success: true,
      message: 'Completed in TaskSweep and Microsoft To Do.',
      syncStatus: 'synced',
    }
  } catch (err) {
    const detail = err instanceof Error ? err.message : 'Sync failed'
    return {
      success: false,
      message: `Marked complete locally. To Do sync failed: ${detail}`,
      syncStatus: 'failed',
    }
  }
}

async function syncM365OutlookComplete(
  task: Task,
  settings: AppSettings,
): Promise<SyncResult> {
  const messageId = task.sourceId ?? task.metadata?.id
  if (!messageId) {
    return {
      success: true,
      message: 'Marked complete locally. Missing email ID — re-sweep from M365.',
      syncStatus: 'skipped',
    }
  }

  if (!settings.m365ClientId || !isM365SignedIn()) {
    return {
      success: true,
      message: 'Marked complete locally. Sign in to M365 in Settings to clear Outlook flag.',
      syncStatus: 'skipped',
    }
  }

  try {
    await clearM365OutlookFlag(settings, messageId)
    return {
      success: true,
      message: 'Completed in TaskSweep and cleared Outlook flag.',
      syncStatus: 'synced',
    }
  } catch (err) {
    const detail = err instanceof Error ? err.message : 'Sync failed'
    return {
      success: false,
      message: `Marked complete locally. Outlook sync failed: ${detail}`,
      syncStatus: 'failed',
    }
  }
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
    syncStatus: 'local',
  }
}

/** Export tasks as CSV for Todoist / Reminders import */
export function exportTasksCsv(tasks: Task[]): string {
  const header = 'title,dueDate,priority,status,source,tags,notes,syncStatus'
  const rows = tasks.map((t) =>
    [
      csvEscape(t.title),
      csvEscape(t.dueDate ?? ''),
      t.priority,
      t.status,
      t.source,
      csvEscape(t.tags.join(';')),
      csvEscape(t.notes ?? ''),
      t.syncStatus ?? '',
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