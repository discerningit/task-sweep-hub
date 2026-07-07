/**
 * Push new tasks to the user's primary task tool after a sweep.
 * MS To Do: creates tasks via Graph API when primary tool is ms-todo.
 */

import {
  createM365TodoTask,
  getM365TodoTaskStatus,
  isM365SignedIn,
} from './connectors/m365'
import type { AppSettings, Task } from '../types/task'

export interface PushResult {
  tasks: Task[]
  pushedCount: number
  failedCount: number
}

function alreadyInMsTodo(task: Task): boolean {
  return task.source === 'm365-todo' || task.metadata?.pushedToMsTodo === 'true'
}

/** Push newly swept tasks to Microsoft To Do when configured as primary tool */
export async function pushNewTasksToPrimaryTool(
  tasks: Task[],
  settings: AppSettings,
): Promise<PushResult> {
  if (settings.primaryTaskTool !== 'ms-todo') {
    return { tasks, pushedCount: 0, failedCount: 0 }
  }

  if (!settings.m365ClientId || !isM365SignedIn()) {
    return {
      tasks: tasks.map((t) =>
        alreadyInMsTodo(t)
          ? t
          : {
              ...t,
              syncStatus: 'skipped',
              syncMessage: 'Sign in to M365 in Settings to push new tasks to To Do.',
            },
      ),
      pushedCount: 0,
      failedCount: 0,
    }
  }

  let pushedCount = 0
  let failedCount = 0
  const updated: Task[] = []

  for (const task of tasks) {
    if (alreadyInMsTodo(task)) {
      updated.push(task)
      continue
    }

    try {
      const created = await createM365TodoTask(settings, task)
      pushedCount++
      updated.push({
        ...task,
        sourceId: created.id,
        sourceUrl: `https://to-do.office.com/tasks/id/${created.id}`,
        metadata: {
          ...task.metadata,
          id: created.id,
          listId: created.listId,
          pushedToMsTodo: 'true',
        },
        syncStatus: 'synced',
        syncMessage: 'Pushed to Microsoft To Do',
        updatedAt: new Date().toISOString(),
      })
    } catch (err) {
      failedCount++
      const detail = err instanceof Error ? err.message : 'Push failed'
      updated.push({
        ...task,
        syncStatus: 'failed',
        syncMessage: `To Do push failed: ${detail}`,
        updatedAt: new Date().toISOString(),
      })
    }
  }

  return { tasks: updated, pushedCount, failedCount }
}

export interface ReconcileResult {
  updated: Task[]
  completedCount: number
}

/** Link to a Microsoft To Do task (swept from M365 or pushed from hub) */
export function getLinkedTodoId(task: Task): { id: string; listId?: string } | null {
  const id = task.sourceId ?? task.metadata?.id
  if (!id) return null
  if (task.source === 'm365-todo' || task.metadata?.pushedToMsTodo === 'true') {
    return { id, listId: task.metadata?.listId }
  }
  return null
}

/**
 * Inbound sync — if a task was completed in Microsoft To Do,
 * mark it completed in TaskSweep too.
 */
export async function reconcileToDoCompletions(
  allTasks: Task[],
  settings: AppSettings,
): Promise<ReconcileResult> {
  if (!settings.m365ClientId || !isM365SignedIn()) {
    return { updated: [], completedCount: 0 }
  }

  const updated: Task[] = []
  let completedCount = 0
  const now = new Date().toISOString()

  for (const task of allTasks) {
    if (task.status !== 'open') continue

    const link = getLinkedTodoId(task)
    if (!link) continue

    try {
      const status = await getM365TodoTaskStatus(settings, link.id, link.listId)
      if (status !== 'completed') continue

      completedCount++
      updated.push({
        ...task,
        status: 'completed',
        completedAt: now,
        updatedAt: now,
        syncStatus: 'synced',
        syncMessage: 'Completed in Microsoft To Do — synced to TaskSweep',
      })
    } catch (err) {
      console.warn('To Do status check failed:', task.title, err)
    }
  }

  return { updated, completedCount }
}