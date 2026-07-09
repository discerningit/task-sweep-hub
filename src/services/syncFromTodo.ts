/**
 * B-48 — Sync from Microsoft To Do (manual, no full sweep).
 *
 * 1. Pull new open tasks from To Do into the hub
 * 2. Mark hub tasks complete when done in To Do
 */

import { getAllTasks, getSettings, saveTasks } from '../db/indexedDb'
import { isM365SignedIn, sweepM365TodoOnly } from './connectors/m365'
import { getSweepAccountIdsForSource } from './m365Accounts'
import { orchestrateExtraction } from './aiOrchestrator'
import { deduplicateAgainstExisting } from './deduplication'
import { reconcileToDoCompletions } from './primaryToolPush'

export interface SyncFromTodoResult {
  newTaskCount: number
  completedCount: number
  message: string
}

export async function runSyncFromTodo(): Promise<SyncFromTodoResult> {
  const settings = await getSettings()

  if (!settings.m365ClientId || !isM365SignedIn()) {
    return {
      newTaskCount: 0,
      completedCount: 0,
      message: 'Sign in to Microsoft 365 in Settings first.',
    }
  }

  let newTaskCount = 0

  try {
    const accountIds = getSweepAccountIdsForSource(settings, 'todo')
    const inputs = (
      await Promise.all(accountIds.map((id) => sweepM365TodoOnly(settings, id)))
    ).flat()
    if (inputs.length > 0) {
      const extraction = await orchestrateExtraction(inputs, settings)
      const existing = await getAllTasks()
      const newTasks = deduplicateAgainstExisting(extraction.tasks, existing)
      if (newTasks.length > 0) {
        await saveTasks(newTasks)
        newTaskCount = newTasks.length
      }
    }
  } catch (err) {
    const detail = err instanceof Error ? err.message : 'To Do read failed'
    return {
      newTaskCount: 0,
      completedCount: 0,
      message: `Sync from To Do failed: ${detail}`,
    }
  }

  const latest = await getAllTasks()
  const reconcile = await reconcileToDoCompletions(latest, settings)
  if (reconcile.updated.length > 0) {
    await saveTasks(reconcile.updated)
  }

  const parts: string[] = []
  if (newTaskCount > 0) parts.push(`${newTaskCount} new from To Do`)
  if (reconcile.completedCount > 0) parts.push(`${reconcile.completedCount} marked done from To Do`)
  if (parts.length === 0) parts.push('Already in sync with Microsoft To Do')

  return {
    newTaskCount,
    completedCount: reconcile.completedCount,
    message: `Sync from To Do: ${parts.join(', ')}.`,
  }
}