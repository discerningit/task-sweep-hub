/**
 * Sweep pipeline — ties connectors → AI extraction → dedup → save.
 * This is the main "Sweep" button logic.
 */

import {
  getAllTasks,
  saveTasks,
  getSettings,
} from '../db/indexedDb'
import type { BeaconHit, RawInput } from '../types/task'
import { getAllConnectors } from './connectors'
import {
  getLastOneNoteSweepResult,
  isM365SignedIn,
  sweepM365OneNote,
  type OneNoteSweepResult,
} from './connectors/m365'
import { orchestrateExtraction, type ExtractionStatus } from './aiOrchestrator'
import { deduplicateAgainstExisting } from './deduplication'
import { scanForBeacons } from './beacon'
import { pushNewTasksToPrimaryTool, reconcileToDoCompletions } from './primaryToolPush'


export interface SweepResult {
  newTaskCount: number
  totalScanned: number
  beacons: BeaconHit[]
  sources: string[]
  pushedToTodoCount: number
  pushFailedCount: number
  completedFromTodoCount: number
  extraction?: ExtractionStatus
  onenotePagesFound?: number
  onenotePagesImported?: number
  onenoteSectionsScanned?: number
  onenoteDetail?: string
  onenoteError?: string
}

async function finalizeSweep(
  allInputs: RawInput[],
  sources: string[],
  settings: Awaited<ReturnType<typeof getSettings>>,
  onenoteStats?: OneNoteSweepResult,
): Promise<SweepResult> {
  const beacons = scanForBeacons(allInputs, settings)

  const extraction = await orchestrateExtraction(allInputs, settings)
  const existing = await getAllTasks()
  const newTasks = deduplicateAgainstExisting(extraction.tasks, existing)

  let pushedToTodoCount = 0
  let pushFailedCount = 0

  if (newTasks.length > 0) {
    const pushResult = await pushNewTasksToPrimaryTool(newTasks, settings)
    pushedToTodoCount = pushResult.pushedCount
    pushFailedCount = pushResult.failedCount
    await saveTasks(pushResult.tasks)
  }

  const latestTasks = await getAllTasks()
  const reconcile = await reconcileToDoCompletions(latestTasks, settings)
  if (reconcile.updated.length > 0) {
    await saveTasks(reconcile.updated)
  }

  const onenoteImported =
    onenoteStats?.pagesImported ??
    allInputs.filter((input) => input.source === 'm365-onenote').length

  return {
    newTaskCount: newTasks.length,
    totalScanned: allInputs.length,
    beacons,
    sources,
    pushedToTodoCount,
    pushFailedCount,
    completedFromTodoCount: reconcile.completedCount,
    extraction: extraction.status,
    onenotePagesFound: onenoteStats?.pagesFound,
    onenotePagesImported: onenoteImported,
    onenoteSectionsScanned: onenoteStats?.sectionsScanned,
    onenoteDetail: onenoteStats?.detail,
    onenoteError: onenoteStats?.error,
  }
}

export async function runSweep(
  connectorIds?: string[],
): Promise<SweepResult> {
  const settings = await getSettings()
  const connectors = getAllConnectors(settings).filter(
    (c) => !connectorIds || connectorIds.includes(c.id),
  )

  const allInputs: RawInput[] = []
  const sources: string[] = []
  let m365Swept = false

  for (const connector of connectors) {
    const available = await connector.isAvailable()
    if (!available) continue

    const inputs = await connector.sweep()
    if (connector.id === 'm365') m365Swept = true
    if (inputs.length > 0) {
      sources.push(connector.name)
      allInputs.push(...inputs)
    }
  }

  const onenoteStats = m365Swept ? getLastOneNoteSweepResult() : undefined

  return finalizeSweep(allInputs, sources, settings, onenoteStats)
}

/** Sweep OneNote only — useful for testing beacon pages and note-based tasks */
export async function runOneNoteSweep(): Promise<SweepResult> {
  const settings = await getSettings()
  if (!settings.m365ClientId || !isM365SignedIn()) {
    throw new Error('Sign in to Microsoft 365 in Settings first.')
  }

  const onenote = await sweepM365OneNote(settings)
  const sources = onenote.inputs.length > 0 ? ['OneNote'] : []

  return finalizeSweep(onenote.inputs, sources, settings, onenote)
}