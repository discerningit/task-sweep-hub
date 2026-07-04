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
import { orchestrateExtraction } from './aiOrchestrator'
import { deduplicateAgainstExisting } from './deduplication'
import { scanForBeacons } from './beacon'

export interface SweepResult {
  newTaskCount: number
  totalScanned: number
  beacons: BeaconHit[]
  sources: string[]
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

  for (const connector of connectors) {
    const available = await connector.isAvailable()
    if (!available) continue

    const inputs = await connector.sweep()
    if (inputs.length > 0) {
      sources.push(connector.name)
      allInputs.push(...inputs)
    }
  }

  const beacons = scanForBeacons(allInputs, settings)

  const extracted = await orchestrateExtraction(allInputs, settings)
  const existing = await getAllTasks()
  const newTasks = deduplicateAgainstExisting(extracted, existing)

  if (newTasks.length > 0) {
    await saveTasks(newTasks)
  }

  return {
    newTaskCount: newTasks.length,
    totalScanned: allInputs.length,
    beacons,
    sources,
  }
}