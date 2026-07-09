/**
 * Beacon system — plant a marker like [TaskSweep-Beacon] in any
 * task list, email subject, or note. When you paste or sweep that
 * ecosystem, TaskSweep finds the beacon and suggests how to connect.
 */

import type { AppSettings, BeaconHit, RawInput, TaskSource } from '../types/task'

const DEFAULT_MARKER = '[TaskSweep-Beacon]'

/** Build beacon text the user can copy into Outlook, To Do, etc. */
export function createBeaconText(settings?: Pick<AppSettings, 'beaconMarker'>): string {
  const marker = settings?.beaconMarker ?? DEFAULT_MARKER
  return `${marker} TaskSweep connectivity test — safe to delete after setup`
}

/** Scan raw inputs for beacon markers */
export function scanForBeacons(
  inputs: RawInput[],
  settings?: Pick<AppSettings, 'beaconMarker'>,
): BeaconHit[] {
  const marker = settings?.beaconMarker ?? DEFAULT_MARKER
  const hits: BeaconHit[] = []
  const seen = new Set<string>()

  for (const input of inputs) {
    if (!input.content.includes(marker)) continue

    const key = `${input.source}:${marker}`
    if (seen.has(key)) continue
    seen.add(key)

    hits.push({
      source: input.source,
      marker,
      context: extractBeaconContext(input.content, marker),
      suggestedConnector: suggestConnector(input.source),
    })
  }

  return hits
}

function extractBeaconContext(content: string, marker: string): string {
  const line = content
    .split('\n')
    .find((l) => l.includes(marker))
  return line?.trim().slice(0, 120) ?? marker
}

function suggestConnector(source: TaskSource): string {
  const map: Partial<Record<TaskSource, string>> = {
    paste: 'Keep pasting from this app — or try File Upload for exports',
    file: 'File connector works — consider M365 if this came from Outlook',
    'proton-mail': 'Proton Mail connector — export .eml files and use Sweep Proton Mail',
    'm365-todo': 'M365 To Do connector is active — run a full M365 sweep',
    'm365-outlook': 'M365 Outlook connector — enable Graph API in Settings',
    'm365-onenote': 'M365 OneNote — notes often hold hidden tasks',
    'm365-teams': 'M365 Teams — check channel tasks and meeting actions',
    beacon: 'Beacon detected — configure the matching connector in Settings',
    manual: 'Manual entry works everywhere, including locked-down VDI',
  }
  return map[source] ?? 'Try Paste or M365 connectors in Settings'
}

/** Quick scan of pasted text for beacons (no full sweep needed) */
export function scanPasteForBeacon(
  text: string,
  settings?: Pick<AppSettings, 'beaconMarker'>,
): BeaconHit | null {
  const marker = settings?.beaconMarker ?? DEFAULT_MARKER
  if (!text.includes(marker)) return null

  return {
    source: 'paste',
    marker,
    context: extractBeaconContext(text, marker),
    suggestedConnector: 'Beacon found in paste — note which app this came from',
  }
}