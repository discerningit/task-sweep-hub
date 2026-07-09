/**
 * Device setup pack — move settings to a new phone or laptop without
 * re-typing the M365 Client ID and other options.
 *
 * Does NOT copy sign-in tokens. You still tap "Sign in to M365" once per device.
 */

import type { AppSettings } from '../types/task'

export const SETUP_PACK_VERSION = 1

export interface SettingsPack {
  version: number
  app: 'tasksweep-hub'
  exportedAt: string
  settings: Pick<
    AppSettings,
    | 'm365ClientId'
    | 'm365ActiveAccountId'
    | 'm365SweepAccountIds'
    | 'protonMailEnabled'
    | 'beaconMarker'
    | 'contextTags'
    | 'primaryTaskTool'
    | 'enabledAiProviders'
    | 'primaryAi'
    | 'grokApiKey'
  >
}

export function createSettingsPack(settings: AppSettings): SettingsPack {
  return {
    version: SETUP_PACK_VERSION,
    app: 'tasksweep-hub',
    exportedAt: new Date().toISOString(),
    settings: {
      m365ClientId: settings.m365ClientId,
      m365ActiveAccountId: settings.m365ActiveAccountId,
      m365SweepAccountIds: settings.m365SweepAccountIds,
      protonMailEnabled: settings.protonMailEnabled,
      beaconMarker: settings.beaconMarker,
      contextTags: settings.contextTags,
      primaryTaskTool: settings.primaryTaskTool,
      enabledAiProviders: settings.enabledAiProviders,
      primaryAi: settings.primaryAi,
      grokApiKey: settings.grokApiKey,
    },
  }
}

export function encodeSettingsPack(settings: AppSettings): string {
  return JSON.stringify(createSettingsPack(settings), null, 2)
}

export function parseSettingsPack(raw: string): SettingsPack {
  const trimmed = raw.trim()
  let parsed: unknown

  try {
    parsed = JSON.parse(trimmed)
  } catch {
    throw new Error('Setup pack is not valid JSON. Copy the full export from your other device.')
  }

  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Setup pack format not recognized.')
  }

  const pack = parsed as Partial<SettingsPack>
  if (pack.app !== 'tasksweep-hub') {
    throw new Error('This file is not a TaskSweep Hub setup pack.')
  }
  if (!pack.settings || typeof pack.settings !== 'object') {
    throw new Error('Setup pack is missing settings.')
  }

  return pack as SettingsPack
}

/** Merge imported pack into current settings and mark setup complete */
export function applySettingsPack(
  current: AppSettings,
  pack: SettingsPack,
): AppSettings {
  return {
    ...current,
    ...pack.settings,
    setupCompleted: true,
  }
}

export function needsDeviceSetup(settings: AppSettings): boolean {
  if (settings.setupCompleted) return false
  return !settings.m365ClientId
}

export function downloadSettingsPack(settings: AppSettings): void {
  const json = encodeSettingsPack(settings)
  const blob = new Blob([json], { type: 'application/json' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `tasksweep-setup-${new Date().toISOString().slice(0, 10)}.json`
  a.click()
  URL.revokeObjectURL(url)
}