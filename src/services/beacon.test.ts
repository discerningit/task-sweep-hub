import { describe, expect, it } from 'vitest'
import {
  createBeaconText,
  scanForBeacons,
  scanPasteForBeacon,
} from './beacon'
import type { RawInput } from '../types/task'

describe('beacon', () => {
  it('creates copyable beacon text with default marker', () => {
    const text = createBeaconText()
    expect(text).toContain('[TaskSweep-Beacon]')
    expect(text).toContain('connectivity test')
  })

  it('respects custom beacon marker from settings', () => {
    const text = createBeaconText({ beaconMarker: '[MyBeacon]' })
    expect(text).toContain('[MyBeacon]')
  })

  it('detects beacon in pasted text', () => {
    const hit = scanPasteForBeacon(
      'Subject: [TaskSweep-Beacon] Test from Outlook\nSome body text',
    )
    expect(hit).not.toBeNull()
    expect(hit?.source).toBe('paste')
    expect(hit?.suggestedConnector).toBeTruthy()
  })

  it('returns null when no beacon present', () => {
    const hit = scanPasteForBeacon('Regular email with no marker')
    expect(hit).toBeNull()
  })

  it('scans multiple raw inputs for beacons', () => {
    const inputs: RawInput[] = [
      {
        id: '1',
        source: 'm365-todo',
        content: '[TaskSweep-Beacon] planted in To Do',
        receivedAt: '2026-01-01T00:00:00.000Z',
      },
      {
        id: '2',
        source: 'paste',
        content: 'Normal pasted tasks only',
        receivedAt: '2026-01-01T00:00:00.000Z',
      },
    ]
    const hits = scanForBeacons(inputs)
    expect(hits).toHaveLength(1)
    expect(hits[0].source).toBe('m365-todo')
  })
})