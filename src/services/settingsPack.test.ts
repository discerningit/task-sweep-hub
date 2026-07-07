import { describe, expect, it } from 'vitest'
import { DEFAULT_SETTINGS } from '../db/indexedDb'
import {
  applySettingsPack,
  encodeSettingsPack,
  needsDeviceSetup,
  parseSettingsPack,
} from './settingsPack'

describe('settingsPack', () => {
  it('encodes and parses a setup pack', () => {
    const settings = {
      ...DEFAULT_SETTINGS,
      m365ClientId: 'test-client-id',
      setupCompleted: true,
    }
    const json = encodeSettingsPack(settings)
    const pack = parseSettingsPack(json)
    expect(pack.app).toBe('tasksweep-hub')
    expect(pack.settings.m365ClientId).toBe('test-client-id')
  })

  it('rejects invalid JSON', () => {
    expect(() => parseSettingsPack('not json')).toThrow(/valid JSON/)
  })

  it('rejects wrong app id', () => {
    expect(() =>
      parseSettingsPack(JSON.stringify({ app: 'other', settings: {} })),
    ).toThrow(/not a TaskSweep/)
  })

  it('applies pack and marks setup complete', () => {
    const pack = parseSettingsPack(
      encodeSettingsPack({
        ...DEFAULT_SETTINGS,
        m365ClientId: 'abc-123',
      }),
    )
    const next = applySettingsPack(DEFAULT_SETTINGS, pack)
    expect(next.m365ClientId).toBe('abc-123')
    expect(next.setupCompleted).toBe(true)
  })

  it('detects when new device needs setup', () => {
    expect(needsDeviceSetup(DEFAULT_SETTINGS)).toBe(true)
    expect(
      needsDeviceSetup({ ...DEFAULT_SETTINGS, m365ClientId: 'x' }),
    ).toBe(false)
    expect(
      needsDeviceSetup({ ...DEFAULT_SETTINGS, setupCompleted: true }),
    ).toBe(false)
  })
})