import { describe, expect, it } from 'vitest'
import type { AppSettings, M365Account } from '../types/task'
import {
  defaultAccountLabel,
  getActiveM365AccountId,
  getSweepAccountIds,
  mergeM365Accounts,
  removeM365AccountFromSettings,
  stampM365Metadata,
} from './m365Accounts'

const personal: M365Account = {
  homeAccountId: 'personal-1',
  username: 'me@outlook.com',
  tenantId: '9188040d-6ce7-4e72-b656-0671adf88c0b',
}

const work: M365Account = {
  homeAccountId: 'work-1',
  username: 'me@contoso.com',
  tenantId: 'org-tenant',
}

const baseSettings: AppSettings = {
  enabledAiProviders: ['local'],
  primaryAi: 'local',
  primaryTaskTool: 'hub-only',
  beaconMarker: '[TaskSweep-Beacon]',
  contextTags: [],
  m365Accounts: [personal, work],
  m365ActiveAccountId: 'work-1',
  m365SweepAccountIds: ['personal-1'],
}

describe('m365Accounts', () => {
  it('labels personal vs work tenants', () => {
    expect(defaultAccountLabel(personal)).toBe('Personal')
    expect(defaultAccountLabel(work)).toBe('Work')
  })

  it('returns sweep-enabled account ids', () => {
    expect(getSweepAccountIds(baseSettings)).toEqual(['personal-1'])
  })

  it('defaults sweep to all accounts when none selected', () => {
    const settings = { ...baseSettings, m365SweepAccountIds: undefined }
    expect(getSweepAccountIds(settings)).toEqual(['personal-1', 'work-1'])
  })

  it('resolves active account for To Do push', () => {
    expect(getActiveM365AccountId(baseSettings)).toBe('work-1')
  })

  it('stamps account metadata on swept inputs', () => {
    const meta = stampM365Metadata({ id: 'page-1' }, { ...personal, label: 'Personal' })
    expect(meta.m365HomeAccountId).toBe('personal-1')
    expect(meta.m365Username).toBe('me@outlook.com')
    expect(meta.m365AccountLabel).toBe('Personal')
  })

  it('merges stored labels with MSAL accounts', () => {
    const merged = mergeM365Accounts(
      [{ ...personal, label: 'My Personal' }],
      [
        {
          homeAccountId: 'personal-1',
          environment: 'login.windows.net',
          tenantId: personal.tenantId!,
          username: personal.username,
          localAccountId: '1',
          name: 'Me',
          idTokenClaims: {},
        },
      ],
    )
    expect(merged[0].label).toBe('My Personal')
  })

  it('removes account from settings on sign-out', () => {
    const next = removeM365AccountFromSettings(baseSettings, 'work-1')
    expect(next.m365Accounts?.map((a) => a.homeAccountId)).toEqual(['personal-1'])
    expect(next.m365ActiveAccountId).toBe('personal-1')
  })
})