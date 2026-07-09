/**
 * Multi-account M365 helpers — personal + work accounts at once.
 */

import type { AccountInfo } from '@azure/msal-browser'
import type {
  AppSettings,
  M365Account,
  M365ConnectorSource,
  Task,
  TaskSource,
} from '../types/task'

/** Metadata key stamped on tasks/inputs from a specific M365 account */
export const M365_HOME_ACCOUNT_ID_KEY = 'm365HomeAccountId'
export const M365_USERNAME_KEY = 'm365Username'
export const M365_ACCOUNT_LABEL_KEY = 'm365AccountLabel'

export const M365_CONNECTOR_SOURCE_LABELS: Record<M365ConnectorSource, string> = {
  todo: 'Microsoft To Do',
  outlook: 'Outlook mail',
  onenote: 'OneNote',
}

export const M365_SOURCE_SCOPES: Record<M365ConnectorSource, string> = {
  todo: 'Tasks.ReadWrite',
  outlook: 'Mail.ReadWrite',
  onenote: 'Notes.Read',
}

export const M365_BASE_SCOPES = ['User.Read'] as const

/** Scopes requested on first sign-in (work-tenant friendly) */
export const M365_LOGIN_SCOPES = ['User.Read', 'Tasks.ReadWrite'] as const

export const ALL_M365_CONNECTOR_SOURCES: M365ConnectorSource[] = [
  'todo',
  'outlook',
  'onenote',
]

/** Azure AD consumer tenant — personal Microsoft accounts */
const PERSONAL_TENANT_ID = '9188040d-6ce7-4e72-b656-0671adf88c0b'

const M365_SOURCES: TaskSource[] = [
  'm365-todo',
  'm365-outlook',
  'm365-onenote',
  'm365-teams',
]

const TASK_SOURCE_TO_CONNECTOR: Partial<Record<TaskSource, M365ConnectorSource>> = {
  'm365-todo': 'todo',
  'm365-outlook': 'outlook',
  'm365-onenote': 'onenote',
}

export function isM365TaskSource(source: TaskSource): boolean {
  return M365_SOURCES.includes(source)
}

export function connectorSourceForTask(source: TaskSource): M365ConnectorSource | undefined {
  return TASK_SOURCE_TO_CONNECTOR[source]
}

export function accountFromMsal(account: AccountInfo): M365Account {
  const base = {
    homeAccountId: account.homeAccountId,
    username: account.username,
    name: account.name,
    tenantId: account.tenantId,
    label: defaultAccountLabel(account),
  }
  return {
    ...base,
    enabledSources: defaultEnabledSourcesForAccount(base),
  }
}

export function defaultAccountLabel(account: Pick<M365Account, 'tenantId' | 'username'>): string {
  if (account.tenantId === PERSONAL_TENANT_ID) return 'Personal'
  if (account.username) {
    const domain = account.username.split('@')[1]?.toLowerCase()
    if (domain && !['outlook.com', 'hotmail.com', 'live.com', 'msn.com'].includes(domain)) {
      return 'Work'
    }
  }
  return 'Personal'
}

/** Work tenants often allow To Do only — personal gets all sources */
export function defaultEnabledSourcesForAccount(
  account: Pick<M365Account, 'tenantId' | 'username' | 'label'>,
): M365ConnectorSource[] {
  if (defaultAccountLabel(account) === 'Work') {
    return ['todo']
  }
  return [...ALL_M365_CONNECTOR_SOURCES]
}

export function getAccountEnabledSources(account?: M365Account): M365ConnectorSource[] {
  if (!account) return [...ALL_M365_CONNECTOR_SOURCES]
  if (account.enabledSources?.length) return account.enabledSources
  return defaultEnabledSourcesForAccount(account)
}

export function isAccountSourceEnabled(
  account: M365Account | undefined,
  source: M365ConnectorSource,
): boolean {
  return getAccountEnabledSources(account).includes(source)
}

export function scopesForSources(sources: M365ConnectorSource[]): string[] {
  const scopes = new Set<string>([...M365_BASE_SCOPES])
  for (const source of sources) {
    scopes.add(M365_SOURCE_SCOPES[source])
  }
  return [...scopes]
}

export function scopesForAccount(account?: M365Account): string[] {
  return scopesForSources(getAccountEnabledSources(account))
}

/** Merge MSAL accounts with stored labels/preferences */
export function mergeM365Accounts(
  stored: M365Account[],
  msalAccounts: AccountInfo[],
): M365Account[] {
  const storedById = new Map(stored.map((a) => [a.homeAccountId, a]))

  return msalAccounts.map((msal) => {
    const prev = storedById.get(msal.homeAccountId)
    const fresh = accountFromMsal(msal)
    return {
      ...fresh,
      label: prev?.label ?? fresh.label,
      enabledSources: prev?.enabledSources ?? fresh.enabledSources,
    }
  })
}

export function getSignedInM365Accounts(settings: AppSettings): M365Account[] {
  return settings.m365Accounts ?? []
}

/** Accounts enabled for sweep — defaults to all signed-in */
export function getSweepAccountIds(settings: AppSettings): string[] {
  const signedIn = getSignedInM365Accounts(settings).map((a) => a.homeAccountId)
  if (signedIn.length === 0) return []

  const selected = settings.m365SweepAccountIds?.filter((id) => signedIn.includes(id))
  return selected?.length ? selected : signedIn
}

/** Sweep-enabled accounts that have a given Graph source turned on */
export function getSweepAccountIdsForSource(
  settings: AppSettings,
  source: M365ConnectorSource,
): string[] {
  return getSweepAccountIds(settings).filter((id) => {
    const account = findM365Account(settings, id)
    return isAccountSourceEnabled(account, source)
  })
}

export function getActiveM365AccountId(settings: AppSettings): string | undefined {
  const accounts = getSignedInM365Accounts(settings)
  if (accounts.length === 0) return undefined

  if (
    settings.m365ActiveAccountId &&
    accounts.some((a) => a.homeAccountId === settings.m365ActiveAccountId)
  ) {
    return settings.m365ActiveAccountId
  }

  return accounts[0]?.homeAccountId
}

export function findM365Account(
  settings: AppSettings,
  homeAccountId?: string,
): M365Account | undefined {
  if (!homeAccountId) return undefined
  return getSignedInM365Accounts(settings).find((a) => a.homeAccountId === homeAccountId)
}

/** Resolve which M365 account a task belongs to (for sync-back) */
export function resolveTaskM365AccountId(task: Task, settings: AppSettings): string | undefined {
  const fromMeta = task.metadata?.[M365_HOME_ACCOUNT_ID_KEY]
  if (fromMeta && findM365Account(settings, fromMeta)) return fromMeta
  return getActiveM365AccountId(settings)
}

/** Whether sync-back to Graph is allowed for this task given per-account source config */
export function isTaskSourceEnabledForAccount(
  task: Task,
  settings: AppSettings,
): boolean {
  const connector = connectorSourceForTask(task.source)
  if (!connector) return true

  const homeAccountId = resolveTaskM365AccountId(task, settings)
  const account = findM365Account(settings, homeAccountId)
  return isAccountSourceEnabled(account, connector)
}

export function stampM365Metadata(
  metadata: Record<string, string> | undefined,
  account: M365Account,
): Record<string, string> {
  return {
    ...metadata,
    [M365_HOME_ACCOUNT_ID_KEY]: account.homeAccountId,
    [M365_USERNAME_KEY]: account.username,
    ...(account.label ? { [M365_ACCOUNT_LABEL_KEY]: account.label } : {}),
  }
}

export function updateM365AccountInSettings(
  settings: AppSettings,
  homeAccountId: string,
  patch: Partial<Pick<M365Account, 'label' | 'enabledSources'>>,
): AppSettings {
  return {
    ...settings,
    m365Accounts: (settings.m365Accounts ?? []).map((a) =>
      a.homeAccountId === homeAccountId ? { ...a, ...patch } : a,
    ),
  }
}

export function syncM365AccountsToSettings(
  settings: AppSettings,
  msalAccounts: AccountInfo[],
): AppSettings {
  const merged = mergeM365Accounts(settings.m365Accounts ?? [], msalAccounts)
  const signedInIds = new Set(merged.map((a) => a.homeAccountId))

  const activeId =
    settings.m365ActiveAccountId && signedInIds.has(settings.m365ActiveAccountId)
      ? settings.m365ActiveAccountId
      : merged[0]?.homeAccountId

  const sweepIds = settings.m365SweepAccountIds?.filter((id) => signedInIds.has(id))
  const sweepAccountIds = sweepIds?.length
    ? sweepIds
    : merged.map((a) => a.homeAccountId)

  return {
    ...settings,
    m365Accounts: merged,
    m365ActiveAccountId: activeId,
    m365SweepAccountIds: sweepAccountIds,
  }
}

export function removeM365AccountFromSettings(
  settings: AppSettings,
  homeAccountId: string,
): AppSettings {
  const accounts = (settings.m365Accounts ?? []).filter((a) => a.homeAccountId !== homeAccountId)
  const sweepIds = (settings.m365SweepAccountIds ?? []).filter((id) => id !== homeAccountId)

  return {
    ...settings,
    m365Accounts: accounts,
    m365ActiveAccountId:
      settings.m365ActiveAccountId === homeAccountId
        ? accounts[0]?.homeAccountId
        : settings.m365ActiveAccountId,
    m365SweepAccountIds: sweepIds.length ? sweepIds : accounts.map((a) => a.homeAccountId),
  }
}