/**
 * Microsoft 365 connector (Graph API) — Outlook, To Do, OneNote, Teams.
 *
 * SETUP (one-time, in Azure Portal):
 * 1. Register an app at https://portal.azure.com → App registrations
 * 2. Add redirect URI: http://localhost:5173 and your production URL
 *    (GitHub Pages: https://<user>.github.io/task-sweep-hub/ — include trailing slash)
 * 3. API permissions (delegated): Tasks.ReadWrite, Mail.ReadWrite, Notes.Read.All, User.Read
 * 4. Paste the Application (client) ID into Settings in TaskSweep Hub
 *
 * Sign-in uses redirect flow (same window) — more reliable than popups on GitHub Pages.
 *
 * In enterprise/VDI: if Graph is blocked, use Paste or File Upload instead.
 */

import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from '@azure/msal-browser'
import type { AppSettings, Connector, RawInput } from '../../types/task'

const SCOPES = [
  'User.Read',
  'Tasks.ReadWrite',
  'Mail.ReadWrite',
  'Notes.Read.All',
]

/** Mirrored from Settings for auth bootstrap on page load */
export const M365_CLIENT_ID_KEY = 'tasksweep_m365_client_id'
export const M365_SIGNED_IN_FLAG = 'tasksweep_m365_signed_in'

let msalInstance: PublicClientApplication | null = null
let msalClientId: string | null = null
let cachedAccount: AccountInfo | null = null

/**
 * MSAL redirect URI must match Azure exactly.
 * GitHub Pages serves from /task-sweep-hub/, not the domain root.
 */
function getRedirectUri(): string {
  const base = import.meta.env.BASE_URL ?? '/'
  if (base === '/') {
    return window.location.origin
  }
  return `${window.location.origin}${base.endsWith('/') ? base : `${base}/`}`
}

function getMsal(clientId: string): PublicClientApplication {
  if (!msalInstance || msalClientId !== clientId) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId,
        redirectUri: getRedirectUri(),
      },
      cache: {
        cacheLocation: 'localStorage',
      },
    })
    msalClientId = clientId
  }
  return msalInstance
}

function refreshCachedAccount(msal: PublicClientApplication): AccountInfo | null {
  const accounts = msal.getAllAccounts()
  cachedAccount = accounts[0] ?? null
  return cachedAccount
}

/** Keep client ID in localStorage so auth works immediately on page load */
export function syncM365ClientId(clientId?: string): void {
  if (clientId) {
    localStorage.setItem(M365_CLIENT_ID_KEY, clientId)
  } else {
    localStorage.removeItem(M365_CLIENT_ID_KEY)
  }
}

/**
 * Run before React mounts. Completes sign-in when Microsoft redirects back
 * to the main app window (not a popup).
 */
export async function bootstrapM365Auth(): Promise<boolean> {
  const clientId = localStorage.getItem(M365_CLIENT_ID_KEY)
  if (!clientId) return false

  const msal = getMsal(clientId)
  await msal.initialize()
  const result = await msal.handleRedirectPromise()
  if (result?.account) {
    cachedAccount = result.account
    sessionStorage.setItem(M365_SIGNED_IN_FLAG, '1')
    return true
  }

  refreshCachedAccount(msal)
  return false
}

export async function initM365(settings: AppSettings): Promise<boolean> {
  if (!settings.m365ClientId) return false
  syncM365ClientId(settings.m365ClientId)
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()
  await msal.handleRedirectPromise()
  refreshCachedAccount(msal)
  return cachedAccount !== null
}

/**
 * Sign in via redirect — the page navigates to Microsoft and back.
 * No popup window.
 */
export async function signInM365(settings: AppSettings): Promise<AuthenticationResult | null> {
  if (!settings.m365ClientId) {
    throw new Error('Add your M365 Client ID in Settings first')
  }
  syncM365ClientId(settings.m365ClientId)
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()

  const redirectResult = await msal.handleRedirectPromise()
  if (redirectResult?.account) {
    cachedAccount = redirectResult.account
    return redirectResult
  }

  refreshCachedAccount(msal)
  if (cachedAccount) return null

  await msal.loginRedirect({ scopes: SCOPES })
  return null
}

export function isM365SignedIn(): boolean {
  if (cachedAccount) return true
  const clientId = localStorage.getItem(M365_CLIENT_ID_KEY)
  if (!clientId || !msalInstance) return false
  return msalInstance.getAllAccounts().length > 0
}

export async function signOutM365(settings: AppSettings): Promise<void> {
  if (!settings.m365ClientId) return
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()
  const account = cachedAccount ?? msal.getAllAccounts()[0]
  cachedAccount = null
  if (account) await msal.logoutRedirect({ account })
}

async function graphGet<T>(token: string, path: string): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${path}`)
  return res.json() as Promise<T>
}

async function graphPatch(token: string, path: string, body: unknown): Promise<void> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${path}`)
}

/** Get a Graph access token (exported for sync-back) */
export async function getM365AccessToken(settings: AppSettings): Promise<string | null> {
  return acquireToken(settings)
}

async function acquireToken(settings: AppSettings): Promise<string | null> {
  if (!settings.m365ClientId) return null
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()
  const account = cachedAccount ?? refreshCachedAccount(msal)
  if (!account) return null

  try {
    const result = await msal.acquireTokenSilent({
      scopes: SCOPES,
      account,
    })
    return result.accessToken
  } catch {
    await msal.acquireTokenRedirect({ scopes: SCOPES, account })
    return null
  }
}

interface GraphTodoList {
  value: { id: string; displayName: string }[]
}

interface GraphTodoTasks {
  value: {
    id: string
    title: string
    status?: string
    body?: { content?: string }
    dueDateTime?: { dateTime: string }
    importance?: string
  }[]
}

interface GraphMessages {
  value: {
    id: string
    subject: string
    bodyPreview: string
    webLink?: string
  }[]
}

/** Pull To Do tasks and flagged emails via Graph */
export async function sweepM365(settings: AppSettings): Promise<RawInput[]> {
  const token = await acquireToken(settings)
  if (!token) return []

  const inputs: RawInput[] = []
  const now = new Date().toISOString()

  try {
    const lists = await graphGet<GraphTodoList>(token, '/me/todo/lists')
    for (const list of lists.value ?? []) {
      const tasks = await graphGet<GraphTodoTasks>(
        token,
        `/me/todo/lists/${list.id}/tasks?$top=50`,
      )
      for (const task of tasks.value ?? []) {
        if (task.status === 'completed') continue
        inputs.push({
          id: crypto.randomUUID(),
          source: 'm365-todo',
          content: [task.title, task.body?.content ?? ''].filter(Boolean).join('\n'),
          receivedAt: now,
          sourceUrl: `https://to-do.office.com/tasks/id/${task.id}`,
          metadata: {
            id: task.id,
            listId: list.id,
            listName: list.displayName,
            dueDate: task.dueDateTime?.dateTime ?? '',
            importance: task.importance ?? 'normal',
          },
        })
      }
    }
  } catch (e) {
    console.warn('M365 To Do sweep failed:', e)
  }

  try {
    const mail = await graphGet<GraphMessages>(
      token,
      '/me/messages?$top=30&$filter=flag/flagStatus eq \'flagged\'',
    )
    for (const msg of mail.value ?? []) {
      inputs.push({
        id: crypto.randomUUID(),
        source: 'm365-outlook',
        content: `${msg.subject}\n${msg.bodyPreview}`,
        receivedAt: now,
        sourceUrl: msg.webLink,
        metadata: { id: msg.id, subject: msg.subject },
      })
    }
  } catch (e) {
    console.warn('M365 Outlook sweep failed:', e)
  }

  return inputs
}

/** Resolve To Do list ID (stored on sweep, or search lists for older tasks) */
async function resolveTodoListId(
  token: string,
  taskId: string,
  listId?: string,
): Promise<string> {
  if (listId) return listId

  const lists = await graphGet<GraphTodoList>(token, '/me/todo/lists')
  for (const list of lists.value ?? []) {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks/${taskId}`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (res.ok) return list.id
  }
  throw new Error('Could not find To Do list for this task. Re-sweep from M365.')
}

/** Mark a Microsoft To Do task completed in Graph */
export async function completeM365TodoTask(
  settings: AppSettings,
  taskId: string,
  listId?: string,
): Promise<void> {
  const token = await acquireToken(settings)
  if (!token) throw new Error('Not signed in to Microsoft 365')

  const resolvedListId = await resolveTodoListId(token, taskId, listId)
  await graphPatch(token, `/me/todo/lists/${resolvedListId}/tasks/${taskId}`, {
    status: 'completed',
  })
}

/** Clear Outlook follow-up flag on a message */
export async function clearM365OutlookFlag(
  settings: AppSettings,
  messageId: string,
): Promise<void> {
  const token = await acquireToken(settings)
  if (!token) throw new Error('Not signed in to Microsoft 365')

  await graphPatch(token, `/me/messages/${messageId}`, {
    flag: { flagStatus: 'notFlagged' },
  })
}

export function createM365Connector(getSettings: () => AppSettings): Connector {
  return {
    id: 'm365',
    name: 'Microsoft 365',
    description: 'Outlook, To Do, OneNote via Graph API',
    requiresAuth: true,
    isAvailable: () => {
      const s = getSettings()
      return Boolean(s.m365ClientId) && isM365SignedIn()
    },
    async sweep() {
      return sweepM365(getSettings())
    },
  }
}