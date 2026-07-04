/**
 * Microsoft 365 connector (Graph API) — Outlook, To Do, OneNote, Teams.
 *
 * SETUP (one-time, in Azure Portal):
 * 1. Register an app at https://portal.azure.com → App registrations
 * 2. Add redirect URI: http://localhost:5173 (and your production URL)
 * 3. API permissions (delegated): Tasks.Read, Mail.Read, Notes.Read, User.Read
 * 4. Paste the Application (client) ID into Settings in TaskSweep Hub
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
  'Tasks.Read',
  'Mail.Read',
  'Notes.Read.All',
]

let msalInstance: PublicClientApplication | null = null
let cachedAccount: AccountInfo | null = null

function getMsal(clientId: string): PublicClientApplication {
  if (!msalInstance) {
    msalInstance = new PublicClientApplication({
      auth: {
        clientId,
        redirectUri: window.location.origin,
      },
      cache: {
        cacheLocation: 'localStorage',
      },
    })
  }
  return msalInstance
}

export async function initM365(settings: AppSettings): Promise<boolean> {
  if (!settings.m365ClientId) return false
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()
  const accounts = msal.getAllAccounts()
  cachedAccount = accounts[0] ?? null
  return true
}

export async function signInM365(settings: AppSettings): Promise<AuthenticationResult | null> {
  if (!settings.m365ClientId) {
    throw new Error('Add your M365 Client ID in Settings first')
  }
  const msal = getMsal(settings.m365ClientId)
  await msal.initialize()

  try {
    const result = await msal.acquireTokenPopup({ scopes: SCOPES })
    cachedAccount = result.account
    return result
  } catch (err) {
    console.error('M365 sign-in failed:', err)
    return null
  }
}

export function isM365SignedIn(): boolean {
  return cachedAccount !== null
}

export async function signOutM365(settings: AppSettings): Promise<void> {
  if (!settings.m365ClientId) return
  const msal = getMsal(settings.m365ClientId)
  const account = cachedAccount ?? msal.getAllAccounts()[0]
  if (account) await msal.logoutPopup({ account })
  cachedAccount = null
}

async function graphGet<T>(token: string, path: string): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${path}`)
  return res.json() as Promise<T>
}

async function acquireToken(settings: AppSettings): Promise<string | null> {
  if (!settings.m365ClientId || !cachedAccount) return null
  const msal = getMsal(settings.m365ClientId)
  try {
    const result = await msal.acquireTokenSilent({
      scopes: SCOPES,
      account: cachedAccount,
    })
    return result.accessToken
  } catch {
    const result = await msal.acquireTokenPopup({ scopes: SCOPES })
    cachedAccount = result.account
    return result.accessToken
  }
}

interface GraphTodoList {
  value: { id: string; displayName: string }[]
}

interface GraphTodoTasks {
  value: {
    id: string
    title: string
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
        if (task.title?.toLowerCase().includes('completed')) continue
        inputs.push({
          id: crypto.randomUUID(),
          source: 'm365-todo',
          content: [task.title, task.body?.content ?? ''].filter(Boolean).join('\n'),
          sourceUrl: `https://to-do.office.com/tasks/id/${task.id}`,
          receivedAt: now,
          metadata: {
            id: task.id,
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
        sourceUrl: msg.webLink,
        receivedAt: now,
        metadata: { id: msg.id, subject: msg.subject },
      })
    }
  } catch (e) {
    console.warn('M365 Outlook sweep failed:', e)
  }

  return inputs
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