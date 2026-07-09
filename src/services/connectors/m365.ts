/**
 * Microsoft 365 connector (Graph API) — Outlook, To Do, OneNote, Teams.
 *
 * SETUP (one-time, in Azure Portal):
 * 1. Register an app at https://portal.azure.com → App registrations
 * 2. Add redirect URI: http://localhost:5173 and your production URL
 *    (GitHub Pages: https://<user>.github.io/task-sweep-hub/ — include trailing slash)
 * 3. API permissions (delegated): Tasks.ReadWrite, Mail.ReadWrite, Notes.Read, User.Read
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
import type { AppSettings, Connector, RawInput, Task, TaskPriority } from '../../types/task'
import { htmlToText } from './onenoteHtml'
import {
  dedupePagesById,
  encodeOneNoteResourceId,
  normalizeOneNotePage,
  normalizeOneNoteSection,
  prioritizeBeaconPages,
  shouldIncludeOneNotePage,
  sortPagesByModified,
  type OneNoteDiscoveryResult,
  type OneNotePageSummary,
  type OneNoteSectionSummary,
} from './onenotePages'

const SCOPES = [
  'User.Read',
  'Tasks.ReadWrite',
  'Mail.ReadWrite',
  'Notes.Read',
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

interface GraphErrorBody {
  error?: { code?: string; message?: string }
}

async function graphGet<T>(token: string, path: string): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) {
    let detail = ''
    try {
      const body = (await res.json()) as GraphErrorBody
      const parts = [body.error?.code, body.error?.message].filter(Boolean)
      detail = parts.join(': ')
    } catch {
      // ignore non-JSON error bodies
    }
    throw new Error(`Graph API ${res.status}: ${path}${detail ? ` — ${detail}` : ''}`)
  }
  return res.json() as Promise<T>
}

async function graphGetText(token: string, path: string): Promise<string> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${path}`)
  return res.text()
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

async function graphPost<T>(token: string, path: string, body: unknown): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  })
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${path}`)
  return res.json() as Promise<T>
}

function toGraphImportance(priority: TaskPriority): string {
  if (priority === 'urgent' || priority === 'high') return 'high'
  if (priority === 'low') return 'low'
  return 'normal'
}

/** Parse informal due dates into Graph dateTime format */
export function parseDueDateForGraph(
  due?: string,
): { dateTime: string; timeZone: string } | undefined {
  if (!due?.trim()) return undefined

  if (/^\d{4}-\d{2}-\d{2}/.test(due)) {
    return { dateTime: `${due.slice(0, 10)}T00:00:00`, timeZone: 'UTC' }
  }

  const match = due.match(/(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?/)
  if (!match) return undefined

  const year = match[3]
    ? match[3].length === 2
      ? `20${match[3]}`
      : match[3]
    : String(new Date().getFullYear())
  const month = match[1].padStart(2, '0')
  const day = match[2].padStart(2, '0')
  return { dateTime: `${year}-${month}-${day}T00:00:00`, timeZone: 'UTC' }
}

async function getDefaultTodoListId(token: string): Promise<string> {
  const lists = await graphGet<GraphTodoList>(token, '/me/todo/lists')
  const defaultList = lists.value?.find((l) => l.wellknownListName === 'defaultList')
  if (defaultList) return defaultList.id
  if (lists.value?.[0]) return lists.value[0].id
  throw new Error('No Microsoft To Do list found on your account')
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

interface GraphTodoListItem {
  id: string
  displayName: string
  wellknownListName?: string
}

interface GraphTodoList {
  value: GraphTodoListItem[]
}

interface GraphTodoTaskCreated {
  id: string
}

interface GraphTodoTaskDetail {
  id: string
  status?: string
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

interface GraphPagePreview {
  previewText?: string
}

export interface OneNoteSweepResult {
  inputs: RawInput[]
  pagesFound: number
  pagesImported: number
  sectionsScanned?: number
  detail?: string
  error?: string
}

let lastOneNoteSweepResult: OneNoteSweepResult = {
  inputs: [],
  pagesFound: 0,
  pagesImported: 0,
}

/** Stats from the most recent OneNote sweep (used by Sweep all sources summary) */
export function getLastOneNoteSweepResult(): OneNoteSweepResult {
  return lastOneNoteSweepResult
}

const ONENOTE_PAGE_LIMIT = 20
const ONENOTE_CONTENT_LIMIT = 8000
const ONENOTE_MAX_SECTIONS = 40

interface GraphListResponse {
  value?: Record<string, unknown>[]
}

function dedupeSections(sections: OneNoteSectionSummary[]): OneNoteSectionSummary[] {
  const seen = new Set<string>()
  return sections.filter((section) => {
    if (seen.has(section.id)) return false
    seen.add(section.id)
    return true
  })
}

function formatOneNoteError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  if (
    message.includes('403') ||
    message.includes('40003') ||
    message.includes('40004') ||
    message.includes('40007')
  ) {
    return 'OneNote permission missing — add Notes.Read in Azure API permissions, then sign out/in in Settings.'
  }
  if (message.includes('401') || message.includes('40001')) {
    return 'OneNote sign-in expired — open Settings and sign in to Microsoft 365 again.'
  }
  if (message.includes('20266')) {
    return 'OneNote has many sections — scanning section by section (this is normal).'
  }
  return message.length > 180 ? `${message.slice(0, 180)}…` : message
}

async function listOneNoteSections(token: string): Promise<OneNoteSectionSummary[]> {
  const sections: OneNoteSectionSummary[] = []

  try {
    const res = await graphGet<GraphListResponse>(token, '/me/onenote/sections')
    for (const raw of res.value ?? []) {
      const section = normalizeOneNoteSection(raw)
      if (section) sections.push(section)
    }
  } catch (e) {
    console.warn('OneNote /sections failed:', e)
  }

  if (sections.length > 0) return dedupeSections(sections)

  const notebooks = await graphGet<GraphListResponse>(token, '/me/onenote/notebooks')
  for (const notebook of notebooks.value ?? []) {
    const notebookId = typeof notebook.id === 'string' ? notebook.id : ''
    if (!notebookId) continue

    const nbSections = await graphGet<GraphListResponse>(
      token,
      `/me/onenote/notebooks/${encodeOneNoteResourceId(notebookId)}/sections`,
    )
    for (const raw of nbSections.value ?? []) {
      const section = normalizeOneNoteSection(raw)
      if (section) sections.push(section)
    }
  }

  return dedupeSections(sections)
}

async function listPagesForSection(
  token: string,
  section: OneNoteSectionSummary,
): Promise<OneNotePageSummary[]> {
  const res = await graphGet<GraphListResponse>(
    token,
    `/me/onenote/sections/${encodeOneNoteResourceId(section.id)}/pages`,
  )

  const pages: OneNotePageSummary[] = []
  for (const raw of res.value ?? []) {
    const page = normalizeOneNotePage(raw, section.displayName)
    if (page) pages.push(page)
  }
  return pages
}

async function discoverOneNotePages(
  token: string,
  beaconMarker?: string,
): Promise<OneNoteDiscoveryResult> {
  const collected: OneNotePageSummary[] = []
  let lastError: string | undefined
  let sectionErrors = 0
  let notebooksFound = 0
  let sectionsScanned = 0

  let sections: OneNoteSectionSummary[] = []
  try {
    sections = await listOneNoteSections(token)
  } catch (e) {
    lastError = formatOneNoteError(e)
    console.warn('OneNote section discovery failed:', e)
  }

  try {
    const notebooks = await graphGet<GraphListResponse>(token, '/me/onenote/notebooks')
    notebooksFound = notebooks.value?.length ?? 0
  } catch (e) {
    if (!lastError) lastError = formatOneNoteError(e)
    console.warn('OneNote /notebooks failed:', e)
  }

  for (const section of sections.slice(0, ONENOTE_MAX_SECTIONS)) {
    sectionsScanned++
    try {
      collected.push(...(await listPagesForSection(token, section)))
      if (collected.length >= ONENOTE_PAGE_LIMIT * 3) break
    } catch (e) {
      sectionErrors++
      if (!lastError) lastError = formatOneNoteError(e)
      console.warn(`OneNote pages failed for section ${section.id}:`, e)
    }
  }

  const pages = prioritizeBeaconPages(
    sortPagesByModified(dedupePagesById(collected)),
    beaconMarker,
  ).slice(0, ONENOTE_PAGE_LIMIT)

  let detail = `Scanned ${sectionsScanned} section(s), ${notebooksFound} notebook(s).`
  if (sectionErrors > 0) detail += ` ${sectionErrors} section(s) could not be read.`

  let error: string | undefined
  if (pages.length === 0) {
    if (notebooksFound === 0 && sections.length === 0) {
      error =
        'No OneNote notebooks found. Open the OneNote app, create a notebook and page, wait for sync, then retry.'
    } else if (sections.length === 0) {
      error = `Found ${notebooksFound} notebook(s) but no sections with pages. Add a page in OneNote, then retry.`
    } else if (sectionsScanned > 0 && sectionErrors === sectionsScanned) {
      error = lastError ?? 'Could not read pages from any section.'
    } else if (sectionsScanned > 0) {
      detail += ' Sections exist but no pages were returned yet.'
      error = lastError
    } else {
      error = lastError
    }
  }

  return { pages, sectionsScanned, error, detail }
}

async function readOneNotePageText(token: string, pageId: string): Promise<string> {
  const encodedId = encodeOneNoteResourceId(pageId)

  try {
    const html = await graphGetText(token, `/me/onenote/pages/${encodedId}/content`)
    const text = htmlToText(html)
    if (text.trim().length > 0) return text
  } catch (e) {
    console.warn(`OneNote page ${pageId} content fetch failed:`, e)
  }

  try {
    const preview = await graphGet<GraphPagePreview>(
      token,
      `/me/onenote/pages/${encodedId}/preview`,
    )
    return preview.previewText?.trim() ?? ''
  } catch (e) {
    console.warn(`OneNote page ${pageId} preview fetch failed:`, e)
    return ''
  }
}

/** Pull open Microsoft To Do tasks only (no Outlook) */
export async function sweepM365TodoOnly(settings: AppSettings): Promise<RawInput[]> {
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
    throw e
  }

  return inputs
}

/** Pull recent OneNote pages and extract text for task mining */
export async function sweepM365OneNote(
  settings: AppSettings,
  existingToken?: string,
): Promise<OneNoteSweepResult> {
  const token = existingToken ?? (await acquireToken(settings))
  if (!token) return { inputs: [], pagesFound: 0, pagesImported: 0 }

  const inputs: RawInput[] = []
  const now = new Date().toISOString()
  const beaconMarker = settings.beaconMarker?.trim()

  const discovery = await discoverOneNotePages(token, beaconMarker)
  const pages = discovery.pages

  for (const page of pages) {
    if (!page.id) continue

    try {
      const title = page.title?.trim() || 'OneNote page'
      const text = (await readOneNotePageText(token, page.id)).slice(0, ONENOTE_CONTENT_LIMIT)
      if (!shouldIncludeOneNotePage(title, text, beaconMarker)) continue

      inputs.push({
        id: crypto.randomUUID(),
        source: 'm365-onenote',
        content: `${title}\n${text}`.trim(),
        receivedAt: now,
        sourceUrl: page.links?.oneNoteWebUrl?.href,
        metadata: {
          id: page.id,
          subject: title,
          lastModified: page.lastModifiedDateTime ?? '',
          ...(page.sectionName ? { sectionName: page.sectionName } : {}),
        },
      })
    } catch (e) {
      console.warn(`OneNote page ${page.id} skipped:`, e)
    }
  }

  const result: OneNoteSweepResult = {
    inputs,
    pagesFound: pages.length,
    pagesImported: inputs.length,
    sectionsScanned: discovery.sectionsScanned,
    detail: discovery.detail,
    error: discovery.error,
  }
  lastOneNoteSweepResult = result
  return result
}

/** Pull To Do tasks and flagged emails via Graph */
export async function sweepM365(settings: AppSettings): Promise<RawInput[]> {
  const token = await acquireToken(settings)
  if (!token) return []

  const inputs: RawInput[] = []
  const now = new Date().toISOString()

  try {
    inputs.push(...(await sweepM365TodoOnly(settings)))
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

  try {
    const onenote = await sweepM365OneNote(settings, token)
    inputs.push(...onenote.inputs)
    if (onenote.pagesFound === 0) {
      console.warn('M365 OneNote sweep: no pages discovered. Check Notes.Read.All permission and sign-in.')
    }
  } catch (e) {
    console.warn('M365 OneNote sweep failed:', e)
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

/** Create a new task in Microsoft To Do (default list) */
export async function createM365TodoTask(
  settings: AppSettings,
  task: Pick<Task, 'title' | 'notes' | 'dueDate' | 'priority' | 'tags'>,
): Promise<{ id: string; listId: string }> {
  const token = await acquireToken(settings)
  if (!token) throw new Error('Not signed in to Microsoft 365')

  const listId = await getDefaultTodoListId(token)
  const body: Record<string, unknown> = {
    title: task.title,
    importance: toGraphImportance(task.priority),
  }

  const due = parseDueDateForGraph(task.dueDate)
  if (due) body.dueDateTime = due

  const noteParts = [task.notes, task.tags.length ? `Tags: ${task.tags.join(', ')}` : '']
    .filter(Boolean)
    .join('\n')
  if (noteParts) {
    body.body = { content: noteParts, contentType: 'text' }
  }

  const created = await graphPost<GraphTodoTaskCreated>(
    token,
    `/me/todo/lists/${listId}/tasks`,
    body,
  )
  return { id: created.id, listId }
}

/** Read current status of a To Do task (for inbound sync) */
export async function getM365TodoTaskStatus(
  settings: AppSettings,
  taskId: string,
  listId?: string,
): Promise<string | null> {
  const token = await acquireToken(settings)
  if (!token) return null

  const resolvedListId = await resolveTodoListId(token, taskId, listId)
  const task = await graphGet<GraphTodoTaskDetail>(
    token,
    `/me/todo/lists/${resolvedListId}/tasks/${taskId}`,
  )
  return task.status ?? null
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