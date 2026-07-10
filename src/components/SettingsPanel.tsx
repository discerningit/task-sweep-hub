import { useEffect, useState } from 'react'
import {
  ensureM365ExtraConsents,
  initM365,
  isM365SignedIn,
  refreshM365AccountSettings,
  signInM365,
  signOutM365,
} from '../services/connectors'
import {
  ALL_M365_CONNECTOR_SOURCES,
  getAccountEnabledSources,
  M365_CONNECTOR_SOURCE_LABELS,
} from '../services/m365Accounts'
import type { AppSettings, M365Account, M365ConnectorSource, PrimaryTaskTool } from '../types/task'
import { DeviceSetup } from './DeviceSetup'

interface SettingsPanelProps {
  settings: AppSettings
  onSave: (settings: AppSettings) => void
  onExport: () => void
}

export function SettingsPanel({ settings, onSave, onExport }: SettingsPanelProps) {
  const [draft, setDraft] = useState(settings)
  const [m365SignedIn, setM365SignedIn] = useState(isM365SignedIn())

  useEffect(() => {
    setDraft(settings)
    setM365SignedIn(isM365SignedIn())
  }, [settings])

  const save = async () => {
    const next = {
      ...draft,
      setupCompleted: Boolean(draft.m365ClientId) || draft.setupCompleted,
    }
    onSave(next)
    await ensureM365ExtraConsents(next)
  }

  const handleImport = async (imported: AppSettings) => {
    setDraft(imported)
    onSave(imported)
    await ensureM365ExtraConsents(imported)
  }

  const syncAccountsFromMsal = async (base: AppSettings): Promise<AppSettings> => {
    if (!base.m365ClientId || !isM365SignedIn()) return base
    return refreshM365AccountSettings(base)
  }

  const handleM365SignIn = async (addAccount = false) => {
    await initM365(draft)
    const saved = await syncAccountsFromMsal(draft)
    setDraft(saved)
    onSave(saved)
    const result = await signInM365(saved, { addAccount })
    if (result || isM365SignedIn()) {
      const refreshed = await syncAccountsFromMsal(saved)
      setDraft(refreshed)
      onSave(refreshed)
      setM365SignedIn(true)
      await ensureM365ExtraConsents(refreshed)
    }
  }

  const handleM365SignOut = async (homeAccountId?: string) => {
    const next = await signOutM365(draft, homeAccountId)
    setDraft(next)
    onSave(next)
    setM365SignedIn(isM365SignedIn())
  }

  const accounts = draft.m365Accounts ?? []

  const toggleSweepAccount = (homeAccountId: string) => {
    const signedIn = accounts.map((a) => a.homeAccountId)
    const current = draft.m365SweepAccountIds ?? signedIn
    const enabled = current.includes(homeAccountId)
    const next = enabled
      ? current.filter((id) => id !== homeAccountId)
      : [...current, homeAccountId]
    setDraft({
      ...draft,
      m365SweepAccountIds: next.length ? next : signedIn,
    })
  }

  const updateAccountLabel = (homeAccountId: string, label: string) => {
    setDraft({
      ...draft,
      m365Accounts: accounts.map((a) =>
        a.homeAccountId === homeAccountId ? { ...a, label: label || undefined } : a,
      ),
    })
  }

  const toggleAccountSource = (homeAccountId: string, source: M365ConnectorSource) => {
    setDraft({
      ...draft,
      m365Accounts: accounts.map((a) => {
        if (a.homeAccountId !== homeAccountId) return a
        const current = getAccountEnabledSources(a)
        const enabled = current.includes(source)
        const nextSources = enabled
          ? current.filter((s) => s !== source)
          : [...current, source]
        return {
          ...a,
          enabledSources: nextSources.length ? nextSources : ['todo'],
        }
      }),
    })
  }

  return (
    <>
      <DeviceSetup settings={draft} onImport={handleImport} />

      <section className="panel settings-panel">
      <h2>Settings</h2>

      <label className="field">
        Beacon marker text
        <input
          value={draft.beaconMarker}
          onChange={(e) => setDraft({ ...draft, beaconMarker: e.target.value })}
        />
      </label>
      <p className="hint">
        Put this text in a OneNote page title to prioritize that page in OneNote sweeps.
        Empty pages with the marker in the title are still imported.
      </p>

      <label className="field">
        M365 Application (Client) ID
        <input
          value={draft.m365ClientId ?? ''}
          onChange={(e) => setDraft({ ...draft, m365ClientId: e.target.value || undefined })}
          placeholder="From Azure App Registration"
        />
      </label>
      <p className="hint">
        Register at Azure Portal → App registrations. Choose <strong>Accounts in any
        organizational directory and personal Microsoft accounts</strong> so one app works
        for both personal and work sign-in.
        Redirect URI must match this app&apos;s URL exactly (e.g. <code>http://localhost:5173</code> for dev, or your Cloudflare /
        GitHub Pages HTTPS URL after deploy — see DEPLOY.md).
        Permissions: Tasks.ReadWrite, Mail.ReadWrite, Notes.Read (OneNote), User.Read.
        Sign-in starts with To Do only; enable Outlook or OneNote per account below, then Save
        to grant additional access when prompted.
        Work accounts often allow <strong>To Do only</strong> — disable mail and notes for those.
      </p>

      <div className="m365-accounts">
        <h3>Microsoft 365 accounts</h3>
        {!m365SignedIn && (
          <p className="hint">Not signed in. Add your Client ID above, then sign in.</p>
        )}
        {accounts.length > 0 && (
          <ul className="m365-account-list">
            {accounts.map((account) => (
              <M365AccountRow
                key={account.homeAccountId}
                account={account}
                sweepEnabled={(draft.m365SweepAccountIds ?? accounts.map((a) => a.homeAccountId)).includes(
                  account.homeAccountId,
                )}
                onLabelChange={(label) => updateAccountLabel(account.homeAccountId, label)}
                onToggleSource={(source) => toggleAccountSource(account.homeAccountId, source)}
                onToggleSweep={() => toggleSweepAccount(account.homeAccountId)}
                onSignOut={() => void handleM365SignOut(account.homeAccountId)}
              />
            ))}
          </ul>
        )}
        <div className="m365-auth">
          {!draft.m365ClientId ? (
            <span>Add Client ID to enable sign-in</span>
          ) : accounts.length === 0 ? (
            <button type="button" onClick={() => void handleM365SignIn(false)}>
              Sign in to Microsoft 365
            </button>
          ) : (
            <button type="button" onClick={() => void handleM365SignIn(true)}>
              Add another account
            </button>
          )}
          {accounts.length > 1 && (
            <button type="button" className="danger" onClick={() => void handleM365SignOut()}>
              Sign out all
            </button>
          )}
        </div>
      </div>

      <div className="proton-settings">
        <h3>Proton Mail</h3>
        <label className="field inline-check">
          <input
            type="checkbox"
            checked={draft.protonMailEnabled !== false}
            onChange={(e) => setDraft({ ...draft, protonMailEnabled: e.target.checked })}
          />
          Enable Proton Mail connector
        </label>
        <p className="hint">
          Proton does not offer a browser API like Microsoft Graph. Use <strong>Sweep Proton Mail</strong>{' '}
          on the Tasks tab with <code>.eml</code> files exported from Proton Mail.
          In Proton: open a message → <strong>More</strong> → <strong>Export</strong> → save as{' '}
          <code>.eml</code>, or bulk-export from Settings → Import/Export.
          Tasks are tagged <code>proton-mail</code> and stay in TaskSweep (no live sync-back).
        </p>
      </div>

      {accounts.length > 0 && draft.primaryTaskTool === 'ms-todo' && (
        <label className="field">
          Push new tasks to To Do account
          <select
            value={draft.m365ActiveAccountId ?? accounts[0]?.homeAccountId ?? ''}
            onChange={(e) => setDraft({ ...draft, m365ActiveAccountId: e.target.value })}
          >
            {accounts.map((a) => (
              <option key={a.homeAccountId} value={a.homeAccountId}>
                {a.label ?? a.username}
              </option>
            ))}
          </select>
        </label>
      )}

      <label className="field">
        Primary task tool (push new tasks, sync done both ways)
        <select
          value={draft.primaryTaskTool}
          onChange={(e) =>
            setDraft({ ...draft, primaryTaskTool: e.target.value as PrimaryTaskTool })
          }
        >
          <option value="hub-only">TaskSweep Hub only</option>
          <option value="ms-todo">Microsoft To Do</option>
          <option value="todoist">Todoist</option>
          <option value="apple-reminders">Apple Reminders</option>
          <option value="jira">Jira</option>
          <option value="export-csv">Export CSV</option>
        </select>
      </label>

      <label className="field">
        Grok API key (xAI)
        <input
          type="password"
          value={draft.grokApiKey ?? ''}
          onChange={(e) => setDraft({ ...draft, grokApiKey: e.target.value || undefined })}
          placeholder="xai-… from console.x.ai"
          autoComplete="off"
        />
      </label>
      <p className="hint">
        Get a key at <a href="https://console.x.ai" target="_blank" rel="noreferrer">console.x.ai</a>.
        Stored on this device only. Enable <strong>Grok</strong> under AI extraction and set it as Primary AI.
        On the live site, Grok calls the xAI API directly; if blocked, use local extraction or run{' '}
        <code>npm run dev</code> locally (includes a dev proxy).
      </p>

      <label className="field">
        Context tags (comma-separated, for AI)
        <input
          value={draft.contextTags.join(', ')}
          onChange={(e) =>
            setDraft({
              ...draft,
              contextTags: e.target.value.split(',').map((t) => t.trim()).filter(Boolean),
            })
          }
        />
      </label>

      <div className="settings-actions">
        <button type="button" className="primary" onClick={() => void save()}>Save settings</button>
        <button type="button" className="secondary" onClick={onExport}>Export tasks CSV</button>
      </div>
    </section>
    </>
  )
}

interface M365AccountRowProps {
  account: M365Account
  sweepEnabled: boolean
  onLabelChange: (label: string) => void
  onToggleSource: (source: M365ConnectorSource) => void
  onToggleSweep: () => void
  onSignOut: () => void
}

function M365AccountRow({
  account,
  sweepEnabled,
  onLabelChange,
  onToggleSource,
  onToggleSweep,
  onSignOut,
}: M365AccountRowProps) {
  const enabledSources = getAccountEnabledSources(account)

  return (
    <li className="m365-account-row">
      <div className="m365-account-info">
        <input
          className="m365-account-label"
          value={account.label ?? ''}
          placeholder={account.username}
          onChange={(e) => onLabelChange(e.target.value)}
          aria-label={`Label for ${account.username}`}
        />
        <span className="m365-account-email">{account.username}</span>
        <div className="m365-source-toggles">
          {ALL_M365_CONNECTOR_SOURCES.map((source) => (
            <label key={source} className="m365-source-toggle">
              <input
                type="checkbox"
                checked={enabledSources.includes(source)}
                onChange={() => onToggleSource(source)}
              />
              {M365_CONNECTOR_SOURCE_LABELS[source]}
            </label>
          ))}
        </div>
      </div>
      <label className="m365-sweep-toggle">
        <input type="checkbox" checked={sweepEnabled} onChange={onToggleSweep} />
        Sweep
      </label>
      <button type="button" className="danger" onClick={onSignOut}>
        Sign out
      </button>
    </li>
  )
}