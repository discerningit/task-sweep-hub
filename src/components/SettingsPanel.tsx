import { useEffect, useState } from 'react'
import {
  initM365,
  isM365SignedIn,
  refreshM365AccountSettings,
  signInM365,
  signOutM365,
} from '../services/connectors'
import type { AppSettings, M365Account, PrimaryTaskTool } from '../types/task'
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

  const save = () => {
    onSave({
      ...draft,
      setupCompleted: Boolean(draft.m365ClientId) || draft.setupCompleted,
    })
  }

  const handleImport = (imported: AppSettings) => {
    setDraft(imported)
    onSave(imported)
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
        After adding Notes.Read, sign out and sign in again once.
        Sign-in opens Microsoft in this same window (no popup).
        After upgrading permissions, sign out and sign in again once.
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
        <button type="button" className="primary" onClick={save}>Save settings</button>
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
  onToggleSweep: () => void
  onSignOut: () => void
}

function M365AccountRow({
  account,
  sweepEnabled,
  onLabelChange,
  onToggleSweep,
  onSignOut,
}: M365AccountRowProps) {
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