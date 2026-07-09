import { useEffect, useState } from 'react'
import {
  initM365,
  isM365SignedIn,
  signInM365,
  signOutM365,
} from '../services/connectors'
import type { AppSettings, PrimaryTaskTool } from '../types/task'
import { DeviceSetup } from './DeviceSetup'

interface SettingsPanelProps {
  settings: AppSettings
  onSave: (settings: AppSettings) => void
  onExport: () => void
}

export function SettingsPanel({ settings, onSave, onExport }: SettingsPanelProps) {
  const [draft, setDraft] = useState(settings)
  const [m365Status, setM365Status] = useState(isM365SignedIn() ? 'signed-in' : 'signed-out')

  useEffect(() => {
    setM365Status(isM365SignedIn() ? 'signed-in' : 'signed-out')
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

  const handleM365SignIn = async () => {
    onSave(draft)
    await initM365(draft)
    const result = await signInM365(draft)
    if (result || isM365SignedIn()) {
      setM365Status('signed-in')
    }
    // loginRedirect navigates away — page reloads when Microsoft sends you back
  }

  const handleM365SignOut = async () => {
    await signOutM365(draft)
    setM365Status('signed-out')
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
        Register at Azure Portal → App registrations. Redirect URI must match this app&apos;s
        URL exactly (e.g. <code>http://localhost:5173</code> for dev, or your Cloudflare /
        GitHub Pages HTTPS URL after deploy — see DEPLOY.md).
        Permissions: Tasks.ReadWrite, Mail.ReadWrite, Notes.Read (OneNote), User.Read.
        After adding Notes.Read, sign out and sign in again once.
        Sign-in opens Microsoft in this same window (no popup).
        After upgrading permissions, sign out and sign in again once.
      </p>

      <div className="m365-auth">
        <span>Status: {m365Status === 'signed-in' ? 'Signed in' : 'Not signed in'}</span>
        {m365Status === 'signed-in' ? (
          <button type="button" onClick={handleM365SignOut}>Sign out M365</button>
        ) : (
          <button type="button" onClick={handleM365SignIn} disabled={!draft.m365ClientId}>
            Sign in to Microsoft 365
          </button>
        )}
      </div>

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