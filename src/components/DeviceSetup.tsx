import { useState } from 'react'
import type { AppSettings } from '../types/task'
import {
  applySettingsPack,
  downloadSettingsPack,
  encodeSettingsPack,
  parseSettingsPack,
} from '../services/settingsPack'

interface DeviceSetupProps {
  settings: AppSettings
  onImport: (settings: AppSettings) => void
}

export function DeviceSetup({ settings, onImport }: DeviceSetupProps) {
  const [importText, setImportText] = useState('')
  const [copied, setCopied] = useState(false)
  const [importError, setImportError] = useState<string | null>(null)
  const [importOk, setImportOk] = useState(false)

  const hasExportableConfig = Boolean(settings.m365ClientId)

  const copyPack = async () => {
    await navigator.clipboard.writeText(encodeSettingsPack(settings))
    setCopied(true)
    setTimeout(() => setCopied(false), 2500)
  }

  const handleImport = () => {
    setImportError(null)
    setImportOk(false)
    try {
      const pack = parseSettingsPack(importText)
      onImport(applySettingsPack(settings, pack))
      setImportText('')
      setImportOk(true)
    } catch (err) {
      setImportError(err instanceof Error ? err.message : 'Import failed')
    }
  }

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setImportText(await file.text())
    e.target.value = ''
  }

  return (
    <section className="panel device-setup">
      <h2>New device setup</h2>
      <p className="hint">
        Copy settings from a device you already configured — no need to re-type the M365 Client ID.
        Sign-in to Microsoft is still required once on each device (for security).
      </p>

      <div className="setup-columns">
        <div className="setup-block">
          <h3>On your configured device</h3>
          <p className="hint">Export and send to your new phone (AirDrop, email, Notes, etc.)</p>
          <div className="input-actions">
            <button type="button" className="primary" onClick={copyPack} disabled={!hasExportableConfig}>
              {copied ? 'Copied!' : 'Copy setup pack'}
            </button>
            <button
              type="button"
              className="secondary"
              onClick={() => downloadSettingsPack(settings)}
              disabled={!hasExportableConfig}
            >
              Download setup file
            </button>
          </div>
          {!hasExportableConfig && (
            <p className="hint">Save your M365 Client ID here first, then export for other devices.</p>
          )}
        </div>

        <div className="setup-block">
          <h3>On this new device</h3>
          <p className="hint">Paste the setup pack or upload the file you exported</p>
          <textarea
            value={importText}
            onChange={(e) => setImportText(e.target.value)}
            placeholder='Paste JSON starting with { "app": "tasksweep-hub" … }'
            rows={5}
          />
          <div className="input-actions">
            <label className="button secondary">
              Upload setup file
              <input type="file" accept=".json,application/json" onChange={handleFile} hidden />
            </label>
            <button type="button" className="primary" onClick={handleImport} disabled={!importText.trim()}>
              Import setup
            </button>
          </div>
          {importError && <p className="setup-error">{importError}</p>}
          {importOk && (
            <p className="setup-ok">
              Imported. Scroll down to <strong>Sign in to Microsoft 365</strong>, then sweep tasks.
            </p>
          )}
        </div>
      </div>
    </section>
  )
}