import { useRef } from 'react'
import { useState } from 'react'
import {
  getDefaultSweepConnectorIds,
  setPasteContent,
  setProtonMailFiles,
  setUploadFiles,
} from '../services/connectors'
import type { AppSettings } from '../types/task'

interface InputAreaProps {
  settings: AppSettings
  onSweep: (connectorIds: string[]) => void
  onSweepOneNote?: () => void
  sweeping: boolean
  m365Ready?: boolean
}

export function InputArea({
  settings,
  onSweep,
  onSweepOneNote,
  sweeping,
  m365Ready,
}: InputAreaProps) {
  const [text, setText] = useState('')
  const protonInputRef = useRef<HTMLInputElement>(null)

  const protonEnabled = settings.protonMailEnabled !== false

  const handleSweepPaste = () => {
    if (!text.trim()) return
    setPasteContent(text)
    onSweep(['paste'])
    setText('')
  }

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      setUploadFiles(e.target.files)
      onSweep(['file'])
      e.target.value = ''
    }
  }

  const handleProtonFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      setProtonMailFiles(e.target.files)
      onSweep(['proton-mail'])
      e.target.value = ''
    }
  }

  return (
    <section className="panel input-area">
      <h2>Add tasks</h2>
      <p className="hint">
        Paste from email, Teams, Reminders, or any app. One task per line works best.
      </p>
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        placeholder={`Example:\n- Call Cedar Ridge contractor re: permit\n- Board meeting prep (nonprofit) — due 3/15\n- Review client ticket #4521`}
        rows={5}
      />
      <div className="input-actions">
        <button
          type="button"
          className="primary"
          onClick={handleSweepPaste}
          disabled={sweeping || !text.trim()}
        >
          {sweeping ? 'Sweeping…' : 'Sweep pasted text'}
        </button>
        <label className="button secondary">
          Upload file
          <input
            type="file"
            accept=".txt,.csv,.eml,.md,.json"
            multiple
            onChange={handleFileChange}
            hidden
          />
        </label>
        {protonEnabled && (
          <>
            <button
              type="button"
              className="secondary"
              onClick={() => protonInputRef.current?.click()}
              disabled={sweeping}
              title="Select .eml files exported from Proton Mail"
            >
              {sweeping ? 'Sweeping…' : 'Sweep Proton Mail'}
            </button>
            <input
              ref={protonInputRef}
              type="file"
              accept=".eml,message/rfc822"
              multiple
              onChange={handleProtonFileChange}
              hidden
            />
          </>
        )}
        <button
          type="button"
          className="secondary"
          onClick={() => onSweep(getDefaultSweepConnectorIds(settings))}
          disabled={sweeping}
        >
          Sweep all sources
        </button>
        {m365Ready && onSweepOneNote && (
          <button
            type="button"
            className="secondary"
            onClick={onSweepOneNote}
            disabled={sweeping}
            title="Pull tasks from recent OneNote pages"
          >
            {sweeping ? 'Sweeping…' : 'Sweep OneNote'}
          </button>
        )}
      </div>
    </section>
  )
}