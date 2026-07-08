import { useState } from 'react'
import { setPasteContent, setUploadFiles } from '../services/connectors'

interface InputAreaProps {
  onSweep: (connectorIds: string[]) => void
  onSweepOneNote?: () => void
  sweeping: boolean
  m365Ready?: boolean
}

export function InputArea({ onSweep, onSweepOneNote, sweeping, m365Ready }: InputAreaProps) {
  const [text, setText] = useState('')

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
        <button
          type="button"
          className="secondary"
          onClick={() => onSweep(['paste', 'file', 'm365'])}
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