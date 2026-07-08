import { useState } from 'react'
import { createBeaconText, scanPasteForBeacon } from '../services/beacon'
import type { AppSettings, BeaconHit } from '../types/task'

interface BeaconToolsProps {
  settings: AppSettings
  onBeaconFound?: (hit: BeaconHit) => void
}

export function BeaconTools({ settings, onBeaconFound }: BeaconToolsProps) {
  const [copied, setCopied] = useState(false)
  const [scanText, setScanText] = useState('')
  const [lastHit, setLastHit] = useState<BeaconHit | null>(null)
  const [scanned, setScanned] = useState(false)

  const beaconText = createBeaconText(settings)

  const copyBeacon = async () => {
    await navigator.clipboard.writeText(beaconText)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  const scanBeacon = () => {
    const hit = scanPasteForBeacon(scanText, settings)
    setLastHit(hit)
    setScanned(true)
    if (hit) onBeaconFound?.(hit)
  }

  return (
    <section className="panel beacon-tools">
      <h2>Beacon tools</h2>
      <p className="hint">
        Plant a beacon titled <code>{settings.beaconMarker}</code> in Outlook,
        To Do, OneNote, or any list. Run an M365 sweep or paste from that app here.
      </p>

      <div className="beacon-copy">
        <pre>{beaconText}</pre>
        <button type="button" onClick={copyBeacon}>
          {copied ? 'Copied!' : 'Copy beacon text'}
        </button>
      </div>

      <label className="field">
        Scan pasted text for beacon
        <textarea
          value={scanText}
          onChange={(e) => setScanText(e.target.value)}
          placeholder="Paste content from an app where you planted the beacon…"
          rows={3}
        />
      </label>
      <button type="button" className="secondary" onClick={scanBeacon}>
        Scan for beacon
      </button>

      {scanned && (
        <div className="beacon-result">
          {lastHit ? (
            <>
              <strong>Beacon found</strong> in {lastHit.source}
              <p>{lastHit.suggestedConnector}</p>
            </>
          ) : (
            <p>No beacon marker found in that text.</p>
          )}
        </div>
      )}
    </section>
  )
}