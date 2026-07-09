/**
 * Proton Mail connector — import exported .eml messages.
 *
 * Proton has no browser OAuth API for third-party web apps. This connector
 * works with .eml files you export from Proton Mail (single message or bulk).
 * Live inbox sync would require Proton Mail Bridge + a backend proxy (not
 * available in the browser-only GitHub Pages deployment).
 */

import type { AppSettings, Connector, RawInput } from '../../types/task'
import { formatEmlForExtraction, parseEml } from './emlParser'

let pendingEmlFiles: File[] = []

export function setProtonMailFiles(files: FileList | File[]): void {
  pendingEmlFiles = Array.from(files).filter(
    (file) =>
      file.name.toLowerCase().endsWith('.eml') ||
      file.type === 'message/rfc822' ||
      file.type === 'application/octet-stream',
  )
}

export function clearProtonMailFiles(): void {
  pendingEmlFiles = []
}

export function createProtonMailConnector(getSettings: () => AppSettings): Connector {
  return {
    id: 'proton-mail',
    name: 'Proton Mail',
    description: 'Import exported .eml messages from Proton Mail',
    requiresAuth: false,
    isAvailable: () => getSettings().protonMailEnabled !== false,
    async sweep() {
      if (pendingEmlFiles.length === 0) return []

      const inputs: RawInput[] = []
      const now = new Date().toISOString()

      for (const file of pendingEmlFiles) {
        try {
          const raw = await file.text()
          const parsed = parseEml(raw)
          const content = formatEmlForExtraction(parsed)
          if (!content.trim()) continue

          inputs.push({
            id: crypto.randomUUID(),
            source: 'proton-mail',
            content,
            receivedAt: now,
            metadata: {
              filename: file.name,
              subject: parsed.subject,
              from: parsed.from,
              ...(parsed.messageId ? { messageId: parsed.messageId } : {}),
              ...(parsed.date ? { date: parsed.date } : {}),
            },
          })
        } catch (err) {
          console.warn(`Proton Mail: skipped ${file.name}:`, err)
        }
      }

      clearProtonMailFiles()
      return inputs
    },
  }
}