/**
 * Paste connector — the simplest way to sweep tasks.
 * Copy text from email, Teams, Reminders, or anywhere and paste here.
 */

import type { Connector, RawInput } from '../../types/task'

let pendingPaste: string | null = null

/** Called from UI when user pastes or types text */
export function setPasteContent(text: string): void {
  pendingPaste = text
}

export function clearPasteContent(): void {
  pendingPaste = null
}

export const pasteConnector: Connector = {
  id: 'paste',
  name: 'Paste text',
  description: 'Copy tasks from any app and paste into TaskSweep',
  requiresAuth: false,
  isAvailable: () => true,
  async sweep() {
    if (!pendingPaste?.trim()) return []

    const input: RawInput = {
      id: crypto.randomUUID(),
      source: 'paste',
      content: pendingPaste.trim(),
      receivedAt: new Date().toISOString(),
      metadata: { method: 'clipboard' },
    }

    clearPasteContent()
    return [input]
  },
}