/**
 * File upload connector — drop .txt, .csv, .eml exports, or
 * iOS Reminders / Shortcuts output files.
 */

import type { Connector, RawInput } from '../../types/task'

let pendingFiles: File[] = []

export function setUploadFiles(files: FileList | File[]): void {
  pendingFiles = Array.from(files)
}

export function clearUploadFiles(): void {
  pendingFiles = []
}

async function readFileAsText(file: File): Promise<string> {
  return file.text()
}

export const fileUploadConnector: Connector = {
  id: 'file',
  name: 'File upload',
  description: 'Upload text exports, CSV, or email .eml files',
  requiresAuth: false,
  isAvailable: () => true,
  async sweep() {
    if (pendingFiles.length === 0) return []

    const inputs: RawInput[] = []

    for (const file of pendingFiles) {
      const content = await readFileAsText(file)
      inputs.push({
        id: crypto.randomUUID(),
        source: 'file',
        content,
        receivedAt: new Date().toISOString(),
        metadata: {
          filename: file.name,
          type: file.type || 'unknown',
        },
      })
    }

    clearUploadFiles()
    return inputs
  },
}