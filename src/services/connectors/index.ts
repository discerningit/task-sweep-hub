/**
 * Connector registry — add new sources here.
 *
 * To add a connector:
 * 1. Create src/services/connectors/yourConnector.ts
 * 2. Implement the Connector interface
 * 3. Register it in ALL_CONNECTORS below
 */

import type { AppSettings, Connector } from '../../types/task'
import { pasteConnector } from './paste'
import { fileUploadConnector } from './fileUpload'
import { createM365Connector } from './m365'
import { createProtonMailConnector } from './protonMail'

export function getAllConnectors(settings: AppSettings): Connector[] {
  return [
    pasteConnector,
    fileUploadConnector,
    createProtonMailConnector(() => settings),
    createM365Connector(() => settings),
    // Future: iosExportConnector, jiraConnector
  ]
}

/** Connector IDs used by the Sweep all sources button */
export function getDefaultSweepConnectorIds(settings: AppSettings): string[] {
  const ids = ['paste', 'file', 'm365']
  if (settings.protonMailEnabled !== false) ids.push('proton-mail')
  return ids
}

export { setPasteContent, clearPasteContent } from './paste'
export { setUploadFiles, clearUploadFiles } from './fileUpload'
export { setProtonMailFiles, clearProtonMailFiles } from './protonMail'
export {
  bootstrapM365Auth,
  clearM365OutlookFlag,
  completeM365TodoTask,
  createM365TodoTask,
  getM365TodoTaskStatus,
  getM365AccessToken,
  initM365,
  M365_SIGNED_IN_FLAG,
  refreshM365AccountSettings,
  requestM365SourceAccess,
  signInM365,
  signOutM365,
  isM365SignedIn,
  syncM365ClientId,
  sweepM365,
  sweepM365TodoOnly,
  sweepM365OneNote,
} from './m365'