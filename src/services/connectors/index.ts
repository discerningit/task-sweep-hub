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

export function getAllConnectors(settings: AppSettings): Connector[] {
  return [
    pasteConnector,
    fileUploadConnector,
    createM365Connector(() => settings),
    // Future: protonMailConnector, iosExportConnector, jiraConnector
  ]
}

export { setPasteContent, clearPasteContent } from './paste'
export { setUploadFiles, clearUploadFiles } from './fileUpload'
export {
  bootstrapM365Auth,
  clearM365OutlookFlag,
  completeM365TodoTask,
  getM365AccessToken,
  initM365,
  M365_SIGNED_IN_FLAG,
  signInM365,
  signOutM365,
  isM365SignedIn,
  syncM365ClientId,
  sweepM365,
} from './m365'