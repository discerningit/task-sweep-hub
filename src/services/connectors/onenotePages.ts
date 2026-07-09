/**
 * OneNote page discovery helpers for Microsoft Graph.
 *
 * Graph best practice: list pages per section instead of the flat /pages
 * endpoint, which can fail when an account has many sections.
 */

export interface OneNotePageSummary {
  id: string
  title?: string
  links?: { oneNoteWebUrl?: { href?: string } }
  lastModifiedDateTime?: string
  sectionName?: string
}

export interface OneNoteSectionSummary {
  id: string
  displayName?: string
}

export interface OneNoteDiscoveryResult {
  pages: OneNotePageSummary[]
  sectionsScanned: number
  error?: string
  detail?: string
}

/** Escape a literal for OData single-quoted strings */
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''")
}

/**
 * Build a Graph OData query string.
 * URLSearchParams encodes "$" as "%24", which Graph rejects — build manually.
 */
export function buildGraphQuery(
  basePath: string,
  params: Record<string, string | number | undefined>,
): string {
  const parts: string[] = []
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined) continue
    parts.push(`${key}=${encodeURIComponent(String(value))}`)
  }
  const qs = parts.join('&')
  return qs ? `${basePath}?${qs}` : basePath
}

/** OneNote section/page ids contain "!" and must be encoded in URL paths */
export function encodeOneNoteResourceId(id: string): string {
  return encodeURIComponent(id).replace(/!/g, '%21')
}

export function dedupePagesById(pages: OneNotePageSummary[]): OneNotePageSummary[] {
  const seen = new Set<string>()
  const result: OneNotePageSummary[] = []
  for (const page of pages) {
    if (!page.id || seen.has(page.id)) continue
    seen.add(page.id)
    result.push(page)
  }
  return result
}

export function sortPagesByModified(pages: OneNotePageSummary[]): OneNotePageSummary[] {
  return [...pages].sort((a, b) =>
    (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? ''),
  )
}

/** Graph returns mixed property names across OneNote OData vs JSON shapes */
export function normalizeOneNoteSection(raw: Record<string, unknown>): OneNoteSectionSummary | null {
  const id = typeof raw.id === 'string' ? raw.id : ''
  if (!id) return null
  const name =
    typeof raw.displayName === 'string'
      ? raw.displayName
      : typeof raw.name === 'string'
        ? raw.name
        : undefined
  return { id, displayName: name }
}

export function normalizeOneNotePage(
  raw: Record<string, unknown>,
  sectionName?: string,
): OneNotePageSummary | null {
  const id = typeof raw.id === 'string' ? raw.id : ''
  if (!id) return null
  const modified = raw.lastModifiedDateTime ?? raw.lastModifiedTime
  return {
    id,
    title: typeof raw.title === 'string' ? raw.title : undefined,
    lastModifiedDateTime: typeof modified === 'string' ? modified : undefined,
    links: raw.links as OneNotePageSummary['links'],
    sectionName,
  }
}

export function prioritizeBeaconPages(
  pages: OneNotePageSummary[],
  beaconMarker?: string,
): OneNotePageSummary[] {
  if (!beaconMarker?.trim()) return pages
  const marker = beaconMarker.trim()
  const matches: OneNotePageSummary[] = []
  const rest: OneNotePageSummary[] = []
  for (const page of pages) {
    if (page.title?.includes(marker)) matches.push(page)
    else rest.push(page)
  }
  return [...matches, ...rest]
}

/**
 * Keep pages that have meaningful text in the title and/or body.
 * Beacon test pages often have the marker in the title with little body text.
 */
export function shouldIncludeOneNotePage(
  title: string,
  bodyText: string,
  beaconMarker?: string,
): boolean {
  const trimmedTitle = title.trim()
  const trimmedBody = bodyText.trim()
  if (beaconMarker && trimmedTitle.includes(beaconMarker)) return true
  if (beaconMarker && trimmedBody.includes(beaconMarker)) return true
  const combined = [trimmedTitle, trimmedBody].filter(Boolean).join('\n')
  return combined.length >= 3
}

export function buildBeaconTitleFilter(beaconMarker: string): string {
  const escaped = escapeODataString(beaconMarker.toLowerCase())
  return `contains(tolower(title),'${escaped}')`
}