/**
 * Minimal .eml (RFC 822) parser for Proton Mail and other exports.
 */

export interface ParsedEml {
  subject: string
  from: string
  date?: string
  messageId?: string
  body: string
}

export function parseEml(raw: string): ParsedEml {
  const normalized = raw.replace(/\r\n/g, '\n')
  const splitIdx = normalized.indexOf('\n\n')
  const headerBlock = splitIdx >= 0 ? normalized.slice(0, splitIdx) : normalized
  const bodyBlock = splitIdx >= 0 ? normalized.slice(splitIdx + 2) : ''

  const headers = parseHeaders(headerBlock)
  const contentType = headers['content-type'] ?? 'text/plain'

  return {
    subject: decodeHeaderValue(headers['subject'] ?? '(no subject)'),
    from: decodeHeaderValue(headers['from'] ?? ''),
    date: headers['date'],
    messageId: stripAngleBrackets(headers['message-id']),
    body: extractBody(bodyBlock, contentType, headers['content-transfer-encoding']),
  }
}

function parseHeaders(block: string): Record<string, string> {
  const lines = block.split('\n')
  const unfolded: string[] = []

  for (const line of lines) {
    if ((line.startsWith(' ') || line.startsWith('\t')) && unfolded.length > 0) {
      unfolded[unfolded.length - 1] += ' ' + line.trim()
    } else {
      unfolded.push(line)
    }
  }

  const headers: Record<string, string> = {}
  for (const line of unfolded) {
    const colon = line.indexOf(':')
    if (colon === -1) continue
    const key = line.slice(0, colon).trim().toLowerCase()
    const value = line.slice(colon + 1).trim()
    headers[key] = headers[key] ? `${headers[key]}\n${value}` : value
  }
  return headers
}

function extractBody(
  body: string,
  contentType: string,
  transferEncoding?: string,
): string {
  const type = contentType.toLowerCase()
  if (type.includes('multipart/')) {
    const boundaryMatch = contentType.match(/boundary="?([^";\s]+)"?/i)
    if (boundaryMatch) {
      return extractMultipart(body, boundaryMatch[1])
    }
  }

  const charset = contentType.match(/charset="?([^";\s]+)"?/i)?.[1]
  return decodePart(body, transferEncoding, charset)
}

function extractMultipart(body: string, boundary: string): string {
  const parts = body.split(`--${boundary}`)
  let plain = ''
  let html = ''

  for (const part of parts) {
    const trimmed = part.trim()
    if (!trimmed || trimmed === '--') continue

    const partSplit = trimmed.indexOf('\n\n')
    if (partSplit === -1) continue

    const partHeaders = parseHeaders(trimmed.slice(0, partSplit))
    const partBody = trimmed.slice(partSplit + 2)
    const partType = partHeaders['content-type'] ?? 'text/plain'

    if (partType.toLowerCase().includes('multipart/')) {
      const nestedBoundary = partType.match(/boundary="?([^";\s]+)"?/i)?.[1]
      if (nestedBoundary) {
        const nested = extractMultipart(partBody, nestedBoundary)
        if (nested) return nested
      }
      continue
    }

    const encoding = partHeaders['content-transfer-encoding']
    const charset = partType.match(/charset="?([^";\s]+)"?/i)?.[1]
    const decoded = decodePart(partBody, encoding, charset)

    if (partType.toLowerCase().includes('text/plain') && !plain) {
      plain = decoded
    } else if (partType.toLowerCase().includes('text/html') && !html) {
      html = htmlToPlain(decoded)
    }
  }

  return plain || html || body.trim()
}

function decodePart(body: string, encoding?: string, _charset?: string): string {
  const enc = encoding?.toLowerCase() ?? ''
  if (enc.includes('base64')) {
    try {
      return atob(body.replace(/\s/g, ''))
    } catch {
      return body.trim()
    }
  }
  if (enc.includes('quoted-printable')) {
    return decodeQuotedPrintable(body)
  }
  return body.trim()
}

function decodeQuotedPrintable(input: string): string {
  const softBreaksRemoved = input.replace(/=\r?\n/g, '')
  return softBreaksRemoved.replace(/=([0-9A-Fa-f]{2})/g, (_, hex: string) =>
    String.fromCharCode(parseInt(hex, 16)),
  )
}

function decodeHeaderValue(value: string): string {
  const encodedWords = value.match(/=\?([^?]+)\?([BQbq])\?([^?]+)\?=/g)
  if (!encodedWords) return value.trim()

  let decoded = value
  for (const word of encodedWords) {
    const match = word.match(/=\?([^?]+)\?([BQbq])\?([^?]+)\?=/)
    if (!match) continue
    const [, , encoding, payload] = match
    let replacement = payload
    if (encoding.toUpperCase() === 'B') {
      try {
        replacement = atob(payload.replace(/\s/g, ''))
      } catch {
        replacement = payload
      }
    } else {
      replacement = decodeQuotedPrintable(payload.replace(/_/g, ' '))
    }
    decoded = decoded.replace(word, replacement)
  }
  return decoded.trim()
}

function stripAngleBrackets(value?: string): string | undefined {
  if (!value) return undefined
  return value.replace(/^<|>$/g, '').trim() || undefined
}

function htmlToPlain(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

/** Build sweep content block from parsed headers + body */
export function formatEmlForExtraction(parsed: ParsedEml): string {
  const lines = [
    parsed.subject,
    parsed.from ? `From: ${parsed.from}` : '',
    parsed.date ? `Date: ${parsed.date}` : '',
    '',
    parsed.body,
  ]
  return lines.filter((line, i) => i < 3 || line.length > 0).join('\n').trim()
}