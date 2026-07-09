import { describe, expect, it } from 'vitest'
import { formatEmlForExtraction, parseEml } from './emlParser'

const SIMPLE_EML = `From: Client <client@example.com>
To: me@proton.me
Subject: Cedar Ridge permit follow-up
Date: Mon, 9 Jul 2026 10:00:00 +0000
Message-ID: <abc123@proton.me>
Content-Type: text/plain; charset=utf-8

Please call the contractor about the permit status by Friday.
`

describe('emlParser', () => {
  it('parses plain-text .eml headers and body', () => {
    const parsed = parseEml(SIMPLE_EML)
    expect(parsed.subject).toBe('Cedar Ridge permit follow-up')
    expect(parsed.from).toContain('client@example.com')
    expect(parsed.messageId).toBe('abc123@proton.me')
    expect(parsed.body).toContain('call the contractor')
  })

  it('formats content for task extraction', () => {
    const parsed = parseEml(SIMPLE_EML)
    const content = formatEmlForExtraction(parsed)
    expect(content).toContain('Cedar Ridge permit follow-up')
    expect(content).toContain('call the contractor')
  })

  it('extracts text/plain from multipart messages', () => {
    const multipart = `From: Board <board@nonprofit.org>
Subject: Meeting prep
Content-Type: multipart/alternative; boundary="boundary42"

--boundary42
Content-Type: text/plain; charset=utf-8

Review agenda before Tuesday board meeting.

--boundary42
Content-Type: text/html; charset=utf-8

<p>Review agenda before Tuesday board meeting.</p>

--boundary42--
`
    const parsed = parseEml(multipart)
    expect(parsed.subject).toBe('Meeting prep')
    expect(parsed.body).toContain('Review agenda')
    expect(parsed.body).not.toContain('<p>')
  })
})