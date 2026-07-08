import { describe, expect, it } from 'vitest'
import { htmlToText } from './onenoteHtml'

describe('htmlToText', () => {
  it('strips HTML tags and preserves line breaks', () => {
    const text = htmlToText(
      '<p>Call contractor</p><ul><li>Permit follow-up</li><li>Due 3/15</li></ul>',
    )
    expect(text).toContain('Call contractor')
    expect(text).toContain('Permit follow-up')
    expect(text).not.toContain('<p>')
  })

  it('decodes basic entities', () => {
    expect(htmlToText('<p>Tom &amp; Jerry</p>')).toContain('Tom & Jerry')
  })
})