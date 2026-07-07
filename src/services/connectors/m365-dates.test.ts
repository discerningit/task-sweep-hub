import { describe, expect, it } from 'vitest'
import { parseDueDateForGraph } from './m365'

describe('parseDueDateForGraph', () => {
  it('parses ISO dates', () => {
    expect(parseDueDateForGraph('2026-03-15')?.dateTime).toBe('2026-03-15T00:00:00')
  })

  it('parses US-style dates', () => {
    expect(parseDueDateForGraph('3/15/2026')?.dateTime).toBe('2026-03-15T00:00:00')
  })

  it('returns undefined for unparseable text', () => {
    expect(parseDueDateForGraph('next Friday')).toBeUndefined()
  })
})