import { describe, expect, it } from 'vitest'
import {
  buildBeaconTitleFilter,
  buildGraphQuery,
  dedupePagesById,
  encodeOneNoteResourceId,
  shouldIncludeOneNotePage,
  sortPagesByModified,
} from './onenotePages'

describe('shouldIncludeOneNotePage', () => {
  it('includes pages when beacon marker is only in the title', () => {
    expect(
      shouldIncludeOneNotePage('[TaskSweep-Beacon] Test page', '', '[TaskSweep-Beacon]'),
    ).toBe(true)
  })

  it('skips empty pages without a beacon marker', () => {
    expect(shouldIncludeOneNotePage('', '', '[TaskSweep-Beacon]')).toBe(false)
  })

  it('includes pages with short but non-empty combined text', () => {
    expect(shouldIncludeOneNotePage('Todo', 'buy milk')).toBe(true)
  })
})

describe('dedupePagesById', () => {
  it('removes duplicate page ids', () => {
    const pages = dedupePagesById([
      { id: 'a', title: 'One' },
      { id: 'a', title: 'One duplicate' },
      { id: 'b', title: 'Two' },
    ])
    expect(pages).toHaveLength(2)
    expect(pages.map((p) => p.id)).toEqual(['a', 'b'])
  })
})

describe('sortPagesByModified', () => {
  it('orders newest pages first', () => {
    const sorted = sortPagesByModified([
      { id: 'old', lastModifiedDateTime: '2026-01-01T00:00:00Z' },
      { id: 'new', lastModifiedDateTime: '2026-07-01T00:00:00Z' },
    ])
    expect(sorted[0].id).toBe('new')
  })
})

describe('buildGraphQuery', () => {
  it('keeps OData $ prefixes and encodes values', () => {
    const path = buildGraphQuery('/me/onenote/pages', {
      $top: 20,
      $select: 'id,title,lastModifiedDateTime',
    })
    expect(path).toBe(
      '/me/onenote/pages?$top=20&$select=id%2Ctitle%2ClastModifiedDateTime',
    )
  })
})

describe('encodeOneNoteResourceId', () => {
  it('encodes exclamation marks in page ids', () => {
    expect(encodeOneNoteResourceId('1-abc!def')).toBe('1-abc%21def')
  })
})

describe('buildBeaconTitleFilter', () => {
  it('lowercases marker and escapes quotes for OData', () => {
    expect(buildBeaconTitleFilter("[TaskSweep-Beacon]")).toBe(
      "contains(tolower(title),'[tasksweep-beacon]')",
    )
    expect(buildBeaconTitleFilter("O'Brien")).toBe("contains(tolower(title),'o''brien')")
  })
})