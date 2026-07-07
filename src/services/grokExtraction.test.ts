import { describe, expect, it } from 'vitest'
import { parseGrokTaskJson } from './grokExtraction'

describe('grokExtraction', () => {
  it('parses a bare JSON array', () => {
    const items = parseGrokTaskJson(
      '[{"title":"Call contractor","priority":"high","tags":["Cedar Ridge"]}]',
    )
    expect(items).toHaveLength(1)
    expect(items[0].title).toBe('Call contractor')
  })

  it('parses JSON inside markdown fences', () => {
    const items = parseGrokTaskJson(
      'Here are tasks:\n```json\n[{"title":"Board prep","dueDate":"3/15"}]\n```',
    )
    expect(items[0].title).toBe('Board prep')
    expect(items[0].dueDate).toBe('3/15')
  })

  it('throws when no JSON array present', () => {
    expect(() => parseGrokTaskJson('No tasks found today.')).toThrow(/JSON array/)
  })
})