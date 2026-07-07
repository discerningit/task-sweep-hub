/**
 * Grok (xAI) task extraction — sends sweep content to the xAI API
 * and parses structured tasks from the response.
 *
 * API key: Settings → Grok API key (stored locally on your device).
 * Local dev: optional VITE_GROK_API_KEY in .env or Vite proxy (see vite.config.ts).
 */

import type { AppSettings, ExtractedTask, RawInput, TaskPriority } from '../types/task'
import { buildContextPrompt } from './aiOrchestrator'
import { extractTasksLocally } from './taskExtraction'

const GROK_API = import.meta.env.DEV
  ? '/api/xai/v1/chat/completions'
  : 'https://api.x.ai/v1/chat/completions'

const GROK_MODEL = 'grok-3-mini'
const MAX_INPUT_CHARS = 12_000

export function getGrokApiKey(settings: AppSettings): string | undefined {
  const fromSettings = settings.grokApiKey?.trim()
  if (fromSettings) return fromSettings
  const fromEnv = import.meta.env.VITE_GROK_API_KEY as string | undefined
  return fromEnv?.trim() || undefined
}

export function isGrokConfigured(settings: AppSettings): boolean {
  return Boolean(getGrokApiKey(settings))
}

interface GrokTaskJson {
  title?: string
  dueDate?: string
  priority?: string
  notes?: string
  tags?: string[]
}

function normalizePriority(value?: string): TaskPriority {
  const p = value?.toLowerCase() ?? ''
  if (p.includes('urgent') || p === 'p0') return 'urgent'
  if (p.includes('high') || p === 'p1') return 'high'
  if (p.includes('low') || p === 'p3') return 'low'
  return 'normal'
}

/** Pull a JSON array from model output (handles markdown fences) */
export function parseGrokTaskJson(text: string): GrokTaskJson[] {
  const trimmed = text.trim()
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i)
  const candidate = fenced?.[1]?.trim() ?? trimmed

  const start = candidate.indexOf('[')
  const end = candidate.lastIndexOf(']')
  if (start === -1 || end === -1 || end <= start) {
    throw new Error('Grok response did not contain a JSON array')
  }

  const parsed = JSON.parse(candidate.slice(start, end + 1)) as unknown
  if (!Array.isArray(parsed)) throw new Error('Grok response JSON is not an array')
  return parsed as GrokTaskJson[]
}

function buildUserPrompt(inputs: RawInput[]): string {
  const blocks = inputs.map((input) => {
    const header = `--- Source: ${input.source}${input.metadata?.subject ? ` | ${input.metadata.subject}` : ''} ---`
    const content = input.content.slice(0, MAX_INPUT_CHARS)
    return `${header}\n${content}`
  })
  return blocks.join('\n\n').slice(0, MAX_INPUT_CHARS)
}

function toExtractedTasks(
  items: GrokTaskJson[],
  inputs: RawInput[],
): ExtractedTask[] {
  const fallbackSource = inputs[0]?.source ?? 'paste'
  const fallbackMeta = inputs[0]?.metadata
  const fallbackUrl = inputs[0]?.sourceUrl

  return items
    .filter((item) => item.title && item.title.trim().length > 2)
    .map((item) => ({
      title: item.title!.trim(),
      dueDate: item.dueDate?.trim() || undefined,
      priority: normalizePriority(item.priority),
      notes: item.notes?.trim() || undefined,
      source: fallbackSource,
      sourceUrl: fallbackUrl,
      sourceId: fallbackMeta?.id,
      metadata: fallbackMeta,
      tags: Array.isArray(item.tags) ? item.tags.map(String) : [],
    }))
}

export async function extractWithGrok(
  inputs: RawInput[],
  settings: AppSettings,
): Promise<ExtractedTask[]> {
  const apiKey = getGrokApiKey(settings)
  if (!apiKey || inputs.length === 0) {
    return inputs.flatMap(extractTasksLocally)
  }

  try {
    const res = await fetch(GROK_API, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: GROK_MODEL,
        temperature: 0.2,
        messages: [
          { role: 'system', content: buildContextPrompt(settings) },
          {
            role: 'user',
            content: `Extract actionable tasks from these inputs:\n\n${buildUserPrompt(inputs)}`,
          },
        ],
      }),
    })

    if (!res.ok) {
      const errText = await res.text().catch(() => '')
      throw new Error(`Grok API ${res.status}${errText ? `: ${errText.slice(0, 120)}` : ''}`)
    }

    const data = (await res.json()) as {
      choices?: { message?: { content?: string } }[]
    }
    const content = data.choices?.[0]?.message?.content
    if (!content) throw new Error('Grok returned an empty response')

    const parsed = parseGrokTaskJson(content)
    const tasks = toExtractedTasks(parsed, inputs)
    if (tasks.length === 0) throw new Error('Grok found no actionable tasks')
    return tasks
  } catch (err) {
    console.warn('Grok extraction failed, using local rules:', err)
    return inputs.flatMap(extractTasksLocally)
  }
}