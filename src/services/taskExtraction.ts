/**
 * Local task extraction (no AI API required).
 *
 * Parses common patterns from pasted email, bullet lists, and
 * todo exports. Tuned for an IT consultant context: Cedar Ridge
 * home build, nonprofit board work, family errands, client tickets.
 */

import type { ExtractedTask, RawInput, TaskPriority } from '../types/task'

const CONTEXT_KEYWORDS: Record<string, string[]> = {
  'Cedar Ridge': ['cedar ridge', 'cedar', 'home build', 'contractor', 'permit'],
  nonprofit: ['nonprofit', 'board', '501c', 'fundraising', 'volunteer', 'donor'],
  family: ['family', 'kids', 'school', 'doctor', 'appointment', 'mom', 'dad'],
  'IT consulting': ['client', 'ticket', 'jira', 'sla', 'on-call', 'deploy', 'incident'],
}

const PRIORITY_PATTERNS: { pattern: RegExp; priority: TaskPriority }[] = [
  { pattern: /\b(urgent|asap|critical|p0|!!!)\b/i, priority: 'urgent' },
  { pattern: /\b(high|important|p1|!!)\b/i, priority: 'high' },
  { pattern: /\b(low|minor|p3|whenever)\b/i, priority: 'low' },
]

const DATE_PATTERNS = [
  /(?:due|by|before|deadline)[:\s]+(\d{1,2}[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)/i,
  /(?:due|by|before|deadline)[:\s]+(\w+ \d{1,2}(?:,?\s*\d{4})?)/i,
  /(\d{4}-\d{2}-\d{2})/,
  /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/,
]

/** Lines that look like tasks (bullets, checkboxes, numbered, ALL CAPS action) */
const TASK_LINE =
  /^(?:[-*•]\s*|\[\s?[xX ]?\]\s*|\d+[.)]\s*|TODO[:\s]|ACTION[:\s]|TASK[:\s])/i

function inferTags(text: string): string[] {
  const lower = text.toLowerCase()
  const tags: string[] = []
  for (const [tag, keywords] of Object.entries(CONTEXT_KEYWORDS)) {
    if (keywords.some((kw) => lower.includes(kw))) tags.push(tag)
  }
  return tags
}

function inferPriority(text: string): TaskPriority {
  for (const { pattern, priority } of PRIORITY_PATTERNS) {
    if (pattern.test(text)) return priority
  }
  return 'normal'
}

function extractDueDate(text: string): string | undefined {
  for (const pattern of DATE_PATTERNS) {
    const match = text.match(pattern)
    if (match?.[1]) return match[1].trim()
  }
  return undefined
}

function cleanTitle(line: string): string {
  return line
    .replace(TASK_LINE, '')
    .replace(/\s*[-–—]\s*(due|by).*/i, '')
    .replace(/\[TaskSweep-Beacon\]/gi, '')
    .trim()
}

function splitIntoLines(content: string): string[] {
  return content
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 2)
}

/**
 * Extract tasks from one raw input chunk using local rules.
 * AI orchestrator can replace/enhance this when an API is configured.
 */
function extractAppleRemindersInput(input: RawInput): ExtractedTask[] {
  const lines = input.content.split(/\r?\n/).map((l) => l.trim()).filter(Boolean)
  const title = lines[0] ?? ''
  if (title.length < 2) return []

  const notes = lines.find((l) => !l.startsWith('Due:') && !l.startsWith('List:') && l !== title)
  const dueFromMeta = input.metadata?.dueDate
  const dueFromLine = lines.find((l) => l.startsWith('Due:'))?.slice(4).trim()

  return [
    {
      title,
      dueDate: dueFromMeta ?? dueFromLine ?? extractDueDate(input.content),
      priority: inferPriority(input.content),
      notes: notes && notes !== title ? notes : undefined,
      source: input.source,
      sourceId: input.metadata?.reminderId,
      metadata: input.metadata,
      tags: inferTags(input.content),
    },
  ]
}

export function extractTasksLocally(input: RawInput): ExtractedTask[] {
  if (input.source === 'apple-reminders') {
    return extractAppleRemindersInput(input)
  }

  const lines = splitIntoLines(input.content)
  const tasks: ExtractedTask[] = []

  for (const line of lines) {
    const looksLikeTask =
      TASK_LINE.test(line) ||
      /\bdue\b/i.test(line) ||
      /\b(todo|action|remind|follow.?up|call|email|review|submit)\b/i.test(line)

    if (!looksLikeTask) continue

    const title = cleanTitle(line)
    if (title.length < 3) continue

    const tags = inferTags(line + ' ' + (input.metadata?.subject ?? ''))

    tasks.push({
      title,
      dueDate: extractDueDate(line),
      priority: inferPriority(line),
      notes: input.metadata?.subject
        ? `From: ${input.metadata.subject}`
        : undefined,
      source: input.source,
      sourceUrl: input.sourceUrl,
      sourceId: input.metadata?.id,
      metadata: input.metadata,
      tags,
    })
  }

  // If no structured lines found, treat short paste as a single task
  if (tasks.length === 0 && input.content.trim().length < 500) {
    const trimmed = input.content.trim()
    if (trimmed.length >= 3 && !trimmed.startsWith('{')) {
      tasks.push({
        title: cleanTitle(trimmed.split('\n')[0]),
        dueDate: extractDueDate(trimmed),
        priority: inferPriority(trimmed),
        source: input.source,
        sourceUrl: input.sourceUrl,
        metadata: input.metadata,
        tags: inferTags(trimmed),
      })
    }
  }

  return tasks
}