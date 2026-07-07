/**
 * AI orchestration — routes input to user-selected AI providers.
 *
 * - Local: rule-based extraction (offline)
 * - Grok: xAI API (Settings → Grok API key)
 */

import type { AppSettings, AiProviderId, ExtractedTask, RawInput } from '../types/task'
import { extractTasksLocally } from './taskExtraction'
import { extractWithGrok, isGrokConfigured } from './grokExtraction'

export interface AiProvider {
  id: AiProviderId
  name: string
  description: string
  isConfigured: (settings: AppSettings) => boolean
  extract: (inputs: RawInput[], settings: AppSettings) => Promise<ExtractedTask[]>
}

/** Context block sent to AI providers */
export function buildContextPrompt(settings: AppSettings): string {
  const tags = settings.contextTags.join(', ')
  return `You are extracting actionable tasks for an IT consultant with irregular hours.
Context areas: ${tags}.
Home building project: "Cedar Ridge". Also nonprofit leadership and family tasks.

Return ONLY a JSON array (no markdown prose) with objects:
{ "title": string, "dueDate": string optional, "priority": "low"|"normal"|"high"|"urgent", "notes": string optional, "tags": string[] optional }

Skip newsletters, signatures, ads, and non-actionable text. Merge duplicate lines.`
}

const localProvider: AiProvider = {
  id: 'local',
  name: 'Local (built-in)',
  description: 'Free, private, works offline. Good for lists and paste.',
  isConfigured: () => true,
  async extract(inputs) {
    return inputs.flatMap(extractTasksLocally)
  },
}

const grokProvider: AiProvider = {
  id: 'grok',
  name: 'Grok',
  description: 'xAI Grok — smarter extraction from messy email and notes.',
  isConfigured: isGrokConfigured,
  async extract(inputs, settings) {
    return extractWithGrok(inputs, settings)
  },
}

function stubProvider(
  id: AiProviderId,
  name: string,
  description: string,
  check: (s: AppSettings) => boolean,
): AiProvider {
  return {
    id,
    name,
    description,
    isConfigured: check,
    async extract(inputs, settings) {
      if (!check(settings)) {
        return localProvider.extract(inputs, settings)
      }
      const tasks = await localProvider.extract(inputs, settings)
      return tasks.map((t) => ({
        ...t,
        tags: [
          ...new Set([
            ...(t.tags ?? []),
            ...settings.contextTags.filter((tag) =>
              (t.title + (t.notes ?? ''))
                .toLowerCase()
                .includes(tag.toLowerCase().split(' ')[0]),
            ),
          ]),
        ],
      }))
    },
  }
}

export const AI_PROVIDERS: AiProvider[] = [
  localProvider,
  grokProvider,
  stubProvider('copilot', 'Microsoft Copilot', 'Best with M365 work account', () => false),
  stubProvider('claude', 'Claude', 'Via API key or work VDI', () => false),
  stubProvider('kiro', 'Kiro', 'When available in your environment', () => false),
  stubProvider('siri', 'Siri / Shortcuts', 'Use iOS export or paste from Reminders', () => false),
]

export function getProvider(id: AiProviderId): AiProvider {
  return AI_PROVIDERS.find((p) => p.id === id) ?? localProvider
}

export async function orchestrateExtraction(
  inputs: RawInput[],
  settings: AppSettings,
): Promise<ExtractedTask[]> {
  const provider = getProvider(settings.primaryAi)
  const enabled = settings.enabledAiProviders.includes(provider.id)

  if (!enabled || !provider.isConfigured(settings)) {
    return localProvider.extract(inputs, settings)
  }

  return provider.extract(inputs, settings)
}