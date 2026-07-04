/**
 * AI orchestration — routes input to user-selected AI providers.
 *
 * MVP: "local" provider uses rule-based extraction (no API key).
 * Other providers are stubbed with clear hooks for API keys later.
 *
 * IT consultant context is injected into prompts when real AI is wired up.
 */

import type { AppSettings, AiProviderId, ExtractedTask, RawInput } from '../types/task'
import { extractTasksLocally } from './taskExtraction'

export interface AiProvider {
  id: AiProviderId
  name: string
  description: string
  /** True if user has configured API/auth for this provider */
  isConfigured: (settings: AppSettings) => boolean
  extract: (inputs: RawInput[], settings: AppSettings) => Promise<ExtractedTask[]>
}

/** Context block sent to AI providers (when implemented) */
export function buildContextPrompt(settings: AppSettings): string {
  const tags = settings.contextTags.join(', ')
  return `You are extracting actionable tasks for an IT consultant with irregular hours.
Context areas: ${tags}.
Home building project: "Cedar Ridge". Also nonprofit leadership and family tasks.
Return JSON array: [{ "title", "dueDate", "priority", "notes", "tags" }].
Skip newsletters, signatures, and non-actionable text.`
}

/** Local rule-based provider — always available, no network */
const localProvider: AiProvider = {
  id: 'local',
  name: 'Local (built-in)',
  description: 'Free, private, works offline. Good for lists and paste.',
  isConfigured: () => true,
  async extract(inputs) {
    return inputs.flatMap(extractTasksLocally)
  },
}

/** Placeholder providers — enable in Settings when you add API keys */
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
        console.warn(`${name} not configured — falling back to local extraction`)
        return localProvider.extract(inputs, settings)
      }
      // TODO: Wire real API call here. For now, local + context tags.
      const tasks = await localProvider.extract(inputs, settings)
      return tasks.map((t) => ({
        ...t,
        tags: [...new Set([...t.tags ?? [], ...settings.contextTags.filter((tag) =>
          (t.title + (t.notes ?? '')).toLowerCase().includes(tag.toLowerCase().split(' ')[0]),
        )])],
      }))
    },
  }
}

export const AI_PROVIDERS: AiProvider[] = [
  localProvider,
  stubProvider('copilot', 'Microsoft Copilot', 'Best with M365 work account', () => false),
  stubProvider('claude', 'Claude', 'Via API key or work VDI', (s) => Boolean(s.m365ClientId)),
  stubProvider('kiro', 'Kiro', 'When available in your environment', () => false),
  stubProvider('grok', 'Grok', 'xAI Grok API', () => false),
  stubProvider('siri', 'Siri / Shortcuts', 'Use iOS export or paste from Reminders', () => false),
]

export function getProvider(id: AiProviderId): AiProvider {
  return AI_PROVIDERS.find((p) => p.id === id) ?? localProvider
}

/** Main entry: extract tasks from raw inputs using selected AI */
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