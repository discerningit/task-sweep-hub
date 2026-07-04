import { AI_PROVIDERS } from '../services/aiOrchestrator'
import type { AppSettings, AiProviderId } from '../types/task'

interface AiSelectorProps {
  settings: AppSettings
  onChange: (settings: AppSettings) => void
}

export function AiSelector({ settings, onChange }: AiSelectorProps) {
  const toggleProvider = (id: AiProviderId) => {
    const enabled = settings.enabledAiProviders.includes(id)
    const next = enabled
      ? settings.enabledAiProviders.filter((p) => p !== id)
      : [...settings.enabledAiProviders, id]
    onChange({ ...settings, enabledAiProviders: next.length ? next : ['local'] })
  }

  return (
    <section className="panel ai-selector">
      <h2>AI extraction</h2>
      <p className="hint">
        <strong>Local</strong> works offline. Enable others when you add API keys in Settings.
      </p>
      <div className="ai-grid">
        {AI_PROVIDERS.map((provider) => {
          const enabled = settings.enabledAiProviders.includes(provider.id)
          const configured = provider.isConfigured(settings)
          return (
            <label key={provider.id} className={`ai-card ${enabled ? 'enabled' : ''}`}>
              <input
                type="checkbox"
                checked={enabled}
                onChange={() => toggleProvider(provider.id)}
              />
              <span className="ai-name">{provider.name}</span>
              {!configured && provider.id !== 'local' && (
                <span className="ai-status">needs setup</span>
              )}
            </label>
          )
        })}
      </div>
      <label className="field">
        Primary AI for sweeps
        <select
          value={settings.primaryAi}
          onChange={(e) =>
            onChange({ ...settings, primaryAi: e.target.value as AiProviderId })
          }
        >
          {settings.enabledAiProviders.map((id) => {
            const p = AI_PROVIDERS.find((x) => x.id === id)
            return (
              <option key={id} value={id}>
                {p?.name ?? id}
              </option>
            )
          })}
        </select>
      </label>
    </section>
  )
}