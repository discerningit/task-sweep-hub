/**
 * TaskSweep Hub — main app shell
 *
 * Tabs: Tasks | Beacon | Settings
 * All data stays local in IndexedDB unless you connect M365.
 */

import { useCallback, useState } from 'react'
import { InputArea } from './components/InputArea'
import { TaskList } from './components/TaskList'
import { AiSelector } from './components/AiSelector'
import { BeaconTools } from './components/BeaconTools'
import { SettingsPanel } from './components/SettingsPanel'
import { useTasks } from './hooks/useTasks'
import { runSweep } from './services/sweepPipeline'
import { exportTasksCsv } from './services/syncBack'
import { initM365 } from './services/connectors'
import type { BeaconHit } from './types/task'
import { needsDeviceSetup } from './services/settingsPack'


type Tab = 'tasks' | 'beacon' | 'settings'

function App() {
  const {
    tasks,
    settings,
    loading,
    message,
    clearMessage,
    refresh,
    markComplete,
    markSnooze,
    removeTask,
    updateSettings,
  } = useTasks()

  const [tab, setTab] = useState<Tab>('tasks')
  const [sweeping, setSweeping] = useState(false)
  const [search, setSearch] = useState('')
  const [filter, setFilter] = useState<'all' | 'open' | 'completed' | 'snoozed'>('open')
  const [sweepSummary, setSweepSummary] = useState<string | null>(null)
  const [beaconAlerts, setBeaconAlerts] = useState<BeaconHit[]>([])


  const handleSweep = useCallback(
    async (connectorIds: string[]) => {
      setSweeping(true)
      setSweepSummary(null)
      try {
        if (settings?.m365ClientId) await initM365(settings)
        const result = await runSweep(connectorIds)
        setSweepSummary(
          `Found ${result.newTaskCount} new task(s) from ${result.sources.join(', ') || 'no sources'}.`,
        )
        if (result.beacons.length > 0) setBeaconAlerts(result.beacons)
        await refresh()
      } catch (err) {
        setSweepSummary(err instanceof Error ? err.message : 'Sweep failed')
      } finally {
        setSweeping(false)
      }
    },
    [settings, refresh],
  )

  const handleExport = () => {
    const csv = exportTasksCsv(tasks)
    const blob = new Blob([csv], { type: 'text/csv' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `tasksweep-export-${new Date().toISOString().slice(0, 10)}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }

  if (loading || !settings) {
    return <div className="loading">Loading TaskSweep Hub…</div>
  }

  const openCount = tasks.filter((t) => t.status === 'open').length

  return (
    <div className="app">
      <header className="header">
        <div className="header-brand">
          <h1>TaskSweep Hub</h1>
          <p className="tagline">Sweep tasks from everywhere into one simple list</p>
        </div>
        <nav className="tabs">
          <button
            type="button"
            className={tab === 'tasks' ? 'active' : ''}
            onClick={() => setTab('tasks')}
          >
            Tasks {openCount > 0 && <span className="badge">{openCount}</span>}
          </button>
          <button
            type="button"
            className={tab === 'beacon' ? 'active' : ''}
            onClick={() => setTab('beacon')}
          >
            Beacon
          </button>
          <button
            type="button"
            className={tab === 'settings' ? 'active' : ''}
            onClick={() => setTab('settings')}
          >
            Settings
          </button>
        </nav>
      </header>

      {needsDeviceSetup(settings) && (
        <div className="onboarding-banner">
          <div className="onboarding-text">
            <strong>New device?</strong> Import settings from a computer where TaskSweep is already set up.
          </div>
          <div className="onboarding-actions">
            <button type="button" className="primary" onClick={() => setTab('settings')}>
              Set up this device
            </button>
            <button
              type="button"
              className="secondary"
              onClick={() => void updateSettings({ ...settings, setupCompleted: true })}
            >
              I&apos;ll enter settings manually
            </button>
          </div>
        </div>
      )}

      {(message || sweepSummary) && (
        <div className="toast" onClick={clearMessage}>
          {sweepSummary ?? message}
          <span className="toast-dismiss">click to dismiss</span>
        </div>
      )}

      {beaconAlerts.length > 0 && tab === 'tasks' && (
        <div className="beacon-alert">
          <strong>Beacon detected</strong> — {beaconAlerts[0].suggestedConnector}
          <button type="button" onClick={() => setBeaconAlerts([])}>Dismiss</button>
        </div>
      )}

      <main className="main">
        {tab === 'tasks' && (
          <>
            <InputArea onSweep={handleSweep} sweeping={sweeping} />
            <AiSelector
              settings={settings}
              onChange={(next) => void updateSettings(next)}
            />
            <section className="panel task-panel">
              <div className="task-toolbar">
                <h2>Your tasks</h2>
                <input
                  type="search"
                  placeholder="Search tasks…"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  className="search"
                />
                <select
                  value={filter}
                  onChange={(e) => setFilter(e.target.value as typeof filter)}
                >
                  <option value="open">Open</option>
                  <option value="all">All</option>
                  <option value="completed">Completed</option>
                  <option value="snoozed">Snoozed</option>
                </select>
              </div>
              <TaskList
                tasks={tasks}
                search={search}
                filter={filter}
                onComplete={(t) => void markComplete(t)}
                onSnooze={(t, d) => void markSnooze(t, d)}
                onDelete={(id) => void removeTask(id)}
              />
            </section>
          </>
        )}

        {tab === 'beacon' && (
          <BeaconTools
            settings={settings}
            onBeaconFound={(hit) => setBeaconAlerts([hit])}
          />
        )}

        {tab === 'settings' && (
          <SettingsPanel
            settings={settings}
            onSave={(s) => void updateSettings(s)}
            onExport={handleExport}
          />
        )}
      </main>

      <footer className="footer">
        Local-first · Data stays on this device · MVP: paste + file + M365
      </footer>
    </div>
  )
}

export default App