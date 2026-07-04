/**
 * React hook — loads tasks from IndexedDB and exposes actions.
 */

import { useCallback, useEffect, useState } from 'react'
import {
  deleteTask,
  getAllTasks,
  getSettings,
  saveSettings,
  saveTask,
} from '../db/indexedDb'
import type { AppSettings, Task } from '../types/task'
import { completeTask, snoozeTask } from '../services/syncBack'

export function useTasks() {
  const [tasks, setTasks] = useState<Task[]>([])
  const [settings, setSettings] = useState<AppSettings | null>(null)
  const [loading, setLoading] = useState(true)
  const [message, setMessage] = useState<string | null>(null)

  const refresh = useCallback(async () => {
    const [allTasks, appSettings] = await Promise.all([
      getAllTasks(),
      getSettings(),
    ])
    setTasks(allTasks.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)))
    setSettings(appSettings)
    setLoading(false)
  }, [])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const updateTask = useCallback(
    async (task: Task) => {
      const updated = { ...task, updatedAt: new Date().toISOString() }
      await saveTask(updated)
      await refresh()
    },
    [refresh],
  )

  const markComplete = useCallback(
    async (task: Task) => {
      if (!settings) return
      const updated: Task = {
        ...task,
        status: 'completed',
        completedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      }
      await saveTask(updated)
      const result = await completeTask(updated, settings)
      setMessage(result.message)
      await refresh()
    },
    [settings, refresh],
  )

  const markSnooze = useCallback(
    async (task: Task, days: number) => {
      if (!settings) return
      const until = new Date()
      until.setDate(until.getDate() + days)
      const updated: Task = {
        ...task,
        status: 'snoozed',
        snoozedUntil: until.toISOString(),
        updatedAt: new Date().toISOString(),
      }
      await saveTask(updated)
      const result = await snoozeTask(updated, until.toLocaleDateString(), settings)
      setMessage(result.message)
      await refresh()
    },
    [settings, refresh],
  )

  const removeTask = useCallback(
    async (id: string) => {
      await deleteTask(id)
      await refresh()
    },
    [refresh],
  )

  const updateSettings = useCallback(
    async (next: AppSettings) => {
      await saveSettings(next)
      setSettings(next)
      setMessage('Settings saved')
    },
    [],
  )

  const clearMessage = useCallback(() => setMessage(null), [])

  return {
    tasks,
    settings,
    loading,
    message,
    clearMessage,
    refresh,
    updateTask,
    markComplete,
    markSnooze,
    removeTask,
    updateSettings,
  }
}