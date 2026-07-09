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
  saveTasks,
} from '../db/indexedDb'
import type { AppSettings, Task } from '../types/task'
import { completeTask, snoozeTask } from '../services/syncBack'
import { initM365, M365_SIGNED_IN_FLAG, syncM365ClientId } from '../services/connectors'
import { reconcileToDoCompletions } from '../services/primaryToolPush'

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
    syncM365ClientId(appSettings.m365ClientId)
    if (appSettings.m365ClientId) await initM365(appSettings)
    if (sessionStorage.getItem(M365_SIGNED_IN_FLAG)) {
      sessionStorage.removeItem(M365_SIGNED_IN_FLAG)
      setMessage('Signed in to Microsoft 365')
    }

    let tasksToShow = allTasks
    if (appSettings.m365ClientId) {
      const reconcile = await reconcileToDoCompletions(allTasks, appSettings)
      if (reconcile.updated.length > 0) {
        await saveTasks(reconcile.updated)
        const completedIds = new Set(reconcile.updated.map((t) => t.id))
        tasksToShow = allTasks.map((t) => completedIds.has(t.id)
          ? reconcile.updated.find((u) => u.id === t.id)!
          : t)
        setMessage(
          `${reconcile.completedCount} task(s) completed in Microsoft To Do — synced to TaskSweep.`,
        )
      }
    }

    setTasks(tasksToShow.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)))
    setSettings(appSettings)
    setLoading(false)
  }, [])

  useEffect(() => {
    void refresh()
  }, [refresh])

  const sortTasks = useCallback(
    (list: Task[]) => [...list].sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)),
    [],
  )

  const updateTask = useCallback(
    async (task: Task) => {
      const updated = { ...task, updatedAt: new Date().toISOString() }
      await saveTask(updated)
      setTasks((prev) => sortTasks(prev.map((t) => (t.id === updated.id ? updated : t))))
    },
    [sortTasks],
  )

  const markComplete = useCallback(
    async (task: Task) => {
      if (!settings) return
      const now = new Date().toISOString()
      const optimistic: Task = {
        ...task,
        status: 'completed',
        completedAt: now,
        updatedAt: now,
      }
      setTasks((prev) => sortTasks(prev.map((t) => (t.id === task.id ? optimistic : t))))

      try {
        const result = await completeTask(optimistic, settings)
        const saved: Task = {
          ...optimistic,
          syncStatus: result.syncStatus,
          syncMessage: result.message,
        }
        await saveTask(saved)
        setTasks((prev) => sortTasks(prev.map((t) => (t.id === saved.id ? saved : t))))
        setMessage(result.message)
      } catch {
        setTasks((prev) => sortTasks(prev.map((t) => (t.id === task.id ? task : t))))
        setMessage('Could not complete task — try again.')
      }
    },
    [settings, sortTasks],
  )

  const markSnooze = useCallback(
    async (task: Task, days: number) => {
      if (!settings) return
      const until = new Date()
      until.setDate(until.getDate() + days)
      const now = new Date().toISOString()
      const optimistic: Task = {
        ...task,
        status: 'snoozed',
        snoozedUntil: until.toISOString(),
        updatedAt: now,
      }
      setTasks((prev) => sortTasks(prev.map((t) => (t.id === task.id ? optimistic : t))))

      try {
        await saveTask(optimistic)
        const result = await snoozeTask(optimistic, until.toLocaleDateString(), settings)
        setMessage(result.message)
      } catch {
        setTasks((prev) => sortTasks(prev.map((t) => (t.id === task.id ? task : t))))
        setMessage('Could not snooze task — try again.')
      }
    },
    [settings, sortTasks],
  )

  const removeTask = useCallback(
    async (id: string) => {
      let removed: Task | undefined
      setTasks((prev) => {
        removed = prev.find((t) => t.id === id)
        return prev.filter((t) => t.id !== id)
      })
      try {
        await deleteTask(id)
      } catch {
        if (removed) {
          setTasks((prev) => sortTasks([...prev, removed!]))
        }
        setMessage('Could not delete task — try again.')
      }
    },
    [sortTasks],
  )

  const updateSettings = useCallback(
    async (next: AppSettings) => {
      await saveSettings(next)
      syncM365ClientId(next.m365ClientId)
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