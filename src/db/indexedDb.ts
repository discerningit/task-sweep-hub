/**
 * Local-first storage using IndexedDB (via the `idb` library).
 *
 * All your tasks and settings stay on THIS device unless you
 * explicitly sync via a connector. Nothing is sent to TaskSweep servers.
 */

import { openDB, type DBSchema, type IDBPDatabase } from 'idb'
import type { AppSettings, Task } from '../types/task'

interface SettingsRecord {
  id: 'app'
  data: AppSettings
}

interface TaskSweepDB extends DBSchema {
  tasks: {
    key: string
    value: Task
    indexes: {
      'by-status': TaskStatus
      'by-source': string
      'by-hash': string
    }
  }
  settings: {
    key: 'app'
    value: SettingsRecord
  }
}

type TaskStatus = Task['status']

const DB_NAME = 'tasksweep-hub'
const DB_VERSION = 1

let dbPromise: Promise<IDBPDatabase<TaskSweepDB>> | null = null

function getDb() {
  if (!dbPromise) {
    dbPromise = openDB<TaskSweepDB>(DB_NAME, DB_VERSION, {
      upgrade(db) {
        const taskStore = db.createObjectStore('tasks', { keyPath: 'id' })
        taskStore.createIndex('by-status', 'status')
        taskStore.createIndex('by-source', 'source')
        taskStore.createIndex('by-hash', 'contentHash')

        db.createObjectStore('settings', { keyPath: 'id' })
      },
    })
  }
  return dbPromise
}

export const DEFAULT_SETTINGS: AppSettings = {
  enabledAiProviders: ['local'],
  primaryAi: 'local',
  primaryTaskTool: 'hub-only',
  protonMailEnabled: true,
  beaconMarker: '[TaskSweep-Beacon]',
  contextTags: [
    'Cedar Ridge',
    'nonprofit',
    'family',
    'IT consulting',
    'irregular hours',
  ],
}

export async function getAllTasks(): Promise<Task[]> {
  const db = await getDb()
  return db.getAll('tasks')
}

export async function getTask(id: string): Promise<Task | undefined> {
  const db = await getDb()
  return db.get('tasks', id)
}

export async function saveTask(task: Task): Promise<void> {
  const db = await getDb()
  await db.put('tasks', task)
}

export async function saveTasks(tasks: Task[]): Promise<void> {
  const db = await getDb()
  const tx = db.transaction('tasks', 'readwrite')
  await Promise.all([...tasks.map((t) => tx.store.put(t)), tx.done])
}

export async function deleteTask(id: string): Promise<void> {
  const db = await getDb()
  await db.delete('tasks', id)
}

export async function getSettings(): Promise<AppSettings> {
  const db = await getDb()
  const stored = await db.get('settings', 'app')
  return stored?.data ?? DEFAULT_SETTINGS
}

export async function saveSettings(settings: AppSettings): Promise<void> {
  const db = await getDb()
  await db.put('settings', { id: 'app', data: settings })
}