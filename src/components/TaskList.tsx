import type { Task } from '../types/task'

interface TaskListProps {
  tasks: Task[]
  search: string
  filter: 'all' | 'open' | 'completed' | 'snoozed'
  onComplete: (task: Task) => void
  onSnooze: (task: Task, days: number) => void
  onDelete: (id: string) => void
}

export function TaskList({
  tasks,
  search,
  filter,
  onComplete,
  onSnooze,
  onDelete,
}: TaskListProps) {
  const filtered = tasks.filter((t) => {
    if (filter !== 'all' && t.status !== filter) return false
    if (!search) return true
    const q = search.toLowerCase()
    return (
      t.title.toLowerCase().includes(q) ||
      t.tags.some((tag) => tag.toLowerCase().includes(q)) ||
      (t.notes?.toLowerCase().includes(q) ?? false)
    )
  })

  if (filtered.length === 0) {
    return (
      <div className="empty-state">
        <p>No tasks yet.</p>
        <p className="hint">Paste text above or click <strong>Sweep</strong> to gather tasks.</p>
      </div>
    )
  }

  return (
    <ul className="task-list">
      {filtered.map((task) => (
        <li key={task.id} className={`task-card status-${task.status}`}>
          <div className="task-main">
            <span className={`priority priority-${task.priority}`} title={task.priority}>
              {priorityIcon(task.priority)}
            </span>
            <div className="task-body">
              <h3>{task.title}</h3>
              <div className="task-meta">
                {task.dueDate && <span className="due">Due {task.dueDate}</span>}
                <span className="source">{task.source}</span>
                {task.tags.map((tag) => (
                  <span key={tag} className="tag">{tag}</span>
                ))}
              </div>
              {task.notes && <p className="notes">{task.notes}</p>}
              {task.sourceUrl && (
                <a href={task.sourceUrl} target="_blank" rel="noreferrer" className="source-link">
                  Open source
                </a>
              )}
            </div>
          </div>
          {task.status === 'open' && (
            <div className="task-actions">
              <button type="button" onClick={() => onComplete(task)} title="Mark complete">
                Done
              </button>
              <button type="button" onClick={() => onSnooze(task, 1)} title="Snooze 1 day">
                +1d
              </button>
              <button type="button" onClick={() => onSnooze(task, 7)} title="Snooze 1 week">
                +1w
              </button>
              <button type="button" className="danger" onClick={() => onDelete(task.id)}>
                Delete
              </button>
            </div>
          )}
          {task.status === 'completed' && (
            <div className="done-row">
              <span className="done-badge">Completed</span>
              {task.syncStatus && (
                <span className={`sync-badge sync-${task.syncStatus}`} title={task.syncMessage}>
                  {syncLabel(task.syncStatus)}
                </span>
              )}
            </div>
          )}
          {task.status === 'snoozed' && task.snoozedUntil && (
            <span className="snooze-badge">
              Snoozed until {new Date(task.snoozedUntil).toLocaleDateString()}
            </span>
          )}
        </li>
      ))}
    </ul>
  )
}

function priorityIcon(p: Task['priority']): string {
  const map = { urgent: '!!!', high: '!!', normal: '•', low: '·' }
  return map[p]
}

function syncLabel(status: NonNullable<Task['syncStatus']>): string {
  const map = {
    synced: 'Synced to M365',
    failed: 'Sync failed',
    skipped: 'Local only',
    local: 'Hub only',
  }
  return map[status]
}