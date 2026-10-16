import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import './styles.css';

const API = '/api';
const STATUSES = ['TODO', 'IN_PROGRESS', 'DONE'];
const STATUS_LABEL = { TODO: 'To do', IN_PROGRESS: 'In progress', DONE: 'Done' };
const PRIORITY_LABEL = { LOW: 'Low', MEDIUM: 'Medium', HIGH: 'High' };
const PRIORITY_RANK = { HIGH: 0, MEDIUM: 1, LOW: 2 };

async function api(path, options = {}) {
  const res = await fetch(`${API}${path}`, {
    headers: options.body ? { 'Content-Type': 'application/json' } : undefined,
    ...options,
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    const detail = Array.isArray(body.detail) ? body.detail.map((d) => d.msg).join(', ') : body.detail;
    throw new Error(detail || `Request failed (${res.status})`);
  }
  return res.status === 204 ? null : res.json();
}

const fmtDate = (iso) =>
  new Date(iso).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });

function App() {
  const [tasks, setTasks] = useState([]);
  const [stats, setStats] = useState({ total: 0, todo: 0, inProgress: 0, done: 0 });
  const [status, setStatus] = useState('ALL');
  const [priority, setPriority] = useState('ALL');
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState('newest');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [editing, setEditing] = useState(null); // null | 'new' | task
  const [toast, setToast] = useState('');

  const notify = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(''), 2500);
  };

  const load = useCallback(async () => {
    try {
      const [list, summary] = await Promise.all([api('/tasks'), api('/tasks/stats')]);
      setTasks(list);
      setStats(summary);
      setError('');
    } catch (e) {
      setError(e.message === 'Failed to fetch' ? 'Cannot reach the TaskBoard API.' : e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const run = async (action, message) => {
    try {
      await action();
      await load();
      if (message) notify(message);
    } catch (e) {
      setError(e.message);
    }
  };

  const setTaskStatus = (task, next) =>
    run(() => api(`/tasks/${task.id}`, { method: 'PUT', body: JSON.stringify({ status: next }) }));

  const remove = (task) => {
    if (!window.confirm(`Delete "${task.title}"?`)) return;
    run(() => api(`/tasks/${task.id}`, { method: 'DELETE' }), 'Task deleted');
  };

  const save = async (data, id) => {
    const isNew = id == null;
    await api(isNew ? '/tasks' : `/tasks/${id}`, {
      method: isNew ? 'POST' : 'PUT',
      body: JSON.stringify(data),
    });
    await load();
    notify(isNew ? 'Task created' : 'Task updated');
  };

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    const list = tasks.filter(
      (t) =>
        (status === 'ALL' || t.status === status) &&
        (priority === 'ALL' || t.priority === priority) &&
        (!q || `${t.title} ${t.description} ${t.assignee}`.toLowerCase().includes(q)),
    );
    const sorters = {
      newest: (a, b) => b.id - a.id,
      oldest: (a, b) => a.id - b.id,
      priority: (a, b) => PRIORITY_RANK[a.priority] - PRIORITY_RANK[b.priority] || b.id - a.id,
      title: (a, b) => a.title.localeCompare(b.title),
    };
    return [...list].sort(sorters[sort]);
  }, [tasks, status, priority, query, sort]);

  const pct = stats.total ? Math.round((stats.done / stats.total) * 100) : 0;

  return (
    <div className="page">
      <header className="topbar">
        <div className="brand">
          <span className="logo">T</span>
          <div>
            <h1>TaskBoard</h1>
            <p>Plan, track and finish work</p>
          </div>
        </div>
        <button className="btn primary" onClick={() => setEditing('new')}>
          + New task
        </button>
      </header>

      {error && (
        <div className="alert" role="alert">
          <span>{error}</span>
          <button className="btn small" onClick={load}>Retry</button>
        </div>
      )}

      <section className="stats" aria-label="Summary">
        <Stat label="Total" value={stats.total} tone="blue" />
        <Stat label="To do" value={stats.todo} tone="amber" />
        <Stat label="In progress" value={stats.inProgress} tone="violet" />
        <Stat label="Done" value={stats.done} tone="green" />
        <div className="stat progress">
          <span className="stat-label">Completion</span>
          <strong>{pct}%</strong>
          <div className="bar"><i style={{ width: `${pct}%` }} /></div>
        </div>
      </section>

      <section className="toolbar">
        <input
          type="search"
          placeholder="Search title, description or owner…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          aria-label="Search tasks"
        />
        <div className="tabs" role="tablist">
          {['ALL', ...STATUSES].map((s) => (
            <button key={s} className={status === s ? 'on' : ''} onClick={() => setStatus(s)}>
              {s === 'ALL' ? 'All' : STATUS_LABEL[s]}
            </button>
          ))}
        </div>
        <select value={priority} onChange={(e) => setPriority(e.target.value)} aria-label="Priority filter">
          <option value="ALL">Any priority</option>
          {Object.entries(PRIORITY_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
        </select>
        <select value={sort} onChange={(e) => setSort(e.target.value)} aria-label="Sort">
          <option value="newest">Newest first</option>
          <option value="oldest">Oldest first</option>
          <option value="priority">Priority</option>
          <option value="title">Title A–Z</option>
        </select>
      </section>

      <section className="list">
        {loading ? (
          <div className="empty">Loading tasks…</div>
        ) : visible.length === 0 ? (
          <div className="empty">
            {tasks.length === 0 ? 'No tasks yet — create your first one.' : 'No tasks match these filters.'}
          </div>
        ) : (
          visible.map((t) => (
            <article className={`card ${t.status.toLowerCase()}`} key={t.id}>
              <div className="card-main">
                <h3>{t.title}</h3>
                {t.description && <p>{t.description}</p>}
                <div className="meta">
                  <span className={`badge ${t.priority.toLowerCase()}`}>{PRIORITY_LABEL[t.priority]}</span>
                  <span>👤 {t.assignee}</span>
                  <span>{fmtDate(t.created_at)}</span>
                </div>
              </div>
              <div className="card-actions">
                <select
                  className={`status ${t.status.toLowerCase()}`}
                  value={t.status}
                  onChange={(e) => setTaskStatus(t, e.target.value)}
                  aria-label={`Status of ${t.title}`}
                >
                  {STATUSES.map((s) => <option key={s} value={s}>{STATUS_LABEL[s]}</option>)}
                </select>
                <button className="btn small" onClick={() => setEditing(t)}>Edit</button>
                <button className="btn small danger" onClick={() => remove(t)}>Delete</button>
              </div>
            </article>
          ))
        )}
      </section>

      {editing && (
        <TaskForm
          task={editing === 'new' ? null : editing}
          onClose={() => setEditing(null)}
          onSave={save}
        />
      )}
      {toast && <div className="toast" role="status">{toast}</div>}
    </div>
  );
}

function Stat({ label, value, tone }) {
  return (
    <div className={`stat ${tone}`}>
      <span className="stat-label">{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function TaskForm({ task, onClose, onSave }) {
  const [form, setForm] = useState({
    title: task?.title ?? '',
    description: task?.description ?? '',
    priority: task?.priority ?? 'MEDIUM',
    status: task?.status ?? 'TODO',
    assignee: task?.assignee ?? '',
  });
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState('');
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  useEffect(() => {
    const onKey = (e) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  const submit = async (e) => {
    e.preventDefault();
    setSaving(true);
    setErr('');
    try {
      await onSave({ ...form, title: form.title.trim(), assignee: form.assignee.trim() || 'Unassigned' }, task?.id);
      onClose();
    } catch (ex) {
      setErr(ex.message);
      setSaving(false);
    }
  };

  return (
    <div className="backdrop" onMouseDown={onClose}>
      <form className="modal" onSubmit={submit} onMouseDown={(e) => e.stopPropagation()}>
        <h2>{task ? 'Edit task' : 'New task'}</h2>
        <label>Title
          <input value={form.title} onChange={set('title')} maxLength={200} required autoFocus />
        </label>
        <label>Description
          <textarea rows={3} value={form.description} onChange={set('description')} />
        </label>
        <div className="row">
          <label>Priority
            <select value={form.priority} onChange={set('priority')}>
              {Object.entries(PRIORITY_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
            </select>
          </label>
          <label>Status
            <select value={form.status} onChange={set('status')}>
              {STATUSES.map((s) => <option key={s} value={s}>{STATUS_LABEL[s]}</option>)}
            </select>
          </label>
        </div>
        <label>Assignee
          <input value={form.assignee} onChange={set('assignee')} placeholder="Unassigned" />
        </label>
        {err && <p className="form-error">{err}</p>}
        <div className="actions">
          <button type="button" className="btn" onClick={onClose}>Cancel</button>
          <button className="btn primary" disabled={saving || !form.title.trim()}>
            {saving ? 'Saving…' : task ? 'Save changes' : 'Create task'}
          </button>
        </div>
      </form>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
