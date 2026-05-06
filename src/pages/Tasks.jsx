import React, { useState, useMemo } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, Modal, Field, Input, Select, Textarea, EmptyState, useConfirm } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import * as XLSX from 'xlsx';

const STATUSES = ADMIN_LISTS.statuses;
const TASK_TYPES = ['BA Checklist Item', 'Dev Task'];

function TaskRow({ task, idx, onEdit }) {
  const statusColor = { 'Done': '#4caf50', 'In Progress': '#2278cf', 'Blocked': '#d32f2f', 'Delayed': '#d32f2f', 'Not Started': '#bbb' };
  return (
    <tr style={{ borderBottom: '1px solid #f0f0f0', background: idx % 2 === 0 ? '#fafafe' : '#fff' }}
      onMouseEnter={e => e.currentTarget.style.background = '#f0f4ff'}
      onMouseLeave={e => e.currentTarget.style.background = idx % 2 === 0 ? '#fafafe' : '#fff'}
    >
      <td style={{ padding: '5px 8px', fontSize: 11, color: '#888' }}>{idx + 1}</td>
      <td style={{ padding: '5px 8px', fontSize: 12, color: '#404041', borderLeft: `3px solid ${task.taskType === 'Dev Task' ? '#da9b38' : '#2278cf'}` }}>
        <span style={{ display: 'inline-block', fontSize: 10, fontWeight: 500, padding: '2px 6px', borderRadius: 3, textTransform: 'uppercase', background: task.taskType === 'Dev Task' ? '#fff3e0' : '#e3f2fd', color: task.taskType === 'Dev Task' ? '#e65100' : '#1565c0', marginRight: 5 }}>{task.taskType}</span>
        {task.task}
      </td>
      <td style={{ padding: '5px 8px', fontSize: 12, color: '#666' }}>{task.phase}</td>
      <td style={{ padding: '5px 8px', fontSize: 12, color: '#666' }}>{task.item}</td>
      <td style={{ padding: '5px 8px' }}><StatusBadge status={task.ownerStatus || 'Not Started'} /></td>
      <td style={{ padding: '5px 8px', fontSize: 12, color: '#666' }}>{task.responsible}</td>
      <td style={{ padding: '5px 8px', fontSize: 12, color: '#666' }}>{task.expectedEffort ? `${task.expectedEffort}h` : '—'}</td>
      <td style={{ padding: '5px 8px' }}>
        <button className="btn btn-ghost btn-xs" onClick={() => onEdit(task)}>✎ Edit</button>
      </td>
    </tr>
  );
}

export default function Tasks() {
  const { activeProject, tasks, updateSingleTask, showToast } = useApp();
  const { ask, dialog } = useConfirm();
  const [filterPhase, setFilterPhase] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterText, setFilterText] = useState('');
  const [editTask, setEditTask] = useState(null);
  const [editForm, setEditForm] = useState({});

  const filtered = useMemo(() => tasks.filter(t =>
    (!filterPhase || t.phase === filterPhase) &&
    (!filterStatus || t.ownerStatus === filterStatus) &&
    (!filterText || t.task?.toLowerCase().includes(filterText.toLowerCase()) || t.item?.toLowerCase().includes(filterText.toLowerCase()))
  ), [tasks, filterPhase, filterStatus, filterText]);

  if (!activeProject) return <EmptyState message="No project selected." />;

  const done = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => t.ownerStatus === 'Blocked').length;

  const openEdit = (task) => { setEditTask(task); setEditForm({ ...task }); };

  const saveEdit = async () => {
    await updateSingleTask(editTask.id, editForm);
    setEditTask(null);
    showToast('Task updated');
  };

  const exportXlsx = () => {
    const rows = [['#', 'Phase', 'Item', 'Task', 'Type', 'Responsible', 'Owner Status', 'Reviewer Status', 'Expected Effort', 'Actual Effort', 'Comments']];
    tasks.forEach((t, i) => rows.push([i + 1, t.phase, t.item, t.task, t.taskType, t.responsible, t.ownerStatus, t.reviewerStatus, t.expectedEffort, t.actualEffort, t.comments]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      {dialog}
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Tasks & Checklist"
        subtitle={`${tasks.length} tasks · ${done} done · ${inprog} in progress · ${blocked} blocked`}
        actions={<button className="btn btn-ghost" onClick={exportXlsx}>↓ Export Excel</button>}
      />

      {/* Toolbar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '8px 20px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', flexShrink: 0 }}>
        <input
          placeholder="Search tasks…"
          value={filterText}
          onChange={e => setFilterText(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, width: 200, fontFamily: 'Roboto,sans-serif' }}
        />
        <select value={filterPhase} onChange={e => setFilterPhase(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Phases</option>
          {ADMIN_LISTS.phases.map(p => <option key={p}>{p}</option>)}
        </select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Statuses</option>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </select>
        <span style={{ fontSize: 12, color: '#888', marginLeft: 'auto' }}>{filtered.length} of {tasks.length} shown</span>
      </div>

      {/* Table */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0' }}>
              {['#', 'Task', 'Phase', 'Item', 'Status', 'Responsible', 'Effort', ''].map(h => (
                <th key={h} style={{ padding: '9px 8px', textAlign: 'left', fontSize: 12, fontWeight: 500, color: '#444', whiteSpace: 'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => <TaskRow key={t.id} task={t} idx={i} onEdit={openEdit} />)}
          </tbody>
        </table>
        {filtered.length === 0 && (
          <div style={{ textAlign: 'center', padding: 40, color: '#999' }}>No tasks match the current filters.</div>
        )}
      </div>

      {/* Footer */}
      <div style={{ height: 26, background: '#f5f5f5', borderTop: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 14px', fontSize: 11, color: '#888', flexShrink: 0 }}>
        <span>{tasks.length} total tasks</span>
        <span>Done: {done} · In Progress: {inprog} · Blocked: {blocked}</span>
      </div>

      {/* Edit Modal */}
      {editTask && (
        <Modal title="Edit Task" onClose={() => setEditTask(null)} onSave={saveEdit} wide>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
            <Field label="Task"><Input value={editForm.task || ''} onChange={e => setEditForm(p => ({ ...p, task: e.target.value }))} style={{ gridColumn: '1/-1' }} /></Field>
            <Field label="Owner Status">
              <Select value={editForm.ownerStatus || ''} onChange={e => setEditForm(p => ({ ...p, ownerStatus: e.target.value }))}>
                {STATUSES.map(s => <option key={s}>{s}</option>)}
              </Select>
            </Field>
            <Field label="Reviewer Status">
              <Select value={editForm.reviewerStatus || ''} onChange={e => setEditForm(p => ({ ...p, reviewerStatus: e.target.value }))}>
                {STATUSES.map(s => <option key={s}>{s}</option>)}
              </Select>
            </Field>
            <Field label="Expected Start"><Input type="date" value={editForm.expectedStart || ''} onChange={e => setEditForm(p => ({ ...p, expectedStart: e.target.value }))} /></Field>
            <Field label="Expected End"><Input type="date" value={editForm.expectedEnd || ''} onChange={e => setEditForm(p => ({ ...p, expectedEnd: e.target.value }))} /></Field>
            <Field label="Actual Start"><Input type="date" value={editForm.actualStart || ''} onChange={e => setEditForm(p => ({ ...p, actualStart: e.target.value }))} /></Field>
            <Field label="Actual End"><Input type="date" value={editForm.actualEnd || ''} onChange={e => setEditForm(p => ({ ...p, actualEnd: e.target.value }))} /></Field>
            <Field label="Expected Effort (h)"><Input type="number" value={editForm.expectedEffort || ''} onChange={e => setEditForm(p => ({ ...p, expectedEffort: e.target.value }))} /></Field>
            <Field label="Actual Effort (h)"><Input type="number" value={editForm.actualEffort || ''} onChange={e => setEditForm(p => ({ ...p, actualEffort: e.target.value }))} /></Field>
            <Field label="Responsible"><Input value={editForm.responsible || ''} onChange={e => setEditForm(p => ({ ...p, responsible: e.target.value }))} /></Field>
            <Field label="Comments" style={{ gridColumn: '1/-1' }}><Textarea value={editForm.comments || ''} onChange={e => setEditForm(p => ({ ...p, comments: e.target.value }))} /></Field>
          </div>
        </Modal>
      )}
    </div>
  );
}
