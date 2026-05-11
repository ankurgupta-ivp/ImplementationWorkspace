import React, { useState, useMemo } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, Modal, Field, Input, Select, Textarea, EmptyState, useConfirm } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import * as XLSX from 'xlsx';

const STATUSES = ADMIN_LISTS.statuses;

function TaskRow({ task, idx, onEdit }) {
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
  const { activeProject, tasks, updateSingleTask, updateTasks, showToast } = useApp();
  const { dialog } = useConfirm();
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

  // ── Export ────────────────────────────────────────────────────
  const exportXlsx = () => {
    const headers = ['#','Phase','Item','Task','Type','Tags','Responsible','Owner','Reviewer',
      'Owner Status','Reviewer Status','Expected Start','Expected End',
      'Actual Start','Actual End','Expected Effort (h)','Actual Effort (h)','Comments'];
    const rows = [headers];
    tasks.forEach((t, i) => rows.push([
      i + 1, t.phase, t.item, t.task, t.taskType, t.tags, t.responsible, t.owner, t.reviewer,
      t.ownerStatus, t.reviewerStatus, t.expectedStart, t.expectedEnd,
      t.actualStart, t.actualEnd, t.expectedEffort, t.actualEffort, t.comments,
    ]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [4,18,22,60,16,8,28,28,22,14,14,13,13,13,13,10,10,40].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  // ── Import ────────────────────────────────────────────────────
  // Matches rows back to existing tasks by Phase+Item+Task text (columns B+C+D).
  // Only updates status, dates, effort and comments — never replaces task structure.
  const importXlsx = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Build lookup: "phase||item||task" → row data (cols 9-17, 0-indexed)
        const lookup = {};
        data.slice(1).forEach(row => {
          const key = `${row[1]||''}||${row[2]||''}||${row[3]||''}`;
          lookup[key] = row;
        });

        let matched = 0;
        const updated = tasks.map(t => {
          const key = `${t.phase}||${t.item}||${t.task}`;
          const row = lookup[key];
          if (!row) return t;
          matched++;
          return {
            ...t,
            ownerStatus:    row[9]  || t.ownerStatus,
            reviewerStatus: row[10] || t.reviewerStatus,
            expectedStart:  row[11] || t.expectedStart,
            expectedEnd:    row[12] || t.expectedEnd,
            actualStart:    row[13] || t.actualStart,
            actualEnd:      row[14] || t.actualEnd,
            expectedEffort: row[15] != null ? String(row[15]) : t.expectedEffort,
            actualEffort:   row[16] != null ? String(row[16]) : t.actualEffort,
            comments:       row[17] || t.comments,
          };
        });

        updateTasks(updated);
        showToast(`Imported updates for ${matched} of ${tasks.length} tasks`);
      } catch {
        showToast('Import failed — check the file format', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      {dialog}
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Tasks & Checklist"
        subtitle={`${tasks.length} tasks · ${done} done · ${inprog} in progress · ${blocked} blocked`}
        actions={
          <div style={{ display: 'flex', gap: 8 }}>
            <label className="btn btn-ghost" style={{ cursor: 'pointer', margin: 0 }}>
              ↑ Import Excel
              <input type="file" accept=".xlsx,.xls" onChange={importXlsx} style={{ display: 'none' }} />
            </label>
            <button className="btn btn-ghost" onClick={exportXlsx}>↓ Export Excel</button>
          </div>
        }
      />

      {/* Import instructions banner */}
      <div style={{ background: '#e3f2fd', borderBottom: '1px solid #bbdefb', padding: '7px 20px', fontSize: 12, color: '#1565c0', flexShrink: 0 }}>
        <strong>Import tip:</strong> Export first, update the <strong>Status, Dates, Effort</strong> and <strong>Comments</strong> columns in Excel, then re-import. Tasks are matched by Phase + Item + Task text.
      </div>

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
