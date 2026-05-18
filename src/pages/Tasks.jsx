import React, { useState, useMemo, useRef } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, Modal, Field, Input, Select, Textarea, EmptyState, useConfirm } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import * as XLSX from 'xlsx';

const STATUSES  = ADMIN_LISTS.statuses;
const PHASES    = ADMIN_LISTS.phases;
const RESP_LIST = ADMIN_LISTS.responsibles;
const TASK_TYPES = ['BA Checklist Item', 'Dev Task'];

const EMPTY_TASK = {
  phase: 'Project Creation', item: '', task: '', taskType: 'BA Checklist Item', tags: '',
  responsible: 'Implementation BA', owner: 'Implementation BA', reviewer: 'Implementation Lead',
  ownerStatus: 'Not Started', reviewerStatus: 'Not Started', comments: '',
  actualStart: '', actualEnd: '', expectedStart: '', expectedEnd: '',
  actualEffort: '', expectedEffort: '', actualElapsed: '', expectedElapsed: '',
};

// Column definitions — matches template library order exactly
const COLS = [
  { key: '_drag',          label: '',              width: 22,  filterable: false },
  { key: '_num',           label: '#',             width: 36,  filterable: false },
  { key: 'phase',          label: 'Phase',         width: 140, filterable: true,  type: 'select', options: PHASES },
  { key: 'item',           label: 'Item',          width: 140, filterable: true,  type: 'text' },
  { key: 'task',           label: 'Task',          width: 320, filterable: true,  type: 'text' },
  { key: 'taskType',       label: 'Type',          width: 100, filterable: true,  type: 'select', options: TASK_TYPES },
  { key: 'responsible',    label: 'Responsible',   width: 160, filterable: true,  type: 'select', options: RESP_LIST },
  { key: 'ownerStatus',    label: 'Owner Status',  width: 110, filterable: true,  type: 'select', options: STATUSES },
  { key: 'reviewerStatus', label: 'Rev. Status',   width: 110, filterable: true,  type: 'select', options: STATUSES },
  { key: 'expectedStart',  label: 'Exp. Start',    width: 96,  filterable: false },
  { key: 'expectedEnd',    label: 'Exp. End',      width: 96,  filterable: false },
  { key: 'actualStart',    label: 'Act. Start',    width: 96,  filterable: false },
  { key: 'actualEnd',      label: 'Act. End',      width: 96,  filterable: false },
  { key: 'expectedEffort', label: 'Exp. Effort',   width: 80,  filterable: false },
  { key: 'actualEffort',   label: 'Act. Effort',   width: 80,  filterable: false },
  { key: '_edit',          label: '',              width: 36,  filterable: false },
];

// Header-aware column map for import.
// Maps normalised Excel header text -> internal field key.
// Handles both the app's own export headers and the original template Excel headers.
const HEADER_MAP = {
  // app export headers
  'phase':                 'phase',
  'item':                  'item',
  'task':                  'task',
  'type':                  'taskType',
  'tags':                  'tags',
  'responsible':           'responsible',
  'owner':                 'owner',
  'reviewer':              'reviewer',
  'owner status':          'ownerStatus',
  'reviewer status':       'reviewerStatus',
  'exp. start':            'expectedStart',
  'exp. end':              'expectedEnd',
  'act. start':            'actualStart',
  'act. end':              'actualEnd',
  'exp. effort (h)':       'expectedEffort',
  'act. effort (h)':       'actualEffort',
  'comments':              'comments',
  // original template Excel headers
  'task type':             'taskType',
  'actual start date':     'actualStart',
  'actual end date':       'actualEnd',
  'expected start date':   'expectedStart',
  'expected end date':     'expectedEnd',
  'actual effort spent':   'actualEffort',
  'expected effort spent': 'expectedEffort',
};

function TypeBadge({ type }) {
  const isDev = type === 'Dev Task';
  return (
    <span style={{
      display: 'inline-block', fontSize: 10, fontWeight: 500,
      padding: '2px 6px', borderRadius: 3, whiteSpace: 'nowrap',
      background: isDev ? '#fff3e0' : '#e3f2fd',
      color:      isDev ? '#e65100' : '#1565c0',
    }}>
      {isDev ? 'Dev' : 'BA'}
    </span>
  );
}

function TaskRow({ task, idx, onEdit, onDragStart, onDragOver, onDrop, onDragEnd, isDragging }) {
  const base = isDragging ? '#eef0ff' : idx % 2 === 0 ? '#fafafe' : '#fff';
  return (
    <tr
      draggable
      onDragStart={() => onDragStart(idx)}
      onDragOver={e => { e.preventDefault(); onDragOver(idx); }}
      onDrop={() => onDrop(idx)}
      onDragEnd={onDragEnd}
      style={{ borderBottom: '1px solid #f0f0f0', background: base, opacity: isDragging ? 0.5 : 1 }}
      onMouseEnter={e => { if (!isDragging) e.currentTarget.style.background = '#f0f4ff'; }}
      onMouseLeave={e => { if (!isDragging) e.currentTarget.style.background = base; }}
    >
      <td style={{ padding: '5px 4px', fontSize: 11, color: '#ccc', cursor: 'grab', userSelect: 'none', textAlign: 'center' }}>⠿</td>
      <td style={{ padding: '5px 4px', fontSize: 11, color: '#aaa', textAlign: 'right' }}>{idx + 1}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#555', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{task.phase}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#555', overflow: 'hidden', textOverflow: 'ellipsis' }}>{task.item}</td>
      <td style={{ padding: '5px 6px', fontSize: 12, color: '#404041', borderLeft: `3px solid ${task.taskType === 'Dev Task' ? '#da9b38' : '#2278cf'}`, overflow: 'hidden', textOverflow: 'ellipsis' }}>{task.task}</td>
      <td style={{ padding: '5px 6px' }}><TypeBadge type={task.taskType} /></td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#555', overflow: 'hidden', textOverflow: 'ellipsis' }}>{task.responsible}</td>
      <td style={{ padding: '5px 6px' }}><StatusBadge status={task.ownerStatus || 'Not Started'} /></td>
      <td style={{ padding: '5px 6px' }}><StatusBadge status={task.reviewerStatus || 'Not Started'} /></td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', whiteSpace: 'nowrap' }}>{task.expectedStart || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', whiteSpace: 'nowrap' }}>{task.expectedEnd   || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', whiteSpace: 'nowrap' }}>{task.actualStart   || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', whiteSpace: 'nowrap' }}>{task.actualEnd     || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', textAlign: 'right' }}>{task.expectedEffort ? `${task.expectedEffort}h` : '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#777', textAlign: 'right' }}>{task.actualEffort   ? `${task.actualEffort}h`  : '—'}</td>
      <td style={{ padding: '5px 4px', textAlign: 'center' }}>
        <button className="btn btn-ghost btn-xs" onClick={() => onEdit(task)}>✎</button>
      </td>
    </tr>
  );
}

function ColFilter({ col, value, onChange }) {
  if (!col.filterable) {
    return <th style={{ padding: '2px 4px', background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', borderRight: '1px solid #d8d8e8' }} />;
  }
  const base = {
    width: '100%', fontSize: 11, border: '1px solid #c8c8d8', borderRadius: 3,
    padding: '3px 5px', fontFamily: 'Roboto,sans-serif', background: '#fff',
    boxSizing: 'border-box', color: '#333',
  };
  if (col.type === 'select') {
    return (
      <th style={{ padding: '2px 4px', background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', borderRight: '1px solid #d8d8e8' }}>
        <select value={value} onChange={e => onChange(e.target.value)} style={base}>
          <option value="">All</option>
          {col.options.map(o => <option key={o}>{o}</option>)}
        </select>
      </th>
    );
  }
  return (
    <th style={{ padding: '2px 4px', background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', borderRight: '1px solid #d8d8e8' }}>
      <input value={value} onChange={e => onChange(e.target.value)} placeholder="Filter…" style={base} />
    </th>
  );
}

export default function Tasks() {
  const { activeProject, tasks, updateTasks, updateSingleTask, showToast } = useApp();
  const { dialog } = useConfirm();

  const [colFilters, setColFilters] = useState({});
  const setFilter = (key, val) => setColFilters(prev => ({ ...prev, [key]: val }));

  const [editTask, setEditTask] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [isNew,    setIsNew]    = useState(false);
  const [newForm,  setNewForm]  = useState({ ...EMPTY_TASK });

  const dragIdx = useRef(null);
  const [dragOver, setDragOver] = useState(null);

  const filtered = useMemo(() => tasks.filter(t => {
    for (const col of COLS) {
      const val = (colFilters[col.key] || '').toLowerCase().trim();
      if (!val) continue;
      const cell = (t[col.key] || '').toString().toLowerCase();
      if (!cell.includes(val)) return false;
    }
    return true;
  }), [tasks, colFilters]);

  const hasFilters = Object.values(colFilters).some(v => v);
  const canDrag    = !hasFilters;

  if (!activeProject) return <EmptyState message="No project selected." />;

  const done    = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog  = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => ['Blocked', 'Delayed'].includes(t.ownerStatus)).length;

  const openEdit = (task) => { setIsNew(false); setEditTask(task); setEditForm({ ...task }); };
  const openNew  = () => { setNewForm({ ...EMPTY_TASK }); setIsNew(true); setEditTask({}); };

  const saveEdit = async () => {
    if (isNew) {
      await updateTasks([...tasks, { ...newForm, id: Date.now(), sortOrder: tasks.length }]);
      setNewForm({ ...EMPTY_TASK });
      showToast('Task added');
    } else {
      await updateSingleTask(editTask.id, editForm);
      showToast('Task updated');
    }
    setEditTask(null); setIsNew(false);
  };

  const handleDragStart = (idx) => { dragIdx.current = idx; };
  const handleDragOver  = (idx) => { setDragOver(idx); };
  const handleDrop      = async (dropIdx) => {
    if (dragIdx.current === null || dragIdx.current === dropIdx) return;
    const reordered = [...tasks];
    const [moved] = reordered.splice(dragIdx.current, 1);
    reordered.splice(dropIdx, 0, moved);
    await updateTasks(reordered);
    dragIdx.current = null; setDragOver(null);
    showToast('Task order updated');
  };
  const handleDragEnd = () => { dragIdx.current = null; setDragOver(null); };

  // Export — column order is canonical; import reads by these header names
  const exportXlsx = () => {
    const headers = ['#', 'Phase', 'Item', 'Task', 'Type', 'Tags', 'Responsible', 'Owner', 'Reviewer',
      'Owner Status', 'Reviewer Status', 'Exp. Start', 'Exp. End', 'Act. Start', 'Act. End',
      'Exp. Effort (h)', 'Act. Effort (h)', 'Comments'];
    const rows = [headers];
    tasks.forEach((t, i) => rows.push([
      i + 1, t.phase, t.item, t.task, t.taskType, t.tags,
      t.responsible, t.owner, t.reviewer,
      t.ownerStatus, t.reviewerStatus,
      t.expectedStart, t.expectedEnd, t.actualStart, t.actualEnd,
      t.expectedEffort, t.actualEffort, t.comments,
    ]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [4, 18, 22, 60, 14, 8, 28, 28, 22, 14, 14, 11, 11, 11, 11, 10, 10, 40].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  // Import — header-aware; handles app export AND original template Excel formats.
  // Excel is source of truth: new rows inserted, deleted rows dropped.
  // ONLY writes to this project's Supabase tasks rows — template library untouched.
  const importXlsx = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (data.length < 2) { showToast('File has no data rows', 'error'); e.target.value = ''; return; }

        // Map column index -> field key from header row
        const headerRow = data[0];
        const colIdx = {};
        headerRow.forEach((h, i) => {
          const norm  = String(h).toLowerCase().trim();
          const field = HEADER_MAP[norm];
          if (field !== undefined && field !== null) colIdx[i] = field;
        });

        if (!Object.values(colIdx).includes('phase') || !Object.values(colIdx).includes('task')) {
          showToast('Could not find Phase or Task columns — check the file format', 'error');
          e.target.value = ''; return;
        }

        const dataRows = data.slice(1).filter(row => row.some(c => c !== '' && c != null));
        if (dataRows.length === 0) { showToast('No data rows found', 'error'); e.target.value = ''; return; }

        const existingByKey = {};
        tasks.forEach(t => { existingByKey[`${t.phase}||${t.item}||${t.task}`] = t; });

        const str = (v) => (v == null ? '' : String(v)).trim();
        const dateStr = (v) => {
          if (!v && v !== 0) return '';
          if (v instanceof Date) {
            const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
            return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
          }
          return str(v);
        };
        const DATE_FIELDS = new Set(['expectedStart', 'expectedEnd', 'actualStart', 'actualEnd']);
        const NUM_FIELDS  = new Set(['expectedEffort', 'actualEffort']);

        let matched = 0, added = 0;

        const importedTasks = dataRows.map((row, i) => {
          const raw = {};
          Object.entries(colIdx).forEach(([idx, field]) => { raw[field] = row[idx]; });

          const phase = str(raw.phase || '');
          const item  = str(raw.item  || '');
          const task  = str(raw.task  || '');
          const ex    = existingByKey[`${phase}||${item}||${task}`];
          if (ex) matched++; else added++;

          const resolve = (field, fallback) => {
            const v = raw[field];
            if (DATE_FIELDS.has(field)) return dateStr(v);
            if (NUM_FIELDS.has(field))  return (v !== '' && v != null) ? str(v) : (ex?.[field] || '');
            return (v !== '' && v != null) ? str(v) : (ex?.[field] || fallback || '');
          };

          return {
            id:              ex ? ex.id : undefined,
            sortOrder:       i,
            phase, item, task,
            taskType:        resolve('taskType',       'BA Checklist Item'),
            tags:            resolve('tags',           ''),
            responsible:     resolve('responsible',    ''),
            owner:           resolve('owner',          ''),
            reviewer:        resolve('reviewer',       ''),
            ownerStatus:     resolve('ownerStatus',    'Not Started'),
            reviewerStatus:  resolve('reviewerStatus', 'Not Started'),
            expectedStart:   dateStr(raw.expectedStart),
            expectedEnd:     dateStr(raw.expectedEnd),
            actualStart:     dateStr(raw.actualStart),
            actualEnd:       dateStr(raw.actualEnd),
            expectedEffort:  resolve('expectedEffort', ''),
            actualEffort:    resolve('actualEffort',   ''),
            actualElapsed:   ex?.actualElapsed   || '',
            expectedElapsed: ex?.expectedElapsed || '',
            comments:        resolve('comments',       ''),
          };
        });

        const removed = tasks.length - matched;
        updateTasks(importedTasks); // scoped to activeProjectId — template library untouched
        showToast(
          `Imported ${importedTasks.length} tasks · ${matched} updated · ${added} new` +
          (removed > 0 ? ` · ${removed} removed` : '')
        );
      } catch (err) {
        console.error('Import error:', err);
        showToast('Import failed — check file format', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const formFields = (form, setForm) => (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
      <Field label="Phase">
        <Select value={form.phase || ''} onChange={e => setForm(p => ({ ...p, phase: e.target.value }))}>
          {PHASES.map(p => <option key={p}>{p}</option>)}
        </Select>
      </Field>
      <Field label="Task Type">
        <Select value={form.taskType || ''} onChange={e => setForm(p => ({ ...p, taskType: e.target.value }))}>
          {TASK_TYPES.map(t => <option key={t}>{t}</option>)}
        </Select>
      </Field>
      <Field label="Item">
        <Input value={form.item || ''} onChange={e => setForm(p => ({ ...p, item: e.target.value }))} placeholder="e.g. BRD Preparation" />
      </Field>
      <Field label="Responsible">
        <Select value={form.responsible || ''} onChange={e => setForm(p => ({ ...p, responsible: e.target.value }))}>
          {RESP_LIST.map(r => <option key={r}>{r}</option>)}
        </Select>
      </Field>
      <Field label="Task Description" style={{ gridColumn: '1/-1' }}>
        <Textarea value={form.task || ''} onChange={e => setForm(p => ({ ...p, task: e.target.value }))} placeholder="Describe the task…" style={{ minHeight: 70 }} />
      </Field>
      <Field label="Owner">
        <Input value={form.owner || ''} onChange={e => setForm(p => ({ ...p, owner: e.target.value }))} />
      </Field>
      <Field label="Reviewer">
        <Input value={form.reviewer || ''} onChange={e => setForm(p => ({ ...p, reviewer: e.target.value }))} />
      </Field>
      <Field label="Owner Status">
        <Select value={form.ownerStatus || ''} onChange={e => setForm(p => ({ ...p, ownerStatus: e.target.value }))}>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Reviewer Status">
        <Select value={form.reviewerStatus || ''} onChange={e => setForm(p => ({ ...p, reviewerStatus: e.target.value }))}>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Expected Start"><Input type="date" value={form.expectedStart || ''} onChange={e => setForm(p => ({ ...p, expectedStart: e.target.value }))} /></Field>
      <Field label="Expected End">  <Input type="date" value={form.expectedEnd   || ''} onChange={e => setForm(p => ({ ...p, expectedEnd:   e.target.value }))} /></Field>
      <Field label="Actual Start">  <Input type="date" value={form.actualStart   || ''} onChange={e => setForm(p => ({ ...p, actualStart:   e.target.value }))} /></Field>
      <Field label="Actual End">    <Input type="date" value={form.actualEnd     || ''} onChange={e => setForm(p => ({ ...p, actualEnd:     e.target.value }))} /></Field>
      <Field label="Expected Effort (h)"><Input type="number" value={form.expectedEffort || ''} onChange={e => setForm(p => ({ ...p, expectedEffort: e.target.value }))} /></Field>
      <Field label="Actual Effort (h)">  <Input type="number" value={form.actualEffort   || ''} onChange={e => setForm(p => ({ ...p, actualEffort:   e.target.value }))} /></Field>
      <Field label="Tags"><Input value={form.tags || ''} onChange={e => setForm(p => ({ ...p, tags: e.target.value }))} /></Field>
      <Field label="Comments" style={{ gridColumn: '1/-1' }}>
        <Textarea value={form.comments || ''} onChange={e => setForm(p => ({ ...p, comments: e.target.value }))} />
      </Field>
    </div>
  );

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
            <button className="btn btn-primary" onClick={openNew}>+ New Task</button>
          </div>
        }
      />

      {/* Summary bar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '6px 20px', display: 'flex', alignItems: 'center', gap: 16, flexShrink: 0 }}>
        <span style={{ fontSize: 12, color: '#888' }}>{filtered.length} of {tasks.length} shown</span>
        {hasFilters && (
          <button onClick={() => setColFilters({})}
            style={{ fontSize: 11, color: '#404789', background: 'none', border: '1px solid #b0b4d4', borderRadius: 3, padding: '2px 8px', cursor: 'pointer' }}>
            ✕ Clear all filters
          </button>
        )}
        <span style={{ marginLeft: 'auto', fontSize: 11, color: canDrag ? '#9799b1' : '#f57f17' }}>
          {canDrag ? '⠿ Drag rows to reorder' : 'Clear filters to enable drag reorder'}
        </span>
      </div>

      {/* Table */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, tableLayout: 'fixed' }}>
          <colgroup>
            {COLS.map(c => <col key={c.key} style={{ width: c.width }} />)}
          </colgroup>
          <thead style={{ position: 'sticky', top: 0, zIndex: 3 }}>
            {/* Label row */}
            <tr style={{ background: '#f0f1f8', borderBottom: '1px solid #d0d0d0' }}>
              {COLS.map(c => (
                <th key={c.key} style={{
                  padding: '7px 6px', textAlign: 'left', fontSize: 11, fontWeight: 600,
                  color: '#404789', whiteSpace: 'nowrap', background: '#f0f1f8',
                  borderRight: '1px solid #e4e4ee',
                }}>{c.label}</th>
              ))}
            </tr>
            {/* Per-column filter row */}
            <tr>
              {COLS.map(c => (
                <ColFilter
                  key={c.key}
                  col={c}
                  value={colFilters[c.key] || ''}
                  onChange={val => setFilter(c.key, val)}
                />
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((t, i) => (
              <TaskRow
                key={t.id || i}
                task={t} idx={i}
                onEdit={openEdit}
                onDragStart={canDrag ? handleDragStart : () => {}}
                onDragOver={canDrag ? handleDragOver  : () => {}}
                onDrop={canDrag ? handleDrop         : () => {}}
                onDragEnd={canDrag ? handleDragEnd   : () => {}}
                isDragging={dragOver === i}
              />
            ))}
          </tbody>
        </table>
        {filtered.length === 0 && (
          <div style={{ textAlign: 'center', padding: 40, color: '#999', fontSize: 13 }}>
            {hasFilters ? 'No tasks match the current filters.' : 'No tasks yet.'}
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{ height: 26, background: '#f5f5f5', borderTop: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 14px', fontSize: 11, color: '#888', flexShrink: 0 }}>
        <span>{tasks.length} total tasks</span>
        <span>Done: {done} · In Progress: {inprog} · Blocked: {blocked}</span>
      </div>

      {/* Edit / New Task Modal */}
      {editTask !== null && (
        <Modal
          title={isNew ? 'Add New Task' : 'Edit Task'}
          onClose={() => { setEditTask(null); setIsNew(false); }}
          onSave={saveEdit}
          wide
        >
          {isNew ? formFields(newForm, setNewForm) : formFields(editForm, setEditForm)}
        </Modal>
      )}
    </div>
  );
}
