import React, { useState, useMemo, useRef } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, Modal, Field, Input, Select, Textarea, EmptyState, useConfirm } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import * as XLSX from 'xlsx';

const STATUSES  = ADMIN_LISTS.statuses;
const PHASES    = ADMIN_LISTS.phases;
const RESP_LIST = ADMIN_LISTS.responsibles;

const EMPTY_TASK = {
  phase: 'Project Creation', item: '', task: '', taskType: 'BA Checklist Item', tags: '',
  responsible: 'Implementation BA', owner: 'Implementation BA', reviewer: 'Implementation Lead',
  ownerStatus: 'Not Started', reviewerStatus: 'Not Started', comments: '',
  actualStart: '', actualEnd: '', expectedStart: '', expectedEnd: '',
  actualEffort: '', expectedEffort: '', actualElapsed: '', expectedElapsed: '',
};

// ── Drag-and-drop row ─────────────────────────────────────────
function TaskRow({ task, idx, onEdit, onDragStart, onDragOver, onDrop, onDragEnd, isDragging }) {
  return (
    <tr
      draggable
      onDragStart={() => onDragStart(idx)}
      onDragOver={e => { e.preventDefault(); onDragOver(idx); }}
      onDrop={() => onDrop(idx)}
      onDragEnd={onDragEnd}
      style={{
        borderBottom: '1px solid #f0f0f0',
        background: isDragging ? '#eef0ff' : idx % 2 === 0 ? '#fafafe' : '#fff',
        opacity: isDragging ? 0.5 : 1,
        cursor: 'grab',
      }}
      onMouseEnter={e => { if (!isDragging) e.currentTarget.style.background = '#f0f4ff'; }}
      onMouseLeave={e => { if (!isDragging) e.currentTarget.style.background = idx % 2 === 0 ? '#fafafe' : '#fff'; }}
    >
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#ccc', cursor: 'grab', userSelect: 'none' }}>⠿</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#888' }}>{idx + 1}</td>
      <td style={{ padding: '5px 6px', fontSize: 12, color: '#404041', borderLeft: `3px solid ${task.taskType === 'Dev Task' ? '#da9b38' : '#2278cf'}`, maxWidth: 340 }}>
        <span style={{ display: 'inline-block', fontSize: 10, fontWeight: 500, padding: '2px 5px', borderRadius: 3, textTransform: 'uppercase', background: task.taskType === 'Dev Task' ? '#fff3e0' : '#e3f2fd', color: task.taskType === 'Dev Task' ? '#e65100' : '#1565c0', marginRight: 5, whiteSpace: 'nowrap' }}>{task.taskType === 'Dev Task' ? 'Dev' : 'BA'}</span>
        {task.task}
      </td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.phase}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666' }}>{task.item}</td>
      <td style={{ padding: '5px 6px' }}><StatusBadge status={task.ownerStatus || 'Not Started'} /></td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.expectedStart || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.expectedEnd   || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.actualStart   || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.actualEnd     || '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.expectedEffort ? `${task.expectedEffort}h` : '—'}</td>
      <td style={{ padding: '5px 6px', fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{task.actualEffort  ? `${task.actualEffort}h`  : '—'}</td>
      <td style={{ padding: '5px 6px' }}>
        <button className="btn btn-ghost btn-xs" onClick={() => onEdit(task)}>✎</button>
      </td>
    </tr>
  );
}

export default function Tasks() {
  const { activeProject, tasks, updateTasks, updateSingleTask, showToast } = useApp();
  const { dialog } = useConfirm();
  const [filterPhase,  setFilterPhase]  = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterText,   setFilterText]   = useState('');
  const [editTask,  setEditTask]  = useState(null);
  const [editForm,  setEditForm]  = useState({});
  const [isNew,     setIsNew]     = useState(false);
  const [newForm,   setNewForm]   = useState({ ...EMPTY_TASK });

  // Drag state
  const dragIdx = useRef(null);
  const [dragOver, setDragOver] = useState(null);

  const filtered = useMemo(() => tasks.filter(t =>
    (!filterPhase  || t.phase === filterPhase) &&
    (!filterStatus || t.ownerStatus === filterStatus) &&
    (!filterText   || t.task?.toLowerCase().includes(filterText.toLowerCase()) || t.item?.toLowerCase().includes(filterText.toLowerCase()))
  ), [tasks, filterPhase, filterStatus, filterText]);

  if (!activeProject) return <EmptyState message="No project selected." />;

  const done    = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog  = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => ['Blocked','Delayed'].includes(t.ownerStatus)).length;

  // ── Edit existing ─────────────────────────────────────────────
  const openEdit = (task) => { setIsNew(false); setEditTask(task); setEditForm({ ...task }); };

  const saveEdit = async () => {
    if (isNew) {
      const updated = [...tasks, { ...newForm, id: Date.now(), sortOrder: tasks.length }];
      await updateTasks(updated);
      setNewForm({ ...EMPTY_TASK });
      showToast('Task added');
    } else {
      await updateSingleTask(editTask.id, editForm);
      setEditTask(null);
      showToast('Task updated');
    }
  };

  // ── Add new ───────────────────────────────────────────────────
  const openNew = () => {
    setNewForm({ ...EMPTY_TASK });
    setIsNew(true);
    setEditTask({});
  };

  // ── Drag and drop reorder ─────────────────────────────────────
  // Only works when no filters active (reorder on full list)
  const canDrag = !filterPhase && !filterStatus && !filterText;

  const handleDragStart = (idx) => { dragIdx.current = idx; };
  const handleDragOver  = (idx) => { setDragOver(idx); };
  const handleDrop      = async (dropIdx) => {
    if (dragIdx.current === null || dragIdx.current === dropIdx) return;
    const reordered = [...tasks];
    const [moved] = reordered.splice(dragIdx.current, 1);
    reordered.splice(dropIdx, 0, moved);
    await updateTasks(reordered);
    dragIdx.current = null;
    setDragOver(null);
    showToast('Task order updated');
  };
  const handleDragEnd = () => { dragIdx.current = null; setDragOver(null); };

  // ── Export ────────────────────────────────────────────────────
  const exportXlsx = () => {
    const headers = ['#','Phase','Item','Task','Type','Tags','Responsible','Owner','Reviewer',
      'Owner Status','Reviewer Status','Exp. Start','Exp. End','Act. Start','Act. End',
      'Exp. Effort (h)','Act. Effort (h)','Comments'];
    const rows = [headers];
    tasks.forEach((t, i) => rows.push([
      i+1, t.phase, t.item, t.task, t.taskType, t.tags, t.responsible, t.owner, t.reviewer,
      t.ownerStatus, t.reviewerStatus, t.expectedStart, t.expectedEnd,
      t.actualStart, t.actualEnd, t.expectedEffort, t.actualEffort, t.comments,
    ]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [4,18,22,60,8,8,28,28,22,14,14,11,11,11,11,10,10,40].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  // ── Import ────────────────────────────────────────────────────
  // ── Import ────────────────────────────────────────────────────
  // Excel is the source of truth for this project's task list.
  // Strategy:
  //   1. Read every data row from the Excel (skip header, skip blank rows).
  //   2. Match each Excel row to an existing DB task by Phase+Item+Task text
  //      — if matched, reuse the DB id so Supabase updates in-place.
  //   3. New rows in Excel (no match) are inserted as fresh tasks.
  //   4. Rows deleted from Excel are dropped from the project.
  //   5. ONLY touches this project's rows via updateTasks() — the global
  //      `templates` table is never written, so the template library is safe.
  const importXlsx = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // header:1 → plain arrays; defval:'' so missing cells are '' not undefined
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        // Skip header row and any completely blank rows
        const dataRows = data.slice(1).filter(row =>
          row.some(cell => cell !== '' && cell != null)
        );

        if (dataRows.length === 0) {
          showToast('No data rows found in the file', 'error');
          e.target.value = '';
          return;
        }

        // Build lookup from existing tasks by Phase||Item||Task for id reuse
        const existingByKey = {};
        tasks.forEach(t => {
          const k = `${t.phase}||${t.item}||${t.task}`;
          existingByKey[k] = t;
        });

        // Normalise a cell to a clean string
        const str = (v) => (v == null ? '' : String(v)).trim();

        // Normalise date cells — XLSX can return JS Date objects (cellDates:true),
        // Excel serials (numbers), or ISO strings
        const dateStr = (v) => {
          if (!v && v !== 0) return '';
          if (v instanceof Date) {
            const y = v.getFullYear(), m = v.getMonth()+1, d = v.getDate();
            return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
          }
          return str(v);
        };

        let matched = 0, added = 0;

        // Export column layout (Tasks.jsx exportXlsx):
        // [0]#  [1]Phase  [2]Item  [3]Task  [4]Type  [5]Tags
        // [6]Responsible  [7]Owner  [8]Reviewer
        // [9]OwnerStatus  [10]ReviewerStatus
        // [11]ExpStart  [12]ExpEnd  [13]ActStart  [14]ActEnd
        // [15]ExpEffort  [16]ActEffort  [17]Comments
        const importedTasks = dataRows.map((row, i) => {
          const phase = str(row[1]);
          const item  = str(row[2]);
          const task  = str(row[3]);
          const key   = `${phase}||${item}||${task}`;
          const ex    = existingByKey[key];

          if (ex) matched++; else added++;

          return {
            id:              ex ? ex.id : undefined,   // reuse DB id to update; undefined = insert
            sortOrder:       i,
            phase,
            item,
            task,
            taskType:        str(row[4]) || ex?.taskType        || 'BA Checklist Item',
            tags:            str(row[5]),
            responsible:     str(row[6]) || ex?.responsible     || '',
            owner:           str(row[7]) || ex?.owner           || '',
            reviewer:        str(row[8]) || ex?.reviewer        || '',
            ownerStatus:     str(row[9])  || ex?.ownerStatus    || 'Not Started',
            reviewerStatus:  str(row[10]) || ex?.reviewerStatus || 'Not Started',
            expectedStart:   dateStr(row[11]),
            expectedEnd:     dateStr(row[12]),
            actualStart:     dateStr(row[13]),
            actualEnd:       dateStr(row[14]),
            expectedEffort:  row[15] !== '' && row[15] != null ? str(row[15]) : (ex?.expectedEffort  || ''),
            actualEffort:    row[16] !== '' && row[16] != null ? str(row[16]) : (ex?.actualEffort    || ''),
            actualElapsed:   ex?.actualElapsed    || '',
            expectedElapsed: ex?.expectedElapsed  || '',
            comments:        str(row[17]),
          };
        });

        const removed = tasks.length - matched;
        // updateTasks writes ONLY to Supabase `tasks` table scoped to activeProjectId
        updateTasks(importedTasks);
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
        <Select value={form.phase||''} onChange={e => setForm(p=>({...p,phase:e.target.value}))}>
          {PHASES.map(p => <option key={p}>{p}</option>)}
        </Select>
      </Field>
      <Field label="Task Type">
        <Select value={form.taskType||''} onChange={e => setForm(p=>({...p,taskType:e.target.value}))}>
          <option>BA Checklist Item</option><option>Dev Task</option>
        </Select>
      </Field>
      <Field label="Item"><Input value={form.item||''} onChange={e => setForm(p=>({...p,item:e.target.value}))} placeholder="e.g. BRD Preparation" /></Field>
      <Field label="Responsible">
        <Select value={form.responsible||''} onChange={e => setForm(p=>({...p,responsible:e.target.value}))}>
          {RESP_LIST.map(r => <option key={r}>{r}</option>)}
        </Select>
      </Field>
      <Field label="Task Description" style={{gridColumn:'1/-1'}}>
        <Textarea value={form.task||''} onChange={e => setForm(p=>({...p,task:e.target.value}))} placeholder="Describe the task…" style={{ minHeight: 70 }} />
      </Field>
      <Field label="Owner Status">
        <Select value={form.ownerStatus||''} onChange={e => setForm(p=>({...p,ownerStatus:e.target.value}))}>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Reviewer Status">
        <Select value={form.reviewerStatus||''} onChange={e => setForm(p=>({...p,reviewerStatus:e.target.value}))}>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </Select>
      </Field>
      <Field label="Expected Start"><Input type="date" value={form.expectedStart||''} onChange={e => setForm(p=>({...p,expectedStart:e.target.value}))} /></Field>
      <Field label="Expected End">  <Input type="date" value={form.expectedEnd||''}   onChange={e => setForm(p=>({...p,expectedEnd:e.target.value}))} /></Field>
      <Field label="Actual Start">  <Input type="date" value={form.actualStart||''}   onChange={e => setForm(p=>({...p,actualStart:e.target.value}))} /></Field>
      <Field label="Actual End">    <Input type="date" value={form.actualEnd||''}     onChange={e => setForm(p=>({...p,actualEnd:e.target.value}))} /></Field>
      <Field label="Expected Effort (h)"><Input type="number" value={form.expectedEffort||''} onChange={e => setForm(p=>({...p,expectedEffort:e.target.value}))} /></Field>
      <Field label="Actual Effort (h)">  <Input type="number" value={form.actualEffort||''}  onChange={e => setForm(p=>({...p,actualEffort:e.target.value}))} /></Field>
      <Field label="Comments" style={{gridColumn:'1/-1'}}><Textarea value={form.comments||''} onChange={e => setForm(p=>({...p,comments:e.target.value}))} /></Field>
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

      {/* Toolbar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '8px 20px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', flexShrink: 0 }}>
        <input placeholder="Search tasks…" value={filterText} onChange={e => setFilterText(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, width: 200, fontFamily: 'Roboto,sans-serif' }} />
        <select value={filterPhase} onChange={e => setFilterPhase(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Phases</option>
          {PHASES.map(p => <option key={p}>{p}</option>)}
        </select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Statuses</option>
          {STATUSES.map(s => <option key={s}>{s}</option>)}
        </select>
        <span style={{ fontSize: 12, color: '#888', marginLeft: 'auto' }}>{filtered.length} of {tasks.length} shown</span>
        {canDrag && <span style={{ fontSize: 11, color: '#9799b1' }}>⠿ Drag rows to reorder</span>}
      </div>

      {!canDrag && (
        <div style={{ background: '#fff8e1', borderBottom: '1px solid #ffe082', padding: '5px 20px', fontSize: 11, color: '#f57f17', flexShrink: 0 }}>
          Clear all filters to enable drag-and-drop reordering.
        </div>
      )}

      {/* Table */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0', position: 'sticky', top: 0, zIndex: 2 }}>
              {['','#','Task','Phase','Item','Status','Exp. Start','Exp. End','Act. Start','Act. End','Exp. Effort','Act. Effort',''].map((h,i) => (
                <th key={i} style={{ padding: '8px 6px', textAlign: 'left', fontSize: 11, fontWeight: 600, color: '#444', whiteSpace: 'nowrap', background: '#f5f5f5' }}>{h}</th>
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
                onDrop={canDrag ? handleDrop        : () => {}}
                onDragEnd={canDrag ? handleDragEnd  : () => {}}
                isDragging={dragOver === i}
              />
            ))}
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

      {/* Edit / New Task Modal */}
      {editTask !== null && (
        <Modal
          title={isNew ? 'Add New Task' : 'Edit Task'}
          onClose={() => { setEditTask(null); setIsNew(false); }}
          onSave={saveEdit}
          wide
        >
          {isNew
            ? formFields(newForm, setNewForm)
            : formFields(editForm, setEditForm)
          }
        </Modal>
      )}
    </div>
  );
}
