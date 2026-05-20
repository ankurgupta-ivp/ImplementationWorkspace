import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, EmptyState } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import { getAppState, setAppState } from '../lib/db';
import * as XLSX from 'xlsx';

// ─── Constants ────────────────────────────────────────────────
const STATUSES   = ADMIN_LISTS.statuses;
const PHASES     = ADMIN_LISTS.phases;
const RESP_LIST  = ADMIN_LISTS.responsibles;
const TASK_TYPES = ['BA Checklist Item', 'Dev Task'];

const EMPTY_TASK = {
  phase: 'Project Creation', item: '', task: '', taskType: 'BA Checklist Item',
  tags: '', responsible: 'Implementation BA', owner: 'Implementation BA',
  reviewer: 'Implementation Lead', ownerStatus: 'Not Started',
  reviewerStatus: 'Not Started', comments: '',
  actualStart: '', actualEnd: '', expectedStart: '', expectedEnd: '',
  actualEffort: '', expectedEffort: '', actualElapsed: '', expectedElapsed: '',
};

// Column definitions — all fields shown, effort in days
const BASE_COLS = [
  { key: '_drag',          label: '',              defaultW: 24,  filterable: false, editable: false  },
  { key: '_num',           label: '#',             defaultW: 38,  filterable: false, editable: false  },
  { key: 'phase',          label: 'Phase',         defaultW: 140, filterable: true,  editable: true,  editType: 'select', options: PHASES },
  { key: 'item',           label: 'Item',          defaultW: 140, filterable: true,  editable: true,  editType: 'text'   },
  { key: 'task',           label: 'Task',          defaultW: 280, filterable: true,  editable: true,  editType: 'textarea' },
  { key: 'taskType',       label: 'Type',          defaultW: 105, filterable: true,  editable: true,  editType: 'select', options: TASK_TYPES },
  { key: 'tags',           label: 'Tags',          defaultW: 120, filterable: true,  editable: true,  editType: 'text'   },
  { key: 'responsible',    label: 'Responsible',   defaultW: 155, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'owner',          label: 'Owner',         defaultW: 155, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'reviewer',       label: 'Reviewer',      defaultW: 155, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'ownerStatus',    label: 'Owner Status',  defaultW: 115, filterable: true,  editable: true,  editType: 'select', options: STATUSES },
  { key: 'reviewerStatus', label: 'Rev. Status',   defaultW: 115, filterable: true,  editable: true,  editType: 'select', options: STATUSES },
  { key: 'expectedStart',  label: 'Exp. Start',    defaultW: 100, filterable: false, editable: true,  editType: 'date'   },
  { key: 'expectedEnd',    label: 'Exp. End',      defaultW: 100, filterable: false, editable: true,  editType: 'date'   },
  { key: 'actualStart',    label: 'Act. Start',    defaultW: 100, filterable: false, editable: true,  editType: 'date'   },
  { key: 'actualEnd',      label: 'Act. End',      defaultW: 100, filterable: false, editable: true,  editType: 'date'   },
  { key: 'expectedEffort', label: 'Exp. Effort (d)',defaultW: 90, filterable: false, editable: true,  editType: 'number' },
  { key: 'actualEffort',   label: 'Act. Effort (d)',defaultW: 90, filterable: false, editable: true,  editType: 'number' },
  { key: 'comments',       label: 'Comments',      defaultW: 220, filterable: true,  editable: true,  editType: 'textarea' },
  { key: '_actions',       label: '',              defaultW: 60,  filterable: false, editable: false  },
];

// Header-aware import map (also update effort labels)
const HEADER_MAP = {
  'phase': 'phase', 'item': 'item', 'task': 'task', 'type': 'taskType',
  'tags': 'tags', 'responsible': 'responsible', 'owner': 'owner',
  'reviewer': 'reviewer', 'owner status': 'ownerStatus',
  'reviewer status': 'reviewerStatus', 'exp. start': 'expectedStart',
  'exp. end': 'expectedEnd', 'act. start': 'actualStart', 'act. end': 'actualEnd',
  'exp. effort (d)': 'expectedEffort', 'act. effort (d)': 'actualEffort',
  'exp. effort (h)': 'expectedEffort', 'act. effort (h)': 'actualEffort',
  'comments': 'comments', 'task type': 'taskType',
  'actual start date': 'actualStart', 'actual end date': 'actualEnd',
  'expected start date': 'expectedStart', 'expected end date': 'expectedEnd',
  'actual effort spent': 'actualEffort', 'expected effort spent': 'expectedEffort',
};

// ─── Helpers ──────────────────────────────────────────────────
function TypeBadge({ type }) {
  const isDev = type === 'Dev Task';
  return (
    <span style={{
      display: 'inline-block', fontSize: 10, fontWeight: 500,
      padding: '2px 6px', borderRadius: 3, whiteSpace: 'nowrap',
      background: isDev ? '#fff3e0' : '#e3f2fd',
      color: isDev ? '#e65100' : '#1565c0',
    }}>{isDev ? 'Dev' : 'BA'}</span>
  );
}

// ─── Inline cell editor ───────────────────────────────────────
function CellEditor({ col, value, onCommit, onCancel }) {
  const [val, setVal] = useState(value ?? '');
  const ref = useRef();

  useEffect(() => { ref.current?.focus(); ref.current?.select?.(); }, []);

  const commit = useCallback(() => onCommit(val), [onCommit, val]);
  const onKey  = (e) => {
    if (e.key === 'Enter' && col.editType !== 'textarea') { e.preventDefault(); commit(); }
    if (e.key === 'Escape') onCancel();
  };

  const base = {
    width: '100%', fontSize: 11, border: '2px solid #404789', borderRadius: 3,
    padding: '3px 5px', fontFamily: 'Roboto,sans-serif', background: '#fffff8',
    boxSizing: 'border-box', outline: 'none',
  };

  if (col.editType === 'select') {
    return (
      <select ref={ref} value={val} onChange={e => setVal(e.target.value)}
        onBlur={commit} onKeyDown={onKey} style={base}>
        {col.options.map(o => <option key={o}>{o}</option>)}
      </select>
    );
  }
  if (col.editType === 'textarea') {
    return (
      <textarea ref={ref} value={val} onChange={e => setVal(e.target.value)}
        onBlur={commit} onKeyDown={onKey}
        style={{ ...base, minHeight: 52, resize: 'vertical' }} />
    );
  }
  return (
    <input ref={ref} type={col.editType === 'date' ? 'date' : col.editType === 'number' ? 'number' : 'text'}
      value={val} onChange={e => setVal(e.target.value)}
      onBlur={commit} onKeyDown={onKey} style={base} />
  );
}

// ─── Read-only cell display ───────────────────────────────────
function CellDisplay({ col, task }) {
  const v = task[col.key];
  if (col.key === 'taskType')       return <TypeBadge type={v} />;
  if (col.key === 'ownerStatus')    return <StatusBadge status={v || 'Not Started'} />;
  if (col.key === 'reviewerStatus') return <StatusBadge status={v || 'Not Started'} />;
  if (col.key === 'expectedEffort' || col.key === 'actualEffort')
    return <span style={{ fontSize: 11, color: '#777' }}>{v ? `${v}d` : '—'}</span>;
  if (!v) return <span style={{ color: '#ccc' }}>—</span>;
  return <span style={{ fontSize: 11, color: '#404041', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: col.editType === 'textarea' ? 'normal' : 'nowrap' }}>{v}</span>;
}

// ─── Resize handle ────────────────────────────────────────────
function ResizeHandle({ onResize }) {
  const dragging = useRef(false);
  const startX   = useRef(0);

  const onMouseDown = (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragging.current = true;
    startX.current   = e.clientX;
    const onMove = (mv) => { if (dragging.current) onResize(mv.clientX - startX.current); startX.current = mv.clientX; };
    const onUp   = () => { dragging.current = false; window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup',   onUp);
  };

  return (
    <span onMouseDown={onMouseDown} style={{
      position: 'absolute', right: 0, top: 0, bottom: 0, width: 6, cursor: 'col-resize',
      background: 'transparent', zIndex: 2,
    }}
      onMouseEnter={e => { e.currentTarget.style.background = '#b0b4ee'; }}
      onMouseLeave={e => { e.currentTarget.style.background = 'transparent'; }}
    />
  );
}

// ─── Column filter row ────────────────────────────────────────
function ColFilter({ col, value, onChange }) {
  if (!col.filterable) return <th style={{ background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', padding: '2px 4px' }} />;
  const base = { width: '100%', fontSize: 11, border: '1px solid #c8c8d8', borderRadius: 3, padding: '3px 5px', fontFamily: 'Roboto,sans-serif', background: '#fff', boxSizing: 'border-box' };
  if (col.editType === 'select') {
    return (
      <th style={{ background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', padding: '2px 4px' }}>
        <select value={value} onChange={e => onChange(e.target.value)} style={base}>
          <option value="">All</option>
          {col.options.map(o => <option key={o}>{o}</option>)}
        </select>
      </th>
    );
  }
  return (
    <th style={{ background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', padding: '2px 4px' }}>
      <input value={value} onChange={e => onChange(e.target.value)} placeholder="Filter…" style={base} />
    </th>
  );
}

// ─── Main component ───────────────────────────────────────────
export default function Tasks() {
  const { activeProject, activeProjectId, tasks, updateTasks, updateSingleTask, showToast } = useApp();

  // Grid state — persisted per project
  const [colWidths,  setColWidths]  = useState(() => Object.fromEntries(BASE_COLS.map(c => [c.key, c.defaultW])));
  const [sortKey,    setSortKey]    = useState(null);   // field key or null
  const [sortDir,    setSortDir]    = useState('asc');  // 'asc' | 'desc'
  const [colFilters, setColFilters] = useState({});
  const [editCell,   setEditCell]   = useState(null);   // { taskId, colKey }
  const [pendingNew, setPendingNew] = useState(null);   // { afterIdx } — new blank row being added inline
  const [newRowData, setNewRowData] = useState({ ...EMPTY_TASK });
  const [deleteConfirm, setDeleteConfirm] = useState(null); // taskId to confirm delete

  const dragIdx  = useRef(null);
  const [dragOver, setDragOver] = useState(null);
  const gridStateKey = activeProjectId ? `tasks_grid_${activeProjectId}` : null;

  // ── Load persisted grid state ──────────────────────────────
  useEffect(() => {
    if (!gridStateKey) return;
    getAppState(gridStateKey).then(result => {
      if (result?.value) {
        try {
          const saved = JSON.parse(result.value);
          if (saved.colWidths) setColWidths(w => ({ ...w, ...saved.colWidths }));
          if (saved.sortKey)   setSortKey(saved.sortKey);
          if (saved.sortDir)   setSortDir(saved.sortDir);
        } catch {}
      }
    });
  }, [gridStateKey]);

  // ── Persist grid state on change ──────────────────────────
  const persistGridState = useCallback((widths, sk, sd) => {
    if (!gridStateKey) return;
    setAppState(gridStateKey, JSON.stringify({ colWidths: widths, sortKey: sk, sortDir: sd }));
  }, [gridStateKey]);

  // ── Filtering ──────────────────────────────────────────────
  const filtered = useMemo(() => {
    let result = tasks.filter(t => {
      for (const col of BASE_COLS) {
        const fv = (colFilters[col.key] || '').toLowerCase().trim();
        if (!fv) continue;
        if (!col.filterable) continue;
        const cell = (t[col.key] || '').toString().toLowerCase();
        if (!cell.includes(fv)) return false;
      }
      return true;
    });
    if (sortKey) {
      result = [...result].sort((a, b) => {
        const av = (a[sortKey] || '').toString().toLowerCase();
        const bv = (b[sortKey] || '').toString().toLowerCase();
        const cmp = av.localeCompare(bv, undefined, { numeric: true });
        return sortDir === 'asc' ? cmp : -cmp;
      });
    }
    return result;
  }, [tasks, colFilters, sortKey, sortDir]);

  const hasFilters = Object.values(colFilters).some(v => v);
  const canDrag    = !hasFilters && !sortKey;

  if (!activeProject) return <EmptyState message="No project selected." />;

  const done    = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog  = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => ['Blocked', 'Delayed'].includes(t.ownerStatus)).length;

  // ── Sort toggle ────────────────────────────────────────────
  const handleSort = (key) => {
    let nk = key, nd = 'asc';
    if (sortKey === key && sortDir === 'asc')  { nd = 'desc'; }
    else if (sortKey === key && sortDir === 'desc') { nk = null; nd = 'asc'; }
    setSortKey(nk); setSortDir(nd);
    persistGridState(colWidths, nk, nd);
  };

  // ── Column resize ──────────────────────────────────────────
  const handleResize = (key, delta) => {
    setColWidths(prev => {
      const col = BASE_COLS.find(c => c.key === key);
      const min = col?.key === '_drag' ? 24 : col?.key === '_num' ? 30 : 50;
      const nw  = { ...prev, [key]: Math.max(min, (prev[key] || col?.defaultW || 80) + delta) };
      persistGridState(nw, sortKey, sortDir);
      return nw;
    });
  };

  // ── Inline cell commit ─────────────────────────────────────
  const commitCell = async (taskId, colKey, value) => {
    setEditCell(null);
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;
    if ((task[colKey] ?? '') === value) return; // no change
    await updateSingleTask(taskId, { [colKey]: value });
  };

  // ── Insert row below ───────────────────────────────────────
  const insertBelow = (idx) => {
    setPendingNew({ afterIdx: idx });
    setNewRowData({ ...EMPTY_TASK, phase: filtered[idx]?.phase || EMPTY_TASK.phase });
  };

  const commitNewRow = async () => {
    if (!newRowData.task.trim() && !newRowData.item.trim()) {
      setPendingNew(null); return;
    }
    const afterTask = filtered[pendingNew.afterIdx];
    const afterRealIdx = afterTask ? tasks.findIndex(t => t.id === afterTask.id) : tasks.length - 1;
    const newTask = { ...newRowData, id: undefined, sortOrder: afterRealIdx + 1 };
    const updated = [...tasks];
    updated.splice(afterRealIdx + 1, 0, newTask);
    updated.forEach((t, i) => { t.sortOrder = i; });
    await updateTasks(updated);
    setPendingNew(null);
    setNewRowData({ ...EMPTY_TASK });
    showToast('Task added');
  };

  // ── Delete task ────────────────────────────────────────────
  const deleteTask = async (taskId) => {
    const updated = tasks.filter(t => t.id !== taskId).map((t, i) => ({ ...t, sortOrder: i }));
    await updateTasks(updated);
    setDeleteConfirm(null);
    showToast('Task deleted');
  };

  // ── Drag reorder ───────────────────────────────────────────
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

  // ── Export Excel ───────────────────────────────────────────
  const exportXlsx = () => {
    const headers = ['#', 'Phase', 'Item', 'Task', 'Type', 'Tags', 'Responsible', 'Owner', 'Reviewer',
      'Owner Status', 'Reviewer Status', 'Exp. Start', 'Exp. End', 'Act. Start', 'Act. End',
      'Exp. Effort (d)', 'Act. Effort (d)', 'Comments'];
    const rows = [headers];
    tasks.forEach((t, i) => rows.push([
      i + 1, t.phase, t.item, t.task, t.taskType, t.tags,
      t.responsible, t.owner, t.reviewer, t.ownerStatus, t.reviewerStatus,
      t.expectedStart, t.expectedEnd, t.actualStart, t.actualEnd,
      t.expectedEffort, t.actualEffort, t.comments,
    ]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [4, 18, 22, 55, 14, 18, 24, 24, 22, 14, 14, 11, 11, 11, 11, 10, 10, 40].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  // ── Import Excel ───────────────────────────────────────────
  const importXlsx = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (data.length < 2) { showToast('File has no data rows', 'error'); return; }
        const colIdx = {};
        data[0].forEach((h, i) => { const f = HEADER_MAP[String(h).toLowerCase().trim()]; if (f) colIdx[i] = f; });
        if (!Object.values(colIdx).includes('phase') || !Object.values(colIdx).includes('task')) {
          showToast('Could not find Phase or Task columns', 'error'); return;
        }
        const dataRows = data.slice(1).filter(row => row.some(c => c !== '' && c != null));
        const existingByKey = {};
        tasks.forEach(t => { existingByKey[`${t.phase}||${t.item}||${t.task}`] = t; });
        const str = (v) => (v == null ? '' : String(v)).trim();
        const dateStr = (v) => {
          if (!v && v !== 0) return '';
          if (v instanceof Date) { const y=v.getFullYear(),m=v.getMonth()+1,d=v.getDate(); return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`; }
          return str(v);
        };
        let matched = 0, added = 0;
        const importedTasks = dataRows.map((row, i) => {
          const raw = {};
          Object.entries(colIdx).forEach(([idx, field]) => { raw[field] = row[idx]; });
          const phase = str(raw.phase||''), item = str(raw.item||''), task = str(raw.task||'');
          const ex = existingByKey[`${phase}||${item}||${task}`];
          if (ex) matched++; else added++;
          const resolve = (field, fallback) => {
            const v = raw[field];
            if (['expectedStart','expectedEnd','actualStart','actualEnd'].includes(field)) return dateStr(v);
            return (v !== '' && v != null) ? str(v) : (ex?.[field] || fallback || '');
          };
          return { id: ex?.id, sortOrder: i, phase, item, task,
            taskType: resolve('taskType','BA Checklist Item'), tags: resolve('tags',''),
            responsible: resolve('responsible',''), owner: resolve('owner',''), reviewer: resolve('reviewer',''),
            ownerStatus: resolve('ownerStatus','Not Started'), reviewerStatus: resolve('reviewerStatus','Not Started'),
            expectedStart: dateStr(raw.expectedStart), expectedEnd: dateStr(raw.expectedEnd),
            actualStart: dateStr(raw.actualStart), actualEnd: dateStr(raw.actualEnd),
            expectedEffort: resolve('expectedEffort',''), actualEffort: resolve('actualEffort',''),
            actualElapsed: ex?.actualElapsed||'', expectedElapsed: ex?.expectedElapsed||'',
            comments: resolve('comments',''),
          };
        });
        updateTasks(importedTasks);
        const removed = tasks.length - matched;
        showToast(`Imported ${importedTasks.length} tasks · ${matched} updated · ${added} new${removed > 0 ? ` · ${removed} removed` : ''}`);
      } catch (err) { console.error(err); showToast('Import failed — check file format', 'error'); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // ── New row inline editor ──────────────────────────────────
  const NewRow = ({ afterIdx }) => (
    <tr style={{ background: '#fffff0', borderBottom: '2px solid #404789' }}>
      <td />
      <td style={{ padding: '3px 4px', fontSize: 10, color: '#999', textAlign: 'center' }}>new</td>
      {BASE_COLS.filter(c => !['_drag','_num','_actions'].includes(c.key)).map(col => (
        <td key={col.key} style={{ padding: '2px 4px' }}>
          {col.editType === 'select' ? (
            <select value={newRowData[col.key]||''} onChange={e => setNewRowData(p => ({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif' }}>
              {col.options.map(o => <option key={o}>{o}</option>)}
            </select>
          ) : col.editType === 'textarea' ? (
            <textarea value={newRowData[col.key]||''} onChange={e => setNewRowData(p => ({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif', minHeight:38, resize:'vertical', boxSizing:'border-box' }} />
          ) : (
            <input type={col.editType === 'date' ? 'date' : col.editType === 'number' ? 'number' : 'text'}
              value={newRowData[col.key]||''} onChange={e => setNewRowData(p => ({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif', boxSizing:'border-box' }} />
          )}
        </td>
      ))}
      <td style={{ padding:'3px 6px', whiteSpace:'nowrap' }}>
        <button onClick={commitNewRow} title="Save new task"
          style={{ background:'#404789', color:'#fff', border:'none', borderRadius:3, padding:'3px 8px', fontSize:11, cursor:'pointer', marginRight:4, fontFamily:'Roboto,sans-serif' }}>✓ Save</button>
        <button onClick={() => setPendingNew(null)} title="Cancel"
          style={{ background:'#f5f5f5', color:'#555', border:'1px solid #ddd', borderRadius:3, padding:'3px 6px', fontSize:11, cursor:'pointer', fontFamily:'Roboto,sans-serif' }}>✕</button>
      </td>
    </tr>
  );

  // ── Sort indicator ─────────────────────────────────────────
  const SortIcon = ({ colKey }) => {
    if (sortKey !== colKey) return <span style={{ color: '#c0c4d8', fontSize: 9, marginLeft: 2 }}>⇅</span>;
    return <span style={{ color: '#da9b38', fontSize: 10, marginLeft: 2 }}>{sortDir === 'asc' ? '↑' : '↓'}</span>;
  };

  // ─── Render ────────────────────────────────────────────────
  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
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

      {/* Summary / filter bar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '5px 16px', display: 'flex', alignItems: 'center', gap: 14, flexShrink: 0, flexWrap: 'wrap' }}>
        <span style={{ fontSize: 11, color: '#888' }}>{filtered.length} of {tasks.length} shown</span>
        {hasFilters && (
          <button onClick={() => setColFilters({})}
            style={{ fontSize: 11, color: '#404789', background: 'none', border: '1px solid #b0b4d4', borderRadius: 3, padding: '2px 8px', cursor: 'pointer' }}>
            ✕ Clear filters
          </button>
        )}
        {sortKey && (
          <button onClick={() => { setSortKey(null); persistGridState(colWidths, null, 'asc'); }}
            style={{ fontSize: 11, color: '#b07000', background: 'none', border: '1px solid #e8c870', borderRadius: 3, padding: '2px 8px', cursor: 'pointer' }}>
            ✕ Clear sort
          </button>
        )}
        <span style={{ marginLeft: 'auto', fontSize: 10, color: canDrag ? '#9799b1' : '#bbb' }}>
          {canDrag ? '⠿ Drag rows to reorder · Click cell to edit' : 'Click cell to edit inline'}
        </span>
      </div>

      {/* Grid */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', fontSize: 12, tableLayout: 'fixed',
          width: BASE_COLS.reduce((s, c) => s + (colWidths[c.key] || c.defaultW), 0) }}>
          <colgroup>
            {BASE_COLS.map(c => <col key={c.key} style={{ width: colWidths[c.key] || c.defaultW }} />)}
          </colgroup>
          <thead style={{ position: 'sticky', top: 0, zIndex: 4 }}>
            {/* Header labels + resize + sort */}
            <tr style={{ background: '#f0f1f8', borderBottom: '1px solid #d0d0d0' }}>
              {BASE_COLS.map(col => {
                const sortable = col.editable && col.key !== 'task' && col.editType !== 'textarea';
                return (
                  <th key={col.key} style={{
                    padding: '6px 6px 6px 6px', textAlign: 'left', fontSize: 11, fontWeight: 600,
                    color: '#404789', whiteSpace: 'nowrap', background: '#f0f1f8',
                    borderRight: '1px solid #e4e4ee', position: 'relative',
                    cursor: sortable ? 'pointer' : 'default', userSelect: 'none',
                  }}
                    onClick={() => sortable && handleSort(col.key)}
                  >
                    {col.label}
                    {sortable && <SortIcon colKey={col.key} />}
                    {!['_drag','_num','_actions'].includes(col.key) && (
                      <ResizeHandle onResize={delta => handleResize(col.key, delta)} />
                    )}
                  </th>
                );
              })}
            </tr>
            {/* Filter row */}
            <tr>
              {BASE_COLS.map(col => (
                <ColFilter key={col.key} col={col}
                  value={colFilters[col.key] || ''}
                  onChange={val => setColFilters(p => ({ ...p, [col.key]: val }))} />
              ))}
            </tr>
          </thead>

          <tbody>
            {filtered.map((task, visIdx) => {
              const base = dragOver === visIdx ? '#e8eaff' : visIdx % 2 === 0 ? '#fafafe' : '#fff';
              const isDeleting = deleteConfirm === task.id;
              return (
                <React.Fragment key={task.id || visIdx}>
                  <tr
                    draggable={canDrag}
                    onDragStart={() => canDrag && handleDragStart(visIdx)}
                    onDragOver={e => { e.preventDefault(); canDrag && handleDragOver(visIdx); }}
                    onDrop={() => canDrag && handleDrop(visIdx)}
                    onDragEnd={handleDragEnd}
                    style={{ background: isDeleting ? '#fff0f0' : base, borderBottom: '1px solid #f0f0f4', transition: 'background 0.1s' }}
                    onMouseEnter={e => { if (dragOver !== visIdx && !isDeleting) e.currentTarget.style.background = '#f0f4ff'; }}
                    onMouseLeave={e => { if (dragOver !== visIdx && !isDeleting) e.currentTarget.style.background = base; }}
                  >
                    {/* drag handle */}
                    <td style={{ padding: '4px 3px', color: '#ccc', cursor: canDrag ? 'grab' : 'default', textAlign: 'center', fontSize: 12, userSelect: 'none' }}>
                      {canDrag ? '⠿' : ''}
                    </td>
                    {/* row number */}
                    <td style={{ padding: '4px 4px', fontSize: 10, color: '#aaa', textAlign: 'right' }}>{visIdx + 1}</td>

                    {/* Data cells */}
                    {BASE_COLS.filter(c => !['_drag','_num','_actions'].includes(c.key)).map(col => {
                      const isEditing = editCell?.taskId === task.id && editCell?.colKey === col.key;
                      const isTaskCol = col.key === 'task';
                      return (
                        <td key={col.key}
                          onClick={() => col.editable && !isEditing && setEditCell({ taskId: task.id, colKey: col.key })}
                          style={{
                            padding: isEditing ? '2px 3px' : '4px 6px',
                            borderLeft: isTaskCol ? `3px solid ${task.taskType === 'Dev Task' ? '#da9b38' : '#2278cf'}` : undefined,
                            cursor: col.editable ? 'text' : 'default',
                            maxWidth: colWidths[col.key] || col.defaultW,
                            overflow: 'hidden',
                            verticalAlign: 'top',
                            background: isEditing ? '#fffff0' : undefined,
                            borderBottom: isEditing ? '2px solid #404789' : undefined,
                          }}>
                          {isEditing ? (
                            <CellEditor
                              col={col}
                              value={task[col.key] ?? ''}
                              onCommit={val => commitCell(task.id, col.key, val)}
                              onCancel={() => setEditCell(null)}
                            />
                          ) : (
                            <CellDisplay col={col} task={task} />
                          )}
                        </td>
                      );
                    })}

                    {/* Actions: insert below + delete */}
                    <td style={{ padding: '2px 4px', textAlign: 'center', whiteSpace: 'nowrap', verticalAlign: 'middle' }}>
                      {isDeleting ? (
                        <span style={{ fontSize: 10 }}>
                          <button onClick={() => deleteTask(task.id)}
                            style={{ background: '#e53935', color: '#fff', border: 'none', borderRadius: 3, padding: '2px 5px', fontSize: 10, cursor: 'pointer', marginRight: 2 }}>Del</button>
                          <button onClick={() => setDeleteConfirm(null)}
                            style={{ background: '#eee', color: '#555', border: '1px solid #ccc', borderRadius: 3, padding: '2px 5px', fontSize: 10, cursor: 'pointer' }}>No</button>
                        </span>
                      ) : (
                        <>
                          <button onClick={() => insertBelow(visIdx)} title="Insert task below"
                            style={{ background: 'none', border: 'none', color: '#404789', cursor: 'pointer', fontSize: 14, padding: '0 2px', lineHeight: 1 }}>⊕</button>
                          <button onClick={() => setDeleteConfirm(task.id)} title="Delete task"
                            style={{ background: 'none', border: 'none', color: '#bbb', cursor: 'pointer', fontSize: 13, padding: '0 2px', lineHeight: 1 }}
                            onMouseEnter={e => { e.currentTarget.style.color = '#e53935'; }}
                            onMouseLeave={e => { e.currentTarget.style.color = '#bbb'; }}>✕</button>
                        </>
                      )}
                    </td>
                  </tr>

                  {/* Inline new row inserted after this row */}
                  {pendingNew?.afterIdx === visIdx && (
                    <NewRow afterIdx={visIdx} />
                  )}
                </React.Fragment>
              );
            })}

            {/* New row at end if inserted after last row */}
            {pendingNew?.afterIdx === filtered.length - 1 + 1 && (
              <NewRow afterIdx={filtered.length} />
            )}
          </tbody>
        </table>

        {filtered.length === 0 && !pendingNew && (
          <div style={{ textAlign: 'center', padding: 40, color: '#999', fontSize: 13 }}>
            {hasFilters ? 'No tasks match the current filters.' : (
              <span>No tasks yet. <button onClick={() => insertBelow(-1)}
                style={{ color: '#404789', background: 'none', border: 'none', cursor: 'pointer', fontSize: 13, textDecoration: 'underline' }}>Add first task</button></span>
            )}
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{ height: 26, background: '#f5f5f5', borderTop: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 14px', fontSize: 11, color: '#888', flexShrink: 0 }}>
        <span>{tasks.length} total tasks</span>
        <span>Done: {done} · In Progress: {inprog} · Blocked: {blocked}</span>
      </div>
    </div>
  );
}
