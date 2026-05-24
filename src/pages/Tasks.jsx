import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, StatusBadge, EmptyState } from '../components/UI';
import { ADMIN_LISTS } from '../lib/defaults';
import { getAppState, setAppState } from '../lib/db';
import * as XLSX from 'xlsx';

// ─── Constants ────────────────────────────────────────────────
const STATUSES   = ADMIN_LISTS.statuses;
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

// Built-in columns — cannot be removed, but can be reordered/resized
const BASE_COLS = [
  { key: '_expand',        label: '',               defaultW: 28,  system: true,  filterable: false, editable: false },
  { key: '_num',           label: '#',              defaultW: 38,  system: true,  filterable: false, editable: false },
  { key: 'task',           label: 'Task',           defaultW: 280, system: true,  filterable: true,  editable: true,  editType: 'textarea' },
  { key: 'taskType',       label: 'Type',           defaultW: 105, system: false, filterable: true,  editable: true,  editType: 'select', options: TASK_TYPES },
  { key: 'tags',           label: 'Tags',           defaultW: 120, system: false, filterable: true,  editable: true,  editType: 'text' },
  { key: 'responsible',    label: 'Responsible',    defaultW: 155, system: false, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'owner',          label: 'Owner',          defaultW: 155, system: false, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'reviewer',       label: 'Reviewer',       defaultW: 155, system: false, filterable: true,  editable: true,  editType: 'select', options: RESP_LIST },
  { key: 'ownerStatus',    label: 'Owner Status',   defaultW: 115, system: false, filterable: true,  editable: true,  editType: 'select', options: STATUSES },
  { key: 'reviewerStatus', label: 'Rev. Status',    defaultW: 115, system: false, filterable: true,  editable: true,  editType: 'select', options: STATUSES },
  { key: 'expectedStart',  label: 'Exp. Start',     defaultW: 100, system: false, filterable: false, editable: true,  editType: 'date' },
  { key: 'expectedEnd',    label: 'Exp. End',       defaultW: 100, system: false, filterable: false, editable: true,  editType: 'date' },
  { key: 'actualStart',    label: 'Act. Start',     defaultW: 100, system: false, filterable: false, editable: true,  editType: 'date' },
  { key: 'actualEnd',      label: 'Act. End',       defaultW: 100, system: false, filterable: false, editable: true,  editType: 'date' },
  { key: 'expectedEffort', label: 'Exp. Effort (d)',defaultW: 90,  system: false, filterable: false, editable: true,  editType: 'number' },
  { key: 'actualEffort',   label: 'Act. Effort (d)',defaultW: 90,  system: false, filterable: false, editable: true,  editType: 'number' },
  { key: 'comments',       label: 'Comments',       defaultW: 220, system: false, filterable: true,  editable: true,  editType: 'textarea' },
  { key: '_actions',       label: '',               defaultW: 60,  system: true,  filterable: false, editable: false },
];

// Extra column templates users can add
const EXTRA_COL_TEMPLATES = [
  { key: 'risk',         label: 'Risk',          defaultW: 140, editType: 'text',   filterable: true  },
  { key: 'blockers',     label: 'Blockers',      defaultW: 160, editType: 'text',   filterable: true  },
  { key: 'notes',        label: 'Notes',         defaultW: 200, editType: 'textarea', filterable: true },
  { key: 'priority',     label: 'Priority',      defaultW: 100, editType: 'select', filterable: true,
    options: ['Critical', 'High', 'Medium', 'Low'] },
  { key: 'complexity',   label: 'Complexity',    defaultW: 110, editType: 'select', filterable: true,
    options: ['High', 'Medium', 'Low'] },
  { key: 'sprint',       label: 'Sprint',        defaultW: 100, editType: 'text',   filterable: true  },
  { key: 'milestone',    label: 'Milestone',     defaultW: 130, editType: 'text',   filterable: true  },
  { key: 'dependency',   label: 'Dependency',    defaultW: 150, editType: 'text',   filterable: true  },
];

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

// ─── Quick filter definitions ─────────────────────────────────
const today = () => new Date().toISOString().slice(0, 10);
const endOfWeek = () => {
  const d = new Date();
  d.setDate(d.getDate() + (6 - d.getDay()));
  return d.toISOString().slice(0, 10);
};

const QUICK_FILTERS = [
  {
    key: 'due_today',
    label: 'Due Today',
    icon: '📅',
    color: '#da9b38',
    test: t => t.expectedEnd === today() && !['Done'].includes(t.ownerStatus),
  },
  {
    key: 'due_week',
    label: 'Due This Week',
    icon: '📆',
    color: '#404789',
    test: t => t.expectedEnd && t.expectedEnd >= today() && t.expectedEnd <= endOfWeek() && !['Done'].includes(t.ownerStatus),
  },
  {
    key: 'overdue',
    label: 'Overdue',
    icon: '🔴',
    color: '#e53935',
    test: t => t.expectedEnd && t.expectedEnd < today() && !['Done'].includes(t.ownerStatus),
  },
  {
    key: 'blocked',
    label: 'Blocked',
    icon: '⛔',
    color: '#c62828',
    test: t => ['Blocked', 'Delayed'].includes(t.ownerStatus),
  },
  {
    key: 'in_progress',
    label: 'In Progress',
    icon: '▶',
    color: '#1565c0',
    test: t => t.ownerStatus === 'In Progress',
  },
];

// ─── Pure helpers ─────────────────────────────────────────────
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

// Derive aggregate status for phase/item rows
const STATUS_PRIORITY = ['Blocked', 'Delayed', 'In Progress', 'Done', 'Not Started'];
function deriveStatus(statuses) {
  for (const s of STATUS_PRIORITY) {
    if (statuses.includes(s)) return s;
  }
  return 'Not Started';
}
const STATUS_COLORS = {
  'Done':        { bg: '#e8f5e9', fg: '#2e7d32' },
  'In Progress': { bg: '#e3f2fd', fg: '#1565c0' },
  'Blocked':     { bg: '#fce4ec', fg: '#c62828' },
  'Delayed':     { bg: '#fff3e0', fg: '#e65100' },
  'Not Started': { bg: '#f5f5f5', fg: '#9e9e9e' },
};

// ─── Cell editor ──────────────────────────────────────────────
function CellEditor({ col, value, onCommit, onCancel }) {
  const [val, setVal] = useState(value ?? '');
  const ref = useRef();
  useEffect(() => { ref.current?.focus(); ref.current?.select?.(); }, []);
  const commit = useCallback(() => onCommit(val), [onCommit, val]);
  const onKey = e => {
    if (e.key === 'Enter' && col.editType !== 'textarea') { e.preventDefault(); commit(); }
    if (e.key === 'Escape') onCancel();
  };
  const base = { width: '100%', fontSize: 11, border: '2px solid #404789', borderRadius: 3, padding: '3px 5px', fontFamily: 'Roboto,sans-serif', background: '#fffff8', boxSizing: 'border-box', outline: 'none' };
  if (col.editType === 'select') {
    return (
      <select ref={ref} value={val} onChange={e => setVal(e.target.value)} onBlur={commit} onKeyDown={onKey} style={base}>
        {(col.options || []).map(o => <option key={o}>{o}</option>)}
      </select>
    );
  }
  if (col.editType === 'textarea') {
    return <textarea ref={ref} value={val} onChange={e => setVal(e.target.value)} onBlur={commit} onKeyDown={onKey} style={{ ...base, minHeight: 52, resize: 'vertical' }} />;
  }
  return <input ref={ref} type={col.editType === 'date' ? 'date' : col.editType === 'number' ? 'number' : 'text'} value={val} onChange={e => setVal(e.target.value)} onBlur={commit} onKeyDown={onKey} style={base} />;
}

// ─── Cell display ─────────────────────────────────────────────
function CellDisplay({ col, task }) {
  const v = task[col.key];
  if (col.key === 'taskType') return <TypeBadge type={v} />;
  if (col.key === 'ownerStatus') return <StatusBadge status={v || 'Not Started'} />;
  if (col.key === 'reviewerStatus') return <StatusBadge status={v || 'Not Started'} />;
  if (col.key === 'expectedEffort' || col.key === 'actualEffort')
    return <span style={{ fontSize: 11, color: '#777' }}>{v ? `${v}d` : '—'}</span>;
  if (!v) return <span style={{ color: '#ccc' }}>—</span>;
  return <span style={{ fontSize: 11, color: '#404041', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: col.editType === 'textarea' ? 'normal' : 'nowrap' }}>{v}</span>;
}

// ─── Resize handle ────────────────────────────────────────────
function ResizeHandle({ onResize }) {
  const dragging = useRef(false);
  const startX = useRef(0);
  const onMouseDown = e => {
    e.preventDefault(); e.stopPropagation();
    dragging.current = true; startX.current = e.clientX;
    let moved = false;
    const onMove = mv => { if (!dragging.current) return; moved = true; onResize(mv.clientX - startX.current); startX.current = mv.clientX; };
    const onUp = () => {
      dragging.current = false;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      if (moved) window.addEventListener('click', ce => { ce.stopPropagation(); ce.preventDefault(); }, { capture: true, once: true });
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };
  return (
    <span onMouseDown={onMouseDown}
      style={{ position: 'absolute', right: 0, top: 0, bottom: 0, width: 6, cursor: 'col-resize', background: 'transparent', zIndex: 2 }}
      onMouseEnter={e => { e.currentTarget.style.background = '#b0b4ee'; }}
      onMouseLeave={e => { e.currentTarget.style.background = 'transparent'; }} />
  );
}

// ─── Column filter cell ───────────────────────────────────────
function ColFilter({ col, value, onChange }) {
  const thStyle = { background: '#e8e8f0', borderBottom: '2px solid #c0c0d0', padding: '2px 4px', borderRight: '1px solid #d8d8e8' };
  if (!col.filterable) return <th style={thStyle} />;
  const base = { width: '100%', fontSize: 11, border: '1px solid #c8c8d8', borderRadius: 3, padding: '3px 5px', fontFamily: 'Roboto,sans-serif', background: '#fff', boxSizing: 'border-box' };
  if (col.editType === 'select') {
    return (
      <th style={thStyle}>
        <select value={value} onChange={e => onChange(e.target.value)} style={base}>
          <option value="">All</option>
          {(col.options || []).map(o => <option key={o}>{o}</option>)}
        </select>
      </th>
    );
  }
  return <th style={thStyle}><input value={value} onChange={e => onChange(e.target.value)} placeholder="Filter…" style={base} /></th>;
}

// ─── Main component ───────────────────────────────────────────
export default function Tasks() {
  const { activeProject, activeProjectId, tasks, updateTasks, updateSingleTask, showToast } = useApp();

  // Grid display state
  const [colWidths,      setColWidths]      = useState(() => Object.fromEntries(BASE_COLS.map(c => [c.key, c.defaultW])));
  const [visibleColKeys, setVisibleColKeys] = useState(() => BASE_COLS.map(c => c.key));
  const [extraCols,      setExtraCols]      = useState([]);   // user-added columns
  const [sortKey,        setSortKey]        = useState(null);
  const [sortDir,        setSortDir]        = useState('asc');
  const [colFilters,     setColFilters]     = useState({});
  const [quickFilter,    setQuickFilter]    = useState(null);  // QUICK_FILTERS key or null
  const [editCell,       setEditCell]       = useState(null);
  const [pendingNew,     setPendingNew]     = useState(null);
  const [newRowData,     setNewRowData]     = useState({ ...EMPTY_TASK });
  const [deleteConfirm,  setDeleteConfirm]  = useState(null);
  const [showAddCol,     setShowAddCol]     = useState(false);
  // Expand/collapse: track which phases and items are collapsed
  // Default: all expanded (undefined = expanded)
  const [collapsed, setCollapsed] = useState({});  // key: 'phase::X' or 'item::X::Y'

  const dragIdx = useRef(null);
  const [dragOver, setDragOver] = useState(null);
  const gridStateKey = activeProjectId ? `tasks_grid_${activeProjectId}` : null;

  // All columns = base + extra
  const allCols = useMemo(() => {
    const extraDefs = extraCols.map(k => {
      const tmpl = EXTRA_COL_TEMPLATES.find(t => t.key === k);
      return tmpl ? { ...tmpl, system: false, editable: true } : null;
    }).filter(Boolean);
    // Insert extra cols before _actions
    const base = [...BASE_COLS];
    const actIdx = base.findIndex(c => c.key === '_actions');
    base.splice(actIdx, 0, ...extraDefs);
    return base;
  }, [extraCols]);

  const visibleCols = useMemo(() =>
    allCols.filter(c => visibleColKeys.includes(c.key) || c.system),
    [allCols, visibleColKeys]
  );

  // ── Load persisted state ───────────────────────────────────
  useEffect(() => {
    if (!gridStateKey) return;
    getAppState(gridStateKey).then(result => {
      if (result) {
        try {
          const saved = JSON.parse(result);
          if (saved.colWidths)      setColWidths(w => ({ ...w, ...saved.colWidths }));
          if (saved.sortKey)        setSortKey(saved.sortKey);
          if (saved.sortDir)        setSortDir(saved.sortDir);
          if (saved.visibleColKeys) setVisibleColKeys(saved.visibleColKeys);
          if (saved.extraCols)      setExtraCols(saved.extraCols);
          if (saved.collapsed)      setCollapsed(saved.collapsed);
        } catch {}
      }
    });
  }, [gridStateKey]);

  // ── Persist grid state ─────────────────────────────────────
  const persist = useCallback((patch) => {
    if (!gridStateKey) return;
    setColWidths(prev => {
      const next = { ...prev, ...(patch.colWidths || {}) };
      const state = {
        colWidths: next,
        sortKey:        patch.sortKey        !== undefined ? patch.sortKey        : sortKey,
        sortDir:        patch.sortDir        !== undefined ? patch.sortDir        : sortDir,
        visibleColKeys: patch.visibleColKeys !== undefined ? patch.visibleColKeys : visibleColKeys,
        extraCols:      patch.extraCols      !== undefined ? patch.extraCols      : extraCols,
        collapsed:      patch.collapsed      !== undefined ? patch.collapsed      : collapsed,
      };
      setAppState(gridStateKey, JSON.stringify(state));
      return next;
    });
  }, [gridStateKey, sortKey, sortDir, visibleColKeys, extraCols, collapsed]);

  // ── Hierarchy: group tasks by phase → item ─────────────────
  const hierarchy = useMemo(() => {
    const qf = QUICK_FILTERS.find(q => q.key === quickFilter);
    let base = tasks;
    // Apply quick filter pre-pass (only filters leaf tasks; phases/items then show if they have matching tasks)
    const matchedIds = qf ? new Set(tasks.filter(qf.test).map(t => t.id)) : null;

    // Apply column filters
    base = base.filter(t => {
      for (const col of visibleCols) {
        const fv = (colFilters[col.key] || '').toLowerCase().trim();
        if (!fv || !col.filterable) continue;
        const cell = (t[col.key] || '').toString().toLowerCase();
        if (!cell.includes(fv)) return false;
      }
      if (matchedIds && !matchedIds.has(t.id)) return false;
      return true;
    });

    if (sortKey) {
      base = [...base].sort((a, b) => {
        const av = (a[sortKey] || '').toString().toLowerCase();
        const bv = (b[sortKey] || '').toString().toLowerCase();
        const cmp = av.localeCompare(bv, undefined, { numeric: true });
        return sortDir === 'asc' ? cmp : -cmp;
      });
    }

    // Build phase → item → task tree
    const phaseOrder = [];
    const phaseMap = {};
    base.forEach(t => {
      if (!phaseMap[t.phase]) { phaseMap[t.phase] = {}; phaseOrder.push(t.phase); }
      if (!phaseMap[t.phase][t.item]) phaseMap[t.phase][t.item] = [];
      phaseMap[t.phase][t.item].push(t);
    });
    return phaseOrder.map(phase => {
      const items = Object.entries(phaseMap[phase]).map(([item, tlist]) => {
        const itemStatus = deriveStatus(tlist.map(t => t.ownerStatus));
        return { item, tasks: tlist, status: itemStatus };
      });
      const phaseStatus = deriveStatus(items.map(i => i.status));
      return { phase, items, status: phaseStatus };
    });
  }, [tasks, colFilters, quickFilter, sortKey, sortDir, visibleCols]);

  const hasFilters  = Object.values(colFilters).some(v => v) || !!quickFilter;
  const canDrag     = !hasFilters && !sortKey;
  const totalVisible = hierarchy.reduce((s, p) => s + p.items.reduce((ss, i) => ss + i.tasks.length, 0), 0);

  // ── Expand / collapse ──────────────────────────────────────
  const toggleCollapse = useCallback((key) => {
    setCollapsed(prev => {
      const next = { ...prev, [key]: !prev[key] };
      if (gridStateKey) {
        setColWidths(w => {
          setAppState(gridStateKey, JSON.stringify({ colWidths: w, sortKey, sortDir, visibleColKeys, extraCols, collapsed: next }));
          return w;
        });
      }
      return next;
    });
  }, [gridStateKey, sortKey, sortDir, visibleColKeys, extraCols]);

  const expandAll  = useCallback(() => { setCollapsed({}); persist({ collapsed: {} }); }, [persist]);
  const collapseAll = useCallback(() => {
    const next = {};
    hierarchy.forEach(p => {
      next[`phase::${p.phase}`] = true;
      p.items.forEach(i => { next[`item::${p.phase}::${i.item}`] = true; });
    });
    setCollapsed(next);
    persist({ collapsed: next });
  }, [hierarchy, persist]);

  // ── Early return ───────────────────────────────────────────
  if (!activeProject) return <EmptyState message="No project selected." />;

  const done    = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog  = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => ['Blocked', 'Delayed'].includes(t.ownerStatus)).length;

  // ── Sort ───────────────────────────────────────────────────
  const handleSort = key => {
    let nk = key, nd = 'asc';
    if (sortKey === key && sortDir === 'asc') nd = 'desc';
    else if (sortKey === key && sortDir === 'desc') { nk = null; nd = 'asc'; }
    setSortKey(nk); setSortDir(nd);
    persist({ sortKey: nk, sortDir: nd });
  };

  // ── Resize ─────────────────────────────────────────────────
  const handleResize = (key, delta) => {
    setColWidths(prev => {
      const col = allCols.find(c => c.key === key);
      const min = 30;
      const nw  = { ...prev, [key]: Math.max(min, (prev[key] || col?.defaultW || 80) + delta) };
      if (gridStateKey) setAppState(gridStateKey, JSON.stringify({ colWidths: nw, sortKey, sortDir, visibleColKeys, extraCols, collapsed }));
      return nw;
    });
  };

  // ── Cell commit ────────────────────────────────────────────
  const commitCell = async (taskId, colKey, value) => {
    setEditCell(null);
    const task = tasks.find(t => t.id === taskId);
    if (!task || (task[colKey] ?? '') === value) return;
    await updateSingleTask(taskId, { [colKey]: value });
  };

  // ── Insert row ─────────────────────────────────────────────
  const insertBelow = (afterTaskId, defaultPhase, defaultItem) => {
    const idx = tasks.findIndex(t => t.id === afterTaskId);
    setPendingNew({ afterTaskId, afterRealIdx: idx });
    setNewRowData({ ...EMPTY_TASK, phase: defaultPhase || EMPTY_TASK.phase, item: defaultItem || '' });
  };

  const commitNewRow = async () => {
    if (!newRowData.task.trim() && !newRowData.item.trim()) { setPendingNew(null); return; }
    const afterRealIdx = pendingNew.afterRealIdx ?? tasks.length - 1;
    const updated = [...tasks];
    updated.splice(afterRealIdx + 1, 0, { ...newRowData, id: undefined, sortOrder: afterRealIdx + 1 });
    updated.forEach((t, i) => { t.sortOrder = i; });
    await updateTasks(updated);
    setPendingNew(null); setNewRowData({ ...EMPTY_TASK });
    showToast('Task added');
  };

  // ── Delete ─────────────────────────────────────────────────
  const deleteTask = async taskId => {
    const updated = tasks.filter(t => t.id !== taskId).map((t, i) => ({ ...t, sortOrder: i }));
    await updateTasks(updated); setDeleteConfirm(null); showToast('Task deleted');
  };

  // ── Drag ───────────────────────────────────────────────────
  const handleDragStart = idx => { dragIdx.current = idx; };
  const handleDragOver  = idx => { setDragOver(idx); };
  const handleDrop      = async dropIdx => {
    if (dragIdx.current === null || dragIdx.current === dropIdx) return;
    const reordered = [...tasks];
    const [moved] = reordered.splice(dragIdx.current, 1);
    reordered.splice(dropIdx, 0, moved);
    await updateTasks(reordered);
    dragIdx.current = null; setDragOver(null);
    showToast('Task order updated');
  };
  const handleDragEnd = () => { dragIdx.current = null; setDragOver(null); };

  // ── Add / remove column ─────────────────────────────────────
  const addColumn = key => {
    if (extraCols.includes(key)) return;
    const next = [...extraCols, key];
    const nextVis = [...visibleColKeys, key];
    setExtraCols(next); setVisibleColKeys(nextVis);
    persist({ extraCols: next, visibleColKeys: nextVis });
    setShowAddCol(false); showToast('Column added');
  };
  const removeColumn = key => {
    const nextVis = visibleColKeys.filter(k => k !== key);
    setVisibleColKeys(nextVis);
    persist({ visibleColKeys: nextVis });
    showToast('Column hidden');
  };

  // ── Export ─────────────────────────────────────────────────
  const exportXlsx = () => {
    const dataCols = visibleCols.filter(c => !['_expand','_num','_actions'].includes(c.key));
    const headers  = ['#', 'Phase', 'Item', ...dataCols.map(c => c.label)];
    const rows = [headers];
    let n = 1;
    hierarchy.forEach(p => {
      p.items.forEach(i => {
        i.tasks.forEach(t => {
          rows.push([n++, p.phase, i.item, ...dataCols.map(c => t[c.key] || '')]);
        });
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tasks');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Tasks.xlsx`);
    showToast('Exported to Excel');
  };

  // ── Import ─────────────────────────────────────────────────
  const importXlsx = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (data.length < 2) { showToast('File has no data rows', 'error'); return; }
        const colIdx = {};
        data[0].forEach((h, i) => { const f = HEADER_MAP[String(h).toLowerCase().trim()]; if (f) colIdx[i] = f; });
        if (!Object.values(colIdx).includes('task')) { showToast('Could not find Task column', 'error'); return; }
        const dataRows = data.slice(1).filter(r => r.some(c => c !== '' && c != null));
        const existingByKey = {};
        tasks.forEach(t => { existingByKey[`${t.phase}||${t.item}||${t.task}`] = t; });
        const str = v => (v == null ? '' : String(v)).trim();
        const dateStr = v => {
          if (!v && v !== 0) return '';
          if (v instanceof Date) { const y=v.getFullYear(),m=v.getMonth()+1,d=v.getDate(); return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`; }
          return str(v);
        };
        let matched = 0, added = 0;
        const importedTasks = dataRows.map((row, i) => {
          const raw = {}; Object.entries(colIdx).forEach(([idx, field]) => { raw[field] = row[idx]; });
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
      } catch (err) { console.error(err); showToast('Import failed', 'error'); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // ── Helpers: sort icon, new row ─────────────────────────────
  const SortIcon = ({ colKey }) =>
    sortKey !== colKey
      ? <span style={{ color: '#c0c4d8', fontSize: 9, marginLeft: 2 }}>⇅</span>
      : <span style={{ color: '#da9b38', fontSize: 10, marginLeft: 2 }}>{sortDir === 'asc' ? '↑' : '↓'}</span>;

  const NewRow = ({ phase, item }) => (
    <tr style={{ background: '#fffff0', borderBottom: '2px solid #404789' }}>
      <td colSpan={2} />
      {visibleCols.filter(c => !['_expand','_num','_actions'].includes(c.key)).map(col => (
        <td key={col.key} style={{ padding: '2px 4px' }}>
          {col.editType === 'select' ? (
            <select value={newRowData[col.key]||''} onChange={e => setNewRowData(p=>({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif' }}>
              {(col.options||[]).map(o => <option key={o}>{o}</option>)}
            </select>
          ) : col.editType === 'textarea' ? (
            <textarea value={newRowData[col.key]||''} onChange={e => setNewRowData(p=>({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif', minHeight:38, resize:'vertical', boxSizing:'border-box' }} />
          ) : (
            <input type={col.editType==='date'?'date':col.editType==='number'?'number':'text'}
              value={newRowData[col.key]||''} onChange={e => setNewRowData(p=>({...p,[col.key]:e.target.value}))}
              style={{ width:'100%', fontSize:11, padding:'3px 4px', border:'1px solid #c0c4d8', borderRadius:3, background:'#fff', fontFamily:'Roboto,sans-serif', boxSizing:'border-box' }} />
          )}
        </td>
      ))}
      <td style={{ padding:'3px 6px', whiteSpace:'nowrap' }}>
        <button onClick={commitNewRow}
          style={{ background:'#404789', color:'#fff', border:'none', borderRadius:3, padding:'3px 8px', fontSize:11, cursor:'pointer', marginRight:4, fontFamily:'Roboto,sans-serif' }}>✓ Save</button>
        <button onClick={() => setPendingNew(null)}
          style={{ background:'#f5f5f5', color:'#555', border:'1px solid #ddd', borderRadius:3, padding:'3px 6px', fontSize:11, cursor:'pointer', fontFamily:'Roboto,sans-serif' }}>✕</button>
      </td>
    </tr>
  );

  // ─── Render ────────────────────────────────────────────────
  let globalRowIdx = 0;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Tasks & Checklist"
        subtitle={`${tasks.length} tasks · ${done} done · ${inprog} in progress · ${blocked} blocked`}
        actions={
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <label className="btn btn-ghost" style={{ cursor: 'pointer', margin: 0 }}>
              ↑ Import Excel
              <input type="file" accept=".xlsx,.xls" onChange={importXlsx} style={{ display: 'none' }} />
            </label>
            <button className="btn btn-ghost" onClick={exportXlsx}>↓ Export Excel</button>
            <div style={{ position: 'relative' }}>
              <button className="btn btn-ghost" onClick={() => setShowAddCol(v => !v)}>⊕ Column</button>
              {showAddCol && (
                <div style={{ position: 'absolute', right: 0, top: '100%', marginTop: 4, background: '#fff', border: '1px solid #d0d4e8', borderRadius: 6, boxShadow: '0 4px 16px rgba(0,0,0,0.12)', zIndex: 20, minWidth: 200, padding: '8px 0' }}>
                  <div style={{ padding: '4px 14px 6px', fontSize: 10, color: '#888', textTransform: 'uppercase', letterSpacing: '.5px', fontWeight: 600 }}>Add Column</div>
                  {EXTRA_COL_TEMPLATES.filter(t => !extraCols.includes(t.key)).map(t => (
                    <div key={t.key} onClick={() => addColumn(t.key)}
                      style={{ padding: '6px 14px', fontSize: 12, cursor: 'pointer', color: '#404041' }}
                      onMouseEnter={e => { e.currentTarget.style.background = '#f0f4ff'; }}
                      onMouseLeave={e => { e.currentTarget.style.background = ''; }}>
                      {t.label}
                    </div>
                  ))}
                  {EXTRA_COL_TEMPLATES.every(t => extraCols.includes(t.key)) && (
                    <div style={{ padding: '6px 14px', fontSize: 11, color: '#aaa' }}>All columns added</div>
                  )}
                  <div style={{ borderTop: '1px solid #eee', marginTop: 6, padding: '6px 14px 0' }}>
                    <div style={{ fontSize: 10, color: '#888', marginBottom: 4, textTransform: 'uppercase', letterSpacing: '.5px', fontWeight: 600 }}>Hide Column</div>
                    {visibleCols.filter(c => !c.system && visibleColKeys.includes(c.key)).map(c => (
                      <div key={c.key} onClick={() => removeColumn(c.key)}
                        style={{ padding: '4px 0', fontSize: 12, cursor: 'pointer', color: '#e53935', display: 'flex', alignItems: 'center', gap: 6 }}
                        onMouseEnter={e => { e.currentTarget.style.opacity = '.7'; }}
                        onMouseLeave={e => { e.currentTarget.style.opacity = '1'; }}>
                        ✕ {c.label}
                      </div>
                    ))}
                  </div>
                  <div style={{ borderTop: '1px solid #eee', padding: '6px 14px 2px' }}>
                    <button onClick={() => setShowAddCol(false)} style={{ fontSize: 11, color: '#888', background: 'none', border: 'none', cursor: 'pointer' }}>Close</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        }
      />

      {/* Quick filters */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e8e8f0', padding: '8px 16px', display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0, flexWrap: 'wrap' }}>
        <span style={{ fontSize: 11, color: '#888', marginRight: 4 }}>Quick:</span>
        {QUICK_FILTERS.map(qf => {
          const count = tasks.filter(qf.test).length;
          const active = quickFilter === qf.key;
          return (
            <button key={qf.key}
              onClick={() => setQuickFilter(active ? null : qf.key)}
              style={{
                padding: '4px 10px', fontSize: 11, borderRadius: 14, cursor: 'pointer',
                border: `1px solid ${active ? qf.color : '#d0d4e8'}`,
                background: active ? qf.color : '#fff',
                color: active ? '#fff' : qf.color,
                fontFamily: 'Roboto,sans-serif', fontWeight: active ? 600 : 400,
                display: 'flex', alignItems: 'center', gap: 5, transition: 'all 0.15s',
              }}>
              <span>{qf.icon}</span>
              <span>{qf.label}</span>
              <span style={{ fontSize: 10, opacity: 0.8 }}>({count})</span>
            </button>
          );
        })}
        {hasFilters && (
          <button onClick={() => { setColFilters({}); setQuickFilter(null); }}
            style={{ marginLeft: 8, fontSize: 11, color: '#e53935', background: 'none', border: '1px solid #ffcdd2', borderRadius: 3, padding: '3px 8px', cursor: 'pointer' }}>
            ✕ Clear all filters
          </button>
        )}
      </div>

      {/* Toolbar: expand/collapse + stats */}
      <div style={{ background: '#f8f9fe', borderBottom: '1px solid #e8e8f0', padding: '4px 16px', display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0, flexWrap: 'wrap' }}>
        <button onClick={expandAll}
          style={{ fontSize: 11, padding: '3px 10px', border: '1px solid #c0c4d8', borderRadius: 3, background: '#fff', cursor: 'pointer', fontFamily: 'Roboto,sans-serif' }}>
          ▼ Expand All
        </button>
        <button onClick={collapseAll}
          style={{ fontSize: 11, padding: '3px 10px', border: '1px solid #c0c4d8', borderRadius: 3, background: '#fff', cursor: 'pointer', fontFamily: 'Roboto,sans-serif' }}>
          ▶ Collapse All
        </button>
        <span style={{ fontSize: 11, color: '#888' }}>{totalVisible} of {tasks.length} shown</span>
        {sortKey && (
          <button onClick={() => { setSortKey(null); persist({ sortKey: null, sortDir: 'asc' }); }}
            style={{ fontSize: 11, color: '#b07000', background: 'none', border: '1px solid #e8c870', borderRadius: 3, padding: '2px 8px', cursor: 'pointer' }}>
            ✕ Clear sort
          </button>
        )}
        <span style={{ marginLeft: 'auto', fontSize: 10, color: canDrag ? '#9799b1' : '#bbb' }}>
          {canDrag ? '⠷ Drag rows to reorder · Click cell to edit' : 'Click cell to edit inline'}
        </span>
      </div>

      {/* Grid */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', fontSize: 12, tableLayout: 'fixed', width: visibleCols.reduce((s, c) => s + (colWidths[c.key] || c.defaultW), 0) }}>
          <colgroup>
            {visibleCols.map(c => <col key={c.key} style={{ width: colWidths[c.key] || c.defaultW }} />)}
          </colgroup>
          <thead style={{ position: 'sticky', top: 0, zIndex: 4 }}>
            <tr style={{ background: '#f0f1f8', borderBottom: '1px solid #d0d0d0' }}>
              {visibleCols.map(col => {
                const sortable = col.editable && col.key !== 'task' && col.editType !== 'textarea';
                return (
                  <th key={col.key} onClick={() => sortable && handleSort(col.key)}
                    style={{ padding: '6px 6px', textAlign: 'left', fontSize: 11, fontWeight: 600, color: '#404789', whiteSpace: 'nowrap', background: '#f0f1f8', borderRight: '1px solid #e4e4ee', position: 'relative', cursor: sortable ? 'pointer' : 'default', userSelect: 'none' }}>
                    {col.label}
                    {sortable && <SortIcon colKey={col.key} />}
                    {!['_expand','_num','_actions'].includes(col.key) && <ResizeHandle onResize={delta => handleResize(col.key, delta)} />}
                  </th>
                );
              })}
            </tr>
            <tr>
              {visibleCols.map(col => (
                <ColFilter key={col.key} col={col} value={colFilters[col.key] || ''} onChange={val => setColFilters(p => ({ ...p, [col.key]: val }))} />
              ))}
            </tr>
          </thead>

          <tbody>
            {hierarchy.map(({ phase, items, status: phaseStatus }) => {
              const phaseKey     = `phase::${phase}`;
              const phaseCollapsed = !!collapsed[phaseKey];
              const phaseStatCol = STATUS_COLORS[phaseStatus] || STATUS_COLORS['Not Started'];
              return (
                <React.Fragment key={phase}>
                  {/* Phase header row */}
                  <tr style={{ background: '#e8eaf6', borderBottom: '1px solid #c5cae9' }}>
                    <td style={{ padding: '5px 4px', textAlign: 'center' }}>
                      <button onClick={() => toggleCollapse(phaseKey)}
                        style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 12, color: '#404789', lineHeight: 1, padding: 0 }}>
                        {phaseCollapsed ? '▶' : '▼'}
                      </button>
                    </td>
                    <td style={{ padding: '5px 4px', fontSize: 10, color: '#7986cb', textAlign: 'right' }}>P</td>
                    <td colSpan={visibleCols.length - 2}
                      style={{ padding: '5px 8px', fontWeight: 700, fontSize: 12, color: '#3949ab', letterSpacing: '.2px' }}>
                      {phase}
                      <span style={{ marginLeft: 10, fontSize: 10, fontWeight: 500, padding: '1px 7px', borderRadius: 3, background: phaseStatCol.bg, color: phaseStatCol.fg }}>{phaseStatus}</span>
                      <span style={{ marginLeft: 8, fontSize: 10, color: '#9fa8da', fontWeight: 400 }}>
                        ({items.reduce((s, i) => s + i.tasks.length, 0)} tasks)
                      </span>
                    </td>
                  </tr>

                  {!phaseCollapsed && items.map(({ item, tasks: itemTasks, status: itemStatus }) => {
                    const itemKey       = `item::${phase}::${item}`;
                    const itemCollapsed = !!collapsed[itemKey];
                    const itemStatCol   = STATUS_COLORS[itemStatus] || STATUS_COLORS['Not Started'];
                    return (
                      <React.Fragment key={itemKey}>
                        {/* Item header row */}
                        <tr style={{ background: '#f3f4fb', borderBottom: '1px solid #e0e0f0' }}>
                          <td style={{ padding: '4px 4px', textAlign: 'center' }}>
                            <button onClick={() => toggleCollapse(itemKey)}
                              style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 11, color: '#7986cb', lineHeight: 1, padding: 0 }}>
                              {itemCollapsed ? '▶' : '▼'}
                            </button>
                          </td>
                          <td style={{ padding: '4px 4px', fontSize: 10, color: '#9fa8da', textAlign: 'right' }}>I</td>
                          <td colSpan={visibleCols.length - 2}
                            style={{ padding: '4px 8px 4px 22px', fontWeight: 600, fontSize: 11, color: '#5c6bc0' }}>
                            ├ {item}
                            <span style={{ marginLeft: 8, fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3, background: itemStatCol.bg, color: itemStatCol.fg }}>{itemStatus}</span>
                            <span style={{ marginLeft: 6, fontSize: 10, color: '#b0b8e8', fontWeight: 400 }}>({itemTasks.length} tasks)</span>
                          </td>
                        </tr>

                        {!itemCollapsed && itemTasks.map(task => {
                          const rowIdx     = globalRowIdx++;
                          const isDeleting = deleteConfirm === task.id;
                          const base       = dragOver === rowIdx ? '#e8eaff' : rowIdx % 2 === 0 ? '#fafafe' : '#fff';
                          return (
                            <React.Fragment key={task.id || rowIdx}>
                              <tr
                                draggable={canDrag}
                                onDragStart={() => canDrag && handleDragStart(rowIdx)}
                                onDragOver={e => { e.preventDefault(); canDrag && handleDragOver(rowIdx); }}
                                onDrop={() => canDrag && handleDrop(rowIdx)}
                                onDragEnd={handleDragEnd}
                                style={{ background: isDeleting ? '#fff0f0' : base, borderBottom: '1px solid #f0f0f4', transition: 'background 0.1s' }}
                                onMouseEnter={e => { if (dragOver !== rowIdx && !isDeleting) e.currentTarget.style.background = '#f0f4ff'; }}
                                onMouseLeave={e => { if (dragOver !== rowIdx && !isDeleting) e.currentTarget.style.background = base; }}
                              >
                                <td style={{ padding: '3px 4px', color: '#ccc', cursor: canDrag ? 'grab' : 'default', textAlign: 'center', fontSize: 12, userSelect: 'none' }}>
                                  {canDrag ? '⠷' : ''}
                                </td>
                                <td style={{ padding: '3px 4px', fontSize: 10, color: '#aaa', textAlign: 'right' }}>{rowIdx + 1}</td>

                                {visibleCols.filter(c => !['_expand','_num','_actions'].includes(c.key)).map(col => {
                                  const isEditing = editCell?.taskId === task.id && editCell?.colKey === col.key;
                                  return (
                                    <td key={col.key}
                                      onClick={() => col.editable && !isEditing && setEditCell({ taskId: task.id, colKey: col.key })}
                                      style={{
                                        padding: isEditing ? '2px 3px' : '4px 6px',
                                        paddingLeft: col.key === 'task' ? (isEditing ? 35 : 38) : undefined,
                                        borderLeft: col.key === 'task' ? `3px solid ${task.taskType === 'Dev Task' ? '#da9b38' : '#2278cf'}` : undefined,
                                        cursor: col.editable ? 'text' : 'default',
                                        maxWidth: colWidths[col.key] || col.defaultW,
                                        overflow: 'hidden', verticalAlign: 'top',
                                        background: isEditing ? '#fffff0' : undefined,
                                        borderBottom: isEditing ? '2px solid #404789' : undefined,
                                      }}>
                                      {isEditing
                                        ? <CellEditor col={col} value={task[col.key] ?? ''} onCommit={val => commitCell(task.id, col.key, val)} onCancel={() => setEditCell(null)} />
                                        : <CellDisplay col={col} task={task} />}
                                    </td>
                                  );
                                })}

                                <td style={{ padding: '2px 4px', textAlign: 'center', whiteSpace: 'nowrap', verticalAlign: 'middle' }}>
                                  {isDeleting ? (
                                    <span style={{ fontSize: 10 }}>
                                      <button onClick={() => deleteTask(task.id)} style={{ background: '#e53935', color: '#fff', border: 'none', borderRadius: 3, padding: '2px 5px', fontSize: 10, cursor: 'pointer', marginRight: 2 }}>Del</button>
                                      <button onClick={() => setDeleteConfirm(null)} style={{ background: '#eee', color: '#555', border: '1px solid #ccc', borderRadius: 3, padding: '2px 5px', fontSize: 10, cursor: 'pointer' }}>No</button>
                                    </span>
                                  ) : (
                                    <>
                                      <button onClick={() => insertBelow(task.id, phase, item)} title="Insert task below"
                                        style={{ background: 'none', border: 'none', color: '#404789', cursor: 'pointer', fontSize: 14, padding: '0 2px', lineHeight: 1 }}>⊕</button>
                                      <button onClick={() => setDeleteConfirm(task.id)} title="Delete task"
                                        style={{ background: 'none', border: 'none', color: '#bbb', cursor: 'pointer', fontSize: 13, padding: '0 2px', lineHeight: 1 }}
                                        onMouseEnter={e => { e.currentTarget.style.color = '#e53935'; }}
                                        onMouseLeave={e => { e.currentTarget.style.color = '#bbb'; }}>✕</button>
                                    </>
                                  )}
                                </td>
                              </tr>
                              {pendingNew?.afterTaskId === task.id && <NewRow phase={phase} item={item} />}
                            </React.Fragment>
                          );
                        })}
                      </React.Fragment>
                    );
                  })}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>

        {hierarchy.length === 0 && (
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
    </div>
  );
}
