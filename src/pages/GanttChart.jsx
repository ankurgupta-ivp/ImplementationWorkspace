import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { useApp } from '../hooks/useApp';
import { EmptyState } from '../components/UI';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import { getAppState, setAppState } from '../lib/db';

// ── Default column widths (px) ────────────────────────────────
const DEFAULT_COL_WIDTHS = { name: 220, status: 90, comments: 170 };
const MIN_CHART_W = 600;
const AXIS_H      = 36;
const ROW_H_BASE  = 32;
const LINE_H      = 14;
const PAD_V       = 8;
const CHARS_PER_PX = 6.5; // approx char width at font-size 10–12

// ── Colours ────────────────────────────────────────────────────
const EXP_COLOR = '#404789';
const ACT_COLOR = '#da9b38';
const EXP_LIGHT = '#d0d4f5';
const PHASE_BG  = '#f0f1f8';
const PHASE_FG  = '#404789';
const ITEM_BG   = '#fafafe';

// ── Status ────────────────────────────────────────────────────
const STATUS_PRIORITY = ['Blocked','Delayed','In Progress','Done','Not Started'];
function deriveStatus(statuses) {
  if (!statuses || !statuses.length) return 'Not Started';
  for (const s of STATUS_PRIORITY) { if (statuses.includes(s)) return s; }
  return statuses[0] || 'Not Started';
}
const STATUS_COLORS = {
  'Done':        { bg: '#e8f5e9', fg: '#2e7d32' },
  'In Progress': { bg: '#fff8e1', fg: '#f57f17' },
  'Blocked':     { bg: '#fce4ec', fg: '#c62828' },
  'Delayed':     { bg: '#fff3e0', fg: '#e65100' },
  'Not Started': { bg: '#f5f5f5', fg: '#9e9e9e' },
};

// ── Pure helpers ───────────────────────────────────────────────
function parseDate(s) { if (!s) return null; const d = new Date(s); return isNaN(d.getTime()) ? null : d; }
function fmtDate(d) { if (!d) return '—'; return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'2-digit' }); }
function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function diffDays(a, b) { return Math.round((b - a) / 86400000); }
function minDate(...ds) { const v = ds.filter(Boolean); return v.length ? new Date(Math.min(...v)) : null; }
function maxDate(...ds) { const v = ds.filter(Boolean); return v.length ? new Date(Math.max(...v)) : null; }

// Estimate how many px of height a text label needs given a column width
function labelHeight(text, colPx, fontSize, paddingV) {
  if (!text) return ROW_H_BASE;
  const charsPerLine = Math.max(1, Math.floor((colPx - 16) / (fontSize * 0.6)));
  const lines = text.split('\n').reduce(
    (sum, ln) => sum + Math.max(1, Math.ceil((ln.length || 1) / charsPerLine)), 0
  );
  return Math.max(ROW_H_BASE, lines * (fontSize + 4) + paddingV);
}

// ── Row key ────────────────────────────────────────────────────
function rowKey(row) {
  if (row.type === 'phase') return 'phase::' + row.data.phase;
  if (row.type === 'item')  return 'item::' + row.phase + '::' + row.data.item;
  return 'task::' + (row.data.id || row.data.task);
}

// ── Hierarchy builder ──────────────────────────────────────────
// IMPORTANT: task rows get parsed date fields (expStart/expEnd/actStart/actEnd)
// mapped from their string fields so Bar component can use them directly.
function buildHierarchy(tasks) {
  const devTasks = tasks.filter(t => t.taskType === 'Dev Task');
  const phaseMap = {};
  devTasks.forEach(t => {
    if (!phaseMap[t.phase]) phaseMap[t.phase] = {};
    if (!phaseMap[t.phase][t.item]) phaseMap[t.phase][t.item] = [];
    phaseMap[t.phase][t.item].push(t);
  });
  const phaseOrder = [];
  tasks.forEach(t => {
    if (t.taskType === 'Dev Task' && !phaseOrder.includes(t.phase)) phaseOrder.push(t.phase);
  });
  return phaseOrder.map(phase => {
    const itemMap = phaseMap[phase] || {};
    const itemOrder = [];
    tasks.forEach(t => {
      if (t.phase === phase && t.taskType === 'Dev Task' && !itemOrder.includes(t.item)) itemOrder.push(t.item);
    });
    const items = itemOrder.map(item => {
      const tlist = itemMap[item] || [];
      // Parse task dates once here so Bar can use them at task level too
      const tasksWithDates = tlist.map(t => ({
        ...t,
        expStart: parseDate(t.expectedStart),
        expEnd:   parseDate(t.expectedEnd),
        actStart: parseDate(t.actualStart),
        actEnd:   parseDate(t.actualEnd),
      }));
      const expStart = minDate(...tasksWithDates.map(t => t.expStart));
      const expEnd   = maxDate(...tasksWithDates.map(t => t.expEnd));
      const actStart = minDate(...tasksWithDates.map(t => t.actStart));
      const actEnd   = maxDate(...tasksWithDates.map(t => t.actEnd));
      const status   = deriveStatus(tlist.map(t => t.ownerStatus));
      return { item, expStart, expEnd, actStart, actEnd, tasks: tasksWithDates, status };
    });
    const allExp = items.flatMap(i => [i.expStart, i.expEnd]).filter(Boolean);
    const allAct = items.flatMap(i => [i.actStart, i.actEnd]).filter(Boolean);
    return {
      phase,
      expStart: minDate(...allExp), expEnd: maxDate(...allExp),
      actStart: minDate(...allAct), actEnd: maxDate(...allAct),
      items, status: deriveStatus(items.map(i => i.status)),
    };
  });
}

// ── Axis builder ───────────────────────────────────────────────
function buildAxis(minD, maxD, chartW) {
  const totalDays = diffDays(minD, maxD) + 1;
  const pxPerDay  = chartW / totalDays;
  let tickDays = 7;
  if (pxPerDay * 7  < 30) tickDays = 14;
  if (pxPerDay * 14 < 30) tickDays = 30;
  if (pxPerDay * 30 < 30) tickDays = 60;
  if (pxPerDay * 60 < 30) tickDays = 90;
  const ticks = [];
  let cur = new Date(minD);
  cur.setDate(cur.getDate() - cur.getDay() + 1);
  while (cur <= maxD) {
    if (cur >= minD) ticks.push({ date: new Date(cur), x: Math.round(diffDays(minD, cur) * pxPerDay) });
    cur = addDays(cur, tickDays);
  }
  return { pxPerDay, ticks };
}

// ── Bar SVG ────────────────────────────────────────────────────
function Bar({ start, end, minD, pxPerDay, y, rowH, color, showBoth, isExp }) {
  if (!start || !end) return null;
  const x    = Math.max(0, diffDays(minD, start) * pxPerDay);
  const w    = Math.max(4, diffDays(start, end) * pxPerDay + pxPerDay);
  const mid  = y + rowH / 2;
  const barH = showBoth ? Math.max(8, rowH * 0.35) : Math.max(10, rowH * 0.55);
  const barY = showBoth ? (isExp ? mid - barH - 1 : mid + 1) : mid - barH / 2;
  return (
    <g>
      <rect x={x} y={barY} width={w} height={barH} rx={3} fill={color} opacity={0.85} />
      {w > 44 && (
        <text x={x + 4} y={barY + barH - 3} fontSize={9} fill="#fff"
          style={{ pointerEvents: 'none', fontFamily: 'Roboto,sans-serif' }}>
          {fmtDate(start)}
        </text>
      )}
    </g>
  );
}

// ── Today line ─────────────────────────────────────────────────
function TodayLine({ minD, pxPerDay, totalH }) {
  const today = new Date();
  if (today < minD) return null;
  const x = diffDays(minD, today) * pxPerDay;
  return (
    <g>
      <line x1={x} y1={0} x2={x} y2={totalH} stroke="#e53935" strokeWidth={1.5} strokeDasharray="4 3" opacity={0.7} />
      <rect x={x - 16} y={0} width={32} height={14} rx={3} fill="#e53935" />
      <text x={x} y={10} textAnchor="middle" fontSize={8} fill="#fff" fontFamily="Roboto,sans-serif">Today</text>
    </g>
  );
}

// ── Resize handle (reused from Tasks pattern) ──────────────────
function ResizeHandle({ onResize }) {
  const dragging = useRef(false);
  const startX   = useRef(0);
  const onMouseDown = e => {
    e.preventDefault(); e.stopPropagation();
    dragging.current = true; startX.current = e.clientX;
    let moved = false;
    const onMove = mv => { if (!dragging.current) return; moved = true; onResize(mv.clientX - startX.current); startX.current = mv.clientX; };
    const onUp   = () => {
      dragging.current = false;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      if (moved) window.addEventListener('click', ce => { ce.stopPropagation(); ce.preventDefault(); }, { capture: true, once: true });
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };
  return (
    <div onMouseDown={onMouseDown}
      style={{ position: 'absolute', right: 0, top: 0, bottom: 0, width: 5, cursor: 'col-resize', zIndex: 2 }}
      onMouseEnter={e => { e.currentTarget.style.background = '#b0b4ee'; }}
      onMouseLeave={e => { e.currentTarget.style.background = 'transparent'; }} />
  );
}

// ── Legend item ────────────────────────────────────────────────
function LegendItem({ color, label }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
      <div style={{ width: 20, height: 10, borderRadius: 2, background: color }} />
      <span style={{ fontSize: 11, color: '#555' }}>{label}</span>
    </div>
  );
}

// ── Toggle group ───────────────────────────────────────────────
function ToggleGroup({ label, value, setValue, options }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
      <span style={{ fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{label}:</span>
      <div style={{ display: 'flex', border: '1px solid #c0c4d8', borderRadius: 5, overflow: 'hidden' }}>
        {options.map(opt => (
          <button key={opt.value} onClick={() => setValue(opt.value)} style={{
            padding: '4px 10px', fontSize: 11, border: 'none', cursor: 'pointer',
            background: value === opt.value ? '#404789' : '#fff',
            color:      value === opt.value ? '#fff'    : '#555',
            fontFamily: 'Roboto,sans-serif', fontWeight: value === opt.value ? 600 : 400,
            borderRight: '1px solid #c0c4d8', transition: 'background 0.15s',
          }}>{opt.label}</button>
        ))}
      </div>
    </div>
  );
}

// ── Page header ────────────────────────────────────────────────
function GanttHeader({ activeProject, devCount, depthMode, setDepthMode,
  timelineMode, setTimelineMode, onExportXlsx, onExportPptx }) {
  return (
    <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '10px 20px', flexShrink: 0 }}>
      <div style={{ fontSize: 11, color: '#888', marginBottom: 4 }}>
        Implementation Hub › {activeProject?.name} › Gantt Chart
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <div>
          <div style={{ fontSize: 18, fontWeight: 700, color: '#404041' }}>Gantt Chart</div>
          <div style={{ fontSize: 11, color: '#888' }}>{devCount} Dev Tasks · {activeProject?.name}</div>
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
          <ToggleGroup label="Show" value={depthMode} setValue={setDepthMode} options={[
            { value: 'items', label: 'Items' },
            { value: 'tasks', label: 'Tasks' },
            { value: 'both',  label: 'Both'  },
          ]} />
          <ToggleGroup label="Timeline" value={timelineMode} setValue={setTimelineMode} options={[
            { value: 'expected', label: 'Expected' },
            { value: 'actual',   label: 'Actual'   },
            { value: 'both',     label: 'Both'     },
          ]} />
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={onExportXlsx} style={{ padding: '5px 12px', fontSize: 11, border: '1px solid #c0c4d8', borderRadius: 4, background: '#fff', cursor: 'pointer', color: '#404041', fontFamily: 'Roboto,sans-serif' }}>↓ Excel</button>
            <button onClick={onExportPptx} style={{ padding: '5px 12px', fontSize: 11, border: '1px solid #c0c4d8', borderRadius: 4, background: '#fff', cursor: 'pointer', color: '#404041', fontFamily: 'Roboto,sans-serif' }}>↓ PowerPoint</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Main component ─────────────────────────────────────────────
export default function GanttChart() {
  const { activeProject, activeProjectId, tasks, showToast } = useApp();

  const [depthMode,      setDepthMode]      = useState('items');
  const [timelineMode,   setTimelineMode]   = useState('both');
  const [ganttComments,  setGanttComments]  = useState({});
  const [editingComment, setEditingComment] = useState(null);
  const [commentDraft,   setCommentDraft]   = useState('');
  const [chartW,         setChartW]         = useState(900);
  const [colWidths,      setColWidths]      = useState({ ...DEFAULT_COL_WIDTHS });

  const chartRef    = useRef(null);
  const stateKey    = activeProjectId ? 'gantt_state_' + activeProjectId : null;
  const cmtKey      = activeProjectId ? 'gantt_comments_' + activeProjectId : null;

  // ── Resize observer ──────────────────────────────────────────
  useEffect(() => {
    const obs = new ResizeObserver(entries => {
      const w = entries[0]?.contentRect?.width;
      const lw = colWidths.name + colWidths.status + colWidths.comments;
      if (w > lw + 100) setChartW(Math.max(MIN_CHART_W, w - lw - 32));
    });
    if (chartRef.current) obs.observe(chartRef.current);
    return () => obs.disconnect();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Load persisted state ─────────────────────────────────────
  useEffect(() => {
    if (!stateKey) return;
    getAppState(stateKey).then(r => {
      if (r) { try { const s = JSON.parse(r); if (s.colWidths) setColWidths(w => ({ ...w, ...s.colWidths })); } catch {} }
    });
  }, [stateKey]);

  useEffect(() => {
    if (!cmtKey) return;
    getAppState(cmtKey).then(r => {
      if (r) { try { setGanttComments(JSON.parse(r)); } catch {} }
    });
  }, [cmtKey]);

  // ── Persist helper ───────────────────────────────────────────
  const persistState = useCallback((newWidths) => {
    if (!stateKey) return;
    setAppState(stateKey, JSON.stringify({ colWidths: newWidths }));
  }, [stateKey]);

  // ── Column resize handler ────────────────────────────────────
  const handleColResize = useCallback((col, delta) => {
    setColWidths(prev => {
      const nw = { ...prev, [col]: Math.max(60, (prev[col] || DEFAULT_COL_WIDTHS[col]) + delta) };
      persistState(nw);
      return nw;
    });
  }, [persistState]);

  // ── Derived data ─────────────────────────────────────────────
  const hierarchy = useMemo(() => buildHierarchy(tasks), [tasks]);

  const { globalMin, globalMax } = useMemo(() => {
    const all = [];
    hierarchy.forEach(p => {
      [p.expStart, p.expEnd, p.actStart, p.actEnd].forEach(d => { if (d) all.push(d); });
    });
    if (!all.length) return { globalMin: null, globalMax: null };
    return { globalMin: addDays(new Date(Math.min(...all)), -7), globalMax: addDays(new Date(Math.max(...all)), 7) };
  }, [hierarchy]);

  const { pxPerDay, ticks } = useMemo(
    () => globalMin && globalMax ? buildAxis(globalMin, globalMax, chartW) : { pxPerDay: 0, ticks: [] },
    [globalMin, globalMax, chartW]
  );

  const showExp  = timelineMode === 'expected' || timelineMode === 'both';
  const showAct  = timelineMode === 'actual'   || timelineMode === 'both';
  const showBoth = showExp && showAct;

  // Flat row list
  const rows = useMemo(() => {
    const out = [];
    hierarchy.forEach(phase => {
      out.push({ type: 'phase', data: phase });
      if (depthMode === 'items' || depthMode === 'both') {
        phase.items.forEach(item => {
          out.push({ type: 'item', data: item, phase: phase.phase });
          if (depthMode === 'both') {
            item.tasks.forEach(task => out.push({ type: 'task', data: task, phase: phase.phase, item: item.item }));
          }
        });
      } else if (depthMode === 'tasks') {
        phase.items.forEach(item => {
          item.tasks.forEach(task => out.push({ type: 'task', data: task, phase: phase.phase, item: item.item }));
        });
      }
    });
    return out;
  }, [hierarchy, depthMode]);

  // ── Variable row heights ─────────────────────────────────────
  // Takes the maximum of: comment wrapping height AND label text wrapping height
  const rowHeights = useMemo(() => rows.map(row => {
    const { type, data } = row;
    const label = type === 'phase' ? data.phase : type === 'item' ? data.item : data.task;
    const fontSize = type === 'phase' ? 12 : type === 'item' ? 11 : 10;
    const indent   = type === 'phase' ? 20 : type === 'item' ? 32 : 46;
    const availNamePx = Math.max(40, colWidths.name - indent);

    const labelH = labelHeight(label, availNamePx, fontSize, PAD_V);

    const cmt = ganttComments[rowKey(row)] || '';
    const cmtH = cmt
      ? (() => {
          const charsPerLine = Math.max(1, Math.floor((colWidths.comments - 12) / CHARS_PER_PX));
          const lines = cmt.split('\n').reduce((sum, ln) => sum + Math.max(1, Math.ceil((ln.length || 1) / charsPerLine)), 0);
          return Math.max(ROW_H_BASE, lines * LINE_H + PAD_V);
        })()
      : ROW_H_BASE;

    return Math.max(labelH, cmtH);
  }), [rows, ganttComments, colWidths]);

  const rowYs  = useMemo(() => { const ys = []; let acc = 0; rowHeights.forEach(h => { ys.push(acc); acc += h; }); return ys; }, [rowHeights]);
  const totalH = rowHeights.reduce((s, h) => s + h, 0);
  const LABEL_W = colWidths.name + colWidths.status + colWidths.comments;

  const devCount = tasks.filter(t => t.taskType === 'Dev Task').length;

  // ── Save comment ─────────────────────────────────────────────
  const saveComment = useCallback((key, value) => {
    let latest = {};
    setGanttComments(prev => { latest = { ...prev, [key]: value }; return latest; });
    setTimeout(() => { if (cmtKey) setAppState(cmtKey, JSON.stringify(latest)); }, 0);
  }, [cmtKey]);

  // ── Excel export ──────────────────────────────────────────────
  const exportXlsx = useCallback(() => {
    const hdrs = ['Level', 'Name', 'Status'];
    if (showExp) { hdrs.push('Exp. Start'); hdrs.push('Exp. End'); }
    if (showAct) { hdrs.push('Act. Start'); hdrs.push('Act. End'); }
    hdrs.push('Comments');
    const wsData = [hdrs];
    rows.forEach(row => {
      const { type, data } = row;
      const cmt    = ganttComments[rowKey(row)] || '';
      const status = data.status || (type === 'task' ? (data.ownerStatus || 'Not Started') : 'Not Started');
      const rowArr = [type.charAt(0).toUpperCase() + type.slice(1), type === 'phase' ? data.phase : type === 'item' ? data.item : data.task, status];
      if (showExp) { rowArr.push(fmtDate(data.expStart)); rowArr.push(fmtDate(data.expEnd)); }
      if (showAct) { rowArr.push(fmtDate(data.actStart)); rowArr.push(fmtDate(data.actEnd)); }
      rowArr.push(cmt);
      wsData.push(rowArr);
    });
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const colW = [8, 50, 12];
    if (showExp) { colW.push(12); colW.push(12); }
    if (showAct) { colW.push(12); colW.push(12); }
    colW.push(36);
    ws['!cols'] = colW.map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Gantt');
    XLSX.writeFile(wb, activeProject.name.replace(/[^a-z0-9]/gi, '_') + '_Gantt.xlsx');
    showToast('Gantt exported to Excel');
  }, [rows, ganttComments, showExp, showAct, activeProject, showToast]);

  // ── PPTX export ───────────────────────────────────────────────
  const exportPptx = useCallback(async () => {
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.title  = activeProject.name + ' — Gantt Chart';
      const SLIDE_W = 13.33, SLIDE_H = 7.5;
      const NAME_W_IN = 2.0, STS_W_IN = 0.85, CMT_W_IN = 1.3;
      const LBL_W_IN  = NAME_W_IN + STS_W_IN + CMT_W_IN;
      const BAR_X0    = LBL_W_IN + 0.12;
      const BAR_AREA  = SLIDE_W - BAR_X0 - 0.15;
      const AXIS_Y    = 0.85, MIN_ROW_H = 0.28, START_Y = AXIS_Y + 0.32, MAX_Y = SLIDE_H - 0.15;
      if (!globalMin || !globalMax) { showToast('No date data to export', 'error'); return; }
      const totalD = diffDays(globalMin, globalMax) + 1;
      const pxPD   = BAR_AREA / totalD;
      const xIn    = d => d ? BAR_X0 + diffDays(globalMin, d) * pxPD : BAR_X0;
      const drawHeader = s => {
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: 'F8F9FE' }, line: { color: 'F8F9FE' } });
        ticks.forEach(tick => {
          const tx = BAR_X0 + tick.x / (chartW / BAR_AREA);
          if (tx < BAR_X0 || tx > SLIDE_W - 0.2) return;
          s.addText(tick.date.toLocaleDateString('en-GB', { day:'2-digit', month:'short' }), { x: tx-0.2, y: AXIS_Y, w: 0.4, h: 0.2, fontSize: 7, color:'888888', align:'center', fontFace:'Calibri' });
          s.addShape(pptx.ShapeType.line, { x: tx, y: AXIS_Y+0.2, w: 0, h: SLIDE_H-AXIS_Y-0.2, line: { color:'E0E0E0', width:0.5 } });
        });
      };
      const slide = pptx.addSlide();
      drawHeader(slide);
      slide.addText(activeProject.name + ' — Gantt Chart', { x:0.2, y:0.08, w:SLIDE_W-0.4, h:0.35, fontSize:16, bold:true, color:'404789', fontFace:'Calibri' });
      slide.addText('Dev Tasks · Generated ' + new Date().toLocaleDateString(), { x:0.2, y:0.42, w:SLIDE_W-0.4, h:0.2, fontSize:9, color:'888888', fontFace:'Calibri' });
      if (showExp) { slide.addShape(pptx.ShapeType.rect, { x:SLIDE_W-2.8, y:0.1, w:0.18, h:0.14, fill:{color:'404789'}, line:{color:'404789'} }); slide.addText('Expected', { x:SLIDE_W-2.55, y:0.08, w:0.9, h:0.18, fontSize:8, color:'404789', fontFace:'Calibri' }); }
      if (showAct)  { slide.addShape(pptx.ShapeType.rect, { x:SLIDE_W-1.5, y:0.1, w:0.18, h:0.14, fill:{color:'DA9B38'}, line:{color:'DA9B38'} }); slide.addText('Actual',   { x:SLIDE_W-1.25, y:0.08, w:0.8, h:0.18, fontSize:8, color:'DA9B38', fontFace:'Calibri' }); }
      const today2 = new Date();
      if (today2 >= globalMin && today2 <= globalMax) {
        const tx = BAR_X0 + diffDays(globalMin, today2) * pxPD;
        slide.addShape(pptx.ShapeType.line, { x:tx, y:AXIS_Y, w:0, h:SLIDE_H-AXIS_Y-0.1, line:{color:'E53935',width:1.5,dashType:'dash'} });
        slide.addText('Today', { x:tx-0.18, y:AXIS_Y-0.02, w:0.36, h:0.14, fontSize:7, color:'E53935', align:'center', fontFace:'Calibri' });
      }
      let curY = START_Y, curSlide = slide;
      rows.forEach(row => {
        const { type, data } = row;
        const isPhase = type === 'phase', isItem = type === 'item';
        const cmt = ganttComments[rowKey(row)] || '';
        const cmtLines = cmt ? Math.max(1, cmt.split('\n').reduce((s, ln) => s + Math.max(1, Math.ceil(ln.length / 30)), 0)) : 0;
        const labelTxt = isPhase ? data.phase : isItem ? data.item : data.task;
        const labelLines = Math.max(1, Math.ceil(labelTxt.length / 28));
        const thisRowH = Math.max(MIN_ROW_H, Math.max(cmtLines, labelLines) * 0.14 + 0.04);
        if (curY + thisRowH > MAX_Y) {
          curSlide = pptx.addSlide(); drawHeader(curSlide);
          curSlide.addText(activeProject.name + ' — Gantt Chart (cont.)', { x:0.2, y:0.08, w:SLIDE_W-0.4, h:0.25, fontSize:12, bold:true, color:'404789', fontFace:'Calibri' });
          curY = 0.5;
        }
        if (isPhase) curSlide.addShape(pptx.ShapeType.rect, { x:0.1, y:curY, w:SLIDE_W-0.2, h:thisRowH, fill:{color:'E8EAF6'}, line:{color:'D0D4F5'} });
        const indent = isPhase ? 0.12 : isItem ? 0.22 : 0.32;
        curSlide.addText(labelTxt, { x:indent, y:curY+0.03, w:NAME_W_IN-indent-0.05, h:thisRowH-0.04, fontSize:isPhase?9:isItem?8:7, bold:isPhase, color:isPhase?'404789':'404041', fontFace:'Calibri', wrap:true, valign:'top' });
        curSlide.addShape(pptx.ShapeType.line, { x:NAME_W_IN+0.02, y:curY, w:0, h:thisRowH, line:{color:'E0E0E0',width:0.5} });
        const rowStatus = data.status || (type==='task'?(data.ownerStatus||'Not Started'):'Not Started');
        const stCol = STATUS_COLORS[rowStatus] || STATUS_COLORS['Not Started'];
        curSlide.addText(rowStatus, { x:NAME_W_IN+0.03, y:curY+0.04, w:STS_W_IN-0.06, h:thisRowH-0.06, fontSize:6.5, color:stCol.fg.replace('#',''), fontFace:'Calibri', align:'center', valign:'top' });
        curSlide.addShape(pptx.ShapeType.line, { x:NAME_W_IN+STS_W_IN+0.02, y:curY, w:0, h:thisRowH, line:{color:'E0E0E0',width:0.5} });
        if (cmt) curSlide.addText(cmt, { x:NAME_W_IN+STS_W_IN+0.03, y:curY+0.03, w:CMT_W_IN-0.06, h:thisRowH-0.04, fontSize:6, color:'666666', fontFace:'Calibri', wrap:true, valign:'top' });
        if (showExp && data.expStart && data.expEnd) {
          const bx=xIn(data.expStart), bw=Math.max(0.05,xIn(data.expEnd)-bx+pxPD), barH=showBoth?thisRowH*0.4:thisRowH*0.55, mid=curY+thisRowH/2, by=showBoth?mid-barH-0.01:mid-barH/2;
          curSlide.addShape(pptx.ShapeType.rect, { x:bx, y:by, w:bw, h:barH, fill:{color:isPhase?'2E3F8F':'404789'}, line:{color:'404789'} });
        }
        if (showAct && data.actStart && data.actEnd) {
          const bx=xIn(data.actStart), bw=Math.max(0.05,xIn(data.actEnd)-bx+pxPD), barH=showBoth?thisRowH*0.4:thisRowH*0.55, mid=curY+thisRowH/2, by=showBoth?mid+0.01:mid-barH/2;
          curSlide.addShape(pptx.ShapeType.rect, { x:bx, y:by, w:bw, h:barH, fill:{color:isPhase?'B8770A':'DA9B38'}, line:{color:'DA9B38'} });
        }
        curY += thisRowH;
      });
      await pptx.writeFile({ fileName: activeProject.name.replace(/[^a-z0-9]/gi, '_') + '_Gantt.pptx' });
      showToast('Gantt exported to PowerPoint');
    } catch (err) { console.error('PPTX error:', err); showToast('PPTX export failed', 'error'); }
  }, [rows, ganttComments, globalMin, globalMax, chartW, ticks, showExp, showAct, showBoth, activeProject, showToast]);

  // ── Early returns ─────────────────────────────────────────────
  if (!activeProject) return <EmptyState message="No project selected." />;

  if (devCount === 0) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
        <GanttHeader activeProject={activeProject} devCount={0}
          depthMode={depthMode} setDepthMode={setDepthMode}
          timelineMode={timelineMode} setTimelineMode={setTimelineMode}
          onExportXlsx={exportXlsx} onExportPptx={exportPptx} />
        <EmptyState message="No Dev Tasks found. Mark tasks as 'Dev Task' type in the Tasks & Checklist page." />
      </div>
    );
  }

  // ── Render ─────────────────────────────────────────────────────
  return (
    <div ref={chartRef} style={{ display:'flex', flexDirection:'column', flex:1, overflow:'hidden', background:'#f8f9fe' }}>
      <GanttHeader activeProject={activeProject} devCount={devCount}
        depthMode={depthMode} setDepthMode={setDepthMode}
        timelineMode={timelineMode} setTimelineMode={setTimelineMode}
        onExportXlsx={exportXlsx} onExportPptx={exportPptx} />

      {/* Legend bar */}
      <div style={{ background:'#fff', borderBottom:'1px solid #e0e0e0', padding:'6px 20px', display:'flex', alignItems:'center', gap:20, flexShrink:0 }}>
        {showExp && <LegendItem color={EXP_COLOR} label="Expected" />}
        {showAct && <LegendItem color={ACT_COLOR} label="Actual" />}
        <span style={{ fontSize:11, color:'#888', marginLeft:8 }}>
          {devCount} dev tasks · {hierarchy.length} phases · {hierarchy.reduce((s,p) => s+p.items.length, 0)} items
        </span>
        <div style={{ marginLeft:'auto', display:'flex', alignItems:'center', gap:6 }}>
          <span style={{ width:10, borderTop:'2px dashed #e53935', display:'inline-block' }} />
          <span style={{ fontSize:10, color:'#e53935' }}>Today</span>
        </div>
      </div>

      {!globalMin ? (
        <div style={{ flex:1, display:'flex', alignItems:'center', justifyContent:'center', color:'#999', fontSize:13 }}>
          No date information found on Dev Tasks. Add Expected Start/End dates in the Tasks &amp; Checklist page.
        </div>
      ) : (
        <div style={{ flex:1, overflow:'auto', position:'relative' }}>
          <div style={{ display:'flex', minWidth: LABEL_W + MIN_CHART_W }}>

            {/* Left panel — sticky, resizable columns */}
            <div style={{ width: LABEL_W, minWidth: LABEL_W, flexShrink:0, background:'#fff', borderRight:'2px solid #e0e4f0', position:'sticky', left:0, zIndex:5, display:'flex', flexDirection:'column' }}>

              {/* Frozen header */}
              <div style={{ height:AXIS_H, background:'#f0f1f8', borderBottom:'2px solid #d0d0d0', display:'flex', flexShrink:0, position:'sticky', top:0, zIndex:6 }}>
                <div style={{ width: colWidths.name, minWidth: colWidths.name, display:'flex', alignItems:'center', paddingLeft:14, borderRight:'1px solid #e0e0e8', position:'relative', overflow:'hidden' }}>
                  <span style={{ fontSize:11, fontWeight:600, color:'#404789' }}>Phase / Item / Task</span>
                  <ResizeHandle onResize={d => handleColResize('name', d)} />
                </div>
                <div style={{ width: colWidths.status, minWidth: colWidths.status, display:'flex', alignItems:'center', justifyContent:'center', borderRight:'1px solid #e0e0e8', position:'relative', overflow:'hidden' }}>
                  <span style={{ fontSize:11, fontWeight:600, color:'#404789' }}>Status</span>
                  <ResizeHandle onResize={d => handleColResize('status', d)} />
                </div>
                <div style={{ flex:1, display:'flex', alignItems:'center', paddingLeft:8, position:'relative', overflow:'hidden' }}>
                  <span style={{ fontSize:11, fontWeight:600, color:'#404789' }}>Comments</span>
                  <ResizeHandle onResize={d => handleColResize('comments', d)} />
                </div>
              </div>

              {/* Data rows */}
              {rows.map((row, i) => {
                const { type, data } = row;
                const isPhase = type === 'phase', isItem = type === 'item', isTask = type === 'task';
                const label   = isPhase ? data.phase : isItem ? data.item : data.task;
                const rk      = rowKey(row);
                const cmt     = ganttComments[rk] || '';
                const rh      = rowHeights[i];
                const isEditingCmt = editingComment === rk;
                const status  = data.status || (isTask ? (data.ownerStatus || 'Not Started') : 'Not Started');
                const stColor = STATUS_COLORS[status] || STATUS_COLORS['Not Started'];
                return (
                  <div key={i} style={{ height:rh, display:'flex', flexShrink:0, background: isPhase ? PHASE_BG : isItem ? ITEM_BG : '#fff', borderBottom:'1px solid #f0f0f0' }}>

                    {/* Name cell */}
                    <div style={{ width: colWidths.name, minWidth: colWidths.name, display:'flex', alignItems:'flex-start', paddingTop:6, paddingLeft: isPhase?10:isItem?22:36, paddingRight:6, fontWeight: isPhase?600:400, fontSize: isPhase?12:isItem?11:10, color: isPhase?PHASE_FG:'#404041', borderLeft: isPhase?'3px solid '+PHASE_FG:isItem?'3px solid '+EXP_LIGHT:'3px solid transparent', borderRight:'1px solid #e8e8f0', overflow:'hidden', gap:5, flexShrink:0 }}>
                      {isPhase && <span style={{ fontSize:10, flexShrink:0, marginTop:1 }}>▶</span>}
                      {isItem  && <span style={{ fontSize:9,  color:'#aaa', flexShrink:0, marginTop:1 }}>├</span>}
                      {isTask  && <span style={{ fontSize:9,  color:'#ccc', flexShrink:0, marginTop:1 }}>└</span>}
                      <span style={{ wordBreak:'break-word', overflowWrap:'break-word' }} title={label}>{label}</span>
                    </div>

                    {/* Status cell */}
                    <div style={{ width: colWidths.status, minWidth: colWidths.status, display:'flex', alignItems:'flex-start', justifyContent:'center', paddingTop:6, borderRight:'1px solid #e8e8f0', flexShrink:0 }}>
                      <span style={{ fontSize:9, fontWeight:500, padding:'2px 6px', borderRadius:3, background:stColor.bg, color:stColor.fg, whiteSpace:'nowrap' }}>{status}</span>
                    </div>

                    {/* Comment cell */}
                    <div style={{ flex:1, display:'flex', alignItems:'flex-start', padding:'4px 6px', overflow:'hidden', cursor:'text', minWidth: colWidths.comments }}
                      onClick={() => { setEditingComment(rk); setCommentDraft(cmt); }}>
                      {isEditingCmt ? (
                        <textarea autoFocus value={commentDraft} onChange={e => setCommentDraft(e.target.value)}
                          onBlur={() => { saveComment(rk, commentDraft); setEditingComment(null); }}
                          onKeyDown={e => {
                            if (e.key === 'Escape') setEditingComment(null);
                            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); saveComment(rk, commentDraft); setEditingComment(null); }
                          }}
                          style={{ width:'100%', minHeight:rh-8, resize:'vertical', fontSize:10, border:'2px solid #404789', borderRadius:3, padding:'3px 5px', fontFamily:'Roboto,sans-serif', background:'#fffff8', outline:'none', boxSizing:'border-box', lineHeight:'1.4' }} />
                      ) : (
                        <span style={{ fontSize:10, color: cmt?'#555':'#ccc', whiteSpace:'pre-wrap', wordBreak:'break-word', width:'100%', lineHeight:'1.4' }}>
                          {cmt || <span style={{ fontStyle:'italic', color:'#ccc' }}>add note…</span>}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Chart area */}
            <div style={{ flex:1, minWidth: MIN_CHART_W, position:'relative' }}>
              {/* Frozen axis SVG */}
              <svg width={chartW} height={AXIS_H}
                style={{ display:'block', fontFamily:'Roboto,sans-serif', position:'sticky', top:0, zIndex:3, background:'#f0f1f8' }}>
                <rect x={0} y={0} width={chartW} height={AXIS_H} fill="#f0f1f8" />
                {ticks.map((tick, i) => (
                  <g key={i}>
                    <line x1={tick.x} y1={AXIS_H-6} x2={tick.x} y2={AXIS_H} stroke="#c0c4d4" />
                    <text x={tick.x+3} y={14} fontSize={9} fill="#666" fontFamily="Roboto,sans-serif">
                      {tick.date.toLocaleDateString('en-GB', { day:'2-digit', month:'short' })}
                    </text>
                    <text x={tick.x+3} y={26} fontSize={8} fill="#aaa" fontFamily="Roboto,sans-serif">
                      {tick.date.getFullYear()}
                    </text>
                  </g>
                ))}
                <line x1={0} y1={AXIS_H-1} x2={chartW} y2={AXIS_H-1} stroke="#d0d4e4" strokeWidth={2} />
              </svg>

              {/* Scrollable rows SVG */}
              <svg width={chartW} height={totalH+1} style={{ display:'block', fontFamily:'Roboto,sans-serif' }}>
                {ticks.map((tick, i) => (
                  <line key={i} x1={tick.x} y1={0} x2={tick.x} y2={totalH} stroke="#eaeaf4" />
                ))}
                {rows.map((row, i) => {
                  const { type, data } = row;
                  const isPhase = type === 'phase';
                  const y  = rowYs[i];
                  const rh = rowHeights[i];
                  return (
                    <g key={i}>
                      <rect x={0} y={y} width={chartW} height={rh} fill={isPhase ? '#edeef8' : i%2===0 ? '#fafafe' : '#fff'} />
                      <line x1={0} y1={y+rh} x2={chartW} y2={y+rh} stroke="#f0f0f0" />
                      {showExp && <Bar start={data.expStart} end={data.expEnd} minD={globalMin} pxPerDay={pxPerDay} y={y} rowH={rh} color={isPhase?'#2e3f8f':EXP_COLOR} showBoth={showBoth} isExp={true} />}
                      {showAct && <Bar start={data.actStart} end={data.actEnd} minD={globalMin} pxPerDay={pxPerDay} y={y} rowH={rh} color={isPhase?'#b8770a':ACT_COLOR} showBoth={showBoth} isExp={false} />}
                    </g>
                  );
                })}
                <TodayLine minD={globalMin} pxPerDay={pxPerDay} totalH={totalH} />
              </svg>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}
