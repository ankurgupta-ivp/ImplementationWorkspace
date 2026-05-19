import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import { useApp } from '../hooks/useApp';
import { EmptyState } from '../components/UI';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';

// ─────────────────────────────────────────────────────────────
//  Constants & helpers
// ─────────────────────────────────────────────────────────────
const EXP_COLOR   = '#404789';   // navy  – expected
const ACT_COLOR   = '#da9b38';   // amber – actual
const EXP_LIGHT   = '#d0d4f5';
const PHASE_BG    = '#f0f1f8';
const PHASE_FG    = '#404789';
const ITEM_BG     = '#fafafe';
const ROW_H       = 28;          // px per row
const LABEL_W     = 300;         // px for left label column
const MIN_CHART_W = 600;

function parseDate(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' });
}
function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function diffDays(a, b) { return Math.round((b - a) / 86400000); }
function minDate(...dates) { const v = dates.filter(Boolean); return v.length ? new Date(Math.min(...v)) : null; }
function maxDate(...dates) { const v = dates.filter(Boolean); return v.length ? new Date(Math.max(...v)) : null; }

// ─────────────────────────────────────────────────────────────
//  Data builder
//  Returns hierarchical structure:
//  [ { phase, expStart, expEnd, actStart, actEnd,
//      items: [ { item, expStart, expEnd, actStart, actEnd,
//                 tasks: [ { ...task } ] } ] } ]
// ─────────────────────────────────────────────────────────────
function buildHierarchy(tasks) {
  const devTasks = tasks.filter(t => t.taskType === 'Dev Task');

  // Group by phase
  const phaseMap = {};
  devTasks.forEach(t => {
    if (!phaseMap[t.phase]) phaseMap[t.phase] = {};
    if (!phaseMap[t.phase][t.item]) phaseMap[t.phase][t.item] = [];
    phaseMap[t.phase][t.item].push(t);
  });

  // Preserve original phase order from full task list
  const phaseOrder = [];
  tasks.forEach(t => { if (t.taskType === 'Dev Task' && !phaseOrder.includes(t.phase)) phaseOrder.push(t.phase); });

  return phaseOrder.map(phase => {
    const itemMap = phaseMap[phase];
    const itemOrder = [];
    tasks.forEach(t => { if (t.phase === phase && t.taskType === 'Dev Task' && !itemOrder.includes(t.item)) itemOrder.push(t.item); });

    const items = itemOrder.map(item => {
      const tlist = itemMap[item] || [];
      const expStart = minDate(...tlist.map(t => parseDate(t.expectedStart)));
      const expEnd   = maxDate(...tlist.map(t => parseDate(t.expectedEnd)));
      const actStart = minDate(...tlist.map(t => parseDate(t.actualStart)));
      const actEnd   = maxDate(...tlist.map(t => parseDate(t.actualEnd)));
      return { item, expStart, expEnd, actStart, actEnd, tasks: tlist };
    });

    const allExp = items.flatMap(i => [i.expStart, i.expEnd]).filter(Boolean);
    const allAct = items.flatMap(i => [i.actStart, i.actEnd]).filter(Boolean);
    return {
      phase,
      expStart: minDate(...allExp),
      expEnd:   maxDate(...allExp),
      actStart: minDate(...allAct),
      actEnd:   maxDate(...allAct),
      items,
    };
  });
}

// ─────────────────────────────────────────────────────────────
//  Timeline axis builder
// ─────────────────────────────────────────────────────────────
function buildAxis(minD, maxD, chartW) {
  const totalDays = diffDays(minD, maxD) + 1;
  const pxPerDay  = chartW / totalDays;

  // Choose tick interval based on zoom
  let tickDays = 7;
  if (pxPerDay * 7  < 30) tickDays = 14;
  if (pxPerDay * 14 < 30) tickDays = 30;
  if (pxPerDay * 30 < 30) tickDays = 60;
  if (pxPerDay * 60 < 30) tickDays = 90;

  const ticks = [];
  let cur = new Date(minD);
  cur.setDate(cur.getDate() - cur.getDay() + 1); // align to Monday
  while (cur <= maxD) {
    if (cur >= minD) {
      const x = Math.round(diffDays(minD, cur) * pxPerDay);
      ticks.push({ date: new Date(cur), x });
    }
    cur = addDays(cur, tickDays);
  }
  return { pxPerDay, ticks, totalDays };
}

// ─────────────────────────────────────────────────────────────
//  Bar component
// ─────────────────────────────────────────────────────────────
function Bar({ start, end, minD, pxPerDay, y, color, showBoth, isExp }) {
  if (!start || !end) return null;
  const x = Math.max(0, diffDays(minD, start) * pxPerDay);
  const w = Math.max(4, diffDays(start, end) * pxPerDay + pxPerDay);
  const barY = showBoth ? (isExp ? y + 4 : y + 14) : y + 6;
  const barH = showBoth ? 10 : 16;

  return (
    <g>
      <rect x={x} y={barY} width={w} height={barH} rx={3} fill={color} opacity={0.85} />
      {w > 40 && (
        <text x={x + 4} y={barY + barH - 3} fontSize={9} fill="#fff" style={{ pointerEvents: 'none', fontFamily: 'Roboto,sans-serif' }}>
          {fmtDate(start)}
        </text>
      )}
    </g>
  );
}

// ─────────────────────────────────────────────────────────────
//  Today line
// ─────────────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────────────
//  Main component
// ─────────────────────────────────────────────────────────────
export default function GanttChart() {
  const { activeProject, tasks, showToast } = useApp();

  // Toggles
  const [depthMode,    setDepthMode]    = useState('items');   // 'items' | 'tasks' | 'both'
  const [timelineMode, setTimelineMode] = useState('both');    // 'expected' | 'actual' | 'both'

  const chartRef = useRef(null);
  const [chartW,  setChartW] = useState(900);

  useEffect(() => {
    const obs = new ResizeObserver(entries => {
      const w = entries[0]?.contentRect?.width;
      if (w > LABEL_W + 100) setChartW(Math.max(MIN_CHART_W, w - LABEL_W - 32));
    });
    if (chartRef.current) obs.observe(chartRef.current);
    return () => obs.disconnect();
  }, []);

  const hierarchy = useMemo(() => buildHierarchy(tasks), [tasks]);

  // Global date range
  const { globalMin, globalMax } = useMemo(() => {
    const all = [];
    hierarchy.forEach(p => {
      if (p.expStart) all.push(p.expStart);
      if (p.expEnd)   all.push(p.expEnd);
      if (p.actStart) all.push(p.actStart);
      if (p.actEnd)   all.push(p.actEnd);
    });
    if (!all.length) return { globalMin: null, globalMax: null };
    const mn = new Date(Math.min(...all));
    const mx = new Date(Math.max(...all));
    return {
      globalMin: addDays(mn, -7),
      globalMax: addDays(mx,  7),
    };
  }, [hierarchy]);

  const { pxPerDay, ticks } = useMemo(
    () => globalMin && globalMax ? buildAxis(globalMin, globalMax, chartW) : { pxPerDay: 0, ticks: [], totalDays: 0 },
    [globalMin, globalMax, chartW]
  );

  const showExp = timelineMode === 'expected' || timelineMode === 'both';
  const showAct = timelineMode === 'actual'   || timelineMode === 'both';
  const showBoth = showExp && showAct;

  // Build flat rows for SVG rendering
  const rows = useMemo(() => {
    const out = [];
    hierarchy.forEach(phase => {
      out.push({ type: 'phase', data: phase });
      if (depthMode === 'items' || depthMode === 'both') {
        phase.items.forEach(item => {
          out.push({ type: 'item', data: item, phase: phase.phase });
          if (depthMode === 'both') {
            item.tasks.forEach(task => {
              out.push({ type: 'task', data: task, phase: phase.phase, item: item.item });
            });
          }
        });
      } else if (depthMode === 'tasks') {
        phase.items.forEach(item => {
          item.tasks.forEach(task => {
            out.push({ type: 'task', data: task, phase: phase.phase, item: item.item });
          });
        });
      }
    });
    return out;
  }, [hierarchy, depthMode]);

  const totalH = rows.length * ROW_H;
  const AXIS_H = 36;
  const devCount = tasks.filter(t => t.taskType === 'Dev Task').length;

  // ── Export: Excel ─────────────────────────────────────────
  // Hooks must all be declared before any early returns (React rules of hooks)
  const exportXlsx = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const wsData = [
      ['Phase', 'Item', 'Task', 'Exp. Start', 'Exp. End', 'Act. Start', 'Act. End', 'Owner Status'],
    ];
    hierarchy.forEach(phase => {
      wsData.push([phase.phase, '', '', fmtDate(phase.expStart), fmtDate(phase.expEnd), fmtDate(phase.actStart), fmtDate(phase.actEnd), '']);
      phase.items.forEach(item => {
        wsData.push(['', item.item, '', fmtDate(item.expStart), fmtDate(item.expEnd), fmtDate(item.actStart), fmtDate(item.actEnd), '']);
        item.tasks.forEach(t => {
          wsData.push(['', '', t.task, t.expectedStart || '', t.expectedEnd || '', t.actualStart || '', t.actualEnd || '', t.ownerStatus]);
        });
      });
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [26, 26, 50, 12, 12, 12, 12, 14].map(wch => ({ wch }));

    // Style header row bold
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[cellRef]) {
        ws[cellRef].s = { font: { bold: true }, fill: { patternType: 'solid', fgColor: { rgb: 'E8EAF6' } } };
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Gantt');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Gantt.xlsx`);
    showToast('Gantt exported to Excel');
  }, [hierarchy, activeProject, showToast]);

  // ── Export: PowerPoint ────────────────────────────────────
  const exportPptx = useCallback(async () => {
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';   // 13.33" × 7.5"
      pptx.title  = `${activeProject.name} — Gantt Chart`;

      const SLIDE_W  = 13.33;  // inches
      const SLIDE_H  = 7.5;
      const LBL_W    = 3.2;    // inches for label column
      const BAR_X0   = LBL_W + 0.15;
      const BAR_AREA = SLIDE_W - BAR_X0 - 0.2;
      const AXIS_Y   = 0.85;
      const ROW_H_IN = 0.28;
      const START_Y  = AXIS_Y + 0.32;

      if (!globalMin || !globalMax) { showToast('No date data to export', 'error'); return; }
      const totalD = diffDays(globalMin, globalMax) + 1;
      const pxPD   = BAR_AREA / totalD;

      // Helper: date -> x inches
      const xIn = (d) => d ? BAR_X0 + diffDays(globalMin, d) * pxPD : BAR_X0;

      const slide = pptx.addSlide();

      // Background
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: 'F8F9FE' }, line: { color: 'F8F9FE' } });

      // Title
      slide.addText(`${activeProject.name} — Gantt Chart`, {
        x: 0.2, y: 0.08, w: SLIDE_W - 0.4, h: 0.35,
        fontSize: 16, bold: true, color: '404789', fontFace: 'Calibri',
      });
      slide.addText(`Dev Tasks · Generated ${new Date().toLocaleDateString()}`, {
        x: 0.2, y: 0.42, w: SLIDE_W - 0.4, h: 0.2,
        fontSize: 9, color: '888888', fontFace: 'Calibri',
      });

      // Legend
      slide.addShape(pptx.ShapeType.rect, { x: SLIDE_W - 2.8, y: 0.1, w: 0.18, h: 0.14, fill: { color: '404789' }, line: { color: '404789' } });
      slide.addText('Expected', { x: SLIDE_W - 2.55, y: 0.08, w: 0.9, h: 0.18, fontSize: 8, color: '404789', fontFace: 'Calibri' });
      slide.addShape(pptx.ShapeType.rect, { x: SLIDE_W - 1.5, y: 0.1, w: 0.18, h: 0.14, fill: { color: 'DA9B38' }, line: { color: 'DA9B38' } });
      slide.addText('Actual', { x: SLIDE_W - 1.25, y: 0.08, w: 0.8, h: 0.18, fontSize: 8, color: 'DA9B38', fontFace: 'Calibri' });

      // Axis ticks
      ticks.forEach(tick => {
        const tx = BAR_X0 + tick.x / (chartW / BAR_AREA);
        if (tx < BAR_X0 || tx > SLIDE_W - 0.2) return;
        slide.addText(tick.date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }), {
          x: tx - 0.2, y: AXIS_Y, w: 0.4, h: 0.2,
          fontSize: 7, color: '888888', align: 'center', fontFace: 'Calibri',
        });
        slide.addShape(pptx.ShapeType.line, { x: tx, y: AXIS_Y + 0.2, w: 0, h: SLIDE_H - AXIS_Y - 0.2, line: { color: 'E0E0E0', width: 0.5 } });
      });

      // Today line
      const today = new Date();
      if (today >= globalMin && today <= globalMax) {
        const tx = BAR_X0 + diffDays(globalMin, today) * pxPD;
        slide.addShape(pptx.ShapeType.line, { x: tx, y: AXIS_Y, w: 0, h: SLIDE_H - AXIS_Y - 0.1, line: { color: 'E53935', width: 1.5, dashType: 'dash' } });
        slide.addText('Today', { x: tx - 0.18, y: AXIS_Y - 0.02, w: 0.36, h: 0.14, fontSize: 7, color: 'E53935', align: 'center', fontFace: 'Calibri' });
      }

      // Rows — auto-paginate when content exceeds slide height
      const MAX_Y = SLIDE_H - 0.15;  // bottom margin
      let curY = START_Y;
      let curSlide = slide;

      // Helper: start a new slide with repeated header + axis
      const addContinuationSlide = () => {
        const s = pptx.addSlide();
        s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: 'F8F9FE' }, line: { color: 'F8F9FE' } });
        s.addText(`${activeProject.name} — Gantt Chart (cont.)`, {
          x: 0.2, y: 0.08, w: SLIDE_W - 0.4, h: 0.25,
          fontSize: 12, bold: true, color: '404789', fontFace: 'Calibri',
        });
        ticks.forEach(tk => {
          const tx = BAR_X0 + tk.x / (chartW / BAR_AREA);
          if (tx < BAR_X0 || tx > SLIDE_W - 0.2) return;
          s.addText(tk.date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }), {
            x: tx - 0.2, y: 0.3, w: 0.4, h: 0.18,
            fontSize: 7, color: '888888', align: 'center', fontFace: 'Calibri',
          });
          s.addShape(pptx.ShapeType.line, { x: tx, y: 0.48, w: 0, h: SLIDE_H - 0.48 - 0.1, line: { color: 'E0E0E0', width: 0.5 } });
        });
        return s;
      };

      const rowsToRender = [];
      hierarchy.forEach(phase => {
        rowsToRender.push({ type: 'phase', data: phase });
        phase.items.forEach(item => {
          rowsToRender.push({ type: 'item', data: item });
          item.tasks.forEach(t => rowsToRender.push({ type: 'task', data: t }));
        });
      });

      rowsToRender.forEach(row => {
        const { type, data } = row;
        const isPhase = type === 'phase';
        const isItem  = type === 'item';

        // Start a new slide if this row would overflow
        if (curY + ROW_H_IN > MAX_Y) {
          curSlide = addContinuationSlide();
          curY = 0.5;
        }

        // Row background for phase
        if (isPhase) {
          curSlide.addShape(pptx.ShapeType.rect, { x: 0.1, y: curY, w: SLIDE_W - 0.2, h: ROW_H_IN, fill: { color: 'E8EAF6' }, line: { color: 'D0D4F5' } });
        }

        // Label
        const label = isPhase ? data.phase : isItem ? `  ${data.item}` : `    ${data.task}`;
        curSlide.addText(label, {
          x: 0.12, y: curY + 0.04, w: LBL_W - 0.1, h: ROW_H_IN - 0.06,
          fontSize: isPhase ? 9 : isItem ? 8 : 7,
          bold: isPhase,
          color: isPhase ? '404789' : '404041',
          fontFace: 'Calibri',
        });

        // Expected bar
        if (showExp && data.expStart && data.expEnd) {
          const bx = xIn(data.expStart);
          const bw = Math.max(0.05, xIn(data.expEnd) - bx + pxPD);
          const by = showBoth ? curY + 0.03 : curY + 0.06;
          const bh = showBoth ? ROW_H_IN * 0.42 : ROW_H_IN * 0.6;
          curSlide.addShape(pptx.ShapeType.rect, { x: bx, y: by, w: bw, h: bh, fill: { color: isPhase ? '2E3F8F' : '404789' }, line: { color: '404789' } });
        }
        // Actual bar
        if (showAct && data.actStart && data.actEnd) {
          const bx = xIn(data.actStart);
          const bw = Math.max(0.05, xIn(data.actEnd) - bx + pxPD);
          const by = showBoth ? curY + ROW_H_IN * 0.5 : curY + 0.06;
          const bh = showBoth ? ROW_H_IN * 0.42 : ROW_H_IN * 0.6;
          curSlide.addShape(pptx.ShapeType.rect, { x: bx, y: by, w: bw, h: bh, fill: { color: isPhase ? 'B8770A' : 'DA9B38' }, line: { color: 'DA9B38' } });
        }

        curY += ROW_H_IN;
      });

      await pptx.writeFile({ fileName: `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Gantt.pptx` });
      showToast('Gantt exported to PowerPoint');
    } catch (err) {
      console.error('PPTX export error:', err);
      showToast('PPTX export failed', 'error');
    }
  }, [hierarchy, globalMin, globalMax, chartW, ticks, showExp, showAct, showBoth, activeProject, showToast]);

  // ── Early returns — must come AFTER all hooks ─────────────
  if (!activeProject) return <EmptyState message="No project selected." />;

  if (devCount === 0) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
        <GanttHeader activeProject={activeProject} devCount={0}
          depthMode={depthMode} setDepthMode={setDepthMode}
          timelineMode={timelineMode} setTimelineMode={setTimelineMode}
          onExportXlsx={exportXlsx} onExportPptx={exportPptx} showToast={showToast} />
        <EmptyState message="No Dev Tasks found. Mark tasks as 'Dev Task' type in the Tasks & Checklist page." />
      </div>
    );
  }

  // ─────────────────────────────────────────────────────────
  //  Render
  // ─────────────────────────────────────────────────────────
  return (
    <div ref={chartRef} style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden', background: '#f8f9fe' }}>
      <GanttHeader
        activeProject={activeProject} devCount={devCount}
        depthMode={depthMode} setDepthMode={setDepthMode}
        timelineMode={timelineMode} setTimelineMode={setTimelineMode}
        onExportXlsx={exportXlsx} onExportPptx={exportPptx}
        showToast={showToast}
      />

      {/* Legend */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '6px 20px', display: 'flex', alignItems: 'center', gap: 20, flexShrink: 0 }}>
        {showExp && <LegendItem color={EXP_COLOR} label="Expected" />}
        {showAct && <LegendItem color={ACT_COLOR} label="Actual" />}
        <span style={{ fontSize: 11, color: '#888', marginLeft: 8 }}>
          {devCount} dev tasks · {hierarchy.length} phases · {hierarchy.reduce((s, p) => s + p.items.length, 0)} items
        </span>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ width: 10, borderTop: '2px dashed #e53935', display: 'inline-block' }} />
          <span style={{ fontSize: 10, color: '#e53935' }}>Today</span>
        </div>
      </div>

      {/* Gantt table */}
      {!globalMin ? (
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999', fontSize: 13 }}>
          No date information found on Dev Tasks. Add Expected Start/End dates in the Tasks &amp; Checklist page.
        </div>
      ) : (
        <div style={{ flex: 1, overflow: 'auto', display: 'flex' }}>
          {/* Left label column — sticky so it stays visible on horizontal scroll */}
          <div style={{ width: LABEL_W, minWidth: LABEL_W, flexShrink: 0, background: '#fff', borderRight: '2px solid #e0e4f0', position: 'sticky', left: 0, zIndex: 4 }}>
            {/* Axis placeholder */}
            <div style={{ height: AXIS_H, background: '#f0f1f8', borderBottom: '1px solid #d0d0d0', display: 'flex', alignItems: 'center', paddingLeft: 14 }}>
              <span style={{ fontSize: 11, fontWeight: 600, color: '#404789' }}>Task / Item / Phase</span>
            </div>
            {rows.map((row, i) => {
              const { type, data } = row;
              const isPhase = type === 'phase';
              const isItem  = type === 'item';
              const isTask  = type === 'task';
              const label   = isPhase ? data.phase : isItem ? data.item : data.task;
              return (
                <div key={i} style={{
                  height: ROW_H,
                  background: isPhase ? PHASE_BG : isItem ? ITEM_BG : '#fff',
                  borderBottom: '1px solid #f0f0f0',
                  display: 'flex', alignItems: 'center',
                  paddingLeft: isPhase ? 10 : isItem ? 22 : 36,
                  paddingRight: 8,
                  fontWeight: isPhase ? 600 : 400,
                  fontSize: isPhase ? 12 : isItem ? 11 : 10,
                  color: isPhase ? PHASE_FG : '#404041',
                  borderLeft: isPhase ? `3px solid ${PHASE_FG}` : isItem ? `3px solid ${EXP_LIGHT}` : '3px solid transparent',
                  overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis',
                  gap: 6,
                }}>
                  {isPhase && <span style={{ fontSize: 10 }}>▶</span>}
                  {isItem  && <span style={{ fontSize: 9, color: '#aaa' }}>├</span>}
                  {isTask  && <span style={{ fontSize: 9, color: '#ccc' }}>└</span>}
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    title={label}>{label}</span>
                </div>
              );
            })}
          </div>

          {/* Right chart area */}
          <div style={{ flex: 1, minWidth: 0 }}>
            <svg
              width={chartW}
              height={AXIS_H + totalH + 1}
              style={{ display: 'block', fontFamily: 'Roboto,sans-serif' }}
            >
              {/* Axis background */}
              <rect x={0} y={0} width={chartW} height={AXIS_H} fill="#f0f1f8" />

              {/* Tick lines & labels */}
              {ticks.map((tick, i) => (
                <g key={i}>
                  <line x1={tick.x} y1={AXIS_H - 6} x2={tick.x} y2={AXIS_H} stroke="#c0c4d4" />
                  <line x1={tick.x} y1={AXIS_H} x2={tick.x} y2={AXIS_H + totalH} stroke="#eaeaf4" />
                  <text x={tick.x + 3} y={14} fontSize={9} fill="#666" fontFamily="Roboto,sans-serif">
                    {tick.date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' })}
                  </text>
                  <text x={tick.x + 3} y={26} fontSize={8} fill="#aaa" fontFamily="Roboto,sans-serif">
                    {tick.date.getFullYear()}
                  </text>
                </g>
              ))}
              <line x1={0} y1={AXIS_H} x2={chartW} y2={AXIS_H} stroke="#d0d4e4" />

              {/* Row backgrounds + bars */}
              {rows.map((row, i) => {
                const { type, data } = row;
                const isPhase = type === 'phase';
                const y = AXIS_H + i * ROW_H;
                return (
                  <g key={i}>
                    {/* Row stripe */}
                    <rect x={0} y={y} width={chartW} height={ROW_H}
                      fill={isPhase ? '#edeef8' : i % 2 === 0 ? '#fafafe' : '#fff'} />
                    <line x1={0} y1={y + ROW_H} x2={chartW} y2={y + ROW_H} stroke="#f0f0f0" />

                    {/* Expected bar */}
                    {showExp && (
                      <Bar start={data.expStart} end={data.expEnd}
                        minD={globalMin} pxPerDay={pxPerDay}
                        y={y} color={isPhase ? '#2e3f8f' : EXP_COLOR}
                        showBoth={showBoth} isExp={true} />
                    )}
                    {/* Actual bar */}
                    {showAct && (
                      <Bar start={data.actStart} end={data.actEnd}
                        minD={globalMin} pxPerDay={pxPerDay}
                        y={y} color={isPhase ? '#b8770a' : ACT_COLOR}
                        showBoth={showBoth} isExp={false} />
                    )}
                  </g>
                );
              })}

              {/* Today line on top */}
              <TodayLine minD={globalMin} pxPerDay={pxPerDay} totalH={AXIS_H + totalH} />
            </svg>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
//  Sub-components
// ─────────────────────────────────────────────────────────────
function LegendItem({ color, label }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
      <div style={{ width: 20, height: 10, borderRadius: 2, background: color }} />
      <span style={{ fontSize: 11, color: '#555' }}>{label}</span>
    </div>
  );
}

function ToggleGroup({ label, value, setValue, options }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
      <span style={{ fontSize: 11, color: '#666', whiteSpace: 'nowrap' }}>{label}:</span>
      <div style={{ display: 'flex', border: '1px solid #c0c4d8', borderRadius: 5, overflow: 'hidden' }}>
        {options.map(opt => (
          <button
            key={opt.value}
            onClick={() => setValue(opt.value)}
            style={{
              padding: '4px 10px', fontSize: 11, border: 'none', cursor: 'pointer',
              background: value === opt.value ? '#404789' : '#fff',
              color:      value === opt.value ? '#fff'    : '#555',
              fontFamily: 'Roboto,sans-serif',
              fontWeight: value === opt.value ? 600 : 400,
              borderRight: '1px solid #c0c4d8',
              transition: 'background 0.15s',
            }}
          >{opt.label}</button>
        ))}
      </div>
    </div>
  );
}

function GanttHeader({ activeProject, devCount, depthMode, setDepthMode, timelineMode, setTimelineMode, onExportXlsx, onExportPptx }) {
  return (
    <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '10px 20px', flexShrink: 0 }}>
      {/* Breadcrumb + title */}
      <div style={{ fontSize: 11, color: '#888', marginBottom: 4 }}>
        Implementation Hub › {activeProject?.name} › Gantt Chart
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <div>
          <div style={{ fontSize: 18, fontWeight: 700, color: '#404041' }}>Gantt Chart</div>
          <div style={{ fontSize: 11, color: '#888' }}>{devCount} Dev Tasks · {activeProject?.name}</div>
        </div>

        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
          {/* Toggle 1: depth */}
          <ToggleGroup
            label="Show"
            value={depthMode}
            setValue={setDepthMode}
            options={[
              { value: 'items', label: 'Items' },
              { value: 'tasks', label: 'Tasks' },
              { value: 'both',  label: 'Both'  },
            ]}
          />
          {/* Toggle 2: timeline */}
          <ToggleGroup
            label="Timeline"
            value={timelineMode}
            setValue={setTimelineMode}
            options={[
              { value: 'expected', label: 'Expected' },
              { value: 'actual',   label: 'Actual'   },
              { value: 'both',     label: 'Both'     },
            ]}
          />
          {/* Exports */}
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={onExportXlsx}
              style={{ padding: '5px 12px', fontSize: 11, border: '1px solid #c0c4d8', borderRadius: 4, background: '#fff', cursor: 'pointer', color: '#404041', fontFamily: 'Roboto,sans-serif' }}>
              ↓ Excel
            </button>
            <button onClick={onExportPptx}
              style={{ padding: '5px 12px', fontSize: 11, border: '1px solid #c0c4d8', borderRadius: 4, background: '#fff', cursor: 'pointer', color: '#404041', fontFamily: 'Roboto,sans-serif' }}>
              ↓ PowerPoint
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
