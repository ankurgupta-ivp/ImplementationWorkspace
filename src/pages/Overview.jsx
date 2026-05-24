import React, { useState } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, KpiCard, ProgressBar, WidgetCard, Modal, Field, Input, Select, EmptyState } from '../components/UI';

const PHASES = ['Project Creation','Requirement Gathering','Implementation','UAT','Prod Live','Support handover'];
const PHASE_LABELS = {
  'Project Creation':   'Project Setup',
  'Requirement Gathering': 'Requirements',
  'Implementation':     'Dev & Config',
  'UAT':                'UAT',
  'Prod Live':          'Prod Live',
  'Support handover':   'Handover',
};

// ── Helpers ───────────────────────────────────────────────────
function countAnswered(proj) {
  return (proj?.questionnaire?.sections || []).reduce(
    (s, sec) => s + (sec.questions || []).filter(q => q.answer?.trim()).length, 0
  );
}
function countTotal(proj) {
  return (proj?.questionnaire?.sections || []).reduce(
    (s, sec) => s + (sec.questions || []).length, 0
  );
}
function calcEstimate(proj) {
  const est = proj?.estimator;
  if (!est?.stepsBase) return 0;
  return Math.round(est.stepsBase.reduce((s, step) => s + (step.base || 0), 0));
}

function parseD(s) {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function fmtShort(d) {
  if (!d) return null;
  return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'2-digit' });
}
function diffDays(a, b) {
  return Math.round((b - a) / 86400000);
}

// Filter tasks by the chosen type toggle
function filterByType(tasks, typeFilter) {
  if (typeFilter === 'dev')       return tasks.filter(t => t.taskType === 'Dev Task');
  if (typeFilter === 'checklist') return tasks.filter(t => t.taskType !== 'Dev Task');
  return tasks; // 'all'
}

// Derive date stats for a set of tasks in one phase
function phaseDateStats(phaseTasks) {
  const dates = arr => arr.map(parseD).filter(Boolean);
  const expStarts = dates(phaseTasks.map(t => t.expectedStart));
  const expEnds   = dates(phaseTasks.map(t => t.expectedEnd));
  const actStarts = dates(phaseTasks.map(t => t.actualStart));
  const actEnds   = dates(phaseTasks.map(t => t.actualEnd));

  const expStart = expStarts.length ? new Date(Math.min(...expStarts)) : null;
  const expEnd   = expEnds.length   ? new Date(Math.max(...expEnds))   : null;
  const actStart = actStarts.length ? new Date(Math.min(...actStarts)) : null;
  const actEnd   = actEnds.length   ? new Date(Math.max(...actEnds))   : null;

  // Delay: positive = delayed by N days, negative = finished early
  let delayDays = null;
  if (expEnd && actEnd) {
    delayDays = diffDays(expEnd, actEnd); // actEnd - expEnd
  } else if (expEnd && !actEnd) {
    // Phase not finished — check if overdue
    const today = new Date();
    today.setHours(0,0,0,0);
    if (today > expEnd) delayDays = diffDays(expEnd, today);
  }

  return { expStart, expEnd, actStart, actEnd, delayDays };
}

// RAID helpers
function isOverdue(item) {
  if (['Completed','Closed'].includes(item.status)) return false;
  if (!item.eta) return false;
  const d = parseD(item.eta);
  if (!d) return false;
  const today = new Date(); today.setHours(0,0,0,0);
  return d < today;
}

// ── TaskType Toggle ────────────────────────────────────────────
function TypeToggle({ value, onChange }) {
  const opts = [
    { key: 'dev',       label: 'Dev Tasks' },
    { key: 'checklist', label: 'Checklist' },
    { key: 'all',       label: 'All' },
  ];
  return (
    <div style={{ display:'flex', alignItems:'center', gap:6 }}>
      <span style={{ fontSize:10, color:'#888', whiteSpace:'nowrap' }}>Show:</span>
      <div style={{ display:'flex', border:'1px solid #c0c4d8', borderRadius:5, overflow:'hidden' }}>
        {opts.map(o => (
          <button key={o.key} onClick={() => onChange(o.key)} style={{
            padding:'3px 10px', fontSize:11, border:'none', cursor:'pointer',
            background: value === o.key ? '#404789' : '#fff',
            color:      value === o.key ? '#fff' : '#555',
            fontFamily:'Roboto,sans-serif', fontWeight: value === o.key ? 600 : 400,
            borderRight:'1px solid #c0c4d8', transition:'background 0.15s',
          }}>{o.label}</button>
        ))}
      </div>
    </div>
  );
}

// ── Main component ─────────────────────────────────────────────
export default function Overview() {
  const { activeProject, tasks, raidItems, updateProject, setCurrentPage, showToast } = useApp();
  const [editOpen,     setEditOpen]     = useState(false);
  const [form,         setForm]         = useState({});
  const [typeFilter,   setTypeFilter]   = useState('dev');

  if (!activeProject) return <EmptyState message="No project selected. Create one to get started." />;

  // Filtered task set for all KPIs
  const ft = filterByType(tasks, typeFilter);

  // ── KPI metrics ───────────────────────────────────────────────
  const total   = ft.length;
  const done    = ft.filter(t => t.ownerStatus === 'Done').length;
  const inprog  = ft.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = ft.filter(t => ['Blocked','Delayed'].includes(t.ownerStatus)).length;
  const progress = total ? Math.round(done / total * 100) : 0;

  const effortSpent   = ft.reduce((s, t) => s + (parseFloat(t.actualEffort)   || 0), 0);
  const effortPlanned = ft.reduce((s, t) => s + (parseFloat(t.expectedEffort) || 0), 0);
  const effortPct     = effortPlanned > 0 ? Math.round(effortSpent / effortPlanned * 100) : null;
  const effortOver    = effortPct !== null && effortPct > 100;

  // ── Phase stats ───────────────────────────────────────────────
  const phaseStats = PHASES.map(ph => {
    const ts  = ft.filter(t => t.phase === ph);
    const dn  = ts.filter(t => t.ownerStatus === 'Done').length;
    const ip  = ts.filter(t => t.ownerStatus === 'In Progress').length;
    const blk = ts.filter(t => ['Blocked','Delayed'].includes(t.ownerStatus)).length;
    const status = ts.length === 0 ? 'notstart'
      : dn === ts.length ? 'done'
      : (ip > 0 || dn > 0) ? 'active'
      : 'notstart';
    const dates = phaseDateStats(ts);
    return { phase: ph, label: PHASE_LABELS[ph], total: ts.length, done: dn, inprog: ip, blocked: blk, status, ...dates };
  });

  // ── RAID counts ───────────────────────────────────────────────
  const openRaids = raidItems.filter(r => !['Completed','Closed'].includes(r.status));
  const today = new Date(); today.setHours(0,0,0,0);

  const raidCount = (classification) => {
    const items    = openRaids.filter(r => r.classification === classification);
    const overdue  = items.filter(r => isOverdue(r));
    return { total: items.length, overdue: overdue.length };
  };
  const riskStats = raidCount('Risk');
  const depsStats = raidCount('Dependency');
  const issuStats = raidCount('Issue');

  // ── Edit helpers ──────────────────────────────────────────────
  const openEdit = () => { setForm({ ...activeProject.metadata }); setEditOpen(true); };
  const saveEdit = async () => { await updateProject({ ...activeProject, metadata: form }); setEditOpen(false); showToast('Project details saved'); };
  const md = activeProject.metadata || {};

  // ── Render ─────────────────────────────────────────────────────
  return (
    <div style={{ display:'flex', flexDirection:'column', flex:1, overflow:'hidden' }}>
      <PageHeader
        breadcrumb={<>Implementation Hub › {activeProject.name} › Overview</>}
        title={activeProject.name}
        subtitle={`Kickoff: ${md.kickoffDate || '—'} · Target Go-Live: ${md.targetGoLive || '—'} · Lead BA: ${md.leadBA || '—'} · Engineers: ${md.engineers || '—'}`}
        actions={
          <div style={{ display:'flex', gap:8, alignItems:'center' }}>
            <button className="btn btn-ghost" onClick={openEdit}>✎ Edit Project Details</button>
            <button className="btn btn-primary" onClick={() => setCurrentPage('tasks')}>→ Go to Tasks</button>
          </div>
        }
      />

      <div style={{ flex:1, overflowY:'auto', padding:'18px 20px' }}>

        {/* Task type toggle */}
        <div style={{ display:'flex', justifyContent:'flex-end', marginBottom:12 }}>
          <TypeToggle value={typeFilter} onChange={setTypeFilter} />
        </div>

        {/* KPI row */}
        <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:14, marginBottom:16 }}>
          <KpiCard
            label="Overall Progress"
            value={`${progress}%`}
            delta={`${done} of ${total} tasks done`}
            variant={progress >= 80 ? 'ok' : progress >= 50 ? 'default' : 'warn'}
          />
          <KpiCard
            label="In Progress"
            value={inprog}
            delta="tasks active"
          />
          <KpiCard
            label="Blocked / Delayed"
            value={blocked}
            delta={`${ft.filter(t => t.ownerStatus === 'Blocked').length} blocked`}
            variant={blocked > 0 ? 'fail' : 'default'}
          />
          {/* Effort card — days, % spent */}
          <div style={{ background:'#fff', borderRadius:8, padding:'14px 16px', boxShadow:'0 1px 4px rgba(0,0,0,.08)' }}>
            <div style={{ fontSize:11, color:'#888', marginBottom:4 }}>Effort</div>
            <div style={{ fontSize:22, fontWeight:700, color:'#404041', lineHeight:1.1 }}>
              {effortSpent}d
            </div>
            <div style={{ fontSize:11, color:'#888', marginTop:4 }}>vs {effortPlanned}d planned</div>
            {effortPct !== null && (
              <div style={{ marginTop:6, display:'flex', alignItems:'center', gap:6 }}>
                <div style={{ flex:1, height:6, background:'#f0f0f0', borderRadius:3, overflow:'hidden' }}>
                  <div style={{ width:`${Math.min(effortPct, 100)}%`, height:'100%', borderRadius:3, background: effortOver ? '#e53935' : '#4caf50', transition:'width .4s' }} />
                </div>
                <span style={{ fontSize:11, fontWeight:600, color: effortOver ? '#e53935' : '#2e7d32', whiteSpace:'nowrap' }}>
                  {effortPct}% {effortOver ? '▲ over' : '▼ within'}
                </span>
              </div>
            )}
          </div>
        </div>

        {/* Phase timeline */}
        <div style={{ background:'#fff', borderRadius:8, boxShadow:'0 1px 4px rgba(0,0,0,.08)', marginBottom:16, overflow:'hidden' }}>
          <div style={{ display:'flex', flexWrap:'nowrap', overflowX:'auto' }}>
            {phaseStats.map(p => {
              const pct      = p.total ? Math.round(p.done / p.total * 100) : 0;
              const isDelayed = p.delayDays !== null && p.delayDays > 0;
              const isEarly   = p.delayDays !== null && p.delayDays < 0;
              const bgColor   = p.status === 'done' ? '#f1f8e9' : p.status === 'active' ? '#f0f4ff' : '#fff';
              const borderClr = p.status === 'done' ? '#4caf50' : p.status === 'active' ? '#404789' : '#e0e0e0';

              return (
                <div key={p.phase} style={{
                  flex:'0 0 auto', minWidth:160, padding:'11px 13px',
                  borderRight:'1px solid #f0f0f0', borderBottom:`3px solid ${borderClr}`,
                  background: bgColor,
                }}>
                  {/* Phase label + status chip */}
                  <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:3 }}>
                    <span style={{ fontSize:10, color:'#999', textTransform:'uppercase', letterSpacing:'.4px' }}>{p.label}</span>
                    {isDelayed && (
                      <span style={{ fontSize:9, fontWeight:600, padding:'1px 5px', borderRadius:3, background:'#fce4ec', color:'#c62828', whiteSpace:'nowrap' }}>
                        +{p.delayDays}d late
                      </span>
                    )}
                    {isEarly && (
                      <span style={{ fontSize:9, fontWeight:600, padding:'1px 5px', borderRadius:3, background:'#e8f5e9', color:'#2e7d32', whiteSpace:'nowrap' }}>
                        {Math.abs(p.delayDays)}d early
                      </span>
                    )}
                  </div>

                  {/* Progress */}
                  <div style={{ fontSize:12, fontWeight:500, color:'#404041', marginBottom:5 }}>
                    {p.done}/{p.total} done
                    {p.blocked > 0 && <span style={{ fontSize:10, color:'#e53935', marginLeft:6 }}>⚑ {p.blocked}</span>}
                  </div>
                  <ProgressBar pct={pct} />

                  {/* Date rows */}
                  {(p.expStart || p.expEnd) && (
                    <div style={{ marginTop:7, fontSize:10, color:'#888' }}>
                      <span style={{ color:'#404789', fontWeight:500 }}>Exp: </span>
                      {fmtShort(p.expStart) || '?'} → {fmtShort(p.expEnd) || '?'}
                    </div>
                  )}
                  {(p.actStart || p.actEnd) && (
                    <div style={{ marginTop:2, fontSize:10, color: isDelayed ? '#c62828' : '#555' }}>
                      <span style={{ fontWeight:500 }}>Act: </span>
                      {fmtShort(p.actStart) || '?'} → {fmtShort(p.actEnd) || '?'}
                    </div>
                  )}
                  {!p.expStart && !p.actStart && (
                    <div style={{ marginTop:7, fontSize:10, color:'#ccc', fontStyle:'italic' }}>No dates set</div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Quick Access + RAID */}
        <div style={{ display:'grid', gridTemplateColumns:'2fr 1fr', gap:14 }}>

          {/* Quick Section Access — Template Library removed */}
          <WidgetCard title="Quick Section Access">
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
              {[
                { page:'kickoff',   icon:'⊕', label:'Kickoff Questionnaire', meta:`${countAnswered(activeProject)} of ${countTotal(activeProject)} answered` },
                { page:'estimator', icon:'⏱', label:'Estimator',             meta:`${calcEstimate(activeProject)}d total estimate` },
                { page:'tasks',     icon:'☑', label:'Tasks & Checklist',     meta:`${total} items · ${progress}% complete` },
                { page:'raidlog',   icon:'⚑', label:'RAID Log',              meta:`${raidItems.length} items tracked` },
                { page:'gantt',     icon:'📊', label:'Gantt Chart',          meta:'Dev task timeline' },
              ].map(s => (
                <button key={s.page} className="btn btn-outlined"
                  onClick={() => setCurrentPage(s.page)}
                  style={{ justifyContent:'flex-start', padding:14, fontSize:13, textAlign:'left', display:'block' }}>
                  {s.icon} {s.label}
                  <span style={{ fontSize:11, color:'#888', fontWeight:400, marginTop:4, display:'block' }}>{s.meta}</span>
                </button>
              ))}
            </div>
          </WidgetCard>

          {/* RAID summary with breakdown by type */}
          <WidgetCard title="Open RAID Items">
            {/* Counts by classification */}
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8, marginBottom:12 }}>
              {[
                { label:'Risks',        ...riskStats, color:'#c62828', bg:'#fce4ec' },
                { label:'Dependencies', ...depsStats, color:'#e65100', bg:'#fff3e0' },
                { label:'Issues',       ...issuStats, color:'#1565c0', bg:'#e3f2fd' },
              ].map(c => (
                <div key={c.label} style={{ background:c.bg, borderRadius:6, padding:'8px 10px', textAlign:'center' }}>
                  <div style={{ fontSize:18, fontWeight:700, color:c.color }}>{c.total}</div>
                  <div style={{ fontSize:10, color:c.color, fontWeight:500 }}>{c.label}</div>
                  {c.overdue > 0 && (
                    <div style={{ fontSize:10, color:'#e53935', marginTop:2, fontWeight:600 }}>
                      {c.overdue} overdue
                    </div>
                  )}
                  {c.overdue === 0 && c.total > 0 && (
                    <div style={{ fontSize:10, color:'#888', marginTop:2 }}>none overdue</div>
                  )}
                </div>
              ))}
            </div>

            {/* Recent open items */}
            {openRaids.slice(0, 4).map(r => (
              <div key={r.id} style={{ padding:'7px 0', borderBottom:'1px solid #f5f5f5', fontSize:12 }}>
                <div style={{ fontWeight:500, color:'#404041', marginBottom:2, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.item}</div>
                <div style={{ display:'flex', alignItems:'center', gap:8, color:'#888', fontSize:11 }}>
                  <span>{r.classification}</span>
                  <span>·</span>
                  <span>{r.status}</span>
                  {isOverdue(r) && <span style={{ color:'#e53935', fontWeight:600 }}>· overdue</span>}
                </div>
              </div>
            ))}
            {openRaids.length === 0 && (
              <div style={{ color:'#888', fontSize:12, textAlign:'center', padding:'20px 0' }}>No open RAID items 🎉</div>
            )}

            <button className="btn btn-ghost"
              style={{ marginTop:10, width:'100%', fontSize:11 }}
              onClick={() => setCurrentPage('raidlog')}>
              View All RAID Items →
            </button>
          </WidgetCard>
        </div>
      </div>

      {/* Edit Project Modal */}
      {editOpen && (
        <Modal title="Edit Project Details" onClose={() => setEditOpen(false)} onSave={saveEdit}>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'14px 20px' }}>
            {[
              { key:'kickoffDate',  label:'Kickoff Date',  type:'date' },
              { key:'targetGoLive', label:'Target Go-Live', type:'date' },
              { key:'leadBA',       label:'Lead BA',        type:'text' },
              { key:'engineers',    label:'Engineers',      type:'text' },
              { key:'currentPhase', label:'Current Phase',  type:'select', options:PHASES },
            ].map(f => (
              <Field key={f.key} label={f.label}>
                {f.type === 'select'
                  ? <Select value={form[f.key] || ''} onChange={e => setForm(p => ({ ...p, [f.key]: e.target.value }))}>
                      {f.options.map(o => <option key={o}>{o}</option>)}
                    </Select>
                  : <Input type={f.type} value={form[f.key] || ''} onChange={e => setForm(p => ({ ...p, [f.key]: e.target.value }))} />
                }
              </Field>
            ))}
          </div>
        </Modal>
      )}
    </div>
  );
}
