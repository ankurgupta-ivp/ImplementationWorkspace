import React, { useState } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, KpiCard, ProgressBar, WidgetCard, Modal, Field, Input, Select, EmptyState } from '../components/UI';

const PHASES = ['Project Creation', 'Requirement Gathering', 'Implementation', 'UAT', 'Prod Live', 'Support handover'];
const PHASE_LABELS = { 'Project Creation': 'Project Setup', 'Requirement Gathering': 'Requirements', 'Implementation': 'Dev & Config', 'UAT': 'UAT', 'Prod Live': 'Prod Live', 'Support handover': 'Handover' };

function countAnswered(proj) {
  return (proj?.questionnaire?.sections || []).reduce((s, sec) => s + (sec.questions || []).filter(q => q.answer?.trim()).length, 0);
}
function countTotal(proj) {
  return (proj?.questionnaire?.sections || []).reduce((s, sec) => s + (sec.questions || []).length, 0);
}
function calcEstimate(proj) {
  const est = proj?.estimator;
  if (!est?.stepsBase) return 0;
  return Math.round(est.stepsBase.reduce((s, step) => s + (step.base || 0), 0));
}

export default function Overview() {
  const { activeProject, tasks, raidItems, updateProject, setCurrentPage, showToast } = useApp();
  const [editOpen, setEditOpen] = useState(false);
  const [form, setForm] = useState({});

  if (!activeProject) return <EmptyState message="No project selected. Create one to get started." />;

  const total = tasks.length;
  const done = tasks.filter(t => t.ownerStatus === 'Done').length;
  const inprog = tasks.filter(t => t.ownerStatus === 'In Progress').length;
  const blocked = tasks.filter(t => ['Blocked', 'Delayed'].includes(t.ownerStatus)).length;
  const progress = total ? Math.round(done / total * 100) : 0;
  const effortBurned = tasks.reduce((s, t) => s + (parseFloat(t.actualEffort) || 0), 0);
  const effortPlanned = tasks.reduce((s, t) => s + (parseFloat(t.expectedEffort) || 0), 0);

  const phaseStats = PHASES.map(ph => {
    const ts = tasks.filter(t => t.phase === ph);
    const dn = ts.filter(t => t.ownerStatus === 'Done').length;
    const ip = ts.filter(t => t.ownerStatus === 'In Progress').length;
    const status = ts.length === 0 ? 'notstart' : dn === ts.length ? 'done' : (ip > 0 || dn > 0) ? 'active' : 'notstart';
    return { phase: ph, label: PHASE_LABELS[ph], total: ts.length, done: dn, status };
  });

  const openEdit = () => {
    setForm({ ...activeProject.metadata });
    setEditOpen(true);
  };

  const saveEdit = async () => {
    await updateProject({ ...activeProject, metadata: form });
    setEditOpen(false);
    showToast('Project details saved');
  };

  const md = activeProject.metadata || {};

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={<>Implementation Hub › {activeProject.name} › Overview</>}
        title={activeProject.name}
        subtitle={`Kickoff: ${md.kickoffDate || '—'} · Target Go-Live: ${md.targetGoLive || '—'} · Lead BA: ${md.leadBA || '—'} · Engineers: ${md.engineers || '—'}`}
        actions={<>
          <button className="btn btn-ghost" onClick={openEdit}>✎ Edit Project Details</button>
          <button className="btn btn-primary" onClick={() => setCurrentPage('tasks')}>→ Go to Tasks</button>
        </>}
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {/* KPIs */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 14, marginBottom: 16 }}>
          <KpiCard label="Overall Progress" value={`${progress}%`} delta={`${done} of ${total} tasks done`} variant={progress >= 80 ? 'ok' : progress >= 50 ? 'default' : 'warn'} />
          <KpiCard label="In Progress" value={inprog} delta="tasks active" />
          <KpiCard label="Blocked / Delayed" value={blocked} delta={`${tasks.filter(t => t.ownerStatus === 'Blocked').length} blocked`} variant={blocked > 0 ? 'fail' : 'default'} />
          <KpiCard label="Effort" value={`${effortBurned}h`} delta={`vs ${effortPlanned}h planned`} />
        </div>

        {/* Phase Timeline */}
        <div style={{ display: 'flex', background: '#fff', borderRadius: 8, overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,.08)', marginBottom: 16 }}>
          {phaseStats.map(p => (
            <div key={p.phase} style={{
              flex: 1, padding: '11px 13px', borderRight: '1px solid #f0f0f0',
              background: p.status === 'done' ? '#f1f8e9' : p.status === 'active' ? '#f0f4ff' : '#fff',
              borderBottom: p.status === 'active' ? '3px solid #404789' : p.status === 'done' ? '3px solid #4caf50' : '3px solid transparent',
            }}>
              <div style={{ fontSize: 10, color: '#999', textTransform: 'uppercase', letterSpacing: '.4px', marginBottom: 3 }}>{p.label}</div>
              <div style={{ fontSize: 12, fontWeight: 500, color: '#404041', marginBottom: 4 }}>{p.done}/{p.total} done</div>
              <ProgressBar pct={p.total ? Math.round(p.done / p.total * 100) : 0} />
            </div>
          ))}
        </div>

        {/* Quick Access + Recent RAID */}
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 14 }}>
          <WidgetCard title="Quick Section Access">
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              {[
                { page: 'kickoff',   icon: '⊕', label: 'Kickoff Questionnaire', meta: `${countAnswered(activeProject)} of ${countTotal(activeProject)} answered` },
                { page: 'estimator', icon: '⏱', label: 'Estimator',             meta: `${calcEstimate(activeProject)}h total estimate` },
                { page: 'tasks',     icon: '☑', label: 'Tasks & Checklist',     meta: `${total} items · ${progress}% complete` },
                { page: 'raidlog',   icon: '⚑', label: 'RAID Log',              meta: `${raidItems.length} items tracked` },
                { page: 'templates', icon: '📋', label: 'Template Library',     meta: '4 global templates' },
              ].map(s => (
                <button key={s.page} className="btn btn-outlined" onClick={() => setCurrentPage(s.page)}
                  style={{ justifyContent: 'flex-start', padding: 14, fontSize: 13, textAlign: 'left', display: 'block' }}>
                  {s.icon} {s.label}
                  <span style={{ fontSize: 11, color: '#888', fontWeight: 400, marginTop: 4, display: 'block' }}>{s.meta}</span>
                </button>
              ))}
            </div>
          </WidgetCard>

          <WidgetCard title="Open RAID Items">
            {raidItems.filter(r => !['Completed', 'Closed'].includes(r.status)).slice(0, 5).map(r => (
              <div key={r.id} style={{ padding: '8px 0', borderBottom: '1px solid #f5f5f5', fontSize: 12 }}>
                <div style={{ fontWeight: 500, color: '#404041', marginBottom: 2 }}>{r.item}</div>
                <div style={{ color: '#888' }}>{r.classification} · {r.status}</div>
              </div>
            ))}
            {raidItems.filter(r => !['Completed', 'Closed'].includes(r.status)).length === 0 && (
              <div style={{ color: '#888', fontSize: 12, textAlign: 'center', padding: '20px 0' }}>No open RAID items 🎉</div>
            )}
            <button className="btn btn-ghost" style={{ marginTop: 10, width: '100%', fontSize: 11 }} onClick={() => setCurrentPage('raidlog')}>
              View All RAID Items →
            </button>
          </WidgetCard>
        </div>
      </div>

      {editOpen && (
        <Modal title="Edit Project Details" onClose={() => setEditOpen(false)} onSave={saveEdit}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
            {[
              { key: 'kickoffDate', label: 'Kickoff Date', type: 'date' },
              { key: 'targetGoLive', label: 'Target Go-Live', type: 'date' },
              { key: 'leadBA', label: 'Lead BA', type: 'text' },
              { key: 'engineers', label: 'Engineers', type: 'text' },
              { key: 'currentPhase', label: 'Current Phase', type: 'select', options: PHASES },
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
