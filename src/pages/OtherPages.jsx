import React from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, KpiCard, ProgressBar, WidgetCard, EmptyState } from '../components/UI';

// ── Portfolio Dashboard ───────────────────────────────────────
export function Dashboard() {
  const { projects, tasks: activeTasks, switchProject, setCurrentPage } = useApp();
  // We only have tasks for the active project; for the dashboard show project-level stats
  const { activeProjectId } = useApp();

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        title="Portfolio Dashboard"
        subtitle={`${projects.length} active implementation${projects.length !== 1 ? 's' : ''}`}
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {projects.length === 0 && <EmptyState message="No projects yet. Create your first project." />}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14 }}>
          {projects.map(proj => {
            const md = proj.metadata || {};
            const isActive = proj.id === activeProjectId;
            return (
              <div key={proj.id} style={{
                background: '#fff', borderRadius: 8, padding: 18,
                boxShadow: '0 1px 4px rgba(0,0,0,.08)',
                borderTop: `3px solid ${isActive ? '#da9b38' : '#404789'}`,
                cursor: 'pointer',
              }} onClick={() => { switchProject(proj.id); setCurrentPage('overview'); }}>
                <div style={{ fontSize: 15, fontWeight: 500, color: '#404041', marginBottom: 5 }}>{proj.name}</div>
                <div style={{ fontSize: 11, color: '#888', marginBottom: 8 }}>
                  {md.kickoffDate ? `Kickoff: ${md.kickoffDate}` : 'Not started'} · {md.targetGoLive ? `Go-Live: ${md.targetGoLive}` : 'No date set'}
                </div>
                <div style={{ fontSize: 11, color: '#666', marginBottom: 10 }}>
                  Lead BA: {md.leadBA || '—'} · Phase: {md.currentPhase || '—'}
                </div>
                {isActive && (
                  <div style={{ fontSize: 11, color: '#da9b38', fontWeight: 500 }}>● Currently Active</div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ── Template Library ──────────────────────────────────────────
export function Templates() {
  const { templates, setCurrentPage } = useApp();
  const cards = [
    { key: 'questionnaire', icon: '⊕', name: 'Kickoff Questionnaire', desc: 'Standard set of discovery questions covering scope, pricing sources, workflows, and compliance requirements.', meta: `${(templates?.questionnaire?.sections || []).reduce((s, sec) => s + sec.questions.length, 0)} questions across ${(templates?.questionnaire?.sections || []).length} sections` },
    { key: 'estimator',     icon: '⏱', name: 'Estimator',             desc: 'Effort model with multipliers per implementation step, risk adjustments, and role-based breakdown.',       meta: `${(templates?.estimator?.steps || []).length} estimation steps` },
    { key: 'tasks',         icon: '☑', name: 'Tasks & Checklist',     desc: 'Master checklist of BA and Dev tasks covering all implementation phases from project creation to handover.', meta: `${(templates?.tasks?.rows || []).length} tasks across 6 phases` },
    { key: 'raidlog',       icon: '⚑', name: 'RAID Log',              desc: 'Seed entries for Risks, Actions, Issues, Dependencies, Decisions, and Assumptions.',                       meta: `${(templates?.raidlog?.rows || []).length} seed entries` },
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        title="Template Library"
        subtitle="Global templates applied to every new project"
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 14, marginBottom: 18 }}>
          {cards.map(c => (
            <div key={c.key} style={{ background: '#fff', borderRadius: 8, padding: 18, boxShadow: '0 1px 4px rgba(0,0,0,.08)', borderTop: '3px solid #404789' }}>
              <div style={{ width: 42, height: 42, borderRadius: 8, background: '#eef0ff', color: '#404789', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, marginBottom: 10 }}>{c.icon}</div>
              <div style={{ fontSize: 15, fontWeight: 500, color: '#404041', marginBottom: 5 }}>{c.name}</div>
              <div style={{ fontSize: 11, color: '#888', marginBottom: 8 }}>{c.meta}</div>
              <div style={{ fontSize: 12, color: '#555', lineHeight: 1.5, marginBottom: 12 }}>{c.desc}</div>
              <div style={{ fontSize: 11, color: '#9799b1', fontStyle: 'italic' }}>
                v{templates?.[c.key]?.version || '1.0'} · Updated {templates?.[c.key]?.updated || '—'}
              </div>
            </div>
          ))}
        </div>
        <WidgetCard title="About Templates">
          <p style={{ fontSize: 13, color: '#555', lineHeight: 1.7 }}>
            Templates are global — they apply to every new project created. Existing projects are not affected when templates change.
            To modify template contents (e.g. add questionnaire questions or seed tasks), contact your Implementation Lead.
          </p>
        </WidgetCard>
      </div>
    </div>
  );
}

// ── Documentation ─────────────────────────────────────────────
export function Docs() {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader title="Documentation Library" subtitle="Cross-implementation runbooks (coming soon)" />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        <WidgetCard>
          <p style={{ color: '#666', fontSize: 13 }}>Documentation library coming in a future iteration. Current focus is on templates, projects, and execution tracking.</p>
        </WidgetCard>
      </div>
    </div>
  );
}

// ── Data Sources ──────────────────────────────────────────────
export function DataSources() {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader title="Data Source Repository" subtitle="Central catalog (coming soon)" />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        <WidgetCard>
          <p style={{ color: '#666', fontSize: 13 }}>Data source library coming in a future iteration.</p>
        </WidgetCard>
      </div>
    </div>
  );
}
