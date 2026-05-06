import React, { useState } from 'react';
import { useApp } from '../hooks/useApp';
import { Modal, Field, Input } from './UI';

const NAV_ITEMS = [
  { page: 'dashboard', icon: '▣', label: 'Portfolio Dashboard', section: 'Portfolio' },
  { page: 'overview',  icon: '⚲', label: 'Overview',            section: 'This Project' },
  { page: 'kickoff',   icon: '⊕', label: 'Kickoff Questionnaire', section: null },
  { page: 'estimator', icon: '⏱', label: 'Estimator',            section: null },
  { page: 'tasks',     icon: '☑', label: 'Tasks & Checklist',    section: null },
  { page: 'raidlog',   icon: '⚑', label: 'RAID Log',             section: null },
  { page: 'templates', icon: '📋', label: 'Template Library',    section: 'Library' },
  { page: 'docs',      icon: '📄', label: 'Documentation',       section: null },
  { page: 'datasources',icon:'🗄', label: 'Data Sources',        section: null },
];

export default function AppShell({ children }) {
  const { projects, activeProjectId, currentPage, setCurrentPage, switchProject, addProject } = useApp();
  const [showNewProject, setShowNewProject] = useState(false);
  const [newName, setNewName] = useState('');

  const handleCreate = async () => {
    if (!newName.trim()) return;
    await addProject(newName.trim());
    setNewName('');
    setShowNewProject(false);
  };

  return (
    <div style={{ fontFamily: "'Roboto',sans-serif", background: '#f0f4f9', color: '#212121', fontSize: 13, height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* Topbar */}
      <div style={{ height: 48, background: '#fff', borderBottom: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', padding: '0 16px', flexShrink: 0, zIndex: 20 }}>
        <div style={{ width: 30, height: 30, background: '#404789', borderRadius: 5, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 13, fontStyle: 'italic', marginRight: 10 }}>IVP</div>
        <div style={{ fontSize: 14, fontWeight: 500, color: '#404041' }}>PriceMaster <strong style={{ color: '#404789' }}>Implementation Hub</strong></div>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{ background: '#404789', color: '#fff', fontSize: 12, padding: '4px 10px', borderRadius: 4 }}>Implementation Team</div>
        </div>
      </div>

      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        {/* Sidebar */}
        <aside style={{ width: 230, background: '#404789', flexShrink: 0, overflowY: 'auto', color: '#fff' }}>
          {/* Project Selector */}
          <div style={{ padding: '10px 12px', borderBottom: '1px solid #676a9a' }}>
            <div style={{ fontSize: 10, color: '#9799b1', textTransform: 'uppercase', letterSpacing: '.5px', marginBottom: 5, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span>Active Project</span>
              <button onClick={() => setShowNewProject(true)} style={{ background: '#da9b38', color: '#fff', border: 'none', padding: '2px 8px', borderRadius: 3, fontSize: 10, cursor: 'pointer', fontFamily: 'Roboto,sans-serif', fontWeight: 500 }}>+ NEW</button>
            </div>
            <select
              value={activeProjectId || ''}
              onChange={e => switchProject(e.target.value)}
              style={{ background: '#505382', border: 'none', color: '#fff', padding: '6px 8px', width: '100%', borderRadius: 3, fontSize: 12, fontFamily: 'Roboto,sans-serif', cursor: 'pointer' }}
            >
              {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>

          {/* Nav Items */}
          {NAV_ITEMS.map((item, i) => {
            const prev = NAV_ITEMS[i - 1];
            const showSection = item.section && (!prev || prev.section !== item.section);
            const showDivider = showSection && i > 0;
            return (
              <React.Fragment key={item.page}>
                {showDivider && <div style={{ borderTop: '1px solid #676a9a', margin: '6px 14px' }} />}
                {showSection && <div style={{ fontSize: 10, color: '#9799b1', textTransform: 'uppercase', letterSpacing: '.6px', padding: '12px 16px 4px' }}>{item.section}</div>}
                <div
                  onClick={() => setCurrentPage(item.page)}
                  style={{
                    padding: '9px 16px 9px 18px', fontSize: 13, color: currentPage === item.page ? '#aeb4ff' : '#fff',
                    cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8,
                    borderLeft: `3px solid ${currentPage === item.page ? '#da9b38' : 'transparent'}`,
                    background: currentPage === item.page ? '#505382' : 'transparent',
                    fontWeight: currentPage === item.page ? 500 : 400,
                    userSelect: 'none',
                  }}
                  onMouseEnter={e => { if (currentPage !== item.page) e.currentTarget.style.background = '#505382'; }}
                  onMouseLeave={e => { if (currentPage !== item.page) e.currentTarget.style.background = 'transparent'; }}
                >
                  <span style={{ width: 18, textAlign: 'center', fontSize: 14 }}>{item.icon}</span>
                  <span>{item.label}</span>
                </div>
              </React.Fragment>
            );
          })}
        </aside>

        {/* Main content */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          {children}
        </main>
      </div>

      {/* New Project Modal */}
      {showNewProject && (
        <Modal title="Create New Project" onClose={() => setShowNewProject(false)} onSave={handleCreate}>
          <Field label="Project Name" required>
            <Input
              value={newName}
              onChange={e => setNewName(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && handleCreate()}
              placeholder="e.g. Apex Capital — Fixed Income"
              autoFocus
            />
          </Field>
          <p style={{ fontSize: 12, color: '#888', marginTop: 10 }}>
            A new project will be created with all standard templates pre-loaded.
          </p>
        </Modal>
      )}
    </div>
  );
}
