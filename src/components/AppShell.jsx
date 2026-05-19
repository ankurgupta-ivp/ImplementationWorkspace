import React, { useState, useRef } from 'react';
import { useApp } from '../hooks/useApp';
import { Modal, Field, Input } from './UI';

const NAV_ITEMS = [
  { page: 'dashboard',   icon: '▣', label: 'Portfolio Dashboard',   section: 'Portfolio' },
  { page: 'overview',    icon: '⚲', label: 'Overview',              section: 'This Project' },
  { page: 'kickoff',     icon: '⊕', label: 'Kickoff Questionnaire', section: null },
  { page: 'estimator',   icon: '⏱', label: 'Estimator',             section: null },
  { page: 'tasks',       icon: '☑', label: 'Tasks & Checklist',     section: null },
  { page: 'gantt',       icon: '📊', label: 'Gantt Chart',            section: null },
  { page: 'raidlog',     icon: '⚑', label: 'RAID Log',              section: null },
  { page: 'templates',   icon: '📋', label: 'Template Library',     section: 'Library' },
  { page: 'docs',        icon: '📄', label: 'Documentation',        section: null },
  { page: 'datasources', icon: '🗄', label: 'Data Sources',         section: null },
];

const SIDEBAR_FULL     = 230;
const SIDEBAR_COLLAPSED = 48;

export default function AppShell({ children }) {
  const { projects, activeProjectId, currentPage, setCurrentPage, switchProject, addProject } = useApp();
  const [showNewProject, setShowNewProject] = useState(false);
  const [newName,        setNewName]        = useState('');
  // collapsed = pinned collapsed state (toggled by icon click)
  const [collapsed,      setCollapsed]      = useState(false);
  // hoverOpen = temporarily expanded because mouse is over sidebar
  const [hoverOpen,      setHoverOpen]      = useState(false);
  const hoverTimer = useRef(null);

  const isOpen = !collapsed || hoverOpen; // sidebar shows full content when open
  const width  = isOpen ? SIDEBAR_FULL : SIDEBAR_COLLAPSED;

  const handleMouseEnter = () => {
    if (!collapsed) return; // already pinned open — nothing to do
    clearTimeout(hoverTimer.current);
    setHoverOpen(true);
  };
  const handleMouseLeave = () => {
    if (!collapsed) return;
    // Small delay so the panel doesn't snap closed if cursor briefly leaves
    hoverTimer.current = setTimeout(() => setHoverOpen(false), 150);
  };

  const handleCreate = async () => {
    if (!newName.trim()) return;
    await addProject(newName.trim());
    setNewName('');
    setShowNewProject(false);
  };

  return (
    <div style={{ fontFamily: "'Roboto',sans-serif", background: '#f0f4f9', color: '#212121', fontSize: 13, height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>

      {/* ── Topbar ── */}
      <div style={{ height: 48, background: '#fff', borderBottom: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', padding: '0 16px', flexShrink: 0, zIndex: 20 }}>
        <div style={{ width: 30, height: 30, background: '#404789', borderRadius: 5, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontWeight: 700, fontSize: 13, fontStyle: 'italic', marginRight: 10 }}>IVP</div>
        <div style={{ fontSize: 14, fontWeight: 500, color: '#404041' }}>
          PriceMaster <strong style={{ color: '#404789' }}>Implementation Hub</strong>
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{ background: '#404789', color: '#fff', fontSize: 12, padding: '4px 10px', borderRadius: 4 }}>Implementation Team</div>
        </div>
      </div>

      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>

        {/* ── Sidebar ── */}
        <aside
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
          style={{
            width,
            minWidth: width,
            background: '#404789',
            flexShrink: 0,
            overflowY: 'auto',
            overflowX: 'hidden',
            color: '#fff',
            transition: 'width 0.22s ease, min-width 0.22s ease',
            position: 'relative',
            zIndex: 10,
            display: 'flex',
            flexDirection: 'column',
          }}
        >
          {/* Collapse / expand toggle button at top */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: isOpen ? 'flex-end' : 'center', padding: isOpen ? '8px 10px 4px' : '8px 0 4px', flexShrink: 0 }}>
            <button
              onClick={() => { setCollapsed(c => !c); setHoverOpen(false); }}
              title={collapsed ? 'Expand sidebar' : 'Collapse sidebar'}
              style={{
                background: 'transparent', border: 'none', color: '#9799b1', cursor: 'pointer',
                fontSize: 18, lineHeight: 1, padding: '4px 6px', borderRadius: 4,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
              }}
              onMouseEnter={e => { e.currentTarget.style.background = '#505382'; e.currentTarget.style.color = '#fff'; }}
              onMouseLeave={e => { e.currentTarget.style.background = 'transparent'; e.currentTarget.style.color = '#9799b1'; }}
            >
              {/* Left/right chevron arrows */}
              {collapsed && !hoverOpen ? '›' : '‹'}
            </button>
          </div>

          {/* Project selector — only when open */}
          {isOpen && (
            <div style={{ padding: '6px 12px 10px', borderBottom: '1px solid #676a9a', flexShrink: 0 }}>
              <div style={{ fontSize: 10, color: '#9799b1', textTransform: 'uppercase', letterSpacing: '.5px', marginBottom: 5, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>Active Project</span>
                <button
                  onClick={() => setShowNewProject(true)}
                  style={{ background: '#da9b38', color: '#fff', border: 'none', padding: '2px 8px', borderRadius: 3, fontSize: 10, cursor: 'pointer', fontFamily: 'Roboto,sans-serif', fontWeight: 500 }}
                >+ NEW</button>
              </div>
              <select
                value={activeProjectId || ''}
                onChange={e => switchProject(e.target.value)}
                style={{ background: '#505382', border: 'none', color: '#fff', padding: '6px 8px', width: '100%', borderRadius: 3, fontSize: 12, fontFamily: 'Roboto,sans-serif', cursor: 'pointer' }}
              >
                {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
          )}

          {/* When collapsed show a small + button for new project */}
          {!isOpen && (
            <div style={{ display: 'flex', justifyContent: 'center', padding: '6px 0', borderBottom: '1px solid #676a9a', flexShrink: 0 }}>
              <button
                onClick={() => setShowNewProject(true)}
                title="New project"
                style={{ background: '#da9b38', color: '#fff', border: 'none', width: 28, height: 28, borderRadius: 4, fontSize: 14, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700 }}
              >+</button>
            </div>
          )}

          {/* Nav items */}
          <div style={{ flex: 1, paddingBottom: 12 }}>
            {NAV_ITEMS.map((item, i) => {
              const prev       = NAV_ITEMS[i - 1];
              const showSection = isOpen && item.section && (!prev || prev.section !== item.section);
              const showDivider = showSection && i > 0;
              const isActive    = currentPage === item.page;

              return (
                <React.Fragment key={item.page}>
                  {showDivider && <div style={{ borderTop: '1px solid #676a9a', margin: '6px 14px' }} />}
                  {showSection && (
                    <div style={{ fontSize: 10, color: '#9799b1', textTransform: 'uppercase', letterSpacing: '.6px', padding: '10px 16px 4px', whiteSpace: 'nowrap' }}>
                      {item.section}
                    </div>
                  )}
                  <div
                    onClick={() => setCurrentPage(item.page)}
                    title={!isOpen ? item.label : undefined}
                    style={{
                      padding: isOpen ? '9px 16px 9px 18px' : '9px 0',
                      fontSize: 13,
                      color: isActive ? '#aeb4ff' : '#fff',
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: isOpen ? 'flex-start' : 'center',
                      gap: 8,
                      borderLeft: isOpen ? `3px solid ${isActive ? '#da9b38' : 'transparent'}` : '3px solid transparent',
                      background: isActive ? '#505382' : 'transparent',
                      fontWeight: isActive ? 500 : 400,
                      userSelect: 'none',
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                    }}
                    onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = '#505382'; }}
                    onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = 'transparent'; }}
                  >
                    <span style={{ width: 18, textAlign: 'center', fontSize: 15, flexShrink: 0 }}>{item.icon}</span>
                    {isOpen && <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.label}</span>}
                  </div>
                </React.Fragment>
              );
            })}
          </div>
        </aside>

        {/* ── Main content ── */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          {children}
        </main>
      </div>

      {/* ── New Project Modal ── */}
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
