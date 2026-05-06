import React, { useState } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, WidgetCard, EmptyState, Modal, Field, Input, Select, Textarea } from '../components/UI';
import { RAID_CLASSIFICATIONS, RAID_STATUSES, RAID_PENDING } from '../lib/defaults';
import { saveTemplate } from '../lib/db';
import * as XLSX from 'xlsx';

// ── Portfolio Dashboard ───────────────────────────────────────
export function Dashboard() {
  const { projects, switchProject, setCurrentPage, activeProjectId } = useApp();
  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader title="Portfolio Dashboard" subtitle={`${projects.length} active implementation${projects.length !== 1 ? 's' : ''}`} />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {projects.length === 0 && <EmptyState message="No projects yet. Create your first project." />}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 14 }}>
          {projects.map(proj => {
            const md = proj.metadata || {};
            const isActive = proj.id === activeProjectId;
            return (
              <div key={proj.id} style={{ background: '#fff', borderRadius: 8, padding: 18, boxShadow: '0 1px 4px rgba(0,0,0,.08)', borderTop: `3px solid ${isActive ? '#da9b38' : '#404789'}`, cursor: 'pointer' }}
                onClick={() => { switchProject(proj.id); setCurrentPage('overview'); }}>
                <div style={{ fontSize: 15, fontWeight: 500, color: '#404041', marginBottom: 5 }}>{proj.name}</div>
                <div style={{ fontSize: 11, color: '#888', marginBottom: 8 }}>{md.kickoffDate ? `Kickoff: ${md.kickoffDate}` : 'Not started'} · {md.targetGoLive ? `Go-Live: ${md.targetGoLive}` : 'No date set'}</div>
                <div style={{ fontSize: 11, color: '#666', marginBottom: 10 }}>Lead BA: {md.leadBA || '—'} · Phase: {md.currentPhase || '—'}</div>
                {isActive && <div style={{ fontSize: 11, color: '#da9b38', fontWeight: 500 }}>● Currently Active</div>}
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
  const { templates, showToast } = useApp();
  const [activeTab, setActiveTab] = useState(null); // null = overview, else 'questionnaire'|'tasks'|'raidlog'

  const cards = [
    { key: 'questionnaire', icon: '⊕', name: 'Kickoff Questionnaire', color: '#404789',
      meta: `${(templates?.questionnaire?.sections || []).reduce((s, sec) => s + sec.questions.length, 0)} questions · ${(templates?.questionnaire?.sections || []).length} sections`,
      desc: 'Discovery questions covering scope, pricing sources, workflows, and compliance.' },
    { key: 'tasks',         icon: '☑', name: 'Tasks & Checklist',     color: '#2e7d32',
      meta: `${(templates?.tasks?.rows || []).length} tasks across 6 phases`,
      desc: 'BA and Dev tasks from project creation to support handover.' },
    { key: 'raidlog',       icon: '⚑', name: 'RAID Log',              color: '#b71c1c',
      meta: `${(templates?.raidlog?.rows || []).length} seed entries`,
      desc: 'Risks, Actions, Issues, Dependencies, Decisions, Assumptions.' },
    { key: 'estimator',     icon: '⏱', name: 'Estimator',             color: '#e65100',
      meta: `${(templates?.estimator?.steps || []).length} estimation steps`,
      desc: 'Effort model with step multipliers and risk adjustments.' },
  ];

  if (activeTab === 'questionnaire') return <QuestionnaireTemplate onBack={() => setActiveTab(null)} />;
  if (activeTab === 'tasks')         return <TasksTemplate         onBack={() => setActiveTab(null)} />;
  if (activeTab === 'raidlog')       return <RaidTemplate          onBack={() => setActiveTab(null)} />;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader title="Template Library" subtitle="Global templates applied to every new project. Click a template to view and edit." />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4,1fr)', gap: 14, marginBottom: 18 }}>
          {cards.map(c => (
            <div key={c.key}
              onClick={() => c.key !== 'estimator' && setActiveTab(c.key)}
              style={{ background: '#fff', borderRadius: 8, padding: 18, boxShadow: '0 1px 4px rgba(0,0,0,.08)', borderTop: `3px solid ${c.color}`, cursor: c.key !== 'estimator' ? 'pointer' : 'default', transition: 'box-shadow .15s' }}
              onMouseEnter={e => { if (c.key !== 'estimator') e.currentTarget.style.boxShadow = '0 4px 14px rgba(0,0,0,.13)'; }}
              onMouseLeave={e => { e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,.08)'; }}
            >
              <div style={{ width: 42, height: 42, borderRadius: 8, background: c.color + '18', color: c.color, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, marginBottom: 10 }}>{c.icon}</div>
              <div style={{ fontSize: 15, fontWeight: 500, color: '#404041', marginBottom: 4 }}>{c.name}</div>
              <div style={{ fontSize: 11, color: '#888', marginBottom: 8 }}>{c.meta}</div>
              <div style={{ fontSize: 12, color: '#555', lineHeight: 1.5, marginBottom: 12 }}>{c.desc}</div>
              <div style={{ fontSize: 11, color: c.key !== 'estimator' ? c.color : '#bbb', fontWeight: 500 }}>
                {c.key !== 'estimator' ? '→ View & Edit' : 'Managed in Estimator page'}
              </div>
            </div>
          ))}
        </div>
        <WidgetCard title="About Templates">
          <p style={{ fontSize: 13, color: '#555', lineHeight: 1.7 }}>
            Templates are global — they pre-populate every new project created. Existing projects are not affected when templates change.
            Click any template card above to view its contents, add or remove rows, and export or import via Excel.
          </p>
        </WidgetCard>
      </div>
    </div>
  );
}

// ── Questionnaire Template Editor ─────────────────────────────
function QuestionnaireTemplate({ onBack }) {
  const { templates, showToast } = useApp();
  const [sections, setSections] = useState(() => JSON.parse(JSON.stringify(templates?.questionnaire?.sections || [])));
  const [editSec, setEditSec] = useState(null);   // { secIdx, qIdx, text } for editing a question
  const [newQ, setNewQ] = useState('');
  const [newSecName, setNewSecName] = useState('');
  const [saving, setSaving] = useState(false);

  const handleSave = async () => {
    setSaving(true);
    const tpl = templates.questionnaire;
    await saveTemplate('tpl-questionnaire', tpl.name, tpl.version, new Date().toISOString().slice(0,10), { sections });
    setSaving(false);
    showToast('Questionnaire template saved');
  };

  const exportExcel = () => {
    const rows = [['Section', 'Question']];
    sections.forEach(s => s.questions.forEach(q => rows.push([s.name, q])));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 22 }, { wch: 80 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Questionnaire');
    XLSX.writeFile(wb, 'Kickoff_Questionnaire_Template.xlsx');
    showToast('Exported to Excel');
  };

  const importExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1).filter(r => r[0] && r[1]);
      const map = {};
      rows.forEach(([sec, q]) => { if (!map[sec]) map[sec] = []; map[sec].push(String(q)); });
      setSections(Object.entries(map).map(([name, questions]) => ({ name, questions })));
      showToast('Imported from Excel');
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const addSection = () => {
    if (!newSecName.trim()) return;
    setSections(s => [...s, { name: newSecName.trim(), questions: [] }]);
    setNewSecName('');
  };

  const removeSection = (si) => setSections(s => s.filter((_, i) => i !== si));

  const addQuestion = (si) => {
    if (!newQ.trim()) return;
    setSections(s => s.map((sec, i) => i !== si ? sec : { ...sec, questions: [...sec.questions, newQ.trim()] }));
    setNewQ('');
  };

  const removeQuestion = (si, qi) => setSections(s => s.map((sec, i) => i !== si ? sec : { ...sec, questions: sec.questions.filter((_, j) => j !== qi) }));

  const saveQuestion = (si, qi, text) => {
    setSections(s => s.map((sec, i) => i !== si ? sec : { ...sec, questions: sec.questions.map((q, j) => j !== qi ? q : text) }));
    setEditSec(null);
  };

  const totalQ = sections.reduce((s, sec) => s + sec.questions.length, 0);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={<span style={{ cursor: 'pointer', color: '#2278cf' }} onClick={onBack}>← Template Library</span>}
        title="Kickoff Questionnaire Template"
        subtitle={`${sections.length} sections · ${totalQ} questions`}
        actions={<>
          <label className="btn btn-ghost" style={{ cursor: 'pointer' }}>↑ Import Excel<input type="file" accept=".xlsx" onChange={importExcel} style={{ display: 'none' }} /></label>
          <button className="btn btn-ghost" onClick={exportExcel}>↓ Export Excel</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? 'Saving…' : '💾 Save Template'}</button>
        </>}
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {/* Add section */}
        <div style={{ background: '#fff', borderRadius: 8, padding: '12px 16px', marginBottom: 14, boxShadow: '0 1px 4px rgba(0,0,0,.08)', display: 'flex', gap: 8 }}>
          <input value={newSecName} onChange={e => setNewSecName(e.target.value)} onKeyDown={e => e.key === 'Enter' && addSection()}
            placeholder="New section name…"
            style={{ flex: 1, border: '1px solid #c4c4c4', borderRadius: 4, padding: '6px 9px', fontFamily: 'Roboto,sans-serif', fontSize: 13 }} />
          <button className="btn btn-primary" onClick={addSection}>+ Add Section</button>
        </div>

        {sections.map((sec, si) => (
          <div key={si} style={{ background: '#fff', borderRadius: 8, boxShadow: '0 1px 4px rgba(0,0,0,.08)', marginBottom: 14, overflow: 'hidden' }}>
            <div style={{ background: '#f4f4f9', padding: '10px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #f1f2fc' }}>
              <span style={{ fontSize: 14, fontWeight: 500, color: '#343c86' }}>{sec.name} <span style={{ fontSize: 11, color: '#888', fontWeight: 400 }}>({sec.questions.length} questions)</span></span>
              <button className="btn btn-danger btn-xs" onClick={() => removeSection(si)}>✕ Remove Section</button>
            </div>
            <div style={{ padding: '12px 16px' }}>
              {sec.questions.map((q, qi) => (
                <div key={qi} style={{ display: 'flex', alignItems: 'flex-start', gap: 8, padding: '7px 0', borderBottom: '1px solid #f5f5f5' }}>
                  <span style={{ fontSize: 11, color: '#bbb', minWidth: 22, paddingTop: 2 }}>{qi + 1}.</span>
                  {editSec && editSec.si === si && editSec.qi === qi
                    ? <div style={{ flex: 1, display: 'flex', gap: 6 }}>
                        <input value={editSec.text} onChange={e => setEditSec(p => ({ ...p, text: e.target.value }))}
                          style={{ flex: 1, border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 8px', fontFamily: 'Roboto,sans-serif', fontSize: 13 }} autoFocus />
                        <button className="btn btn-primary btn-xs" onClick={() => saveQuestion(si, qi, editSec.text)}>✓</button>
                        <button className="btn btn-ghost btn-xs" onClick={() => setEditSec(null)}>✕</button>
                      </div>
                    : <>
                        <span style={{ flex: 1, fontSize: 13, color: '#404041', lineHeight: 1.5 }}>{q}</span>
                        <button className="btn btn-ghost btn-xs" onClick={() => setEditSec({ si, qi, text: q })}>✎</button>
                        <button className="btn btn-danger btn-xs" onClick={() => removeQuestion(si, qi)}>✕</button>
                      </>
                  }
                </div>
              ))}
              <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
                <input value={newQ} onChange={e => setNewQ(e.target.value)} onKeyDown={e => e.key === 'Enter' && addQuestion(si)}
                  placeholder="Add new question…"
                  style={{ flex: 1, border: '1px solid #c4c4c4', borderRadius: 4, padding: '6px 9px', fontFamily: 'Roboto,sans-serif', fontSize: 12 }} />
                <button className="btn btn-ghost btn-sm" onClick={() => addQuestion(si)}>+ Add</button>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Tasks Template Editor ─────────────────────────────────────
function TasksTemplate({ onBack }) {
  const { templates, showToast } = useApp();
  const [rows, setRows] = useState(() => JSON.parse(JSON.stringify(templates?.tasks?.rows || [])));
  const [filterPhase, setFilterPhase] = useState('');
  const [editRow, setEditRow] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [saving, setSaving] = useState(false);

  const phases = [...new Set(rows.map(r => r.phase))];
  const filtered = rows.filter(r => !filterPhase || r.phase === filterPhase);

  const handleSave = async () => {
    setSaving(true);
    const tpl = templates.tasks;
    await saveTemplate('tpl-tasks', tpl.name, tpl.version, new Date().toISOString().slice(0,10), { rows });
    setSaving(false);
    showToast('Tasks template saved');
  };

  const exportExcel = () => {
    const headers = ['Phase','Item','Task','Task Type','Tags','Responsible','Owner','Reviewer'];
    const data = [headers, ...rows.map(r => [r.phase, r.item, r.task, r.taskType, r.tags, r.responsible, r.owner, r.reviewer])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [18,22,70,16,8,28,28,22].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Checklist');
    XLSX.writeFile(wb, 'Tasks_Checklist_Template.xlsx');
    showToast('Exported to Excel');
  };

  const importExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1).filter(r => r[0] && r[2]);
      setRows(data.map(r => ({ phase: r[0]||'', item: r[1]||'', task: r[2]||'', taskType: r[3]||'BA Checklist Item', tags: r[4]||'', responsible: r[5]||'', owner: r[6]||'', reviewer: r[7]||'' })));
      showToast('Imported from Excel');
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const openEdit = (r, i) => { setEditRow(i); setEditForm({ ...r }); };
  const saveEdit = () => { setRows(prev => prev.map((r, i) => i === editRow ? editForm : r)); setEditRow(null); };
  const deleteRow = (i) => setRows(prev => prev.filter((_, j) => j !== i));
  const addRow = () => setRows(prev => [...prev, { phase: filterPhase || 'Project Creation', item: '', task: 'New task', taskType: 'BA Checklist Item', tags: '', responsible: 'Implementation BA', owner: 'Implementation BA', reviewer: 'Implementation Lead' }]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={<span style={{ cursor: 'pointer', color: '#2278cf' }} onClick={onBack}>← Template Library</span>}
        title="Tasks & Checklist Template"
        subtitle={`${rows.length} tasks across ${phases.length} phases`}
        actions={<>
          <label className="btn btn-ghost" style={{ cursor: 'pointer' }}>↑ Import Excel<input type="file" accept=".xlsx" onChange={importExcel} style={{ display: 'none' }} /></label>
          <button className="btn btn-ghost" onClick={exportExcel}>↓ Export Excel</button>
          <button className="btn btn-ghost" onClick={addRow}>+ Add Task</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? 'Saving…' : '💾 Save Template'}</button>
        </>}
      />
      {/* Filter bar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '8px 20px', display: 'flex', gap: 10, flexShrink: 0 }}>
        <select value={filterPhase} onChange={e => setFilterPhase(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Phases ({rows.length})</option>
          {phases.map(p => <option key={p}>{p} ({rows.filter(r => r.phase === p).length})</option>)}
        </select>
        <span style={{ fontSize: 12, color: '#888', alignSelf: 'center' }}>{filtered.length} tasks shown</span>
      </div>
      {/* Table */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0' }}>
              {['#','Phase','Item','Task','Type','Responsible',''].map(h => (
                <th key={h} style={{ padding: '8px', textAlign: 'left', fontWeight: 500, color: '#444', whiteSpace: 'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((r, i) => {
              const realIdx = rows.indexOf(r);
              return (
                <tr key={i} style={{ borderBottom: '1px solid #f0f0f0', background: i % 2 === 0 ? '#fafafe' : '#fff' }}>
                  <td style={{ padding: '5px 8px', color: '#bbb', fontSize: 11 }}>{realIdx + 1}</td>
                  <td style={{ padding: '5px 8px', color: '#666', whiteSpace: 'nowrap' }}>{r.phase}</td>
                  <td style={{ padding: '5px 8px', color: '#666' }}>{r.item}</td>
                  <td style={{ padding: '5px 8px', color: '#404041', maxWidth: 400 }}>{r.task}</td>
                  <td style={{ padding: '5px 8px' }}>
                    <span style={{ fontSize: 10, padding: '2px 6px', borderRadius: 3, background: r.taskType === 'Dev Task' ? '#fff3e0' : '#e3f2fd', color: r.taskType === 'Dev Task' ? '#e65100' : '#1565c0', fontWeight: 500 }}>{r.taskType}</span>
                  </td>
                  <td style={{ padding: '5px 8px', color: '#666' }}>{r.responsible}</td>
                  <td style={{ padding: '5px 8px' }}>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button className="btn btn-ghost btn-xs" onClick={() => openEdit(r, realIdx)}>✎</button>
                      <button className="btn btn-danger btn-xs" onClick={() => deleteRow(realIdx)}>✕</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {editRow !== null && (
        <Modal title="Edit Task" onClose={() => setEditRow(null)} onSave={saveEdit} wide>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
            <Field label="Phase"><Input value={editForm.phase||''} onChange={e => setEditForm(p=>({...p,phase:e.target.value}))} /></Field>
            <Field label="Item"><Input value={editForm.item||''} onChange={e => setEditForm(p=>({...p,item:e.target.value}))} /></Field>
            <Field label="Task" style={{gridColumn:'1/-1'}}><Textarea value={editForm.task||''} onChange={e => setEditForm(p=>({...p,task:e.target.value}))} /></Field>
            <Field label="Task Type">
              <Select value={editForm.taskType||''} onChange={e => setEditForm(p=>({...p,taskType:e.target.value}))}>
                <option>BA Checklist Item</option><option>Dev Task</option>
              </Select>
            </Field>
            <Field label="Tags"><Input value={editForm.tags||''} onChange={e => setEditForm(p=>({...p,tags:e.target.value}))} /></Field>
            <Field label="Responsible"><Input value={editForm.responsible||''} onChange={e => setEditForm(p=>({...p,responsible:e.target.value}))} /></Field>
            <Field label="Owner"><Input value={editForm.owner||''} onChange={e => setEditForm(p=>({...p,owner:e.target.value}))} /></Field>
            <Field label="Reviewer"><Input value={editForm.reviewer||''} onChange={e => setEditForm(p=>({...p,reviewer:e.target.value}))} /></Field>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── RAID Template Editor ──────────────────────────────────────
function RaidTemplate({ onBack }) {
  const { templates, showToast } = useApp();
  const [rows, setRows] = useState(() => JSON.parse(JSON.stringify(templates?.raidlog?.rows || [])));
  const [editRow, setEditRow] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [saving, setSaving] = useState(false);

  const EMPTY = { classification: 'Risk', item: '', details: '', raisedOn: '', pendingWith: '', updates: '', status: 'Not Initiated', eta: '', comments: '' };

  const handleSave = async () => {
    setSaving(true);
    const tpl = templates.raidlog;
    await saveTemplate('tpl-raidlog', tpl.name, tpl.version, new Date().toISOString().slice(0,10), { rows });
    setSaving(false);
    showToast('RAID Log template saved');
  };

  const exportExcel = () => {
    const headers = ['Classification','Item','Details','Raised On','Pending With','Updates','Status','ETA','Comments'];
    const data = [headers, ...rows.map(r => [r.classification,r.item,r.details,r.raisedOn,r.pendingWith,r.updates,r.status,r.eta,r.comments])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [14,22,50,12,14,40,14,12,30].map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'RAID Log');
    XLSX.writeFile(wb, 'RAID_Log_Template.xlsx');
    showToast('Exported to Excel');
  };

  const importExcel = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1).filter(r => r[0] && r[1]);
      setRows(data.map(r => ({ classification: r[0]||'Risk', item: r[1]||'', details: r[2]||'', raisedOn: r[3]||'', pendingWith: r[4]||'', updates: r[5]||'', status: r[6]||'Not Initiated', eta: r[7]||'', comments: r[8]||'' })));
      showToast('Imported from Excel');
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const CLASS_COLORS = { Risk:'#c62828', Action:'#2e7d32', Issue:'#e65100', Dependency:'#1565c0', Decision:'#6a1b9a', Assumption:'#f57f17' };
  const openEdit = (r, i) => { setEditRow(i); setEditForm({ ...r }); };
  const saveEdit = () => { setRows(prev => prev.map((r, i) => i === editRow ? editForm : r)); setEditRow(null); };
  const deleteRow = (i) => setRows(prev => prev.filter((_, j) => j !== i));
  const addRow = () => { setRows(prev => [...prev, { ...EMPTY }]); };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={<span style={{ cursor: 'pointer', color: '#2278cf' }} onClick={onBack}>← Template Library</span>}
        title="RAID Log Template"
        subtitle={`${rows.length} seed entries · Applied to every new project`}
        actions={<>
          <label className="btn btn-ghost" style={{ cursor: 'pointer' }}>↑ Import Excel<input type="file" accept=".xlsx" onChange={importExcel} style={{ display: 'none' }} /></label>
          <button className="btn btn-ghost" onClick={exportExcel}>↓ Export Excel</button>
          <button className="btn btn-ghost" onClick={addRow}>+ Add Entry</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? 'Saving…' : '💾 Save Template'}</button>
        </>}
      />
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0' }}>
              {['#','Type','Item','Details','Raised On','Pending With','Status','ETA',''].map(h => (
                <th key={h} style={{ padding: '8px', textAlign: 'left', fontWeight: 500, color: '#444', whiteSpace: 'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i} style={{ borderBottom: '1px solid #f0f0f0', background: i % 2 === 0 ? '#fafafe' : '#fff' }}>
                <td style={{ padding: '6px 8px', color: '#bbb', fontSize: 11 }}>{i+1}</td>
                <td style={{ padding: '6px 8px' }}>
                  <span style={{ fontSize: 10, padding: '2px 7px', borderRadius: 3, background: (CLASS_COLORS[r.classification]||'#888')+'18', color: CLASS_COLORS[r.classification]||'#888', fontWeight: 500, textTransform: 'uppercase' }}>{r.classification}</span>
                </td>
                <td style={{ padding: '6px 8px', fontWeight: 500, color: '#404041' }}>{r.item}</td>
                <td style={{ padding: '6px 8px', color: '#555', maxWidth: 280 }}>{r.details}</td>
                <td style={{ padding: '6px 8px', color: '#666', whiteSpace: 'nowrap' }}>{r.raisedOn}</td>
                <td style={{ padding: '6px 8px', color: '#666' }}>{r.pendingWith}</td>
                <td style={{ padding: '6px 8px', color: '#666' }}>{r.status}</td>
                <td style={{ padding: '6px 8px', color: '#666', whiteSpace: 'nowrap' }}>{r.eta||'—'}</td>
                <td style={{ padding: '6px 8px' }}>
                  <div style={{ display: 'flex', gap: 4 }}>
                    <button className="btn btn-ghost btn-xs" onClick={() => openEdit(r, i)}>✎</button>
                    <button className="btn btn-danger btn-xs" onClick={() => deleteRow(i)}>✕</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length === 0 && <div style={{ textAlign: 'center', padding: 40, color: '#999' }}>No seed entries. Click + Add Entry to start.</div>}
      </div>

      {editRow !== null && (
        <Modal title="Edit RAID Entry" onClose={() => setEditRow(null)} onSave={saveEdit} wide>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
            <Field label="Classification">
              <Select value={editForm.classification||''} onChange={e => setEditForm(p=>({...p,classification:e.target.value}))}>
                {RAID_CLASSIFICATIONS.map(c => <option key={c}>{c}</option>)}
              </Select>
            </Field>
            <Field label="Status">
              <Select value={editForm.status||''} onChange={e => setEditForm(p=>({...p,status:e.target.value}))}>
                {RAID_STATUSES.map(s => <option key={s}>{s}</option>)}
              </Select>
            </Field>
            <Field label="Item"><Input value={editForm.item||''} onChange={e => setEditForm(p=>({...p,item:e.target.value}))} /></Field>
            <Field label="Pending With">
              <Select value={editForm.pendingWith||''} onChange={e => setEditForm(p=>({...p,pendingWith:e.target.value}))}>
                <option value="">—</option>
                {RAID_PENDING.map(x => <option key={x}>{x}</option>)}
              </Select>
            </Field>
            <Field label="Raised On"><Input type="date" value={editForm.raisedOn||''} onChange={e => setEditForm(p=>({...p,raisedOn:e.target.value}))} /></Field>
            <Field label="ETA"><Input type="date" value={editForm.eta||''} onChange={e => setEditForm(p=>({...p,eta:e.target.value}))} /></Field>
            <Field label="Details"><Textarea value={editForm.details||''} onChange={e => setEditForm(p=>({...p,details:e.target.value}))} /></Field>
            <Field label="Updates"><Textarea value={editForm.updates||''} onChange={e => setEditForm(p=>({...p,updates:e.target.value}))} /></Field>
            <Field label="Comments"><Textarea value={editForm.comments||''} onChange={e => setEditForm(p=>({...p,comments:e.target.value}))} /></Field>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── Documentation ─────────────────────────────────────────────
export function Docs() {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader title="Documentation Library" subtitle="Cross-implementation runbooks (coming soon)" />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        <WidgetCard><p style={{ color: '#666', fontSize: 13 }}>Documentation library coming in a future iteration.</p></WidgetCard>
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
        <WidgetCard><p style={{ color: '#666', fontSize: 13 }}>Data source library coming in a future iteration.</p></WidgetCard>
      </div>
    </div>
  );
}
